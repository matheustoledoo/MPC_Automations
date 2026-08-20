#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Automação 3 da Central — Relatórios PGFN / Regularize (Dívida Ativa da União).

Fluxo automatizado, exatamente como descrito na especificação:

    eCAC -> Representar empresa -> Dívida Ativa da União ->
    PGFN (Todos os serviços do Regularize) -> Ciente -> (Fechar aviso) ->
    Consultar Dívida Ativa -> Relatório consolidado ->
    Natureza "Todos" + Situação "Todos" -> Gerar relatório ->
    scroll até o fim -> Imprimir -> Salvar como PDF na pasta escolhida.

FASE 1 — PRIMEIRA EMPRESA, APRENDIZADO:
    O usuário faz o fluxo clique a clique. Cada clique executa a ação real e,
    ao mesmo tempo, tem a coordenada gravada. O CNPJ é digitado manualmente e
    o ritmo das 14 teclas é registrado. As cinco posições possíveis do botão
    "Imprimir" são marcadas com F8, sem clicar.

FASE 2 — SEGUNDA EMPRESA EM DIANTE, AUTOMÁTICO:
    Todas as posições aprendidas são repetidas fisicamente para as demais
    empresas da mesma planilha usada pela automação de relatórios do eCAC.

Empresas sem cadastro no Regularize (a tela de criação de cadastro, ou a URL
com "/cadastro") não interrompem a execução: elas são registradas, puladas e
listadas em um PDF final.

Esta versão:
- Não usa Selenium.
- Não usa ChromeDriver.
- Não usa depuração remota.
- Não usa PyAutoGUI, PyScreeze ou Pillow.
- Usa pynput para mouse e teclado, exatamente como ecac_download.py.
- Reaproveita ecac_download.py: planilha, clientes, representação,
  aprendizado de cliques, ritmo de digitação, CSV de resultado e logs.
"""

from __future__ import annotations

import ctypes
import logging
import os
import random
import re
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any, Optional


APP_DIR = Path(__file__).resolve().parent

if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))


# A automação usa a MESMA base de clientes da automação do eCAC.
# O caminho é repassado antes do import porque ecac_download.py lê a
# variável de ambiente no momento em que é carregado.
_PGFN_EXCEL_PATH = os.environ.get("PGFN_EXCEL_PATH", "").strip()

if _PGFN_EXCEL_PATH:
    os.environ["ECAC_EXCEL_PATH"] = _PGFN_EXCEL_PATH


import ecac_download as ecac  # noqa: E402

from ecac_download import (  # noqa: E402
    Client,
    PortalError,
    format_cnpj,
    normalize_text,
    only_digits,
    random_sleep,
    sanitize_filename,
)

from pynput import keyboard, mouse  # noqa: E402
from pypdf import PdfReader  # noqa: E402


# =============================================================================
# CONFIGURAÇÕES PRINCIPAIS
# =============================================================================

DEFAULT_PGFN_OUTPUT_DIR = r"C:\Users\z0057vzy\OneDrive - Siemens Healthineers\Desktop\Saidas PGFN"
PGFN_OUTPUT_DIR = Path(
    os.environ.get("PGFN_OUTPUT_DIR", DEFAULT_PGFN_OUTPUT_DIR)
)

# Página do eCAC aberta depois que a empresa já está representada.
ECAC_URL = os.environ.get(
    "PGFN_ECAC_URL",
    "https://cav.receita.fazenda.gov.br/ecac/",
)

PORTAL_HOST_MARKERS = ("servicos.receitafederal.gov.br",)
ECAC_HOST_MARKERS = ("cav.receita.fazenda.gov.br",)
REGULARIZE_HOST_MARKERS = ("regularize.pgfn.gov.br",)

# Empresa sem cadastro no Regularize: a própria URL denuncia a tela de cadastro.
NO_REGISTRATION_URL_MARKERS = (
    "/cadastro",
    "cadastro-contribuinte",
    "primeiro-acesso",
    "criar-cadastro",
)

# Esperas de carregamento de cada etapa.
WAIT_AFTER_ECAC_OPEN_SECONDS = (12.0, 15.0)
WAIT_AFTER_DIVIDA_ATIVA_SECONDS = (6.0, 9.0)
WAIT_AFTER_REGULARIZE_LINK_SECONDS = 20
WAIT_AFTER_CIENTE_SECONDS = (3.0, 5.0)
WAIT_AFTER_MODAL_CLOSE_SECONDS = (1.5, 2.5)

# Espera obrigatória pedida na especificação.
WAIT_AFTER_CONSULTAR_DIVIDA_SECONDS = 25

WAIT_AFTER_RELATORIO_CONSOLIDADO_SECONDS = (8.0, 12.0)
WAIT_AFTER_FILTER_CLICK_SECONDS = (1.2, 2.0)

# Espera obrigatória pedida na especificação, antes do scroll final.
WAIT_AFTER_GERAR_RELATORIO_SECONDS = 20

WAIT_AFTER_SCROLL_END_SECONDS = (1.5, 2.5)

# O botão "Imprimir" pode aparecer em cinco alturas diferentes.
PRINT_POSITION_COUNT = 5
WAIT_AFTER_PRINT_CLICK_SECONDS = 6

# Janela "Salvar como" do Windows.
SAVE_DIALOG_TIMEOUT_SECONDS = 25
SAVED_FILE_TIMEOUT_SECONDS = 120

# Como lidar com o aviso que pode ou não aparecer depois do "Ciente".
# esc_then_click : envia ESC e, se a posição foi aprendida, também clica.
# esc_only       : envia somente ESC.
# click_only     : usa apenas a posição aprendida.
MODAL_CLOSE_STRATEGY = os.environ.get(
    "PGFN_MODAL_CLOSE_STRATEGY",
    "esc_then_click",
).strip().casefold()

# O Chrome memoriza o destino da impressão. Deixe True para que o usuário
# escolha "Salvar como PDF" durante o aprendizado da primeira empresa.
LEARN_PRINT_DESTINATION = (
    os.environ.get("PGFN_LEARN_PRINT_DESTINATION", "1").strip() != "0"
)

MAX_TAB_CLOSE_ATTEMPTS = 6
MAX_RETRIES_AFTER_LEARNING = ecac.MAX_RETRIES_AFTER_LEARNING
DELAY_BETWEEN_COMPANIES_SECONDS = ecac.DELAY_BETWEEN_COMPANIES_SECONDS
LONG_PAUSE_EVERY = ecac.LONG_PAUSE_EVERY
LONG_PAUSE_SECONDS = ecac.LONG_PAUSE_SECONDS
MANUAL_ACTION_TIMEOUT_SECONDS = ecac.MANUAL_ACTION_TIMEOUT_SECONDS

LOCAL_APP_DATA = ecac.LOCAL_APP_DATA

# Arquivos técnicos ficam fora da pasta de PDFs, igual à automação do eCAC.
RUNTIME_DIR = LOCAL_APP_DATA / "ReceitaFederalPGFN" / "runtime"

LOGGER = logging.getLogger("pgfn_regularize_loop")
RUN_STAMP = datetime.now().strftime("%Y%m%d_%H%M%S")

LOG_FILE: Optional[Path] = None

# Posições possíveis do botão "Imprimir", marcadas com F8 na primeira empresa.
PRINT_POSITIONS: list[tuple[int, int]] = []


# =============================================================================
# AÇÕES APRENDIDAS DO FLUXO PGFN
# =============================================================================

PGFN_ACTION_LABELS = {
    "ecac_divida_ativa": "menu 'Dívida Ativa da União' do eCAC",
    "ecac_pgfn_regularize": (
        "opção 'PGFN - Todos os serviços do Regularize'"
    ),
    "regularize_ciente": "botão 'Ciente' do Regularize",
    "regularize_fechar_aviso": "botão 'Fechar' do aviso do Regularize",
    "regularize_consultar_divida": "opção 'Consultar Dívida Ativa'",
    "regularize_relatorio_consolidado": "opção 'Relatório consolidado'",
    "regularize_natureza_todos": "opção 'Todos' do filtro Natureza",
    "regularize_situacao_todos": "opção 'Todos' do filtro Situação",
    "regularize_gerar_relatorio": "botão 'Gerar relatório'",
    "print_destination_dropdown": (
        "campo 'Destino' da janela de impressão do Chrome"
    ),
    "print_destination_pdf": "opção 'Salvar como PDF' do destino",
    "print_save_button": (
        "botão azul 'Salvar' da janela de impressão do Chrome"
    ),
}

# Os rótulos são registrados no módulo do eCAC para reaproveitar exatamente
# o mesmo mecanismo de aprendizado (capture_mouse_action/perform_mouse_action).
# A alteração vale somente para este processo: a automação do eCAC roda em
# outro processo e continua com o conjunto original de ações obrigatórias.
ecac.ACTION_LABELS.update(PGFN_ACTION_LABELS)

REQUIRED_PGFN_ACTION_KEYS = [
    "ecac_divida_ativa",
    "ecac_pgfn_regularize",
    "regularize_consultar_divida",
    "regularize_relatorio_consolidado",
    "regularize_natureza_todos",
    "regularize_situacao_todos",
    "regularize_gerar_relatorio",
    "print_save_button",
]


class SemCadastroRegularize(PortalError):
    """A empresa não possui cadastro utilizável no Regularize."""


# =============================================================================
# LOG E PASTAS
# =============================================================================

class PgfnPrefixFilter(logging.Filter):
    """Deixa todas as linhas do console no padrão pedido: [PGFN] ..."""

    def filter(self, record: logging.LogRecord) -> bool:
        message = str(record.msg)

        if not message.startswith("[PGFN]"):
            record.msg = "[PGFN] " + message

        return True


def initialize_folders_and_logging() -> None:
    global LOG_FILE

    # A pasta de saída recebe SOMENTE os PDFs finais do Regularize.
    PGFN_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    RUNTIME_DIR.mkdir(parents=True, exist_ok=True)

    LOG_FILE = RUNTIME_DIR / f"execucao_pgfn_regularize_{RUN_STAMP}.log"
    ecac.RESULT_CSV_FILE = (
        RUNTIME_DIR / f"resultado_pgfn_regularize_{RUN_STAMP}.csv"
    )

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
    file_handler.setFormatter(formatter)

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)

    prefix_filter = PgfnPrefixFilter()

    LOGGER.setLevel(logging.INFO)
    LOGGER.handlers.clear()
    LOGGER.filters.clear()
    LOGGER.addHandler(file_handler)
    LOGGER.addHandler(console_handler)
    LOGGER.addFilter(prefix_filter)

    # As funções reaproveitadas do eCAC escrevem no logger delas.
    # Aqui elas passam a escrever no mesmo console, com o mesmo prefixo.
    ecac.LOGGER.setLevel(logging.INFO)
    ecac.LOGGER.handlers.clear()
    ecac.LOGGER.filters.clear()
    ecac.LOGGER.addHandler(file_handler)
    ecac.LOGGER.addHandler(console_handler)
    ecac.LOGGER.addFilter(prefix_filter)


# =============================================================================
# JANELAS E ÁREA DE TRANSFERÊNCIA (somente ctypes, sem bibliotecas extras)
# =============================================================================

def _user32() -> Any:
    return ctypes.windll.user32


def foreground_window_title() -> str:
    if os.name != "nt":
        return ""

    try:
        user32 = _user32()
        handle = user32.GetForegroundWindow()

        if not handle:
            return ""

        length = int(user32.GetWindowTextLengthW(handle))
        buffer = ctypes.create_unicode_buffer(length + 2)
        user32.GetWindowTextW(handle, buffer, length + 2)

        return buffer.value or ""
    except Exception:
        return ""


def foreground_window_class() -> str:
    if os.name != "nt":
        return ""

    try:
        user32 = _user32()
        handle = user32.GetForegroundWindow()

        if not handle:
            return ""

        buffer = ctypes.create_unicode_buffer(260)
        user32.GetClassNameW(handle, buffer, 260)

        return buffer.value or ""
    except Exception:
        return ""


def clear_clipboard() -> None:
    if os.name != "nt":
        return

    try:
        user32 = _user32()

        for _ in range(6):
            if user32.OpenClipboard(0):
                try:
                    user32.EmptyClipboard()
                finally:
                    user32.CloseClipboard()
                return

            time.sleep(0.12)
    except Exception:
        pass


def read_clipboard_text() -> str:
    if os.name != "nt":
        return ""

    CF_UNICODETEXT = 13

    try:
        user32 = _user32()
        kernel32 = ctypes.windll.kernel32

        user32.GetClipboardData.restype = ctypes.c_void_p
        kernel32.GlobalLock.restype = ctypes.c_void_p
        kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
        kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]

        opened = False

        for _ in range(6):
            if user32.OpenClipboard(0):
                opened = True
                break

            time.sleep(0.12)

        if not opened:
            return ""

        try:
            handle = user32.GetClipboardData(CF_UNICODETEXT)

            if not handle:
                return ""

            pointer = kernel32.GlobalLock(ctypes.c_void_p(handle))

            if not pointer:
                return ""

            try:
                return ctypes.c_wchar_p(pointer).value or ""
            finally:
                kernel32.GlobalUnlock(ctypes.c_void_p(handle))
        finally:
            user32.CloseClipboard()
    except Exception as exc:
        LOGGER.warning(
            "Não foi possível ler a área de transferência: %s",
            exc,
        )
        return ""


# =============================================================================
# TECLADO FÍSICO
# =============================================================================

def _hotkey(
    controller: keyboard.Controller,
    character: str,
) -> None:
    controller.press(keyboard.Key.ctrl)
    time.sleep(0.05)
    controller.press(character)
    time.sleep(0.07)
    controller.release(character)
    time.sleep(0.03)
    controller.release(keyboard.Key.ctrl)
    time.sleep(0.12)


def keyboard_ctrl_end() -> None:
    """FULL SCROLL DOWN: força a página até o fim, como pede o processo."""
    controller = keyboard.Controller()

    for _ in range(3):
        ctrl_pressed = False
        end_pressed = False

        try:
            controller.press(keyboard.Key.ctrl)
            ctrl_pressed = True
            time.sleep(0.05)

            controller.press(keyboard.Key.end)
            end_pressed = True
            time.sleep(0.08)

            controller.release(keyboard.Key.end)
            end_pressed = False
            time.sleep(0.03)

            controller.release(keyboard.Key.ctrl)
            ctrl_pressed = False
        finally:
            if end_pressed:
                try:
                    controller.release(keyboard.Key.end)
                except Exception:
                    pass

            if ctrl_pressed:
                try:
                    controller.release(keyboard.Key.ctrl)
                except Exception:
                    pass

        time.sleep(0.15)

    # Complementa com PageDown para páginas que rolam por container interno.
    for _ in range(6):
        ecac.press_and_release_key(
            controller,
            keyboard.Key.page_down,
            0.05,
        )
        time.sleep(0.08)

    random_sleep(WAIT_AFTER_SCROLL_END_SECONDS)


def press_escape(times: int = 1) -> None:
    controller = keyboard.Controller()

    for _ in range(max(1, times)):
        ecac.press_and_release_key(
            controller,
            keyboard.Key.esc,
            0.06,
        )
        time.sleep(0.25)


def type_text_slowly(text: str) -> None:
    controller = keyboard.Controller()

    for character in text:
        try:
            ecac.press_and_release_key(controller, character, 0.012)
        except Exception:
            # Caracteres acentuados/especiais vão pelo type() do pynput.
            controller.type(character)
            time.sleep(0.02)

        time.sleep(0.012)


def read_current_url() -> str:
    """Lê a URL da aba ativa com Ctrl+L, Ctrl+C e a área de transferência."""
    controller = keyboard.Controller()

    clear_clipboard()

    _hotkey(controller, "l")
    time.sleep(0.45)
    _hotkey(controller, "c")
    time.sleep(0.55)

    press_escape()

    url = read_clipboard_text().strip()

    LOGGER.info(
        "Endereço da aba ativa: %s",
        url or "(não identificado)",
    )

    return url


def open_url_in_new_tab(url: str) -> None:
    controller = keyboard.Controller()

    _hotkey(controller, "t")
    time.sleep(0.9)

    type_text_slowly(url)

    ecac.press_and_release_key(
        controller,
        keyboard.Key.enter,
        0.05,
    )


def close_current_tab() -> None:
    controller = keyboard.Controller()
    _hotkey(controller, "w")
    time.sleep(1.2)


def return_to_portal_tab() -> None:
    """
    Fecha as abas abertas durante o fluxo até voltar à aba do portal.

    A aba do portal nunca é fechada: o fechamento para assim que a URL lida
    pertence ao domínio do portal. Se a URL não puder ser lida, nada é
    fechado e a navegação é refeita pelo teclado.
    """
    for _ in range(MAX_TAB_CLOSE_ATTEMPTS):
        url = normalize_text(read_current_url())

        if not url:
            break

        if any(marker in url for marker in PORTAL_HOST_MARKERS):
            break

        LOGGER.info("Fechando a aba auxiliar do fluxo PGFN.")
        close_current_tab()

    ecac.navigate_to_portal_with_keyboard()


# =============================================================================
# PDF FINAL: IMPRESSÃO E VALIDAÇÃO
# =============================================================================

def save_dialog_is_open() -> bool:
    window_class = foreground_window_class()
    title = normalize_text(foreground_window_title())

    if window_class == "#32770":
        return True

    return any(
        marker in title
        for marker in ("salvar como", "save as", "salvar arquivo")
    )


def wait_for_save_dialog(
    timeout: int = SAVE_DIALOG_TIMEOUT_SECONDS,
) -> bool:
    deadline = time.time() + timeout

    while time.time() < deadline:
        if save_dialog_is_open():
            return True

        time.sleep(0.5)

    return False


def wait_for_saved_file(
    target_path: Path,
    timeout: int = SAVED_FILE_TIMEOUT_SECONDS,
) -> bool:
    deadline = time.time() + timeout
    last_size = -1
    stable = 0

    while time.time() < deadline:
        if target_path.exists():
            try:
                size = target_path.stat().st_size
            except OSError:
                size = -1

            if size > 0 and size == last_size:
                stable += 1

                if stable >= 2:
                    return True
            else:
                stable = 0

            last_size = size

        time.sleep(1)

    return False


def pdf_matches_client(
    pdf_path: Path,
    client: Client,
) -> tuple[bool, str]:
    """
    Impede que o PDF de uma empresa seja associado a outra.

    O arquivo é gravado por nós, com caminho e nome definidos antes da
    impressão. Ainda assim o conteúdo é conferido: se aparecer um CNPJ
    diferente e o da empresa não aparecer, o arquivo é recusado.
    """
    try:
        reader = PdfReader(str(pdf_path))
        text = "".join(
            (page.extract_text() or "")
            for page in reader.pages[:5]
        )
    except Exception as exc:
        return True, f"Não foi possível reler o PDF gerado ({exc})."

    if client.cnpj in only_digits(text):
        return True, "CNPJ conferido dentro do PDF."

    other_documents = {
        only_digits(found)
        for found in re.findall(
            r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}",
            text,
        )
    } - {client.cnpj}

    if other_documents:
        return False, (
            "O PDF gerado contém outro CNPJ ("
            + ", ".join(sorted(other_documents)[:3])
            + ") e não o da empresa atual."
        )

    if not text.strip():
        return True, "PDF sem texto extraível; conteúdo não conferido."

    return True, "CNPJ não localizado no texto do PDF (aviso)."


def build_target_pdf_path(client: Client) -> Path:
    """Nome final do relatório: nome da empresa + CNPJ, como pede o processo."""
    company_name = client.spreadsheet_name.strip() or "EMPRESA SEM NOME"

    final_name = sanitize_filename(
        (
            f"PGFN Regularize - {company_name} - "
            f"{format_cnpj(client.cnpj)} - "
            f"{datetime.now().strftime('%Y-%m-%d')}"
        ),
        max_length=190,
    )

    return PGFN_OUTPUT_DIR / f"{final_name}.pdf"


def save_pdf_from_print_dialog(target_path: Path) -> bool:
    """
    Conclui a janela de impressão do Chrome salvando em um caminho exato.

    O nome do arquivo é digitado com o caminho completo, então o PDF já nasce
    na pasta escolhida na interface e com o nome da empresa correta.
    """
    if LEARN_PRINT_DESTINATION:
        # Só é perguntado na primeira empresa; depois é repetido sozinho.
        ecac.perform_mouse_action(
            "print_destination_dropdown",
            (
                "Na janela de impressão do Chrome, clique no campo "
                "'Destino'."
            ),
        )
        time.sleep(1.2)

        ecac.perform_mouse_action(
            "print_destination_pdf",
            "Clique na opção 'Salvar como PDF'.",
        )
        time.sleep(1.2)

    ecac.perform_mouse_action(
        "print_save_button",
        "Clique no botão azul 'Salvar' da janela de impressão.",
    )

    if not wait_for_save_dialog():
        LOGGER.warning(
            "A janela 'Salvar como' do Windows não apareceu."
        )
        return False

    LOGGER.info("Janela 'Salvar como' identificada. Informando o caminho.")

    if target_path.exists():
        try:
            target_path.unlink()
        except OSError:
            pass

    controller = keyboard.Controller()

    _hotkey(controller, "a")
    time.sleep(0.3)

    controller.type(str(target_path))
    time.sleep(0.8)

    ecac.press_and_release_key(
        controller,
        keyboard.Key.enter,
        0.06,
    )

    if not wait_for_saved_file(target_path):
        LOGGER.warning(
            "O arquivo não apareceu em: %s",
            target_path,
        )
        return False

    return True


def capture_print_position(position_number: int) -> None:
    """Marca uma posição possível do botão 'Imprimir' com F8, sem clicar."""
    print()
    print("=" * 94)
    print(
        f"MARCAR POSIÇÃO {position_number}/{PRINT_POSITION_COUNT} "
        "DO BOTÃO IMPRIMIR"
    )
    print("=" * 94)
    print(
        "A página já foi rolada até o fim. Posicione o ponteiro sobre uma "
        "das posições possíveis do botão 'Imprimir'."
    )
    print("NÃO clique. Com o ponteiro no lugar, pressione F8.")
    print(
        "Marque cinco posições diferentes, variando principalmente a "
        "altura em que o botão pode aparecer."
    )
    print("=" * 94)

    captured: dict[str, Optional[tuple[int, int]]] = {"value": None}

    def on_press(
        key: keyboard.Key | keyboard.KeyCode,
    ) -> Optional[bool]:
        if key == keyboard.Key.f8:
            x, y = mouse.Controller().position
            captured["value"] = (int(x), int(y))
            return False

        if key == keyboard.Key.esc:
            return False

        return None

    LOGGER.warning(
        "Aguardando F8 para marcar a posição %d/%d por até %d segundos.",
        position_number,
        PRINT_POSITION_COUNT,
        MANUAL_ACTION_TIMEOUT_SECONDS,
    )

    listener = keyboard.Listener(on_press=on_press)
    listener.start()
    listener.join(MANUAL_ACTION_TIMEOUT_SECONDS)

    if listener.is_alive():
        listener.stop()
        raise PortalError(
            "Tempo esgotado durante a marcação das posições de impressão."
        )

    position = captured["value"]

    if position is None:
        raise PortalError(
            "A marcação foi cancelada ou F8 não foi pressionado."
        )

    PRINT_POSITIONS.append(position)

    LOGGER.info(
        "Posição de impressão %d/%d marcada: X=%d, Y=%d.",
        position_number,
        PRINT_POSITION_COUNT,
        position[0],
        position[1],
    )


def capture_all_print_positions() -> None:
    if len(PRINT_POSITIONS) == PRINT_POSITION_COUNT:
        return

    PRINT_POSITIONS.clear()

    print()
    print("=" * 94)
    print("APRENDIZADO DAS CINCO POSIÇÕES DO BOTÃO IMPRIMIR")
    print("=" * 94)
    print(
        "Marque cinco pontos possíveis para o botão 'Imprimir' que fica "
        "no fim do relatório consolidado."
    )
    print("Não clique durante a marcação. Use somente F8.")
    print("=" * 94)

    for position_number in range(1, PRINT_POSITION_COUNT + 1):
        capture_print_position(position_number)

    LOGGER.info(
        "As %d posições possíveis do botão 'Imprimir' foram registradas.",
        PRINT_POSITION_COUNT,
    )


def click_print_and_save(
    client: Client,
    target_path: Path,
    learning_mode: bool,
) -> Path:
    """
    Testa as posições do botão 'Imprimir' até a janela de impressão aparecer
    e o PDF ser efetivamente salvo na pasta escolhida.
    """
    keyboard_ctrl_end()

    if learning_mode:
        capture_all_print_positions()

    if len(PRINT_POSITIONS) != PRINT_POSITION_COUNT:
        raise PortalError(
            "As posições do botão 'Imprimir' ainda não foram aprendidas."
        )

    controller = mouse.Controller()

    for index, (target_x, target_y) in enumerate(
        PRINT_POSITIONS,
        start=1,
    ):
        # Cada tentativa parte sempre do mesmo fim de página.
        keyboard_ctrl_end()

        LOGGER.info(
            "Testando posição %d/%d de 'Imprimir': X=%d, Y=%d.",
            index,
            PRINT_POSITION_COUNT,
            target_x,
            target_y,
        )

        ecac.move_mouse_smoothly(controller, target_x, target_y)

        controller.press(mouse.Button.left)
        time.sleep(random.uniform(*ecac.MOUSE_CLICK_HOLD_SECONDS))
        controller.release(mouse.Button.left)

        time.sleep(WAIT_AFTER_PRINT_CLICK_SECONDS)

        LOGGER.info("Gerando o PDF pela janela de impressão do Chrome.")

        if save_pdf_from_print_dialog(target_path):
            LOGGER.info("Download identificado: %s", target_path.name)
            break

        LOGGER.warning(
            "A janela de impressão não foi concluída na posição %d/%d. "
            "Tentando a próxima.",
            index,
            PRINT_POSITION_COUNT,
        )

        # Fecha qualquer janela/diálogo aberto antes da próxima tentativa.
        press_escape(2)
        time.sleep(1.0)
    else:
        raise PortalError(
            "As posições do botão 'Imprimir' foram testadas, mas o PDF "
            "não foi salvo."
        )

    LOGGER.info("Validando PDF gerado...")

    if not target_path.exists() or target_path.stat().st_size <= 0:
        raise PortalError(
            f"O PDF do relatório ficou vazio: {target_path.name}"
        )

    belongs, detail = pdf_matches_client(target_path, client)

    if not belongs:
        try:
            target_path.unlink()
        except OSError:
            pass

        raise PortalError(detail)

    LOGGER.info("Validação do PDF: %s", detail)
    LOGGER.info("PDF salvo com sucesso: %s", target_path)

    return target_path


# =============================================================================
# FLUXO PGFN DE UMA EMPRESA
# =============================================================================

def ensure_regularize_is_available() -> None:
    """
    Confere, pela URL da aba ativa, se a empresa possui cadastro utilizável.

    Sem DOM e sem Selenium, o endereço é o sinal disponível: a tela de criar
    cadastro do Regularize responde em uma URL com "/cadastro".
    """
    url = read_current_url()
    normalized = normalize_text(url)

    if not normalized:
        LOGGER.warning(
            "Não foi possível ler o endereço do Regularize. "
            "O fluxo continuará normalmente."
        )
        return

    if any(
        marker in normalized
        for marker in NO_REGISTRATION_URL_MARKERS
    ):
        raise SemCadastroRegularize(
            "Empresa sem cadastro no Regularize "
            f"(tela de cadastro detectada em: {url})."
        )


def close_regularize_notice() -> None:
    """
    Fecha o aviso que pode ou não aparecer depois do 'Ciente'.

    A tela nem sempre exibe esse aviso e não há DOM para consultar, então a
    estratégia padrão é enviar ESC (inofensivo quando não há aviso) e, em
    seguida, repetir a posição do 'Fechar' aprendida na primeira empresa.
    """
    if MODAL_CLOSE_STRATEGY in ("esc_then_click", "esc_only"):
        LOGGER.info("Enviando ESC para fechar um possível aviso.")
        press_escape(2)
        random_sleep(WAIT_AFTER_MODAL_CLOSE_SECONDS)

    if MODAL_CLOSE_STRATEGY == "esc_only":
        return

    if "regularize_fechar_aviso" in ecac.MOUSE_ACTIONS:
        LOGGER.info("Repetindo a posição aprendida do botão 'Fechar'.")
        ecac.replay_mouse_action("regularize_fechar_aviso")
        random_sleep(WAIT_AFTER_MODAL_CLOSE_SECONDS)
        return

    # Primeira empresa: o aviso pode simplesmente não existir.
    print()
    print("=" * 94)
    print("AVISO DO REGULARIZE")
    print("=" * 94)
    print(
        "Se apareceu a janela de aviso, clique agora no botão 'Fechar'. "
        "A posição será gravada."
    )
    print(
        "Se NÃO apareceu nenhuma janela, apenas volte ao terminal e "
        "pressione ENTER para seguir sem gravar posição."
    )
    print("=" * 94)

    captured: dict[str, Optional[tuple[int, int]]] = {"value": None}

    def on_click(
        x: int,
        y: int,
        button: mouse.Button,
        pressed: bool,
    ) -> Optional[bool]:
        if pressed and button == mouse.Button.left:
            captured["value"] = (int(x), int(y))
            return False

        return None

    listener = mouse.Listener(on_click=on_click)
    listener.start()

    try:
        input(
            "\nClique em 'Fechar' na tela OU pressione ENTER "
            "se o aviso não apareceu..."
        )
    finally:
        if listener.is_alive():
            listener.stop()

        listener.join(1.0)

    position = captured["value"]

    if position is None:
        LOGGER.info(
            "Nenhum aviso foi fechado. As próximas empresas usarão "
            "somente o ESC."
        )
        return

    ecac.MOUSE_ACTIONS["regularize_fechar_aviso"] = {"position": position}

    LOGGER.info(
        "Posição aprendida para o botão 'Fechar': X=%d, Y=%d.",
        position[0],
        position[1],
    )

    random_sleep(WAIT_AFTER_MODAL_CLOSE_SECONDS)


def wait_with_countdown(total_seconds: int, reason: str) -> None:
    LOGGER.info("Aguardando %d segundos: %s", total_seconds, reason)

    remaining = total_seconds

    while remaining > 0:
        step = min(10, remaining)
        time.sleep(step)
        remaining -= step

        if remaining > 0:
            LOGGER.info("Faltam %d segundo(s): %s", remaining, reason)


def process_pgfn_client(
    client: Client,
    learning_mode: bool,
) -> Path:
    """Executa o fluxo completo do Regularize para uma empresa."""
    LOGGER.info("Representando empresa...")

    # Etapas 1 a 8 reaproveitadas integralmente da automação do eCAC.
    ecac.represent_client(client, learning_mode=learning_mode)

    LOGGER.info("Acessando o eCAC da empresa representada...")
    open_url_in_new_tab(ECAC_URL)
    random_sleep(WAIT_AFTER_ECAC_OPEN_SECONDS)
    ecac.keyboard_ctrl_home()

    LOGGER.info("Acessando Dívida Ativa da União...")
    ecac.perform_mouse_action(
        "ecac_divida_ativa",
        (
            "No eCAC, com a empresa representada, clique em "
            "'Dívida Ativa da União'."
        ),
    )
    random_sleep(WAIT_AFTER_DIVIDA_ATIVA_SECONDS)

    LOGGER.info("Acessando Regularize...")
    ecac.perform_mouse_action(
        "ecac_pgfn_regularize",
        "Clique em 'PGFN - Todos os serviços do Regularize'.",
    )
    wait_with_countdown(
        WAIT_AFTER_REGULARIZE_LINK_SECONDS,
        "carregamento do Regularize",
    )

    # Empresa sem cadastro não trava a automação: é registrada e pulada.
    ensure_regularize_is_available()

    LOGGER.info("Confirmando o aviso inicial (Ciente)...")
    ecac.perform_mouse_action(
        "regularize_ciente",
        "Clique no botão 'Ciente' do Regularize.",
    )
    random_sleep(WAIT_AFTER_CIENTE_SECONDS)

    close_regularize_notice()

    # A tela de cadastro também pode aparecer depois do Ciente.
    ensure_regularize_is_available()

    LOGGER.info("Consultando Dívida Ativa...")
    ecac.perform_mouse_action(
        "regularize_consultar_divida",
        "Clique em 'Consultar Dívida Ativa'.",
    )
    wait_with_countdown(
        WAIT_AFTER_CONSULTAR_DIVIDA_SECONDS,
        "carregamento da consulta de dívida ativa",
    )

    LOGGER.info("Abrindo o relatório consolidado...")
    ecac.perform_mouse_action(
        "regularize_relatorio_consolidado",
        "Clique em 'Relatório consolidado'.",
    )
    random_sleep(WAIT_AFTER_RELATORIO_CONSOLIDADO_SECONDS)

    LOGGER.info("Selecionando natureza e situação...")
    ecac.perform_mouse_action(
        "regularize_natureza_todos",
        "Clique em 'Todos' no filtro de Natureza.",
    )
    random_sleep(WAIT_AFTER_FILTER_CLICK_SECONDS)

    ecac.perform_mouse_action(
        "regularize_situacao_todos",
        "Clique em 'Todos' no filtro de Situação.",
    )
    random_sleep(WAIT_AFTER_FILTER_CLICK_SECONDS)

    LOGGER.info("Gerando relatório...")
    ecac.perform_mouse_action(
        "regularize_gerar_relatorio",
        "Clique no botão 'Gerar relatório'.",
    )
    wait_with_countdown(
        WAIT_AFTER_GERAR_RELATORIO_SECONDS,
        "geração do relatório consolidado",
    )

    target_path = build_target_pdf_path(client)

    return click_print_and_save(
        client,
        target_path,
        learning_mode=learning_mode,
    )


def learning_is_complete() -> bool:
    required = set(ecac.REQUIRED_ACTION_KEYS) | set(
        REQUIRED_PGFN_ACTION_KEYS
    )

    return (
        required.issubset(ecac.MOUSE_ACTIONS)
        and len(ecac.TYPING_PATTERN) == 14
        and len(PRINT_POSITIONS) == PRINT_POSITION_COUNT
    )


# =============================================================================
# PDF DAS EMPRESAS QUE NÃO PUDERAM SER PROCESSADAS
# =============================================================================

def _pdf_escape(text: str) -> str:
    cleaned = (
        str(text)
        .encode("latin-1", "replace")
        .decode("latin-1")
    )

    for source, target in (
        ("\\", r"\\"),
        ("(", r"\("),
        (")", r"\)"),
    ):
        cleaned = cleaned.replace(source, target)

    return cleaned


def build_simple_pdf(
    output_path: Path,
    title: str,
    subtitle: str,
    lines: list[str],
) -> Path:
    """
    Gera um PDF simples sem nenhuma dependência nova.

    É usado apenas para o relatório final das empresas que não puderam ser
    processadas, conforme pedido no processo.
    """
    page_width = 595
    page_height = 842
    left = 45
    top = page_height - 70
    line_height = 15
    bottom = 55

    usable_lines = max(
        1,
        int((top - 40 - bottom) // line_height),
    )

    chunks = [
        lines[index:index + usable_lines]
        for index in range(0, max(len(lines), 1), usable_lines)
    ] or [[]]

    contents: list[str] = []

    for page_number, chunk in enumerate(chunks, start=1):
        parts = [
            "BT /F2 15 Tf "
            f"{left} {top} Td ({_pdf_escape(title)}) Tj ET",
            "BT /F1 10 Tf "
            f"{left} {top - 20} Td ({_pdf_escape(subtitle)}) Tj ET",
        ]

        y = top - 45

        for line in chunk:
            parts.append(
                "BT /F1 10 Tf "
                f"{left} {y} Td ({_pdf_escape(line)}) Tj ET"
            )
            y -= line_height

        parts.append(
            "BT /F1 8 Tf "
            f"{left} {bottom - 20} Td "
            f"({_pdf_escape(f'Página {page_number} de {len(chunks)}')}) "
            "Tj ET"
        )

        contents.append("\n".join(parts))

    objects: list[bytes] = []

    def add_object(payload: bytes) -> int:
        objects.append(payload)
        return len(objects)

    font_regular = add_object(
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica "
        b"/Encoding /WinAnsiEncoding >>"
    )
    font_bold = add_object(
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold "
        b"/Encoding /WinAnsiEncoding >>"
    )

    pages_object_number = len(objects) + 2 * len(contents) + 1
    page_numbers: list[int] = []

    for content in contents:
        stream = content.encode("latin-1", "replace")
        stream_number = add_object(
            b"<< /Length "
            + str(len(stream)).encode("ascii")
            + b" >>\nstream\n"
            + stream
            + b"\nendstream"
        )

        page_numbers.append(
            add_object(
                (
                    "<< /Type /Page /Parent "
                    f"{pages_object_number} 0 R "
                    f"/MediaBox [0 0 {page_width} {page_height}] "
                    "/Resources << /Font << "
                    f"/F1 {font_regular} 0 R /F2 {font_bold} 0 R "
                    ">> >> "
                    f"/Contents {stream_number} 0 R >>"
                ).encode("latin-1")
            )
        )

    pages_number = add_object(
        (
            "<< /Type /Pages /Count "
            f"{len(page_numbers)} /Kids ["
            + " ".join(f"{number} 0 R" for number in page_numbers)
            + "] >>"
        ).encode("latin-1")
    )

    catalog_number = add_object(
        f"<< /Type /Catalog /Pages {pages_number} 0 R >>".encode("latin-1")
    )

    output = bytearray(b"%PDF-1.4\n")
    offsets: list[int] = []

    for index, payload in enumerate(objects, start=1):
        offsets.append(len(output))
        output += str(index).encode("ascii") + b" 0 obj\n"
        output += payload
        output += b"\nendobj\n"

    xref_offset = len(output)
    output += (
        b"xref\n0 " + str(len(objects) + 1).encode("ascii") + b"\n"
    )
    output += b"0000000000 65535 f \n"

    for offset in offsets:
        output += f"{offset:010d} 00000 n \n".encode("ascii")

    output += (
        b"trailer\n<< /Size "
        + str(len(objects) + 1).encode("ascii")
        + b" /Root "
        + str(catalog_number).encode("ascii")
        + b" 0 R >>\nstartxref\n"
        + str(xref_offset).encode("ascii")
        + b"\n%%EOF\n"
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(bytes(output))

    return output_path


def build_pending_companies_report(
    pending: list[dict[str, str]],
) -> Optional[Path]:
    if not pending:
        LOGGER.info(
            "Todas as empresas processadas possuíam cadastro no Regularize."
        )
        return None

    lines: list[str] = [
        f"Total de empresas sem relatório: {len(pending)}",
        "",
    ]

    for index, item in enumerate(pending, start=1):
        lines.append(
            f"{index:03d}. {item['nome']}"
        )
        lines.append(
            f"      CNPJ: {item['cnpj']}   |   Linha da planilha: "
            f"{item['linha']}"
        )
        lines.append(f"      Motivo: {item['motivo']}")
        lines.append("")

    output_path = PGFN_OUTPUT_DIR / (
        f"PGFN Regularize - Empresas sem relatorio - {RUN_STAMP}.pdf"
    )

    build_simple_pdf(
        output_path,
        "Empresas sem relatório no Regularize",
        (
            "Gerado pela Central de Automações em "
            f"{datetime.now().strftime('%d/%m/%Y %H:%M')}"
        ),
        lines,
    )

    LOGGER.info(
        "PDF com as empresas sem relatório criado em: %s",
        output_path,
    )

    return output_path


# =============================================================================
# MAIN
# =============================================================================

def main() -> int:
    initialize_folders_and_logging()

    excel_path = ecac.resolve_excel_path()

    LOGGER.info("Iniciando automação...")
    LOGGER.info("Planilha: %s", excel_path)
    LOGGER.info("Pasta dos relatórios PGFN: %s", PGFN_OUTPUT_DIR)

    clients = ecac.load_clients(excel_path)

    if not clients:
        LOGGER.warning("Nenhum CNPJ ativo foi encontrado.")
        return 0

    LOGGER.info(
        "%d CNPJs serão processados no Regularize. "
        "A META PLANO foi ignorada.",
        len(clients),
    )

    csv_handle, csv_writer = ecac.create_result_csv()

    successes = 0
    errors = 0
    pending: list[dict[str, str]] = []
    fatal_error: Optional[str] = None

    try:
        ecac.open_regular_chrome()

        print()
        print("=" * 94)
        print("LOGIN MANUAL")
        print("=" * 94)
        print(
            "1. Use o Chrome REGULAR aberto pelo programa, com seu perfil normal."
        )
        print(
            "2. Faça o login pelo Gov.br usando o certificado da META PLANO."
        )
        print(
            "3. Deixe o Chrome maximizado no monitor que será usado."
        )
        print(
            "4. Volte ao terminal e pressione ENTER "
            "(ou use o botão CONTINUAR / ENTER da Central)."
        )
        print("=" * 94)

        input("\nPressione ENTER somente após concluir o login...")

        LOGGER.info(
            "Login concluído. Abrindo a página de representação."
        )
        ecac.navigate_to_portal_with_keyboard()
        time.sleep(5)

        print()
        print("=" * 94)
        print("DUAS FASES DO PROCESSAMENTO")
        print("=" * 94)
        print(
            "1ª empresa: você faz o fluxo clique a clique e o programa grava "
            "cada posição, o ritmo da digitação do CNPJ e as cinco posições "
            "possíveis do botão 'Imprimir'."
        )
        print(
            "2ª empresa em diante: o programa repete tudo sozinho, empresa "
            "por empresa."
        )
        print()
        print(
            "Em TODAS as empresas há a espera obrigatória de 30 segundos "
            "antes de 'Representar'."
        )
        print("=" * 94)

        for position, client in enumerate(clients, start=1):
            started_at = datetime.now()

            learning_mode = position == 1

            LOGGER.info(
                "Empresa [%d/%d] | Linha %d | %s | %s",
                position,
                len(clients),
                client.excel_row,
                format_cnpj(client.cnpj),
                client.spreadsheet_name,
            )

            max_attempts = (
                1 if learning_mode else MAX_RETRIES_AFTER_LEARNING
            )

            completed = False
            skipped = False
            last_message = ""

            for attempt in range(1, max_attempts + 1):
                try:
                    final_path = process_pgfn_client(
                        client,
                        learning_mode=learning_mode,
                    )

                    if learning_mode and not learning_is_complete():
                        raise PortalError(
                            "O relatório da primeira empresa foi gerado, "
                            "mas os cliques, a digitação ou as cinco "
                            "posições de impressão não ficaram completos."
                        )

                    ecac.write_result(
                        csv_handle,
                        csv_writer,
                        client,
                        status="SUCESSO",
                        attempt=attempt,
                        message=(
                            "Relatório consolidado do Regularize salvo "
                            "com sucesso."
                        ),
                        started_at=started_at,
                        report_name=client.spreadsheet_name,
                        final_path=str(final_path),
                    )

                    successes += 1
                    completed = True

                    if learning_mode:
                        LOGGER.info(
                            "APRENDIZADO DA PRIMEIRA EMPRESA CONCLUÍDO. "
                            "O loop automático começa na segunda empresa."
                        )

                    break

                except SemCadastroRegularize as exc:
                    last_message = str(exc)
                    skipped = True

                    LOGGER.warning(
                        "Empresa pulada (%s): %s",
                        format_cnpj(client.cnpj),
                        last_message,
                    )

                    pending.append(
                        {
                            "nome": client.spreadsheet_name
                            or "EMPRESA SEM NOME",
                            "cnpj": format_cnpj(client.cnpj),
                            "linha": str(client.excel_row),
                            "motivo": last_message,
                        }
                    )

                    ecac.write_result(
                        csv_handle,
                        csv_writer,
                        client,
                        status="SEM CADASTRO",
                        attempt=attempt,
                        message=last_message,
                        started_at=started_at,
                    )

                    break

                except Exception as exc:
                    last_message = str(exc)

                    LOGGER.error(
                        "Erro na linha %d, CNPJ %s, tentativa %d/%d: %s",
                        client.excel_row,
                        format_cnpj(client.cnpj),
                        attempt,
                        max_attempts,
                        exc,
                    )

                    if learning_mode:
                        raise RuntimeError(
                            "O aprendizado da primeira empresa não foi "
                            "concluído. A execução foi interrompida para "
                            "não iniciar o loop com posições incompletas."
                        ) from exc

                    if attempt < max_attempts:
                        LOGGER.info(
                            "Voltando à página do serviço antes "
                            "da nova tentativa."
                        )
                        return_to_portal_tab()

            if not completed and not skipped:
                errors += 1

                pending.append(
                    {
                        "nome": client.spreadsheet_name
                        or "EMPRESA SEM NOME",
                        "cnpj": format_cnpj(client.cnpj),
                        "linha": str(client.excel_row),
                        "motivo": last_message or "Falha sem mensagem.",
                    }
                )

                ecac.write_result(
                    csv_handle,
                    csv_writer,
                    client,
                    status="ERRO",
                    attempt=max_attempts,
                    message=last_message or "Falha sem mensagem.",
                    started_at=started_at,
                )

            # Sempre volta à aba do portal antes da próxima empresa.
            return_to_portal_tab()

            if position < len(clients):
                delay = random.uniform(*DELAY_BETWEEN_COMPANIES_SECONDS)

                LOGGER.info(
                    "Aguardando %.1f segundos antes da próxima empresa.",
                    delay,
                )
                time.sleep(delay)

            if position % LONG_PAUSE_EVERY == 0:
                LOGGER.info(
                    "Pausa longa de %d segundos após %d empresas.",
                    LONG_PAUSE_SECONDS,
                    position,
                )
                time.sleep(LONG_PAUSE_SECONDS)

    except KeyboardInterrupt:
        fatal_error = "Execução interrompida pelo usuário."
        LOGGER.warning(fatal_error)

    except Exception as exc:
        fatal_error = str(exc)
        LOGGER.critical(
            "Erro fatal: %s\n%s",
            exc,
            traceback.format_exc(),
        )

    finally:
        csv_handle.close()

    try:
        build_pending_companies_report(pending)
    except Exception as exc:
        LOGGER.error(
            "Não foi possível gerar o PDF das empresas sem relatório: %s",
            exc,
        )

    LOGGER.info(
        "Execução finalizada. Sucessos: %d | Erros: %d | "
        "Sem cadastro/puladas: %d.",
        successes,
        errors,
        len(pending) - errors if len(pending) >= errors else 0,
    )

    if fatal_error:
        LOGGER.error("Motivo da interrupção: %s", fatal_error)
        return 1

    return 0 if errors == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
