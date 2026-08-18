#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Funções auxiliares da automação PGFN / Regularize.

Reúne tudo o que não é regra de negócio do fluxo:

- esperas inteligentes e cliques seguros;
- localização de elementos por vários seletores diferentes;
- troca de abas;
- logs e eventos para a interface gráfica;
- impressão do relatório em PDF;
- registro das empresas não processadas;
- geração do relatório final de falhas (PDF + CSV).

As funções de CNPJ, nome de arquivo, planilha e navegador são reaproveitadas
do `ecac_download.py`. Nada é reimplementado aqui.
"""

from __future__ import annotations

import base64
import csv
import importlib.util
import json
import logging
import os
import queue
import re
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.request
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable, Optional, Sequence

import config


APP_DIR = Path(__file__).resolve().parent

if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))


# =============================================================================
# DEPENDÊNCIAS
# =============================================================================

AUTO_INSTALAR_DEPENDENCIAS = (
    os.environ.get("PGFN_AUTO_INSTALL", "1").strip() != "0"
)

PACOTES_NECESSARIOS = {
    "selenium": "selenium>=4.15",
}


def instalar_dependencias_ausentes() -> None:
    """Instala o Selenium quando ele ainda não estiver disponível."""
    ausentes = [
        requisito
        for modulo, requisito in PACOTES_NECESSARIOS.items()
        if importlib.util.find_spec(modulo) is None
    ]

    if not ausentes or not AUTO_INSTALAR_DEPENDENCIAS:
        return

    print("Instalando bibliotecas necessárias: " + ", ".join(ausentes))

    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--upgrade", *ausentes]
        )
    except Exception as exc:  # pragma: no cover - depende do ambiente
        print(f"Falha ao instalar automaticamente: {exc}")


instalar_dependencias_ausentes()

try:
    from selenium import webdriver
    from selenium.common.exceptions import (
        ElementClickInterceptedException,
        ElementNotInteractableException,
        JavascriptException,
        NoSuchElementException,
        StaleElementReferenceException,
        TimeoutException,
        WebDriverException,
    )
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    from selenium.webdriver.common.action_chains import ActionChains
    from selenium.webdriver.common.by import By
    from selenium.webdriver.remote.webelement import WebElement
    from selenium.webdriver.support.ui import Select

    SELENIUM_DISPONIVEL = True
    ERRO_SELENIUM: Optional[Exception] = None
except Exception as exc:  # pragma: no cover - depende do ambiente
    SELENIUM_DISPONIVEL = False
    ERRO_SELENIUM = exc
    webdriver = None  # type: ignore[assignment]
    WebElement = Any  # type: ignore[misc,assignment]


def garantir_selenium() -> None:
    if SELENIUM_DISPONIVEL:
        return

    raise RuntimeError(
        "O Selenium não está instalado. Execute:\n"
        "    pip install selenium>=4.15\n"
        f"Detalhe técnico: {ERRO_SELENIUM}"
    )


# =============================================================================
# REAPROVEITAMENTO DO ecac_download.py
# =============================================================================

try:
    import ecac_download as ecac

    ERRO_IMPORT_ECAC: Optional[Exception] = None
except Exception as exc:  # pragma: no cover - depende do ambiente
    ecac = None  # type: ignore[assignment]
    ERRO_IMPORT_ECAC = exc


def garantir_ecac() -> Any:
    """Devolve o módulo `ecac_download` ou explica por que ele não carregou."""
    if ecac is not None:
        return ecac

    raise RuntimeError(
        "Não foi possível importar o ecac_download.py, que contém as funções "
        "de planilha, CNPJ e navegador reaproveitadas por esta automação.\n"
        f"Detalhe técnico: {ERRO_IMPORT_ECAC}"
    )


def normalizar_texto(valor: object) -> str:
    """Minúsculas, sem acento e sem espaços repetidos (função do e-CAC)."""
    return garantir_ecac().normalize_text(valor)


def so_digitos(valor: object) -> str:
    return garantir_ecac().only_digits(valor)


def formatar_cnpj(cnpj: str) -> str:
    return garantir_ecac().format_cnpj(cnpj)


def sanitizar_nome_arquivo(texto: str) -> str:
    return garantir_ecac().sanitize_filename(texto)


def resolver_planilha(caminho: str = "") -> Path:
    """
    Resolve a planilha de clientes.

    Com caminho informado, valida a existência. Sem caminho, usa exatamente
    a mesma resolução do `ecac_download.py`.
    """
    modulo = garantir_ecac()

    if caminho.strip():
        planilha = Path(caminho.strip()).expanduser()

        if not planilha.exists():
            raise FileNotFoundError(f"Planilha não encontrada: {planilha}")

        return planilha

    return modulo.resolve_excel_path()


def carregar_empresas(planilha: Path) -> list[Any]:
    """Lê a mesma base de CNPJs usada pela automação do e-CAC."""
    return garantir_ecac().load_clients(planilha)


# =============================================================================
# ERROS
# =============================================================================

class EmpresaNaoProcessada(Exception):
    """
    Falha de UMA empresa.

    Nunca interrompe o lote: a empresa é registrada e o fluxo segue para a
    próxima.

    definitivo=True marca o que não adianta tentar de novo (a empresa não
    tem cadastro no Regularize, por exemplo). Repetir o fluxo custaria a
    espera obrigatória de 30 segundos da representação para chegar ao mesmo
    resultado.
    """

    def __init__(
        self,
        motivo: str,
        etapa: str = "",
        definitivo: bool = False,
    ) -> None:
        super().__init__(motivo)
        self.motivo = motivo
        self.etapa = etapa
        self.definitivo = definitivo


class ElementoNaoEncontrado(EmpresaNaoProcessada):
    """Nenhum dos seletores configurados localizou o elemento."""


# =============================================================================
# LOG E EVENTOS PARA A INTERFACE
# =============================================================================

LOGGER = logging.getLogger("pgfn_regularize")

MARCADOR_EVENTO = "@@PGFN@@"


def configurar_log(pasta_logs: Path, carimbo: str) -> Path:
    """
    Console no formato pedido na especificação ([HH:MM:SS] mensagem) e
    arquivo completo com data e nível.
    """
    pasta_logs.mkdir(parents=True, exist_ok=True)
    arquivo = pasta_logs / f"pgfn_regularize_{carimbo}.log"

    LOGGER.setLevel(logging.INFO)
    LOGGER.handlers.clear()
    LOGGER.propagate = False

    console = logging.StreamHandler(sys.stdout)
    console.setFormatter(
        logging.Formatter("[%(asctime)s] %(message)s", datefmt="%H:%M:%S")
    )

    arquivo_handler = logging.FileHandler(arquivo, encoding="utf-8")
    arquivo_handler.setFormatter(
        logging.Formatter(
            "%(asctime)s | %(levelname)s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
    )

    LOGGER.addHandler(console)
    LOGGER.addHandler(arquivo_handler)

    # As mensagens do ecac_download (planilha, CNPJs inválidos) aparecem
    # no mesmo log, sem precisar de configuração extra lá.
    if ecac is not None:
        ecac.LOGGER.setLevel(logging.INFO)
        ecac.LOGGER.handlers.clear()
        ecac.LOGGER.propagate = False
        ecac.LOGGER.addHandler(console)
        ecac.LOGGER.addHandler(arquivo_handler)

    return arquivo


def emitir_evento(evento: str, **dados: Any) -> None:
    """
    Publica um evento estruturado no stdout.

    A interface gráfica lê essas linhas para atualizar empresa atual,
    progresso, status e contadores. No terminal elas são apenas uma linha
    técnica a mais.
    """
    payload = {"evento": evento}
    payload.update(dados)

    try:
        print(
            MARCADOR_EVENTO + json.dumps(payload, ensure_ascii=False),
            flush=True,
        )
    except Exception:
        pass


def registrar_etapa(mensagem: str, *args: Any) -> None:
    """Log de etapa + atualização do status na interface."""
    texto = mensagem % args if args else mensagem
    LOGGER.info(texto)
    emitir_evento("status", texto=texto)


# =============================================================================
# ESPERAS GENÉRICAS
# =============================================================================

def esperar_condicao(
    condicao: Any,
    timeout: float,
    intervalo: float = config.INTERVALO_POLLING,
) -> Any:
    """
    Espera inteligente: devolve o primeiro resultado verdadeiro de
    `condicao()` ou None quando o tempo acabar.
    """
    limite = time.monotonic() + timeout

    while True:
        try:
            resultado = condicao()

            if resultado:
                return resultado
        except Exception:
            resultado = None

        if time.monotonic() >= limite:
            return None

        time.sleep(intervalo)


_ENTRADA_DO_USUARIO: "queue.Queue[str]" = queue.Queue()
_LEITOR_DA_ENTRADA: Optional[threading.Thread] = None


def _iniciar_leitor_da_entrada() -> None:
    """
    Mantém UM único leitor da entrada padrão para o programa inteiro.

    Com uma thread nova a cada pausa, a que tivesse estourado o tempo
    continuaria bloqueada e engoliria o ENTER da pausa seguinte.
    """
    global _LEITOR_DA_ENTRADA

    if _LEITOR_DA_ENTRADA is not None and _LEITOR_DA_ENTRADA.is_alive():
        return

    def ler() -> None:
        try:
            for linha in sys.stdin:
                _ENTRADA_DO_USUARIO.put(linha)
        except Exception:
            pass

    _LEITOR_DA_ENTRADA = threading.Thread(target=ler, daemon=True)
    _LEITOR_DA_ENTRADA.start()


def esperar_enter(
    instrucoes: Sequence[str],
    timeout: float = config.TIMEOUT_ACAO_MANUAL,
) -> bool:
    """
    Pausa o fluxo aguardando o ENTER do usuário.

    Funciona tanto no terminal quanto pela interface gráfica, que envia a
    quebra de linha pelo botão CONTINUAR. Retorna False no estouro do tempo,
    para que o lote nunca fique preso indefinidamente.
    """
    print()
    print("=" * 94)

    for linha in instrucoes:
        print(linha)

    print("=" * 94, flush=True)

    emitir_evento("aguardando_usuario", instrucoes=list(instrucoes))

    _iniciar_leitor_da_entrada()

    # Descarta confirmações antigas: só vale o ENTER dado depois de o
    # usuário ver estas instruções.
    while True:
        try:
            _ENTRADA_DO_USUARIO.get_nowait()
        except queue.Empty:
            break

    try:
        _ENTRADA_DO_USUARIO.get(timeout=timeout)
        confirmado = True
    except queue.Empty:
        confirmado = False

    emitir_evento("usuario_respondeu", ok=confirmado)

    if not confirmado:
        LOGGER.warning(
            "Tempo esgotado (%d s) aguardando a confirmação do usuário.",
            int(timeout),
        )

    return confirmado


def contagem_regressiva(segundos: int, motivo: str) -> None:
    """Espera fixa exigida pelo portal, com aviso a cada 10 segundos."""
    restante = int(segundos)

    if restante <= 0:
        return

    LOGGER.info("Aguardando %d segundo(s): %s.", restante, motivo)

    while restante > 0:
        passo = min(10, restante)
        time.sleep(passo)
        restante -= passo

        if restante > 0:
            LOGGER.info("Faltam %d segundo(s) (%s).", restante, motivo)


# =============================================================================
# LOCALIZADORES
# =============================================================================

Localizador = tuple  # ("tipo", "valor") ou ("texto", "valor", "tags")

TAGS_PADRAO_TEXTO = ("a", "button", "span", "div", "li", "input", "label")


def _por_selenium(tipo: str) -> Any:
    mapa = {
        "id": By.ID,
        "css": By.CSS_SELECTOR,
        "xpath": By.XPATH,
        "name": By.NAME,
        "tag": By.TAG_NAME,
        "class": By.CLASS_NAME,
        "link": By.LINK_TEXT,
    }

    if tipo not in mapa:
        raise ValueError(f"Tipo de localizador desconhecido: {tipo!r}")

    return mapa[tipo]


def _texto_do_elemento(elemento: Any) -> str:
    texto = ""

    try:
        texto = elemento.text or ""
    except Exception:
        texto = ""

    if not texto.strip():
        for atributo in ("value", "aria-label", "title"):
            try:
                texto = elemento.get_attribute(atributo) or ""
            except Exception:
                texto = ""

            if texto.strip():
                break

    return texto


def _elementos_por_texto(
    contexto: Any,
    texto: str,
    tags: Sequence[str],
) -> list[Any]:
    """
    Procura pelo texto visível ignorando maiúsculas/minúsculas e acentos.

    Entre vários candidatos, o elemento com menos texto vence: ele é o mais
    específico (o botão, e não a <div> que contém o botão).
    """
    alvo = normalizar_texto(texto)

    if not alvo:
        return []

    candidatos: list[tuple[int, Any]] = []

    for tag in tags:
        try:
            elementos = contexto.find_elements(By.TAG_NAME, tag)
        except Exception:
            continue

        for elemento in elementos:
            conteudo = normalizar_texto(_texto_do_elemento(elemento))

            if conteudo and alvo in conteudo:
                candidatos.append((len(conteudo), elemento))

    candidatos.sort(key=lambda item: item[0])
    return [elemento for _, elemento in candidatos]


def encontrar_elementos(
    contexto: Any,
    localizador: Localizador,
) -> list[Any]:
    """Aplica um único localizador e devolve todos os elementos achados."""
    tipo = localizador[0]
    valor = localizador[1]

    if tipo == "texto":
        tags = (
            [t.strip() for t in str(localizador[2]).split(",") if t.strip()]
            if len(localizador) > 2
            else list(TAGS_PADRAO_TEXTO)
        )
        return _elementos_por_texto(contexto, valor, tags)

    try:
        return contexto.find_elements(_por_selenium(tipo), valor)
    except Exception:
        return []


def _elemento_satisfaz(elemento: Any, condicao: str) -> bool:
    try:
        if condicao == "presente":
            return True

        if not elemento.is_displayed():
            return False

        if condicao == "visivel":
            return True

        return bool(elemento.is_enabled())
    except StaleElementReferenceException:
        return False
    except Exception:
        return False


def wait_for_element(
    contexto: Any,
    localizadores: Sequence[Localizador],
    timeout: float = config.TIMEOUT_PADRAO,
    condicao: str = "clicavel",
    descricao: str = "",
    etapa: str = "",
    obrigatorio: bool = True,
) -> Optional[Any]:
    """
    Espera o elemento aparecer testando os localizadores na ordem.

    condicao:
        "presente" -> existe no DOM (usada em checkbox estilizado);
        "visivel"  -> existe e está visível;
        "clicavel" -> visível e habilitado (padrão).

    Com obrigatorio=False devolve None em vez de levantar erro.
    """
    garantir_selenium()

    limite = time.monotonic() + timeout

    while True:
        for localizador in localizadores:
            for elemento in encontrar_elementos(contexto, localizador):
                if _elemento_satisfaz(elemento, condicao):
                    return elemento

        if time.monotonic() >= limite:
            break

        time.sleep(config.INTERVALO_POLLING)

    if not obrigatorio:
        return None

    raise ElementoNaoEncontrado(
        f"Elemento não encontrado: {descricao or localizadores}",
        etapa=etapa or descricao,
    )


def check_element_exists(
    contexto: Any,
    localizadores: Sequence[Localizador],
    timeout: float = 0.0,
    condicao: str = "visivel",
) -> Optional[Any]:
    """Versão silenciosa de wait_for_element, para elementos opcionais."""
    return wait_for_element(
        contexto,
        localizadores,
        timeout=timeout,
        condicao=condicao,
        obrigatorio=False,
    )


def rolar_ate(driver: Any, elemento: Any) -> None:
    try:
        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center', "
            "inline: 'center'});",
            elemento,
        )
        time.sleep(0.25)
    except Exception:
        pass


def safe_click(
    driver: Any,
    localizadores: Sequence[Localizador],
    timeout: float = config.TIMEOUT_PADRAO,
    descricao: str = "",
    etapa: str = "",
    obrigatorio: bool = True,
    tentativas: int = 3,
) -> Optional[Any]:
    """
    Clique resistente.

    Espera o elemento ficar clicável, rola até ele e tenta, em ordem:
    clique normal, clique via JavaScript e clique via ActionChains. Um
    elemento que "envelhece" no meio do caminho é procurado de novo.
    """
    garantir_selenium()

    ultimo_erro: Optional[Exception] = None

    for tentativa in range(1, max(1, tentativas) + 1):
        elemento = wait_for_element(
            driver,
            localizadores,
            timeout=timeout if tentativa == 1 else config.TIMEOUT_OPCIONAL,
            condicao="clicavel",
            descricao=descricao,
            etapa=etapa,
            obrigatorio=obrigatorio and tentativa == max(1, tentativas),
        )

        if elemento is None:
            if not obrigatorio:
                return None
            continue

        rolar_ate(driver, elemento)

        try:
            elemento.click()
            LOGGER.info("Clique realizado: %s.", descricao or "elemento")
            return elemento
        except (
            ElementClickInterceptedException,
            ElementNotInteractableException,
            StaleElementReferenceException,
            WebDriverException,
        ) as exc:
            ultimo_erro = exc

        try:
            driver.execute_script("arguments[0].click();", elemento)
            LOGGER.info(
                "Clique via JavaScript: %s.", descricao or "elemento"
            )
            return elemento
        except Exception as exc:
            ultimo_erro = exc

        try:
            ActionChains(driver).move_to_element(elemento).pause(
                0.2
            ).click().perform()
            LOGGER.info(
                "Clique via ActionChains: %s.", descricao or "elemento"
            )
            return elemento
        except Exception as exc:
            ultimo_erro = exc

        time.sleep(0.6)

    if not obrigatorio:
        return None

    raise EmpresaNaoProcessada(
        f"Não foi possível clicar em '{descricao or localizadores}'. "
        f"Detalhe: {ultimo_erro}",
        etapa=etapa or descricao,
    )


def marcar_checkbox(
    driver: Any,
    localizadores: Sequence[Localizador],
    descricao: str,
    etapa: str = "",
    timeout: float = config.TIMEOUT_PADRAO,
) -> None:
    """
    Marca um checkbox somente se ele ainda não estiver marcado.

    O checkbox do Regularize costuma ficar visualmente escondido atrás de um
    <label>, por isso a busca aceita elementos apenas presentes no DOM e o
    clique tem alternativas.
    """
    elemento = wait_for_element(
        driver,
        localizadores,
        timeout=timeout,
        condicao="presente",
        descricao=descricao,
        etapa=etapa,
    )

    try:
        if elemento.is_selected():
            LOGGER.info("%s já estava marcado.", descricao)
            return
    except StaleElementReferenceException:
        elemento = wait_for_element(
            driver,
            localizadores,
            timeout=config.TIMEOUT_OPCIONAL,
            condicao="presente",
            descricao=descricao,
            etapa=etapa,
        )

    rolar_ate(driver, elemento)

    for estrategia in ("clique", "label", "javascript"):
        try:
            if estrategia == "clique":
                elemento.click()
            elif estrategia == "label":
                identificador = elemento.get_attribute("id") or ""

                if not identificador:
                    continue

                rotulos = driver.find_elements(
                    By.CSS_SELECTOR, f"label[for='{identificador}']"
                )

                if not rotulos:
                    continue

                rotulos[0].click()
            else:
                driver.execute_script(
                    "arguments[0].click();"
                    "arguments[0].dispatchEvent("
                    "new Event('change', {bubbles: true}));",
                    elemento,
                )
        except Exception:
            continue

        try:
            if elemento.is_selected():
                LOGGER.info("%s selecionado.", descricao)
                return
        except StaleElementReferenceException:
            elemento = wait_for_element(
                driver,
                localizadores,
                timeout=config.TIMEOUT_OPCIONAL,
                condicao="presente",
                descricao=descricao,
                etapa=etapa,
                obrigatorio=False,
            )

            if elemento is not None and elemento.is_selected():
                LOGGER.info("%s selecionado.", descricao)
                return

    raise EmpresaNaoProcessada(
        f"Não foi possível marcar a opção '{descricao}'.",
        etapa=etapa or descricao,
    )


def fechar_modal_opcional(
    driver: Any,
    localizadores: Sequence[Localizador],
    descricao: str,
    timeout: float = config.TIMEOUT_OPCIONAL,
) -> bool:
    """
    Fecha um modal que pode ou não aparecer.

    A ausência do modal nunca é tratada como erro.
    """
    elemento = check_element_exists(driver, localizadores, timeout=timeout)

    if elemento is None:
        LOGGER.info("Modal '%s' não apareceu. Seguindo.", descricao)
        return False

    clicado = safe_click(
        driver,
        localizadores,
        timeout=config.TIMEOUT_OPCIONAL,
        descricao=f"botão {descricao}",
        obrigatorio=False,
    )

    if clicado is None:
        LOGGER.warning(
            "O modal '%s' apareceu, mas o clique não funcionou. Seguindo.",
            descricao,
        )
        return False

    LOGGER.info("Modal '%s' fechado.", descricao)
    time.sleep(0.8)
    return True


# =============================================================================
# PÁGINAS E ABAS
# =============================================================================

def wait_for_page(
    driver: Any,
    timeout: float = config.TIMEOUT_CARREGAMENTO_PAGINA,
) -> bool:
    """Aguarda o carregamento da página e o desaparecimento dos spinners."""
    garantir_selenium()

    pronto = esperar_condicao(
        lambda: driver.execute_script("return document.readyState")
        == "complete",
        timeout,
    )

    if not pronto:
        LOGGER.warning(
            "A página não sinalizou carregamento completo em %.0f s.",
            timeout,
        )

    esperar_carregamento_sumir(driver, timeout=min(timeout, 20))
    return bool(pronto)


def esperar_carregamento_sumir(driver: Any, timeout: float = 20) -> None:
    """Espera os indicadores de carregamento (spinners) sumirem."""

    def sem_spinner() -> bool:
        try:
            elementos = driver.find_elements(
                By.CSS_SELECTOR, config.SELETORES_CARREGANDO
            )
        except Exception:
            return True

        for elemento in elementos:
            try:
                if elemento.is_displayed():
                    return False
            except Exception:
                continue

        return True

    esperar_condicao(sem_spinner, timeout)


def esperar_mudanca_de_url(
    driver: Any,
    url_anterior: str,
    timeout: float = config.TIMEOUT_CARREGAMENTO_PAGINA,
) -> bool:
    resultado = esperar_condicao(
        lambda: driver.current_url != url_anterior,
        timeout,
    )
    return bool(resultado)


def switch_to_new_tab(
    driver: Any,
    abas_antes: Sequence[str],
    timeout: float = config.TIMEOUT_NOVA_ABA,
) -> Optional[str]:
    """
    Troca para a aba aberta por um link com target="_blank".

    Devolve None quando nenhuma aba nova aparecer (alguns navegadores
    reaproveitam a aba atual).
    """
    garantir_selenium()

    conhecidas = set(abas_antes)

    nova = esperar_condicao(
        lambda: next(
            (aba for aba in driver.window_handles if aba not in conhecidas),
            None,
        ),
        timeout,
    )

    if not nova:
        return None

    driver.switch_to.window(nova)
    wait_for_page(driver)
    LOGGER.info("Nova aba aberta e selecionada.")
    return nova


def fechar_abas_extras(driver: Any, aba_principal: str) -> None:
    """Fecha tudo o que foi aberto durante a empresa e volta à aba original."""
    try:
        for aba in list(driver.window_handles):
            if aba == aba_principal:
                continue

            try:
                driver.switch_to.window(aba)
                driver.close()
            except Exception:
                continue

        if aba_principal in driver.window_handles:
            driver.switch_to.window(aba_principal)
        elif driver.window_handles:
            driver.switch_to.window(driver.window_handles[0])
    except Exception as exc:
        LOGGER.warning("Não foi possível organizar as abas: %s", exc)


def texto_da_pagina(driver: Any) -> str:
    """Texto visível da página, normalizado (minúsculo e sem acento)."""
    try:
        corpo = driver.find_element(By.TAG_NAME, "body").text or ""
    except Exception:
        try:
            corpo = driver.execute_script("return document.body.innerText;")
        except Exception:
            corpo = ""

    return normalizar_texto(corpo)


def pagina_contem(driver: Any, textos: Iterable[str]) -> Optional[str]:
    """Devolve o primeiro texto da lista presente na página, ou None."""
    conteudo = texto_da_pagina(driver)

    if not conteudo:
        return None

    for texto in textos:
        if normalizar_texto(texto) in conteudo:
            return texto

    return None


def neutralizar_dialogo_de_impressao(driver: Any) -> None:
    """
    Impede que o botão "Imprimir" abra a janela de impressão do sistema.

    O clique continua acontecendo (a página executa a lógica dela), mas o
    diálogo do Windows não trava a automação: o PDF é gerado logo em
    seguida pelo próprio navegador.
    """
    try:
        driver.execute_script(
            "if (!window.__pgfnPrintPatched) {"
            "  window.__pgfnPrintCalled = false;"
            "  window.__pgfnPrintOriginal = window.print;"
            "  window.print = function () {"
            "    window.__pgfnPrintCalled = true;"
            "  };"
            "  window.__pgfnPrintPatched = true;"
            "}"
        )
    except Exception as exc:
        LOGGER.warning(
            "Não foi possível neutralizar a janela de impressão: %s", exc
        )


def impressao_foi_disparada(driver: Any) -> bool:
    try:
        return bool(
            driver.execute_script("return !!window.__pgfnPrintCalled;")
        )
    except Exception:
        return False


# =============================================================================
# PDF DO RELATÓRIO
# =============================================================================

def pdf_valido(caminho: Path) -> bool:
    """Confere se o arquivo existe, tem tamanho plausível e é mesmo um PDF."""
    try:
        if not caminho.exists():
            return False

        if caminho.stat().st_size < config.TAMANHO_MINIMO_PDF_BYTES:
            return False

        with caminho.open("rb") as arquivo:
            return arquivo.read(5).startswith(b"%PDF")
    except OSError:
        return False


def caminho_unico(destino: Path) -> Path:
    """Evita sobrescrever um PDF já existente da mesma empresa."""
    if not destino.exists():
        return destino

    contador = 2

    while True:
        candidato = destino.with_name(
            f"{destino.stem} ({contador}){destino.suffix}"
        )

        if not candidato.exists():
            return candidato

        contador += 1


def montar_nome_do_pdf(pasta: Path, cnpj: str, nome: str) -> Path:
    """Monta "CNPJ - NOME DA EMPRESA.pdf" com nome válido no Windows."""
    digitos = so_digitos(cnpj)
    cnpj_texto = (
        digitos
        if config.NOME_ARQUIVO_COM_CNPJ_SEM_PONTUACAO
        else formatar_cnpj(digitos).replace("/", "-")
    )

    nome_limpo = sanitizar_nome_arquivo(nome or "EMPRESA SEM NOME")

    arquivo = config.MODELO_NOME_ARQUIVO.format(
        cnpj=cnpj_texto,
        nome=nome_limpo,
    )

    sufixo = Path(arquivo).suffix or ".pdf"
    base = sanitizar_nome_arquivo(Path(arquivo).stem)

    return caminho_unico(pasta / f"{base}{sufixo}")


def save_pdf(
    driver: Any,
    destino: Path,
    etapa: str = "Salvar PDF",
    tentativas: int = 2,
) -> Path:
    """
    Salva a página atual em PDF usando a impressão do próprio navegador.

    O arquivo nasce direto no caminho final, com o nome definitivo, sem
    depender da janela de impressão do Windows nem da pasta de downloads.
    """
    garantir_selenium()

    destino.parent.mkdir(parents=True, exist_ok=True)
    ultimo_erro: Any = None

    for tentativa in range(1, max(1, tentativas) + 1):
        try:
            resultado = driver.execute_cdp_cmd(
                "Page.printToPDF", dict(config.OPCOES_IMPRESSAO)
            )
            conteudo = base64.b64decode(resultado["data"])
            destino.write_bytes(conteudo)

            if pdf_valido(destino):
                LOGGER.info(
                    "PDF salvo (%d KB): %s",
                    destino.stat().st_size // 1024,
                    destino,
                )
                return destino

            ultimo_erro = "o arquivo gerado não é um PDF válido"
        except Exception as exc:
            ultimo_erro = exc

        LOGGER.warning(
            "Tentativa %d de salvar o PDF falhou: %s", tentativa, ultimo_erro
        )
        time.sleep(1.5)

    raise EmpresaNaoProcessada(
        f"Não foi possível salvar o PDF do relatório. Detalhe: {ultimo_erro}",
        etapa=etapa,
    )


# =============================================================================
# REGISTRO DAS EMPRESAS NÃO PROCESSADAS
# =============================================================================

@dataclass
class FalhaEmpresa:
    cnpj: str
    nome: str
    etapa: str
    motivo: str
    momento: str = ""


@dataclass
class RegistroFalhas:
    """Lista das empresas que não puderam ser processadas."""

    itens: list[FalhaEmpresa] = field(default_factory=list)

    def registrar(
        self,
        cnpj: str,
        nome: str,
        etapa: str,
        motivo: str,
    ) -> FalhaEmpresa:
        falha = FalhaEmpresa(
            cnpj=formatar_cnpj(cnpj),
            nome=nome or "",
            etapa=etapa or "Não informada",
            motivo=re.sub(r"\s+", " ", motivo or "").strip()
            or "Motivo não informado",
            momento=datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        )

        self.itens.append(falha)

        LOGGER.error(
            "ERRO\nEmpresa: %s\nCNPJ: %s\nEtapa: %s\nMotivo: %s\n"
            "Status: Não processada",
            falha.nome,
            falha.cnpj,
            falha.etapa,
            falha.motivo,
        )

        emitir_evento(
            "falha",
            cnpj=falha.cnpj,
            nome=falha.nome,
            etapa=falha.etapa,
            motivo=falha.motivo,
        )

        return falha

    def __len__(self) -> int:
        return len(self.itens)

    def gerar_csv(self, destino: Path) -> Optional[Path]:
        if not self.itens:
            return None

        destino.parent.mkdir(parents=True, exist_ok=True)

        with destino.open("w", newline="", encoding="utf-8-sig") as arquivo:
            escritor = csv.writer(arquivo, delimiter=";")
            escritor.writerow(
                ["CNPJ", "Empresa", "Motivo", "Etapa", "Data e hora"]
            )

            for falha in self.itens:
                escritor.writerow(
                    [
                        falha.cnpj,
                        falha.nome,
                        falha.motivo,
                        falha.etapa,
                        falha.momento,
                    ]
                )

        return destino

    def gerar_pdf(self, destino: Path, resumo: Sequence[str] = ()) -> Optional[Path]:
        if not self.itens:
            return None

        # Largura pensada para a folha A4 com a fonte Courier 8 pt.
        larguras = (20, 30, 30, 16)
        cabecalho = _linha_tabela(
            ("CNPJ", "EMPRESA", "MOTIVO", "ETAPA"), larguras
        )

        linhas: list[str] = []
        linhas.extend(resumo)

        if resumo:
            linhas.append("")

        linhas.append(cabecalho)
        linhas.append("-" * len(cabecalho))

        for falha in self.itens:
            partes = _quebrar_linha_tabela(
                (falha.cnpj, falha.nome, falha.motivo, falha.etapa),
                larguras,
            )
            linhas.extend(partes)

        linhas.append("")
        linhas.append(f"Total de empresas não processadas: {len(self.itens)}")

        return escrever_pdf_texto(
            destino,
            "EMPRESAS NÃO PROCESSADAS",
            linhas,
            subtitulo=(
                "Automação PGFN / Regularize — "
                + datetime.now().strftime("%d/%m/%Y %H:%M")
            ),
        )


def register_error(
    registro: RegistroFalhas,
    cnpj: str,
    nome: str,
    etapa: str,
    motivo: str,
) -> FalhaEmpresa:
    """Atalho no formato pedido na especificação."""
    return registro.registrar(cnpj, nome, etapa, motivo)


def _ajustar(texto: str, largura: int) -> str:
    texto = re.sub(r"\s+", " ", str(texto or "")).strip()
    return texto[:largura].ljust(largura)


def _linha_tabela(colunas: Sequence[str], larguras: Sequence[int]) -> str:
    return "  ".join(
        _ajustar(coluna, largura)
        for coluna, largura in zip(colunas, larguras)
    ).rstrip()


def _quebrar_linha_tabela(
    colunas: Sequence[str],
    larguras: Sequence[int],
) -> list[str]:
    """Distribui textos longos em várias linhas mantendo as colunas."""
    pedacos: list[list[str]] = []

    for coluna, largura in zip(colunas, larguras):
        texto = re.sub(r"\s+", " ", str(coluna or "")).strip()
        partes: list[str] = []
        atual = ""

        for palavra in texto.split(" "):
            while len(palavra) > largura:
                if atual:
                    partes.append(atual)
                    atual = ""
                partes.append(palavra[:largura])
                palavra = palavra[largura:]

            if not atual:
                atual = palavra
            elif len(atual) + 1 + len(palavra) <= largura:
                atual = f"{atual} {palavra}"
            else:
                partes.append(atual)
                atual = palavra

        if atual:
            partes.append(atual)

        pedacos.append(partes or [""])

    altura = max(len(parte) for parte in pedacos)
    linhas: list[str] = []

    for indice in range(altura):
        colunas_linha = [
            pedaco[indice] if indice < len(pedaco) else ""
            for pedaco in pedacos
        ]
        linhas.append(_linha_tabela(colunas_linha, larguras))

    return linhas


# =============================================================================
# GERADOR DE PDF SIMPLES (SEM DEPENDÊNCIAS EXTERNAS)
# =============================================================================

LARGURA_A4 = 595.28
ALTURA_A4 = 841.89
MARGEM_PDF = 40.0


SUBSTITUICOES_PDF = {
    "\u2014": "-",
    "\u2013": "-",
    "\u2018": "'",
    "\u2019": "'",
    "\u201c": '"',
    "\u201d": '"',
    "\u2026": "...",
    "\u00a0": " ",
}


def _escapar_pdf(texto: str) -> bytes:
    """Converte o texto para WinAnsi e escapa os caracteres especiais."""
    for origem, destino in SUBSTITUICOES_PDF.items():
        texto = texto.replace(origem, destino)

    dados = texto.encode("latin-1", errors="replace")
    dados = dados.replace(b"\\", b"\\\\")
    dados = dados.replace(b"(", b"\\(")
    dados = dados.replace(b")", b"\\)")
    return dados


def escrever_pdf_texto(
    destino: Path,
    titulo: str,
    linhas: Sequence[str],
    subtitulo: str = "",
    tamanho_fonte: float = 8.0,
) -> Path:
    """
    Gera um PDF de texto puro, sem bibliotecas externas.

    É usado no relatório final de empresas não processadas. A fonte Courier
    mantém as colunas alinhadas e a paginação é automática.
    """
    destino.parent.mkdir(parents=True, exist_ok=True)

    entrelinha = tamanho_fonte * 1.35
    topo = ALTURA_A4 - MARGEM_PDF
    altura_cabecalho = 46.0
    area_util = topo - altura_cabecalho - MARGEM_PDF
    por_pagina = max(1, int(area_util // entrelinha))

    blocos: list[list[str]] = [
        list(linhas[inicio : inicio + por_pagina])
        for inicio in range(0, max(len(linhas), 1), por_pagina)
    ] or [[]]

    total_paginas = len(blocos)
    objetos: list[bytes] = []

    def adicionar(conteudo: bytes) -> int:
        objetos.append(conteudo)
        return len(objetos)

    # 1 catálogo, 2 páginas, 3 e 4 fontes. As páginas começam no objeto 5.
    adicionar(b"<< /Type /Catalog /Pages 2 0 R >>")
    adicionar(b"")  # substituído depois, quando as páginas existirem
    adicionar(
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold "
        b"/Encoding /WinAnsiEncoding >>"
    )
    adicionar(
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Courier "
        b"/Encoding /WinAnsiEncoding >>"
    )

    referencias_paginas: list[int] = []

    for indice, bloco in enumerate(blocos, start=1):
        partes: list[bytes] = []

        partes.append(
            b"BT /F1 15 Tf "
            + f"{MARGEM_PDF:.2f} {topo - 12:.2f}".encode("ascii")
            + b" Td ("
            + _escapar_pdf(titulo)
            + b") Tj ET\n"
        )

        rodape_cabecalho = subtitulo or ""
        if rodape_cabecalho:
            partes.append(
                b"BT /F2 9 Tf "
                + f"{MARGEM_PDF:.2f} {topo - 28:.2f}".encode("ascii")
                + b" Td ("
                + _escapar_pdf(rodape_cabecalho)
                + b") Tj ET\n"
            )

        partes.append(
            b"BT /F2 "
            + f"{tamanho_fonte:.2f}".encode("ascii")
            + b" Tf "
            + f"{entrelinha:.2f}".encode("ascii")
            + b" TL "
            + f"{MARGEM_PDF:.2f} {topo - altura_cabecalho:.2f}".encode("ascii")
            + b" Td\n"
        )

        for linha in bloco:
            partes.append(b"(" + _escapar_pdf(linha) + b") Tj T*\n")

        partes.append(b"ET\n")

        partes.append(
            b"BT /F2 8 Tf "
            + f"{LARGURA_A4 - MARGEM_PDF - 90:.2f} {MARGEM_PDF - 12:.2f}".encode(
                "ascii"
            )
            + b" Td ("
            + _escapar_pdf(f"Página {indice} de {total_paginas}")
            + b") Tj ET\n"
        )

        conteudo = b"".join(partes)
        numero_conteudo = adicionar(
            b"<< /Length "
            + str(len(conteudo)).encode("ascii")
            + b" >>\nstream\n"
            + conteudo
            + b"endstream"
        )

        numero_pagina = adicionar(
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 "
            + f"{LARGURA_A4:.2f} {ALTURA_A4:.2f}".encode("ascii")
            + b"] /Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> "
            b"/Contents "
            + str(numero_conteudo).encode("ascii")
            + b" 0 R >>"
        )

        referencias_paginas.append(numero_pagina)

    objetos[1] = (
        b"<< /Type /Pages /Count "
        + str(total_paginas).encode("ascii")
        + b" /Kids ["
        + b" ".join(
            str(numero).encode("ascii") + b" 0 R"
            for numero in referencias_paginas
        )
        + b"] >>"
    )

    saida = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    posicoes: list[int] = []

    for numero, corpo in enumerate(objetos, start=1):
        posicoes.append(len(saida))
        saida += str(numero).encode("ascii") + b" 0 obj\n" + corpo + b"\nendobj\n"

    inicio_xref = len(saida)
    saida += b"xref\n0 " + str(len(objetos) + 1).encode("ascii") + b"\n"
    saida += b"0000000000 65535 f \n"

    for posicao in posicoes:
        saida += f"{posicao:010d} 00000 n \n".encode("ascii")

    saida += (
        b"trailer\n<< /Size "
        + str(len(objetos) + 1).encode("ascii")
        + b" /Root 1 0 R >>\nstartxref\n"
        + str(inicio_xref).encode("ascii")
        + b"\n%%EOF\n"
    )

    destino.write_bytes(bytes(saida))
    return destino


# =============================================================================
# NAVEGADOR
# =============================================================================

def caminho_do_chrome() -> Path:
    """Usa a mesma detecção do ecac_download.py, com override opcional."""
    if config.CAMINHO_CHROME:
        caminho = Path(config.CAMINHO_CHROME).expanduser()

        if not caminho.exists():
            raise RuntimeError(
                f"Chrome não encontrado no caminho informado: {caminho}"
            )

        return caminho

    return garantir_ecac().find_chrome_executable()


def porta_de_depuracao_ativa(porta: int) -> bool:
    try:
        with urllib.request.urlopen(
            f"http://127.0.0.1:{porta}/json/version", timeout=2
        ) as resposta:
            return resposta.status == 200
    except (urllib.error.URLError, OSError, ValueError):
        return False


def abrir_chrome_com_depuracao(
    porta: int,
    perfil: Path,
    url: str,
) -> Optional[subprocess.Popen]:
    """
    Abre o Chrome do usuário com depuração remota e perfil próprio.

    O perfil dedicado é obrigatório: versões recentes do Chrome recusam a
    depuração remota no perfil padrão. Como a pasta é reaproveitada entre
    execuções, o login do Gov.br costuma continuar válido.
    """
    if porta_de_depuracao_ativa(porta):
        LOGGER.info(
            "Já existe um Chrome com depuração na porta %d. Reaproveitando.",
            porta,
        )
        return None

    perfil.mkdir(parents=True, exist_ok=True)
    executavel = caminho_do_chrome()

    comando = [
        str(executavel),
        f"--remote-debugging-port={porta}",
        f"--user-data-dir={perfil}",
        *config.ARGUMENTOS_CHROME,
        url,
    ]

    LOGGER.info("Abrindo o Chrome com depuração remota na porta %d.", porta)

    processo = subprocess.Popen(comando, close_fds=True)

    ativo = esperar_condicao(
        lambda: porta_de_depuracao_ativa(porta),
        config.TIMEOUT_ABERTURA_CHROME,
        intervalo=0.5,
    )

    if not ativo:
        raise RuntimeError(
            "O Chrome não respondeu na porta de depuração "
            f"{porta}. Feche todas as janelas do Chrome e tente novamente."
        )

    return processo


def _opcoes_chrome(pasta_download: Path) -> Any:
    opcoes = ChromeOptions()

    if config.CAMINHO_CHROME:
        opcoes.binary_location = str(caminho_do_chrome())

    pasta_download.mkdir(parents=True, exist_ok=True)

    opcoes.add_experimental_option(
        "prefs",
        {
            "download.default_directory": str(pasta_download),
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True,
            "printing.print_preview_sticky_settings.appState": json.dumps(
                {
                    "recentDestinations": [
                        {
                            "id": "Save as PDF",
                            "origin": "local",
                            "account": "",
                        }
                    ],
                    "selectedDestinationId": "Save as PDF",
                    "version": 2,
                }
            ),
        },
    )

    return opcoes


def criar_driver(
    porta: int = config.PORTA_DEPURACAO,
    perfil: Path = config.PASTA_PERFIL_CHROME,
    pasta_download: Path = config.PASTA_DOWNLOAD_TEMPORARIA,
    modo: str = config.MODO_NAVEGADOR,
    url_inicial: str = config.URL_LOGIN,
) -> tuple[Any, Optional[subprocess.Popen]]:
    """
    Cria o WebDriver.

    modo "anexar"  : a automação abre o Chrome e o Selenium se conecta a ele.
                     O login manual continua sendo feito na janela normal.
    modo "selenium": o Selenium abre e controla o Chrome diretamente.
    """
    garantir_selenium()

    from selenium.webdriver.chrome.service import Service

    caminho_driver = os.environ.get("PGFN_CHROMEDRIVER", "").strip()
    servico = Service(executable_path=caminho_driver) if caminho_driver else None

    processo: Optional[subprocess.Popen] = None

    if modo == "anexar":
        processo = abrir_chrome_com_depuracao(porta, perfil, url_inicial)

        opcoes = ChromeOptions()
        opcoes.add_experimental_option(
            "debuggerAddress", f"127.0.0.1:{porta}"
        )
    else:
        opcoes = _opcoes_chrome(pasta_download)
        opcoes.add_argument(f"--user-data-dir={perfil}")

        for argumento in config.ARGUMENTOS_CHROME:
            opcoes.add_argument(argumento)

        opcoes.add_argument("--kiosk-printing")

    driver = (
        webdriver.Chrome(options=opcoes, service=servico)
        if servico is not None
        else webdriver.Chrome(options=opcoes)
    )

    try:
        driver.set_page_load_timeout(config.TIMEOUT_CARREGAMENTO_PAGINA)
    except Exception:
        pass

    return driver, processo


def abrir_url(driver: Any, url: str) -> None:
    LOGGER.info("Abrindo: %s", url)

    try:
        driver.get(url)
    except TimeoutException:
        LOGGER.warning("Tempo de carregamento excedido em %s.", url)
    except WebDriverException as exc:
        raise EmpresaNaoProcessada(
            f"Não foi possível abrir {url}: {exc}", etapa="Navegação"
        ) from exc

    wait_for_page(driver)
