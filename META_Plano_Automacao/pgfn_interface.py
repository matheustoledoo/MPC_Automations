#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Interface gráfica da automação PGFN / Regularize.

Tela mínima pedida na especificação:

    - seleção da planilha de clientes;
    - seleção da pasta de destino dos PDFs;
    - botão iniciar;
    - empresa atual, CNPJ, progresso e status;
    - contadores (total, processadas, sucesso, não processadas, pendentes);
    - console com o log completo da execução.

A automação roda como processo separado (`pgfn_automation.py`), da mesma
forma que o `app.py` faz com o `ecac_download.py`. Assim a janela nunca
congela e o botão CONTINUAR consegue responder aos momentos em que o
portal exige uma ação manual (login e primeira representação).
"""

from __future__ import annotations

import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


APP_DIR = Path(__file__).resolve().parent
SCRIPT_AUTOMACAO = APP_DIR / "pgfn_automation.py"
LOGO_PATH = APP_DIR / "assets" / "logo.png"

if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

LOCAL_APP_DATA = Path(os.environ.get("LOCALAPPDATA", str(Path.home())))
CONFIG_DIR = LOCAL_APP_DATA / "METAPlanoAutomacao"
CONFIG_FILE = CONFIG_DIR / "config_pgfn.json"

MARCADOR_EVENTO = "@@PGFN@@"

# Paleta compartilhada com a Central de Automações.
try:
    from app import (
        BG,
        BORDER,
        CARD,
        CONSOLE_BG,
        CONSOLE_FG,
        DARK,
        GOLD,
        MAROON,
        MAROON_DARK,
        MUTED,
        SUCCESS,
        TEXT,
        WARNING,
    )
except Exception:  # a interface funciona mesmo sozinha nesta pasta
    BG = "#F5F5F4"
    CARD = "#FFFFFF"
    MAROON = "#8B0000"
    MAROON_DARK = "#650000"
    GOLD = "#D4AF37"
    TEXT = "#252525"
    MUTED = "#777777"
    BORDER = "#E3E3E3"
    DARK = "#171717"
    CONSOLE_BG = "#101010"
    CONSOLE_FG = "#E8E8E8"
    SUCCESS = "#277A3F"
    WARNING = "#A15C00"

try:
    import config as pgfn_config

    PASTA_PADRAO = str(pgfn_config.PASTA_DESTINO_PADRAO)
except Exception:
    PASTA_PADRAO = str(Path.home() / "Desktop" / "Relatorios_PGFN")


class PGFNApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("META PLANO | PGFN / Regularize - Dívida Ativa da União")
        self.geometry("1220x820")
        self.minsize(1080, 720)
        self.configure(bg=BG)

        self.process: subprocess.Popen | None = None
        self.worker: threading.Thread | None = None
        self.eventos: queue.Queue = queue.Queue()
        self.parada_solicitada = False

        dados = self.carregar_config()
        self.planilha_var = tk.StringVar(value=dados.get("planilha", ""))
        self.pasta_var = tk.StringVar(
            value=dados.get("pasta_destino", PASTA_PADRAO)
        )

        self.empresa_var = tk.StringVar(value="-")
        self.cnpj_var = tk.StringVar(value="-")
        self.progresso_texto_var = tk.StringVar(value="0 / 0")
        self.status_var = tk.StringVar(value="Pronto para iniciar")
        self.progresso_var = tk.DoubleVar(value=0)

        self.contadores = {
            "total": tk.StringVar(value="0"),
            "processadas": tk.StringVar(value="0"),
            "sucesso": tk.StringVar(value="0"),
            "falhas": tk.StringVar(value="0"),
            "pendentes": tk.StringVar(value="0"),
        }

        self.logo_image = None
        self._montar_estilos()
        self._montar_tela()

        self.after(100, self._consumir_eventos)
        self.protocol("WM_DELETE_WINDOW", self.ao_fechar)

    # ------------------------------------------------------------------
    # TELA
    # ------------------------------------------------------------------

    def _montar_estilos(self) -> None:
        estilo = ttk.Style(self)

        try:
            estilo.theme_use("clam")
        except Exception:
            pass

        estilo.configure(
            "PGFN.Horizontal.TProgressbar",
            troughcolor="#E6E6E6",
            background=MAROON,
            bordercolor="#E6E6E6",
            lightcolor=MAROON,
            darkcolor=MAROON,
        )

    def _montar_tela(self) -> None:
        raiz = tk.Frame(self, bg=BG)
        raiz.pack(fill="both", expand=True)

        lateral = tk.Frame(raiz, bg=MAROON_DARK, width=275)
        lateral.pack(side="left", fill="y")
        lateral.pack_propagate(False)
        self._montar_lateral(lateral)

        conteudo = tk.Frame(raiz, bg=BG)
        conteudo.pack(side="left", fill="both", expand=True)
        self._montar_conteudo(conteudo)

    def _montar_lateral(self, pai: tk.Frame) -> None:
        caixa_logo = tk.Frame(pai, bg="white", height=170)
        caixa_logo.pack(fill="x")
        caixa_logo.pack_propagate(False)

        try:
            self.logo_image = tk.PhotoImage(file=str(LOGO_PATH))
            largura, altura = 215, 145
            fator = max(
                max(1, (self.logo_image.width() + largura - 1) // largura),
                max(1, (self.logo_image.height() + altura - 1) // altura),
            )

            if fator > 1:
                self.logo_image = self.logo_image.subsample(fator, fator)

            tk.Label(caixa_logo, image=self.logo_image, bg="white").pack(
                expand=True
            )
        except Exception:
            tk.Label(
                caixa_logo,
                text="META PLANO\nCONTÁBIL",
                bg="white",
                fg=MAROON,
                font=("Georgia", 22, "bold"),
            ).pack(expand=True)

        tk.Label(
            pai,
            text="PGFN / REGULARIZE",
            bg=MAROON_DARK,
            fg="white",
            font=("Segoe UI", 13, "bold"),
        ).pack(anchor="w", padx=24, pady=(26, 4))

        tk.Label(
            pai,
            text="Relatório Consolidado da\nDívida Ativa da União",
            justify="left",
            bg=MAROON_DARK,
            fg="#EADDDD",
            font=("Segoe UI", 10),
        ).pack(anchor="w", padx=24)

        tk.Frame(pai, bg="#8F2A2A", height=1).pack(fill="x", padx=22, pady=22)

        tk.Label(
            pai,
            text="COMO FUNCIONA",
            bg=MAROON_DARK,
            fg="#EADDDD",
            font=("Segoe UI", 9, "bold"),
        ).pack(anchor="w", padx=24, pady=(0, 8))

        passos = (
            "1. Escolha a planilha e a pasta.\n"
            "2. Clique em INICIAR.\n"
            "3. Faça o login na Receita.\n"
            "4. Represente a 1ª empresa.\n"
            "5. Clique em CONTINUAR.\n"
            "6. O robô faz o restante."
        )

        tk.Label(
            pai,
            text=passos,
            justify="left",
            bg=MAROON_DARK,
            fg="#EADDDD",
            font=("Segoe UI", 9),
        ).pack(anchor="w", padx=24)

        tk.Label(
            pai,
            text=(
                "Empresas sem cadastro no\nRegularize não param o lote:\n"
                "elas entram no relatório final\nde não processadas."
            ),
            justify="left",
            bg=MAROON_DARK,
            fg="#C9AAAA",
            font=("Segoe UI", 9),
        ).pack(anchor="w", padx=24, pady=(20, 0))

        tk.Label(
            pai,
            text="v1.0  •  META PLANO",
            bg=MAROON_DARK,
            fg="#C9AAAA",
            font=("Segoe UI", 8),
        ).pack(side="bottom", pady=18)

    def _montar_conteudo(self, pai: tk.Frame) -> None:
        cabecalho = tk.Frame(pai, bg=BG)
        cabecalho.pack(fill="x", padx=30, pady=(25, 14))

        tk.Label(
            cabecalho,
            text="Situação Fiscal PGFN",
            bg=BG,
            fg=TEXT,
            font=("Segoe UI", 22, "bold"),
        ).pack(anchor="w")

        tk.Label(
            cabecalho,
            text=(
                "Emite o Relatório Consolidado da Dívida Ativa da União para "
                "todas as empresas da base."
            ),
            bg=BG,
            fg=MUTED,
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(4, 0))

        configuracao = tk.Frame(
            pai, bg=CARD, highlightthickness=1, highlightbackground=BORDER
        )
        configuracao.pack(fill="x", padx=30, pady=(0, 14))

        tk.Label(
            configuracao,
            text="Configuração",
            bg=CARD,
            fg=TEXT,
            font=("Segoe UI", 12, "bold"),
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=18, pady=(16, 12))

        self._linha_caminho(
            configuracao,
            1,
            "Planilha de clientes",
            self.planilha_var,
            self.escolher_planilha,
            "A mesma base usada na automação do e-CAC",
        )
        self._linha_caminho(
            configuracao,
            2,
            "Pasta de destino",
            self.pasta_var,
            self.escolher_pasta,
            "Onde os PDFs serão salvos (CNPJ + nome da empresa)",
        )
        configuracao.grid_columnconfigure(1, weight=1)

        barra = tk.Frame(pai, bg=BG)
        barra.pack(fill="x", padx=30, pady=(0, 14))

        self.botao_iniciar = tk.Button(
            barra,
            text="▶  INICIAR AUTOMAÇÃO",
            command=self.iniciar,
            bg=MAROON,
            fg="white",
            activebackground=MAROON_DARK,
            activeforeground="white",
            relief="flat",
            bd=0,
            padx=24,
            pady=12,
            font=("Segoe UI", 10, "bold"),
            cursor="hand2",
        )
        self.botao_iniciar.pack(side="left")

        self.botao_continuar = tk.Button(
            barra,
            text="↵  CONTINUAR / ENTER",
            command=self.enviar_enter,
            bg=GOLD,
            fg="#2B2200",
            activebackground="#C6A331",
            relief="flat",
            bd=0,
            padx=20,
            pady=12,
            font=("Segoe UI", 10, "bold"),
            cursor="hand2",
            state="disabled",
        )
        self.botao_continuar.pack(side="left", padx=10)

        self.botao_parar = tk.Button(
            barra,
            text="■  PARAR",
            command=self.parar,
            bg="#555555",
            fg="white",
            activebackground="#333333",
            relief="flat",
            bd=0,
            padx=18,
            pady=12,
            font=("Segoe UI", 10, "bold"),
            cursor="hand2",
            state="disabled",
        )
        self.botao_parar.pack(side="left")

        self._montar_status(pai)
        self._montar_contadores(pai)
        self._montar_console(pai)

    def _linha_caminho(self, pai, linha, titulo, variavel, comando, ajuda):
        caixa = tk.Frame(pai, bg=CARD)
        caixa.grid(row=linha, column=0, sticky="w", padx=(18, 12), pady=8)

        tk.Label(
            caixa, text=titulo, bg=CARD, fg=TEXT, font=("Segoe UI", 9, "bold")
        ).pack(anchor="w")
        tk.Label(
            caixa, text=ajuda, bg=CARD, fg=MUTED, font=("Segoe UI", 8)
        ).pack(anchor="w")

        entrada = tk.Entry(
            pai,
            textvariable=variavel,
            bg="#FAFAFA",
            fg=TEXT,
            relief="solid",
            bd=1,
            font=("Segoe UI", 9),
        )
        entrada.grid(row=linha, column=1, sticky="ew", pady=8, ipady=8)

        tk.Button(
            pai,
            text="Selecionar",
            command=comando,
            bg="#EEEEEE",
            fg=TEXT,
            activebackground="#E0E0E0",
            relief="flat",
            bd=0,
            padx=14,
            pady=8,
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
        ).grid(row=linha, column=2, padx=(10, 18), pady=8)

    def _montar_status(self, pai: tk.Frame) -> None:
        cartao = tk.Frame(
            pai, bg=CARD, highlightthickness=1, highlightbackground=BORDER
        )
        cartao.pack(fill="x", padx=30, pady=(0, 12))

        linha = tk.Frame(cartao, bg=CARD)
        linha.pack(fill="x", padx=16, pady=(14, 6))

        self._campo_status(linha, "Empresa atual", self.empresa_var, 0, 320)
        self._campo_status(linha, "CNPJ", self.cnpj_var, 1, 170)
        self._campo_status(linha, "Progresso", self.progresso_texto_var, 2, 110)
        self._campo_status(linha, "Status", self.status_var, 3, 330)

        ttk.Progressbar(
            cartao,
            variable=self.progresso_var,
            maximum=100,
            style="PGFN.Horizontal.TProgressbar",
        ).pack(fill="x", padx=16, pady=(4, 14))

    def _campo_status(self, pai, titulo, variavel, coluna, largura) -> None:
        caixa = tk.Frame(pai, bg=CARD, width=largura)
        caixa.grid(row=0, column=coluna, sticky="w", padx=(0, 18))
        caixa.grid_propagate(False)

        tk.Label(
            caixa, text=titulo, bg=CARD, fg=MUTED, font=("Segoe UI", 8, "bold")
        ).pack(anchor="w")
        tk.Label(
            caixa,
            textvariable=variavel,
            bg=CARD,
            fg=TEXT,
            font=("Segoe UI", 10, "bold"),
            anchor="w",
            wraplength=largura,
            justify="left",
        ).pack(anchor="w", fill="x")

    def _montar_contadores(self, pai: tk.Frame) -> None:
        cartao = tk.Frame(pai, bg=BG)
        cartao.pack(fill="x", padx=30, pady=(0, 12))

        rotulos = (
            ("Total de empresas", "total", TEXT),
            ("Processadas", "processadas", TEXT),
            ("Sucesso", "sucesso", SUCCESS),
            ("Não processadas", "falhas", MAROON),
            ("Pendentes", "pendentes", WARNING),
        )

        for coluna, (titulo, chave, cor) in enumerate(rotulos):
            caixa = tk.Frame(
                cartao, bg=CARD, highlightthickness=1, highlightbackground=BORDER
            )
            caixa.grid(row=0, column=coluna, sticky="ew", padx=(0, 10))
            cartao.grid_columnconfigure(coluna, weight=1)

            tk.Label(
                caixa,
                text=titulo.upper(),
                bg=CARD,
                fg=MUTED,
                font=("Segoe UI", 8, "bold"),
            ).pack(anchor="w", padx=14, pady=(10, 0))
            tk.Label(
                caixa,
                textvariable=self.contadores[chave],
                bg=CARD,
                fg=cor,
                font=("Segoe UI", 18, "bold"),
            ).pack(anchor="w", padx=14, pady=(0, 10))

    def _montar_console(self, pai: tk.Frame) -> None:
        cartao = tk.Frame(
            pai, bg=CARD, highlightthickness=1, highlightbackground=BORDER
        )
        cartao.pack(fill="both", expand=True, padx=30, pady=(0, 25))

        cabecalho = tk.Frame(cartao, bg=DARK)
        cabecalho.pack(fill="x")

        tk.Label(
            cabecalho,
            text="Console de execução",
            bg=DARK,
            fg="white",
            font=("Consolas", 10, "bold"),
        ).pack(side="left", padx=14, pady=9)

        tk.Button(
            cabecalho,
            text="Limpar",
            command=self.limpar_console,
            bg=DARK,
            fg="#BBBBBB",
            activebackground=DARK,
            activeforeground="white",
            relief="flat",
            bd=0,
            cursor="hand2",
        ).pack(side="right", padx=12)

        moldura = tk.Frame(cartao, bg=CONSOLE_BG)
        moldura.pack(fill="both", expand=True)

        barra = tk.Scrollbar(moldura)
        barra.pack(side="right", fill="y")

        self.console = tk.Text(
            moldura,
            bg=CONSOLE_BG,
            fg=CONSOLE_FG,
            insertbackground="white",
            font=("Consolas", 9),
            relief="flat",
            bd=0,
            wrap="word",
            yscrollcommand=barra.set,
            padx=12,
            pady=10,
        )
        self.console.pack(fill="both", expand=True)
        barra.config(command=self.console.yview)

        self.console.tag_configure("success", foreground="#6FD28A")
        self.console.tag_configure("warning", foreground="#F3C969")
        self.console.tag_configure("error", foreground="#FF7777")
        self.console.tag_configure(
            "stage", foreground="#F4D35E", font=("Consolas", 9, "bold")
        )

        self.log("Automação PGFN / Regularize pronta.\n", "stage")
        self.log(
            "Os PDFs são salvos como CNPJ + nome da empresa na pasta "
            "escolhida.\n"
        )

    # ------------------------------------------------------------------
    # AÇÕES
    # ------------------------------------------------------------------

    def escolher_planilha(self) -> None:
        caminho = filedialog.askopenfilename(
            title="Selecione a planilha de clientes",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos os arquivos", "*.*")],
        )

        if caminho:
            self.planilha_var.set(caminho)
            self.salvar_config()

    def escolher_pasta(self) -> None:
        caminho = filedialog.askdirectory(
            title="Selecione a pasta de destino dos PDFs"
        )

        if caminho:
            self.pasta_var.set(caminho)
            self.salvar_config()

    def validar(self) -> bool:
        planilha = self.planilha_var.get().strip()

        if planilha and not Path(planilha).exists():
            messagebox.showerror(
                "Planilha", "A planilha selecionada não existe."
            )
            return False

        pasta = self.pasta_var.get().strip()

        if not pasta:
            messagebox.showerror(
                "Pasta de destino",
                "Escolha a pasta onde os PDFs serão salvos.",
            )
            return False

        try:
            Path(pasta).mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            messagebox.showerror(
                "Pasta de destino", f"Não foi possível usar a pasta:\n{exc}"
            )
            return False

        return True

    def iniciar(self) -> None:
        if self.worker and self.worker.is_alive():
            return

        if not self.validar():
            return

        self.salvar_config()
        self.parada_solicitada = False
        self.progresso_var.set(0)
        self.botao_iniciar.config(state="disabled")
        self.botao_parar.config(state="normal")
        self.botao_continuar.config(state="normal")
        self.status_var.set("Iniciando...")

        for variavel in self.contadores.values():
            variavel.set("0")

        self.log("\n" + "=" * 76 + "\n", "stage")
        self.log("NOVA EXECUÇÃO — PGFN / REGULARIZE\n", "stage")
        self.log("=" * 76 + "\n", "stage")

        self.worker = threading.Thread(target=self.executar, daemon=True)
        self.worker.start()

    def executar(self) -> None:
        try:
            ambiente = os.environ.copy()
            ambiente["PYTHONUNBUFFERED"] = "1"
            ambiente["PYTHONIOENCODING"] = "utf-8"

            comando = [
                sys.executable,
                "-u",
                str(SCRIPT_AUTOMACAO),
                "--pasta-destino",
                self.pasta_var.get().strip(),
            ]

            planilha = self.planilha_var.get().strip()

            if planilha:
                comando += ["--planilha", planilha]

            codigo = self.rodar_processo(comando, ambiente)

            if self.parada_solicitada:
                self.eventos.put(("warning", "Execução interrompida."))
            elif codigo == 0:
                self.eventos.put(
                    ("done", "Processamento concluído. Todas as empresas OK.")
                )
            elif codigo == 2:
                self.eventos.put(
                    (
                        "warning",
                        "Processamento concluído com empresas não "
                        "processadas. Veja o relatório final.",
                    )
                )
            else:
                self.eventos.put(
                    ("error", f"A automação terminou com código {codigo}.")
                )
        except Exception as exc:
            self.eventos.put(("error", f"Falha inesperada: {exc}"))
        finally:
            self.process = None
            self.eventos.put(("finished", None))

    def rodar_processo(self, comando: list[str], ambiente: dict) -> int:
        criacao = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0

        self.process = subprocess.Popen(
            comando,
            cwd=str(APP_DIR),
            env=ambiente,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            creationflags=criacao,
        )

        assert self.process.stdout is not None

        for linha in iter(self.process.stdout.readline, ""):
            if not linha and self.process.poll() is not None:
                break

            if linha.startswith(MARCADOR_EVENTO):
                self.eventos.put(("evento", linha[len(MARCADOR_EVENTO) :]))
            elif linha:
                self.eventos.put(("log", linha))

            if self.parada_solicitada:
                break

        if self.parada_solicitada and self.process.poll() is None:
            self.process.terminate()

        return self.process.wait()

    def enviar_enter(self) -> None:
        processo = self.process

        if (
            not processo
            or processo.poll() is not None
            or processo.stdin is None
        ):
            return

        try:
            processo.stdin.write("\n")
            processo.stdin.flush()
            self.log("[APP] ENTER enviado para a automação.\n", "stage")
            self.botao_continuar.config(bg=GOLD)
        except Exception as exc:
            self.log(f"[APP] Não foi possível enviar ENTER: {exc}\n", "error")

    def parar(self) -> None:
        if not self.process and not (self.worker and self.worker.is_alive()):
            return

        if not messagebox.askyesno(
            "Parar automação", "Deseja realmente interromper a execução?"
        ):
            return

        self.parada_solicitada = True
        self.status_var.set("Interrompendo...")
        self.log("[APP] Interrupção solicitada pelo usuário.\n", "warning")

        if self.process and self.process.poll() is None:
            try:
                self.process.terminate()
            except Exception:
                pass

    # ------------------------------------------------------------------
    # EVENTOS
    # ------------------------------------------------------------------

    def _consumir_eventos(self) -> None:
        try:
            while True:
                tipo, dado = self.eventos.get_nowait()

                if tipo == "log":
                    self.log(str(dado), self._marcar(str(dado)))
                elif tipo == "evento":
                    self._tratar_evento(str(dado))
                elif tipo == "warning":
                    self.status_var.set(str(dado))
                    self.log(str(dado) + "\n", "warning")
                elif tipo == "error":
                    self.status_var.set("Execução finalizada com erro")
                    self.log(str(dado) + "\n", "error")
                elif tipo == "done":
                    self.status_var.set(str(dado))
                    self.progresso_var.set(100)
                    self.log(str(dado) + "\n", "success")
                elif tipo == "finished":
                    self.botao_iniciar.config(state="normal")
                    self.botao_parar.config(state="disabled")
                    self.botao_continuar.config(state="disabled", bg=GOLD)
        except queue.Empty:
            pass

        self.after(100, self._consumir_eventos)

    def _tratar_evento(self, bruto: str) -> None:
        try:
            dados = json.loads(bruto)
        except Exception:
            return

        evento = dados.get("evento", "")

        if evento == "inicio":
            total = int(dados.get("total", 0))
            self.contadores["total"].set(str(total))
            self.contadores["pendentes"].set(str(total))
            self.progresso_texto_var.set(f"0 / {total}")
            self.log(
                f"[APP] {total} empresa(s) na base. Pasta: "
                f"{dados.get('pasta', '')}\n",
                "stage",
            )
        elif evento == "empresa":
            indice = int(dados.get("indice", 0))
            total = int(dados.get("total", 0)) or 1
            self.empresa_var.set(str(dados.get("nome", "-")))
            self.cnpj_var.set(str(dados.get("cnpj", "-")))
            self.progresso_texto_var.set(f"{indice} / {total}")
            self.progresso_var.set((indice - 1) / total * 100)
        elif evento == "status":
            self.status_var.set(str(dados.get("texto", "")))
        elif evento == "contadores":
            for chave in ("total", "processadas", "sucesso", "falhas", "pendentes"):
                if chave in dados:
                    self.contadores[chave].set(str(dados[chave]))

            total = int(dados.get("total", 0)) or 1
            self.progresso_var.set(
                int(dados.get("processadas", 0)) / total * 100
            )
        elif evento == "aguardando_usuario":
            self.status_var.set("AGUARDANDO VOCÊ — use o botão CONTINUAR")
            self.botao_continuar.config(state="normal", bg="#FF9F1C")
            self.log(
                "[APP] A automação está esperando uma ação sua. "
                "Faça o que o console pede e clique em CONTINUAR.\n",
                "warning",
            )
        elif evento == "usuario_respondeu":
            self.botao_continuar.config(bg=GOLD)
        elif evento == "pdf":
            self.log(f"[APP] PDF salvo: {dados.get('arquivo', '')}\n", "success")
        elif evento == "falha":
            self.log(
                f"[APP] Não processada: {dados.get('cnpj', '')} "
                f"{dados.get('nome', '')} — {dados.get('motivo', '')}\n",
                "error",
            )
        elif evento == "relatorio_falhas":
            for arquivo in dados.get("arquivos", []):
                self.log(f"[APP] Relatório de falhas: {arquivo}\n", "warning")
        elif evento == "fim":
            self.progresso_var.set(100)
            self.status_var.set(
                f"Fim: {dados.get('sucesso', 0)} com sucesso, "
                f"{dados.get('falhas', 0)} não processadas."
            )

    @staticmethod
    def _marcar(texto: str) -> str | None:
        baixo = texto.lower()

        if "erro" in baixo or "critical" in baixo or "não processada" in baixo:
            return "error"

        if "aviso" in baixo or "warning" in baixo or "aguardando" in baixo:
            return "warning"

        if "pdf salvo" in baixo or "concluí" in baixo or "sucesso" in baixo:
            return "success"

        return None

    # ------------------------------------------------------------------
    # CONSOLE E CONFIGURAÇÃO
    # ------------------------------------------------------------------

    def log(self, texto: str, tag: str | None = None) -> None:
        self.console.configure(state="normal")

        if tag:
            self.console.insert("end", texto, tag)
        else:
            self.console.insert("end", texto)

        self.console.see("end")
        self.console.configure(state="disabled")

    def limpar_console(self) -> None:
        self.console.configure(state="normal")
        self.console.delete("1.0", "end")
        self.console.configure(state="disabled")

    def carregar_config(self) -> dict:
        try:
            if CONFIG_FILE.exists():
                return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass

        return {}

    def salvar_config(self) -> None:
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            CONFIG_FILE.write_text(
                json.dumps(
                    {
                        "planilha": self.planilha_var.get(),
                        "pasta_destino": self.pasta_var.get(),
                    },
                    ensure_ascii=False,
                    indent=2,
                ),
                encoding="utf-8",
            )
        except Exception:
            pass

    def ao_fechar(self) -> None:
        if self.worker and self.worker.is_alive():
            if not messagebox.askyesno(
                "Sair", "Há uma automação em execução. Deseja interromper?"
            ):
                return

            self.parada_solicitada = True

            if self.process and self.process.poll() is None:
                try:
                    self.process.terminate()
                except Exception:
                    pass

        self.salvar_config()
        self.destroy()


def main() -> None:
    PGFNApp().mainloop()


if __name__ == "__main__":
    main()
