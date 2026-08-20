#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import json
import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


APP_DIR = Path(__file__).resolve().parent
DOWNLOAD_SCRIPT = APP_DIR / "ecac_download.py"
ANALYZER_SCRIPT = APP_DIR / "pdf_inaptos.py"
PGFN_SCRIPT = APP_DIR / "pgfn_regularize.py"
LOGO_PATH = APP_DIR / "assets" / "logo.png"

LOCAL_APP_DATA = Path(os.environ.get("LOCALAPPDATA", str(Path.home())))
CONFIG_DIR = LOCAL_APP_DATA / "METAPlanoAutomacao"
CONFIG_FILE = CONFIG_DIR / "config.json"

DEFAULT_EXCEL = r"C:\Users\z0057vzy\OneDrive - Siemens Healthineers\Desktop\Teste\Planilha_Clientes-1784744405619.xlsx"
DEFAULT_PDF_DIR = r"C:\Users\z0057vzy\OneDrive - Siemens Healthineers\Desktop\Saidas"
DEFAULT_REPORT = r"C:\Users\z0057vzy\OneDrive - Siemens Healthineers\Desktop\Relatorio_Empresas_Inaptas.xlsx"
DEFAULT_PGFN_DIR = r"C:\Users\z0057vzy\OneDrive - Siemens Healthineers\Desktop\Saidas PGFN"

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


class ToggleSwitch(tk.Canvas):
    def __init__(self, master, text: str, variable: tk.BooleanVar, command=None, **kwargs):
        super().__init__(
            master,
            width=270,
            height=52,
            bg=CARD,
            highlightthickness=0,
            bd=0,
            cursor="hand2",
            **kwargs,
        )
        self.variable = variable
        self.command = command
        self.label = text
        self.bind("<Button-1>", self.toggle)
        self.draw()

    def toggle(self, _event=None):
        self.variable.set(not self.variable.get())
        self.draw()
        if self.command:
            self.command()

    def draw(self):
        self.delete("all")
        value = self.variable.get()
        self.create_text(
            2,
            26,
            anchor="w",
            text=self.label,
            fill=TEXT,
            font=("Segoe UI", 10, "bold"),
        )
        x1, y1, x2, y2 = 215, 14, 265, 38
        fill = MAROON if value else "#C8C8C8"
        self.create_oval(x1, y1, x1 + 24, y2, fill=fill, outline=fill)
        self.create_oval(x2 - 24, y1, x2, y2, fill=fill, outline=fill)
        self.create_rectangle(x1 + 12, y1, x2 - 12, y2, fill=fill, outline=fill)
        knob_x = x2 - 19 if value else x1 + 5
        self.create_oval(
            knob_x,
            y1 + 5,
            knob_x + 14,
            y1 + 19,
            fill="white",
            outline="white",
        )


class MetaPlanoApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("META PLANO | Automação Receita Federal")
        self.geometry("1220x860")
        self.minsize(1080, 740)
        self.configure(bg=BG)

        self.process: subprocess.Popen | None = None
        self.worker_thread: threading.Thread | None = None
        self.events: queue.Queue = queue.Queue()
        self.stop_requested = False
        self.pipeline_stage = ""

        self.config_data = self.load_config()
        self.excel_var = tk.StringVar(value=self.config_data.get("excel", DEFAULT_EXCEL))
        self.pdf_dir_var = tk.StringVar(value=self.config_data.get("pdf_dir", DEFAULT_PDF_DIR))
        self.report_var = tk.StringVar(value=self.config_data.get("report", DEFAULT_REPORT))
        self.pgfn_dir_var = tk.StringVar(value=self.config_data.get("pgfn_dir", DEFAULT_PGFN_DIR))
        self.run_download_var = tk.BooleanVar(value=self.config_data.get("run_download", True))
        self.run_analysis_var = tk.BooleanVar(value=self.config_data.get("run_analysis", True))
        self.run_pgfn_var = tk.BooleanVar(value=self.config_data.get("run_pgfn", False))
        self.status_var = tk.StringVar(value="Pronto para iniciar")
        self.progress_var = tk.DoubleVar(value=0)

        self.logo_image = None
        self._build_styles()
        self._build_ui()
        self.after(100, self.process_events)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def _build_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(
            "Main.Horizontal.TProgressbar",
            troughcolor="#E6E6E6",
            background=MAROON,
            bordercolor="#E6E6E6",
            lightcolor=MAROON,
            darkcolor=MAROON,
        )

    def _build_ui(self):
        root = tk.Frame(self, bg=BG)
        root.pack(fill="both", expand=True)

        sidebar = tk.Frame(root, bg=MAROON_DARK, width=275)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        self._build_sidebar(sidebar)

        content = tk.Frame(root, bg=BG)
        content.pack(side="left", fill="both", expand=True)
        self._build_content(content)

    def _build_sidebar(self, parent):
        logo_holder = tk.Frame(parent, bg="white", height=170)
        logo_holder.pack(fill="x")
        logo_holder.pack_propagate(False)

        try:
            self.logo_image = tk.PhotoImage(file=str(LOGO_PATH))
            max_w, max_h = 215, 145
            sx = max(1, (self.logo_image.width() + max_w - 1) // max_w)
            sy = max(1, (self.logo_image.height() + max_h - 1) // max_h)
            factor = max(sx, sy)
            if factor > 1:
                self.logo_image = self.logo_image.subsample(factor, factor)
            tk.Label(logo_holder, image=self.logo_image, bg="white").pack(expand=True)
        except Exception:
            tk.Label(
                logo_holder,
                text="META PLANO\nCONTÁBIL",
                bg="white",
                fg=MAROON,
                font=("Georgia", 22, "bold"),
            ).pack(expand=True)

        tk.Label(
            parent,
            text="AUTOMAÇÃO FISCAL",
            bg=MAROON_DARK,
            fg="white",
            font=("Segoe UI", 13, "bold"),
        ).pack(anchor="w", padx=24, pady=(26, 4))
        tk.Label(
            parent,
            text="Receita Federal, análise de empresas\ninaptas e PGFN / Regularize",
            justify="left",
            bg=MAROON_DARK,
            fg="#EADDDD",
            font=("Segoe UI", 10),
        ).pack(anchor="w", padx=24)

        tk.Frame(parent, bg="#8F2A2A", height=1).pack(fill="x", padx=22, pady=22)

        tk.Label(
            parent,
            text="MODO DE EXECUÇÃO",
            bg=MAROON_DARK,
            fg="#EADDDD",
            font=("Segoe UI", 9, "bold"),
        ).pack(anchor="w", padx=24, pady=(0, 8))

        toggle_box = tk.Frame(parent, bg="white")
        toggle_box.pack(fill="x", padx=16, pady=4)
        ToggleSwitch(toggle_box, "1. Baixar relatórios", self.run_download_var, self.save_config).pack(padx=6)
        tk.Frame(toggle_box, height=1, bg=BORDER).pack(fill="x", padx=10)
        ToggleSwitch(toggle_box, "2. Analisar INAPTOS", self.run_analysis_var, self.save_config).pack(padx=6)
        tk.Frame(toggle_box, height=1, bg=BORDER).pack(fill="x", padx=10)
        ToggleSwitch(
            toggle_box,
            "3. Relatórios PGFN / Regularize",
            self.run_pgfn_var,
            self.save_config,
        ).pack(padx=6)

        tk.Label(
            parent,
            text="As etapas ligadas são executadas na ordem\n1, 2 e 3, uma depois da outra.",
            justify="left",
            bg=MAROON_DARK,
            fg="#EADDDD",
            font=("Segoe UI", 9),
        ).pack(anchor="w", padx=24, pady=(12, 0))

        tk.Label(
            parent,
            text="Pasta Saídas",
            bg=MAROON_DARK,
            fg="white",
            font=("Segoe UI", 9, "bold"),
        ).pack(anchor="w", padx=24, pady=(26, 2))
        tk.Label(
            parent,
            text="Somente PDFs finais renomeados.\nO PGFN usa uma pasta separada.",
            justify="left",
            bg=MAROON_DARK,
            fg="#EADDDD",
            font=("Segoe UI", 9),
        ).pack(anchor="w", padx=24)

        tk.Label(
            parent,
            text="v1.0  •  META PLANO",
            bg=MAROON_DARK,
            fg="#C9AAAA",
            font=("Segoe UI", 8),
        ).pack(side="bottom", pady=18)

    def _build_content(self, parent):
        header = tk.Frame(parent, bg=BG)
        header.pack(fill="x", padx=30, pady=(25, 14))
        tk.Label(
            header,
            text="Central de Automações",
            bg=BG,
            fg=TEXT,
            font=("Segoe UI", 22, "bold"),
        ).pack(anchor="w")
        tk.Label(
            header,
            text="Configure os arquivos, escolha as etapas e acompanhe toda a execução pelo console.",
            bg=BG,
            fg=MUTED,
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(4, 0))

        settings = tk.Frame(parent, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        settings.pack(fill="x", padx=30, pady=(0, 14))

        tk.Label(
            settings,
            text="Configuração",
            bg=CARD,
            fg=TEXT,
            font=("Segoe UI", 12, "bold"),
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=18, pady=(16, 12))

        self._path_row(settings, 1, "Planilha de clientes", self.excel_var, self.choose_excel, "Arquivo usado para os CNPJs ativos")
        self._path_row(settings, 2, "Pasta dos PDFs", self.pdf_dir_var, self.choose_pdf_dir, "Destino da automação 1 e fonte da automação 2")
        self._path_row(settings, 3, "Excel de empresas inaptas", self.report_var, self.choose_report, "Salvo fora da pasta de PDFs")
        self._path_row(settings, 4, "Pasta dos relatórios PGFN", self.pgfn_dir_var, self.choose_pgfn_dir, "Destino da automação 3 (Regularize)")

        settings.grid_columnconfigure(1, weight=1)

        action_bar = tk.Frame(parent, bg=BG)
        action_bar.pack(fill="x", padx=30, pady=(0, 14))

        self.run_button = tk.Button(
            action_bar,
            text="▶  INICIAR AUTOMAÇÃO",
            command=self.start,
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
        self.run_button.pack(side="left")

        self.continue_button = tk.Button(
            action_bar,
            text="↵  CONTINUAR / ENTER",
            command=self.send_enter,
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
        self.continue_button.pack(side="left", padx=10)

        self.stop_button = tk.Button(
            action_bar,
            text="■  PARAR",
            command=self.stop,
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
        self.stop_button.pack(side="left")

        status_card = tk.Frame(parent, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        status_card.pack(fill="x", padx=30, pady=(0, 14))
        tk.Label(
            status_card,
            textvariable=self.status_var,
            bg=CARD,
            fg=TEXT,
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", padx=16, pady=(12, 7))
        ttk.Progressbar(
            status_card,
            variable=self.progress_var,
            maximum=100,
            style="Main.Horizontal.TProgressbar",
        ).pack(fill="x", padx=16, pady=(0, 12))

        console_card = tk.Frame(parent, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        console_card.pack(fill="both", expand=True, padx=30, pady=(0, 25))

        console_header = tk.Frame(console_card, bg=DARK)
        console_header.pack(fill="x")
        tk.Label(
            console_header,
            text="Console de execução",
            bg=DARK,
            fg="white",
            font=("Consolas", 10, "bold"),
        ).pack(side="left", padx=14, pady=9)
        tk.Button(
            console_header,
            text="Limpar",
            command=self.clear_console,
            bg=DARK,
            fg="#BBBBBB",
            activebackground=DARK,
            activeforeground="white",
            relief="flat",
            bd=0,
            cursor="hand2",
        ).pack(side="right", padx=12)

        console_frame = tk.Frame(console_card, bg=CONSOLE_BG)
        console_frame.pack(fill="both", expand=True)
        scrollbar = tk.Scrollbar(console_frame)
        scrollbar.pack(side="right", fill="y")
        self.console = tk.Text(
            console_frame,
            bg=CONSOLE_BG,
            fg=CONSOLE_FG,
            insertbackground="white",
            font=("Consolas", 9),
            relief="flat",
            bd=0,
            wrap="word",
            yscrollcommand=scrollbar.set,
            padx=12,
            pady=10,
        )
        self.console.pack(fill="both", expand=True)
        scrollbar.config(command=self.console.yview)
        self.console.tag_configure("success", foreground="#6FD28A")
        self.console.tag_configure("warning", foreground="#F3C969")
        self.console.tag_configure("error", foreground="#FF7777")
        self.console.tag_configure("stage", foreground="#F4D35E", font=("Consolas", 9, "bold"))

        self.log("META PLANO Automation Center iniciado.\n", "stage")
        self.log("A pasta de PDFs deve permanecer contendo somente relatórios PDF finais.\n")
        self.log(
            "Os relatórios do PGFN / Regularize são salvos na pasta própria da automação 3.\n"
        )

    def _path_row(self, parent, row, title, var, command, subtitle):
        label_box = tk.Frame(parent, bg=CARD)
        label_box.grid(row=row, column=0, sticky="w", padx=(18, 12), pady=8)
        tk.Label(label_box, text=title, bg=CARD, fg=TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(label_box, text=subtitle, bg=CARD, fg=MUTED, font=("Segoe UI", 8)).pack(anchor="w")

        entry = tk.Entry(
            parent,
            textvariable=var,
            bg="#FAFAFA",
            fg=TEXT,
            relief="solid",
            bd=1,
            font=("Segoe UI", 9),
        )
        entry.grid(row=row, column=1, sticky="ew", pady=8, ipady=8)

        tk.Button(
            parent,
            text="Selecionar",
            command=command,
            bg="#EEEEEE",
            fg=TEXT,
            activebackground="#E0E0E0",
            relief="flat",
            bd=0,
            padx=14,
            pady=8,
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
        ).grid(row=row, column=2, padx=(10, 18), pady=8)

    def choose_excel(self):
        path = filedialog.askopenfilename(
            title="Selecione a planilha de clientes",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos os arquivos", "*.*")],
        )
        if path:
            self.excel_var.set(path)
            self.save_config()

    def choose_pdf_dir(self):
        path = filedialog.askdirectory(title="Selecione a pasta dos PDFs")
        if path:
            self.pdf_dir_var.set(path)
            current_report = Path(self.report_var.get()) if self.report_var.get() else None
            if not current_report or not current_report.parent.exists():
                self.report_var.set(str(Path(path).parent / "Relatorio_Empresas_Inaptas.xlsx"))
            self.save_config()

    def choose_pgfn_dir(self):
        path = filedialog.askdirectory(title="Selecione a pasta dos relatórios PGFN / Regularize")
        if path:
            self.pgfn_dir_var.set(path)
            self.save_config()

    def choose_report(self):
        path = filedialog.asksaveasfilename(
            title="Salvar relatório de empresas inaptas",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="Relatorio_Empresas_Inaptas.xlsx",
        )
        if path:
            self.report_var.set(path)
            self.save_config()

    def validate(self) -> bool:
        if (
            not self.run_download_var.get()
            and not self.run_analysis_var.get()
            and not self.run_pgfn_var.get()
        ):
            messagebox.showwarning("Selecione uma etapa", "Ative pelo menos uma das automações.")
            return False

        pdf_dir = Path(self.pdf_dir_var.get().strip())
        if not self.pdf_dir_var.get().strip():
            if self.run_download_var.get() or self.run_analysis_var.get():
                messagebox.showerror("Pasta de PDFs", "Selecione a pasta onde os PDFs serão salvos/analisados.")
                return False

        if self.run_download_var.get():
            excel = Path(self.excel_var.get().strip())
            if not excel.exists():
                messagebox.showerror("Planilha", "A planilha de clientes selecionada não existe.")
                return False
            pdf_dir.mkdir(parents=True, exist_ok=True)

        if self.run_analysis_var.get():
            if not pdf_dir.exists():
                messagebox.showerror("Pasta de PDFs", "A pasta de PDFs selecionada não existe.")
                return False
            report = self.report_var.get().strip()
            if not report:
                messagebox.showerror("Excel de inaptos", "Escolha onde salvar o Excel de empresas inaptas.")
                return False
            if Path(report).resolve().parent == pdf_dir.resolve():
                messagebox.showerror(
                    "Pasta de PDFs deve ficar limpa",
                    "O Excel de empresas inaptas não pode ser salvo dentro da pasta dos PDFs.\n\n"
                    "Escolha outro local para o Excel.",
                )
                return False

        if self.run_pgfn_var.get():
            excel = Path(self.excel_var.get().strip())
            if not excel.exists():
                messagebox.showerror("Planilha", "A planilha de clientes selecionada não existe.")
                return False

            pgfn_dir_text = self.pgfn_dir_var.get().strip()
            if not pgfn_dir_text:
                messagebox.showerror(
                    "Pasta dos relatórios PGFN",
                    "Selecione a pasta onde os relatórios do Regularize serão salvos.",
                )
                return False

            pgfn_dir = Path(pgfn_dir_text)
            if self.pdf_dir_var.get().strip() and pgfn_dir.resolve() == pdf_dir.resolve():
                messagebox.showerror(
                    "Pastas separadas",
                    "Os relatórios do PGFN / Regularize precisam de uma pasta própria.\n\n"
                    "A pasta dos PDFs do eCAC é usada pela análise de INAPTOS e deve conter "
                    "somente os relatórios da Receita Federal.",
                )
                return False

            pgfn_dir.mkdir(parents=True, exist_ok=True)

        return True

    def start(self):
        if self.worker_thread and self.worker_thread.is_alive():
            return
        if not self.validate():
            return

        self.save_config()
        self.stop_requested = False
        self.progress_var.set(0)
        self.run_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.continue_button.config(state="disabled")
        self.status_var.set("Iniciando...")
        self.log("\n" + "=" * 76 + "\n", "stage")
        self.log("NOVA EXECUÇÃO\n", "stage")
        self.log("=" * 76 + "\n", "stage")

        self.worker_thread = threading.Thread(target=self.run_pipeline, daemon=True)
        self.worker_thread.start()

    def enabled_stages(self) -> list[str]:
        stages = []
        if self.run_download_var.get():
            stages.append("download")
        if self.run_analysis_var.get():
            stages.append("analysis")
        if self.run_pgfn_var.get():
            stages.append("pgfn")
        return stages

    def stage_label(self, stage: str, stages: list[str]) -> str:
        titles = {
            "download": "Baixando relatórios da Receita Federal",
            "analysis": "Analisando situação cadastral dos PDFs",
            "pgfn": "Gerando relatórios PGFN / Regularize",
        }
        return (
            f"Automação {stages.index(stage) + 1}/{len(stages)} - {titles[stage]}"
        )

    def run_pipeline(self):
        stages = self.enabled_stages()
        try:
            if "download" in stages and not self.stop_requested:
                self.pipeline_stage = "download"
                self.events.put(("status", self.stage_label("download", stages)))
                self.events.put(("continue_enabled", True))
                code = self.run_download_process()
                self.events.put(("continue_enabled", False))

                if self.stop_requested:
                    return
                if code not in (0, 2):
                    self.events.put(("error", f"A automação de download terminou com código {code}. A sequência foi interrompida."))
                    return
                if code == 2:
                    self.events.put(("warning", "A automação 1 terminou com algumas empresas em erro. A análise seguirá com os PDFs disponíveis."))

            if "analysis" in stages and not self.stop_requested:
                self.pipeline_stage = "analysis"
                self.events.put(("progress", 0))
                self.events.put(("status", self.stage_label("analysis", stages)))
                code = self.run_analysis_process()
                if code != 0:
                    self.events.put(("error", f"A análise de PDFs terminou com código {code}."))
                    return

            if "pgfn" in stages and not self.stop_requested:
                self.pipeline_stage = "pgfn"
                self.events.put(("progress", 0))
                self.events.put(("status", self.stage_label("pgfn", stages)))
                self.events.put(("continue_enabled", True))
                code = self.run_pgfn_process()
                self.events.put(("continue_enabled", False))

                if self.stop_requested:
                    return
                if code not in (0, 2):
                    self.events.put(("error", f"A automação PGFN / Regularize terminou com código {code}."))
                    return
                if code == 2:
                    self.events.put((
                        "warning",
                        "A automação PGFN terminou com empresas em erro ou sem cadastro no Regularize. "
                        "Confira o PDF gerado com a lista dessas empresas.",
                    ))

            if not self.stop_requested:
                self.events.put(("progress", 100))
                self.events.put(("done", "Processamento concluído com sucesso."))
        except Exception as exc:
            self.events.put(("error", f"Falha inesperada: {exc}"))
        finally:
            self.process = None
            self.pipeline_stage = ""
            self.events.put(("finished", None))

    def run_download_process(self) -> int:
        env = os.environ.copy()
        env["PYTHONUNBUFFERED"] = "1"
        env["ECAC_EXCEL_PATH"] = self.excel_var.get().strip()
        env["ECAC_OUTPUT_DIR"] = self.pdf_dir_var.get().strip()

        cmd = [sys.executable, "-u", str(DOWNLOAD_SCRIPT)]
        return self.run_subprocess(cmd, env)

    def run_pgfn_process(self) -> int:
        env = os.environ.copy()
        env["PYTHONUNBUFFERED"] = "1"
        env["PGFN_EXCEL_PATH"] = self.excel_var.get().strip()
        env["PGFN_OUTPUT_DIR"] = self.pgfn_dir_var.get().strip()

        cmd = [sys.executable, "-u", str(PGFN_SCRIPT)]
        return self.run_subprocess(cmd, env)

    def run_analysis_process(self) -> int:
        env = os.environ.copy()
        env["PYTHONUNBUFFERED"] = "1"
        cmd = [
            sys.executable,
            "-u",
            str(ANALYZER_SCRIPT),
            "--pdf-dir",
            self.pdf_dir_var.get().strip(),
            "--output",
            self.report_var.get().strip(),
        ]
        return self.run_subprocess(cmd, env)

    def run_subprocess(self, cmd: list[str], env: dict[str, str]) -> int:
        creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        self.process = subprocess.Popen(
            cmd,
            cwd=str(APP_DIR),
            env=env,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            creationflags=creationflags,
        )

        assert self.process.stdout is not None
        for line in iter(self.process.stdout.readline, ""):
            if not line and self.process.poll() is not None:
                break
            if line:
                self.events.put(("log", line))
                self.parse_progress(line)
            if self.stop_requested:
                break

        if self.stop_requested and self.process.poll() is None:
            self.process.terminate()

        return self.process.wait()

    def parse_progress(self, line: str):
        match = re.search(r"\[(\d+)/(\d+)\]", line)
        if not match:
            match = re.search(r"\[SCAN\s+(\d+)/(\d+)\]", line, flags=re.IGNORECASE)
        if match:
            current = int(match.group(1))
            total = int(match.group(2))
            if total:
                self.events.put(("progress", current / total * 100))

    def send_enter(self):
        proc = self.process
        if not proc or proc.poll() is not None or proc.stdin is None:
            return
        try:
            proc.stdin.write("\n")
            proc.stdin.flush()
            self.log("[APP] ENTER enviado para a automação.\n", "stage")
        except Exception as exc:
            self.log(f"[APP] Não foi possível enviar ENTER: {exc}\n", "error")

    def stop(self):
        if not self.process and not (self.worker_thread and self.worker_thread.is_alive()):
            return
        if not messagebox.askyesno("Parar automação", "Deseja realmente interromper a execução atual?"):
            return
        self.stop_requested = True
        self.status_var.set("Interrompendo...")
        self.log("[APP] Interrupção solicitada pelo usuário.\n", "warning")
        if self.process and self.process.poll() is None:
            try:
                self.process.terminate()
            except Exception:
                pass

    def process_events(self):
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "log":
                    tag = None
                    low = str(payload).lower()
                    if "erro" in low or "critical" in low:
                        tag = "error"
                    elif "warning" in low or "aviso" in low:
                        tag = "warning"
                    elif "sucesso" in low or "conclu" in low or "pdf salvo" in low:
                        tag = "success"
                    self.log(str(payload), tag)
                elif kind == "status":
                    self.status_var.set(str(payload))
                    self.log(f"\n>>> {payload}\n", "stage")
                elif kind == "progress":
                    self.progress_var.set(float(payload))
                elif kind == "continue_enabled":
                    self.continue_button.config(state="normal" if payload else "disabled")
                elif kind == "warning":
                    self.log(str(payload) + "\n", "warning")
                elif kind == "error":
                    self.status_var.set("Execução finalizada com erro")
                    self.log(str(payload) + "\n", "error")
                elif kind == "done":
                    self.status_var.set(str(payload))
                    self.log(str(payload) + "\n", "success")
                elif kind == "finished":
                    self.run_button.config(state="normal")
                    self.stop_button.config(state="disabled")
                    self.continue_button.config(state="disabled")
        except queue.Empty:
            pass
        self.after(100, self.process_events)

    def log(self, text: str, tag=None):
        self.console.configure(state="normal")
        if tag:
            self.console.insert("end", text, tag)
        else:
            self.console.insert("end", text)
        self.console.see("end")
        self.console.configure(state="disabled")

    def clear_console(self):
        self.console.configure(state="normal")
        self.console.delete("1.0", "end")
        self.console.configure(state="disabled")

    def load_config(self) -> dict:
        try:
            if CONFIG_FILE.exists():
                return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
        return {}

    def save_config(self):
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            data = {
                "excel": self.excel_var.get(),
                "pdf_dir": self.pdf_dir_var.get(),
                "report": self.report_var.get(),
                "pgfn_dir": self.pgfn_dir_var.get(),
                "run_download": self.run_download_var.get(),
                "run_analysis": self.run_analysis_var.get(),
                "run_pgfn": self.run_pgfn_var.get(),
            }
            CONFIG_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def on_close(self):
        if self.worker_thread and self.worker_thread.is_alive():
            if not messagebox.askyesno("Sair", "Há uma automação em execução. Deseja interromper e sair?"):
                return
            self.stop_requested = True
            if self.process and self.process.poll() is None:
                try:
                    self.process.terminate()
                except Exception:
                    pass
        self.save_config()
        self.destroy()


def main():
    app = MetaPlanoApp()
    app.mainloop()


if __name__ == "__main__":
    main()
