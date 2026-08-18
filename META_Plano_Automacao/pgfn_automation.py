#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Automação PGFN / Regularize — Relatório Consolidado da Dívida Ativa da União.

O que este arquivo faz, para cada empresa da MESMA base de CNPJs usada pela
automação do e-CAC (`ecac_download.py`):

    representar a empresa
        -> acessar o e-CAC
        -> Dívida Ativa da União
        -> PGFN, todos os serviços do Regularize
        -> Consultar Dívida Ativa
        -> Relatório Consolidado
        -> marcar todas as naturezas e todas as situações
        -> Gerar Relatório
        -> Imprimir
        -> salvar o PDF como "CNPJ - EMPRESA.pdf"
        -> próxima empresa

O login e a representação da PRIMEIRA empresa são manuais, exatamente como
na automação do e-CAC. Depois disso o robô assume o controle.

Uma empresa com problema (sem cadastro no Regularize, página indisponível,
elemento que não apareceu) nunca derruba o lote: ela é registrada e o fluxo
segue. No fim é gerado um PDF com todas as empresas não processadas.

Reaproveitamento do `ecac_download.py`:
    - leitura da base de CNPJs (load_clients / resolve_excel_path);
    - modelo Client, validação e formatação de CNPJ;
    - saneamento de nomes de arquivo;
    - localização do executável do Chrome.

Uso:
    python pgfn_automation.py --pasta-destino "C:\\Relatorios\\PGFN"
"""

from __future__ import annotations

import argparse
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

import config
import pgfn_utils as utils
from pgfn_utils import EmpresaNaoProcessada, LOGGER


ETAPA_REPRESENTACAO = "Representação"
ETAPA_ECAC = "e-CAC"
ETAPA_DIVIDA_ATIVA = "Dívida Ativa da União"
ETAPA_REGULARIZE = "Regularize"
ETAPA_CONSULTA = "Consultar Dívida Ativa"
ETAPA_RELATORIO = "Relatório Consolidado"
ETAPA_FILTROS = "Naturezas e situações"
ETAPA_GERACAO = "Gerar relatório"
ETAPA_IMPRESSAO = "Imprimir"
ETAPA_PDF = "Salvar PDF"


class AutomacaoPGFN:
    """Controla o lote inteiro: navegador, empresas, relatórios e contadores."""

    def __init__(
        self,
        pasta_destino: Path,
        planilha: str = "",
        limite: int = 0,
        apenas_cnpj: str = "",
        pular_login_manual: bool = False,
        fechar_navegador: bool = False,
        modo_navegador: str = config.MODO_NAVEGADOR,
    ) -> None:
        self.pasta_destino = Path(pasta_destino).expanduser()
        self.planilha = planilha
        self.limite = max(0, int(limite))
        self.apenas_cnpj = utils.so_digitos(apenas_cnpj)
        self.pular_login_manual = pular_login_manual
        self.fechar_navegador = fechar_navegador
        self.modo_navegador = modo_navegador

        self.carimbo = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.driver: Any = None
        self.processo_chrome: Any = None
        self.aba_principal: str = ""

        self.registro = utils.RegistroFalhas()
        self.resultados: list[dict[str, str]] = []

        self.total = 0
        self.processadas = 0
        self.sucessos = 0

    # -------------------------------------------------------------------
    # CICLO PRINCIPAL
    # -------------------------------------------------------------------

    def executar(self) -> int:
        arquivo_log = utils.configurar_log(config.PASTA_LOGS, self.carimbo)

        LOGGER.info("Iniciando processamento")
        LOGGER.info("Log desta execução: %s", arquivo_log)

        self.pasta_destino.mkdir(parents=True, exist_ok=True)
        LOGGER.info("Pasta de destino dos PDFs: %s", self.pasta_destino)

        planilha = utils.resolver_planilha(self.planilha)
        LOGGER.info("Planilha de clientes: %s", planilha)

        empresas = self._selecionar_empresas(utils.carregar_empresas(planilha))

        if not empresas:
            LOGGER.warning("Nenhuma empresa ativa encontrada na planilha.")
            return 0

        self.total = len(empresas)
        LOGGER.info("%d empresa(s) serão processadas.", self.total)

        utils.emitir_evento(
            "inicio",
            total=self.total,
            pasta=str(self.pasta_destino),
            planilha=str(planilha),
        )
        self._publicar_contadores()

        fatal: Optional[str] = None

        try:
            self._abrir_navegador()
            self._login_manual(empresas[0])

            for indice, empresa in enumerate(empresas, start=1):
                self._processar_com_tentativas(empresa, indice)

                if indice < self.total:
                    time.sleep(config.ESPERA_ENTRE_EMPRESAS)

        except KeyboardInterrupt:
            fatal = "Execução interrompida pelo usuário."
            LOGGER.warning(fatal)
        except Exception as exc:
            fatal = str(exc)
            LOGGER.critical("Erro fatal: %s\n%s", exc, traceback.format_exc())
        finally:
            self._encerrar_navegador()

        self._gerar_relatorios_finais()

        LOGGER.info(
            "Execução finalizada. Sucesso: %d | Não processadas: %d | "
            "Total: %d.",
            self.sucessos,
            len(self.registro),
            self.total,
        )

        utils.emitir_evento(
            "fim",
            total=self.total,
            sucesso=self.sucessos,
            falhas=len(self.registro),
            interrompido=bool(fatal),
        )

        if fatal:
            LOGGER.error("Motivo da interrupção: %s", fatal)
            return 1

        return 0 if not self.registro.itens else 2

    def _selecionar_empresas(self, empresas: list[Any]) -> list[Any]:
        if self.apenas_cnpj:
            empresas = [
                empresa
                for empresa in empresas
                if utils.so_digitos(empresa.cnpj) == self.apenas_cnpj
            ]

        if self.limite:
            empresas = empresas[: self.limite]

        return empresas

    def _processar_com_tentativas(self, empresa: Any, indice: int) -> None:
        nome = empresa.spreadsheet_name or "(sem nome na planilha)"

        LOGGER.info("[%d/%d] Empresa: %s", indice, self.total, nome)
        LOGGER.info("CNPJ: %s", utils.formatar_cnpj(empresa.cnpj))

        utils.emitir_evento(
            "empresa",
            indice=indice,
            total=self.total,
            cnpj=utils.formatar_cnpj(empresa.cnpj),
            nome=nome,
        )

        primeira = indice == 1
        tentativas = 1 if primeira else max(1, config.TENTATIVAS_POR_EMPRESA)
        etapa_falha = ""
        motivo_falha = ""

        for tentativa in range(1, tentativas + 1):
            try:
                if tentativa > 1:
                    LOGGER.warning(
                        "Nova tentativa (%d/%d) para %s.",
                        tentativa,
                        tentativas,
                        nome,
                    )
                    self._recuperar_tela()

                caminho = self._processar_empresa(empresa, primeira=primeira)

                self.sucessos += 1
                self.processadas += 1
                self._registrar_resultado(
                    empresa, "SUCESSO", "", "PDF salvo", str(caminho)
                )

                LOGGER.info("Empresa concluída")
                self._publicar_contadores()
                return

            except EmpresaNaoProcessada as exc:
                etapa_falha = exc.etapa or "Não informada"
                motivo_falha = exc.motivo

                if exc.definitivo:
                    LOGGER.info(
                        "Situação definitiva desta empresa: não há o que "
                        "tentar de novo."
                    )
                    self._limpar_abas()
                    break
            except Exception as exc:  # falha inesperada não derruba o lote
                etapa_falha = etapa_falha or "Erro inesperado"
                motivo_falha = str(exc)
                LOGGER.debug(traceback.format_exc())
            finally:
                self._limpar_abas()

        self.processadas += 1
        self.registro.registrar(
            cnpj=empresa.cnpj,
            nome=nome,
            etapa=etapa_falha,
            motivo=motivo_falha,
        )
        self._registrar_resultado(
            empresa, "NAO PROCESSADA", etapa_falha, motivo_falha, ""
        )
        self._publicar_contadores()

    def _registrar_resultado(
        self,
        empresa: Any,
        status: str,
        etapa: str,
        mensagem: str,
        arquivo: str,
    ) -> None:
        self.resultados.append(
            {
                "linha_excel": str(getattr(empresa, "excel_row", "")),
                "codigo": str(getattr(empresa, "code", "")),
                "cnpj": utils.formatar_cnpj(empresa.cnpj),
                "empresa": empresa.spreadsheet_name or "",
                "status": status,
                "etapa": etapa,
                "mensagem": mensagem,
                "arquivo": arquivo,
                "momento": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            }
        )

    def _publicar_contadores(self) -> None:
        utils.emitir_evento(
            "contadores",
            total=self.total,
            processadas=self.processadas,
            sucesso=self.sucessos,
            falhas=len(self.registro),
            pendentes=max(0, self.total - self.processadas),
        )

    # -------------------------------------------------------------------
    # NAVEGADOR E LOGIN
    # -------------------------------------------------------------------

    def _abrir_navegador(self) -> None:
        utils.registrar_etapa("Abrindo o navegador")

        self.driver, self.processo_chrome = utils.criar_driver(
            modo=self.modo_navegador,
            url_inicial=config.URL_LOGIN,
        )

        try:
            self.aba_principal = self.driver.current_window_handle
        except Exception:
            self.aba_principal = ""

        if self.modo_navegador != "anexar":
            utils.abrir_url(self.driver, config.URL_LOGIN)
            self.aba_principal = self.driver.current_window_handle

    def _encerrar_navegador(self) -> None:
        if self.driver is None:
            return

        if not self.fechar_navegador:
            LOGGER.info(
                "O navegador continuará aberto para conferência manual."
            )
            return

        try:
            self.driver.quit()
        except Exception:
            pass

    def _login_manual(self, primeira_empresa: Any) -> None:
        """
        Login e primeira representação: sempre manuais.

        Nenhuma senha, certificado ou CAPTCHA é automatizado.
        """
        if self.pular_login_manual:
            LOGGER.info(
                "Login manual ignorado por opção de linha de comando."
            )
            return

        utils.registrar_etapa("Aguardando o login manual do usuário")

        instrucoes = [
            "LOGIN MANUAL NA RECEITA FEDERAL",
            "",
            "1. Use a janela do Chrome aberta pela automação.",
            "2. Faça o login pelo Gov.br (certificado, senha ou app).",
            f"3. Abra: {config.URL_LOGIN}",
            "4. Represente MANUALMENTE a primeira empresa da lista:",
            f"     Empresa: {primeira_empresa.spreadsheet_name}",
            f"     CNPJ:    {utils.formatar_cnpj(primeira_empresa.cnpj)}",
            "     (informe o CNPJ, escolha o perfil Procurador e",
            "      clique em Representar)",
            "5. Volte aqui e pressione ENTER.",
            "",
            "A partir daí o robô assume o controle das demais empresas.",
        ]

        if not utils.esperar_enter(instrucoes):
            raise RuntimeError(
                "O login manual não foi confirmado dentro do tempo limite."
            )

        LOGGER.info("Login confirmado pelo usuário.")

    def _recuperar_tela(self) -> None:
        """Volta o navegador para um estado conhecido antes de tentar de novo."""
        try:
            self._limpar_abas()
            utils.abrir_url(self.driver, config.URL_PORTAL)
        except Exception as exc:
            LOGGER.warning("Não foi possível recuperar a tela: %s", exc)

    def _limpar_abas(self) -> None:
        if self.driver is None or not self.aba_principal:
            return

        utils.fechar_abas_extras(self.driver, self.aba_principal)

    # -------------------------------------------------------------------
    # FLUXO DE UMA EMPRESA
    # -------------------------------------------------------------------

    def _processar_empresa(self, empresa: Any, primeira: bool) -> Path:
        if primeira:
            LOGGER.info(
                "Primeira empresa: usando a representação feita manualmente."
            )
        else:
            self._representar(empresa)

        # O e-CAC é aberto primeiro: é o cabeçalho dele que mostra a
        # empresa representada e permite a conferência da sessão.
        self._abrir_ecac()
        self._conferir_empresa_representada(empresa)
        self._abrir_divida_ativa()
        self._abrir_regularize()
        self._consultar_divida_ativa()
        self._abrir_relatorio_consolidado()
        self._selecionar_filtros()
        self._gerar_relatorio()

        return self._imprimir_e_salvar(empresa)

    # ---------------------------------------------------------------
    # 1. REPRESENTAÇÃO
    # ---------------------------------------------------------------

    def _representar(self, empresa: Any) -> None:
        utils.registrar_etapa("Representando a empresa")

        try:
            self._representar_automaticamente(empresa)
            return
        except EmpresaNaoProcessada as exc:
            LOGGER.warning("Representação automática falhou: %s", exc.motivo)

            if not config.PERMITIR_REPRESENTACAO_MANUAL:
                raise

        self._representar_manualmente(empresa)

    def _representar_automaticamente(self, empresa: Any) -> None:
        driver = self.driver

        self._limpar_abas()
        utils.abrir_url(driver, config.URL_PORTAL)
        self._fechar_modais_opcionais()

        utils.safe_click(
            driver,
            config.LOC_ABRIR_PAINEL_REPRESENTACAO,
            timeout=config.TIMEOUT_OPCIONAL,
            descricao="painel de troca de perfil",
            etapa=ETAPA_REPRESENTACAO,
            obrigatorio=False,
        )

        campo = utils.wait_for_element(
            driver,
            config.LOC_CAMPO_CNPJ,
            timeout=config.TIMEOUT_PADRAO,
            condicao="clicavel",
            descricao="campo de CNPJ",
            etapa=ETAPA_REPRESENTACAO,
        )

        self._preencher_cnpj(campo, empresa.cnpj)
        self._selecionar_procurador()

        # A mesma espera obrigatória usada pelo ecac_download.py. Sem ela o
        # portal costuma pedir para aguardar antes de uma nova representação.
        utils.contagem_regressiva(
            config.ESPERA_ANTES_REPRESENTAR,
            "espera obrigatória do portal antes de Representar",
        )

        utils.safe_click(
            driver,
            config.LOC_BOTAO_REPRESENTAR,
            timeout=config.TIMEOUT_PADRAO,
            descricao="botão Representar",
            etapa=ETAPA_REPRESENTACAO,
        )

        LOGGER.info("Aguardando o portal concluir a representação.")
        utils.wait_for_page(driver)
        utils.esperar_condicao(
            lambda: self._cnpj_na_pagina(empresa.cnpj),
            config.TIMEOUT_APOS_REPRESENTAR,
        )

    def _preencher_cnpj(self, campo: Any, cnpj: str) -> None:
        from selenium.webdriver.common.keys import Keys

        digitos = utils.so_digitos(cnpj)

        try:
            campo.click()
        except Exception:
            pass

        try:
            campo.clear()
        except Exception:
            pass

        try:
            campo.send_keys(Keys.CONTROL, "a")
            campo.send_keys(Keys.DELETE)
        except Exception:
            pass

        campo.send_keys(digitos)
        LOGGER.info("CNPJ informado: %s", utils.formatar_cnpj(digitos))
        time.sleep(0.8)

    def _selecionar_procurador(self) -> None:
        """Seleciona o perfil Procurador, seja <select> ou lista customizada."""
        from selenium.webdriver.support.ui import Select

        driver = self.driver

        seletor = utils.check_element_exists(
            driver,
            config.LOC_SELECT_PERFIL,
            timeout=2,
            condicao="presente",
        )

        if seletor is not None:
            try:
                lista = Select(seletor)

                for opcao in lista.options:
                    if "procurador" in utils.normalizar_texto(opcao.text):
                        lista.select_by_visible_text(opcao.text)
                        LOGGER.info("Perfil selecionado: %s", opcao.text.strip())
                        time.sleep(0.5)
                        return
            except Exception as exc:
                LOGGER.warning("Lista de perfis padrão não funcionou: %s", exc)

        utils.safe_click(
            driver,
            config.LOC_ABRIR_LISTA_PERFIL,
            timeout=config.TIMEOUT_OPCIONAL,
            descricao="campo de perfil",
            etapa=ETAPA_REPRESENTACAO,
            obrigatorio=False,
        )

        utils.safe_click(
            driver,
            config.LOC_OPCAO_PROCURADOR,
            timeout=config.TIMEOUT_OPCIONAL,
            descricao="opção Procurador",
            etapa=ETAPA_REPRESENTACAO,
        )
        time.sleep(0.5)

    def _representar_manualmente(self, empresa: Any) -> None:
        """Plano B: pede a representação ao usuário sem abortar o lote."""
        utils.registrar_etapa("Aguardando representação manual da empresa")

        instrucoes = [
            "REPRESENTAÇÃO MANUAL NECESSÁRIA",
            "",
            "A automação não conseguiu representar esta empresa sozinha.",
            "",
            f"Empresa: {empresa.spreadsheet_name}",
            f"CNPJ:    {utils.formatar_cnpj(empresa.cnpj)}",
            "",
            "Represente a empresa no portal e pressione ENTER para continuar.",
            "Se preferir pular esta empresa, aguarde o tempo limite.",
        ]

        if not utils.esperar_enter(instrucoes):
            raise EmpresaNaoProcessada(
                "Representação automática falhou e não houve confirmação "
                "manual dentro do tempo limite.",
                etapa=ETAPA_REPRESENTACAO,
            )

        utils.wait_for_page(self.driver)

    def _cnpj_na_pagina(self, cnpj: str) -> bool:
        conteudo = utils.texto_da_pagina(self.driver)

        if not conteudo:
            return False

        digitos = utils.so_digitos(cnpj)
        formatado = utils.normalizar_texto(utils.formatar_cnpj(digitos))

        return digitos in conteudo.replace(".", "").replace(
            "/", ""
        ).replace("-", "") or formatado in conteudo

    def _conferir_empresa_representada(self, empresa: Any) -> None:
        """Confirma que a sessão corresponde à empresa que será processada."""
        utils.registrar_etapa("Conferindo a empresa representada")

        confere = utils.esperar_condicao(
            lambda: self._cnpj_na_pagina(empresa.cnpj),
            config.TIMEOUT_OPCIONAL,
        )

        if confere:
            LOGGER.info("Empresa representada confirmada na tela.")
            return

        mensagem = (
            "O CNPJ da empresa não foi localizado na tela após a "
            "representação."
        )

        if config.EXIGIR_CONFERENCIA_DO_CNPJ:
            raise EmpresaNaoProcessada(mensagem, etapa=ETAPA_REPRESENTACAO)

        LOGGER.warning("%s Seguindo mesmo assim.", mensagem)

    # ---------------------------------------------------------------
    # 2. e-CAC E DÍVIDA ATIVA
    # ---------------------------------------------------------------

    def _fechar_modais_opcionais(
        self,
        rodadas: int = 2,
        timeout: float = config.TIMEOUT_MODAL_RAPIDO,
    ) -> None:
        """
        Fecha os modais 'Ciente' e 'Fechar', que podem ou não aparecer.

        A verificação é rápida por padrão: em telas onde nenhum modal
        costuma surgir, esperar seria só tempo perdido em cada empresa.
        Nos pontos em que o modal é esperado, o chamador aumenta o tempo.
        """
        for _ in range(max(1, rodadas)):
            achou_ciente = utils.fechar_modal_opcional(
                self.driver,
                config.LOC_BOTAO_CIENTE,
                "Ciente",
                timeout=timeout,
            )
            achou_fechar = utils.fechar_modal_opcional(
                self.driver,
                config.LOC_BOTAO_FECHAR,
                "Fechar",
                timeout=min(timeout, config.TIMEOUT_MODAL_RAPIDO),
            )

            if not achou_ciente and not achou_fechar:
                return

    def _abrir_ecac(self) -> None:
        utils.registrar_etapa("Acessando o e-CAC")

        driver = self.driver

        if utils.check_element_exists(
            driver, config.LOC_DIVIDA_ATIVA, timeout=2, condicao="presente"
        ):
            LOGGER.info("O menu de serviços do e-CAC já está aberto.")
            return

        for url in (config.URL_ECAC, *config.URLS_ECAC_ALTERNATIVAS):
            utils.abrir_url(driver, url)
            self._fechar_modais_opcionais(rodadas=1)

            if utils.check_element_exists(
                driver,
                config.LOC_DIVIDA_ATIVA,
                timeout=config.TIMEOUT_OPCIONAL,
                condicao="presente",
            ):
                LOGGER.info("e-CAC aberto como procurador da empresa.")
                return

        raise EmpresaNaoProcessada(
            "Não foi possível abrir o menu de serviços do e-CAC.",
            etapa=ETAPA_ECAC,
        )

    def _abrir_divida_ativa(self) -> None:
        utils.registrar_etapa("Acessando Dívida Ativa")

        utils.safe_click(
            self.driver,
            config.LOC_DIVIDA_ATIVA,
            timeout=config.TIMEOUT_PADRAO,
            descricao="Dívida Ativa da União",
            etapa=ETAPA_DIVIDA_ATIVA,
        )

        utils.esperar_carregamento_sumir(self.driver, timeout=10)

        encontrado = utils.check_element_exists(
            self.driver,
            config.LOC_PGFN_REGULARIZE,
            timeout=config.TIMEOUT_OPCIONAL,
            condicao="presente",
        )

        if encontrado is None:
            raise EmpresaNaoProcessada(
                "O bloco de serviços da Dívida Ativa da União não abriu.",
                etapa=ETAPA_DIVIDA_ATIVA,
            )

    def _abrir_regularize(self) -> None:
        utils.registrar_etapa("Acessando Regularize")

        driver = self.driver
        abas_antes = list(driver.window_handles)

        utils.safe_click(
            driver,
            config.LOC_PGFN_REGULARIZE,
            timeout=config.TIMEOUT_PADRAO,
            descricao="PGFN - Todos os serviços do Regularize",
            etapa=ETAPA_REGULARIZE,
        )

        # O link usa target="_blank": normalmente abre uma nova aba, mas
        # também pode reaproveitar a atual.
        nova = utils.switch_to_new_tab(
            driver, abas_antes, timeout=config.TIMEOUT_NOVA_ABA
        )

        if nova is None:
            LOGGER.info(
                "Nenhuma aba nova foi aberta. Continuando na aba atual."
            )
            utils.wait_for_page(driver)

        self._fechar_modais_opcionais(timeout=config.TIMEOUT_OPCIONAL)
        self._verificar_bloqueios(ETAPA_REGULARIZE)

    def _verificar_bloqueios(self, etapa: str) -> None:
        """
        Detecta empresas sem cadastro no Regularize e páginas indisponíveis.

        Nesses casos a empresa é registrada e o lote continua.
        """
        driver = self.driver

        try:
            url = (driver.current_url or "").lower()
        except Exception:
            url = ""

        if config.FRAGMENTO_URL_CADASTRO in url:
            raise EmpresaNaoProcessada(
                "Cadastro necessário: o Regularize redirecionou para a "
                "tela de cadastro.",
                etapa=etapa,
                definitivo=True,
            )

        encontrado = utils.pagina_contem(
            driver, config.TEXTOS_CADASTRO_NECESSARIO
        )

        if encontrado:
            raise EmpresaNaoProcessada(
                f"Cadastro necessário no Regularize ('{encontrado}').",
                etapa=etapa,
                definitivo=True,
            )

        indisponivel = utils.pagina_contem(driver, config.TEXTOS_INDISPONIVEL)

        if indisponivel:
            raise EmpresaNaoProcessada(
                f"Página do Regularize indisponível ('{indisponivel}').",
                etapa=etapa,
            )

    # ---------------------------------------------------------------
    # 3. CONSULTA E RELATÓRIO
    # ---------------------------------------------------------------

    def _consultar_divida_ativa(self) -> None:
        utils.registrar_etapa("Consultando Dívida Ativa")

        driver = self.driver

        utils.safe_click(
            driver,
            config.LOC_CONSULTAR_DIVIDA_ATIVA,
            timeout=config.TIMEOUT_PADRAO,
            descricao="Consultar Dívida Ativa",
            etapa=ETAPA_CONSULTA,
        )

        # Espera inteligente: segue assim que o botão do relatório aparecer,
        # respeitando o teto de 25 segundos previsto para esta tela.
        pronto = utils.esperar_condicao(
            lambda: utils.check_element_exists(
                driver,
                config.LOC_RELATORIO_CONSOLIDADO,
                timeout=0,
                condicao="visivel",
            ),
            config.TIMEOUT_CONSULTA_DIVIDA,
        )

        self._fechar_modais_opcionais(rodadas=1)

        if pronto is None:
            self._verificar_bloqueios(ETAPA_CONSULTA)

            sem_debitos = utils.pagina_contem(
                driver, config.TEXTOS_SEM_DEBITOS
            )

            if sem_debitos:
                raise EmpresaNaoProcessada(
                    "A consulta não retornou débitos inscritos em dívida "
                    f"ativa ('{sem_debitos}').",
                    etapa=ETAPA_CONSULTA,
                    definitivo=True,
                )

            raise EmpresaNaoProcessada(
                "A tela de consulta da dívida ativa não carregou dentro de "
                f"{config.TIMEOUT_CONSULTA_DIVIDA} segundos.",
                etapa=ETAPA_CONSULTA,
            )

    def _abrir_relatorio_consolidado(self) -> None:
        utils.registrar_etapa("Abrindo Relatório Consolidado")

        utils.safe_click(
            self.driver,
            config.LOC_RELATORIO_CONSOLIDADO,
            timeout=config.TIMEOUT_PADRAO,
            descricao="Relatório Consolidado",
            etapa=ETAPA_RELATORIO,
        )

        utils.wait_for_page(self.driver, timeout=config.TIMEOUT_PADRAO)
        self._fechar_modais_opcionais(rodadas=1)

        encontrado = utils.check_element_exists(
            self.driver,
            config.LOC_TODAS_NATUREZAS,
            timeout=config.TIMEOUT_PADRAO,
            condicao="presente",
        )

        if encontrado is None:
            raise EmpresaNaoProcessada(
                "A tela do Relatório Consolidado não apresentou os filtros "
                "de natureza e situação.",
                etapa=ETAPA_RELATORIO,
            )

    def _selecionar_filtros(self) -> None:
        utils.marcar_checkbox(
            self.driver,
            config.LOC_TODAS_NATUREZAS,
            "Natureza: Todos",
            etapa=ETAPA_FILTROS,
        )
        LOGGER.info("Natureza: Todos selecionado")

        utils.marcar_checkbox(
            self.driver,
            config.LOC_TODAS_SITUACOES,
            "Situação: Todos",
            etapa=ETAPA_FILTROS,
        )
        LOGGER.info("Situação: Todos selecionado")

    def _gerar_relatorio(self) -> None:
        utils.registrar_etapa("Gerando relatório")

        driver = self.driver

        utils.safe_click(
            driver,
            config.LOC_GERAR_RELATORIO,
            timeout=config.TIMEOUT_PADRAO,
            descricao="Gerar Relatório",
            etapa=ETAPA_GERACAO,
        )

        self._fechar_modais_opcionais(rodadas=1)

        # Espera inteligente com teto de 20 segundos: o botão Imprimir só
        # existe depois que o relatório termina de ser montado.
        pronto = utils.esperar_condicao(
            lambda: utils.check_element_exists(
                driver,
                config.LOC_IMPRIMIR,
                timeout=0,
                condicao="clicavel",
            ),
            config.TIMEOUT_GERACAO_RELATORIO,
        )

        if pronto is None:
            self._verificar_bloqueios(ETAPA_GERACAO)

            raise EmpresaNaoProcessada(
                "O relatório não terminou de ser gerado em "
                f"{config.TIMEOUT_GERACAO_RELATORIO} segundos "
                "(o botão Imprimir não apareceu).",
                etapa=ETAPA_GERACAO,
            )

        utils.esperar_carregamento_sumir(driver, timeout=10)
        LOGGER.info("Relatório gerado")

    # ---------------------------------------------------------------
    # 4. IMPRESSÃO E PDF
    # ---------------------------------------------------------------

    def _imprimir_e_salvar(self, empresa: Any) -> Path:
        utils.registrar_etapa("Imprimindo")

        driver = self.driver

        # O botão Imprimir chama window.print(). A janela de impressão do
        # Windows é neutralizada para não travar o robô; o PDF é gerado
        # logo em seguida pelo próprio navegador, já com o nome final.
        utils.neutralizar_dialogo_de_impressao(driver)

        clicado = utils.safe_click(
            driver,
            config.LOC_IMPRIMIR,
            timeout=config.TIMEOUT_PADRAO,
            descricao="Imprimir",
            etapa=ETAPA_IMPRESSAO,
            obrigatorio=False,
        )

        if clicado is None:
            LOGGER.warning(
                "O botão Imprimir não respondeu. O relatório será salvo "
                "diretamente da tela."
            )
        else:
            time.sleep(config.ESPERA_APOS_IMPRIMIR)

        self._fechar_modais_opcionais(rodadas=1)

        nome = empresa.spreadsheet_name or self._nome_na_tela(empresa)
        destino = utils.montar_nome_do_pdf(
            self.pasta_destino, empresa.cnpj, nome
        )

        caminho = utils.save_pdf(driver, destino, etapa=ETAPA_PDF)

        if not utils.pdf_valido(caminho):
            raise EmpresaNaoProcessada(
                f"O PDF salvo não passou na validação: {caminho}",
                etapa=ETAPA_PDF,
            )

        LOGGER.info("PDF salvo")
        utils.emitir_evento("pdf", arquivo=str(caminho))
        return caminho

    def _nome_na_tela(self, empresa: Any) -> str:
        """Nome de reserva quando a planilha não traz a razão social."""
        try:
            titulo = self.driver.title or ""
        except Exception:
            titulo = ""

        return titulo.strip() or f"EMPRESA {utils.so_digitos(empresa.cnpj)}"

    # -------------------------------------------------------------------
    # RELATÓRIOS FINAIS
    # -------------------------------------------------------------------

    def _gerar_relatorios_finais(self) -> None:
        config.PASTA_RUNTIME.mkdir(parents=True, exist_ok=True)

        csv_resultados = (
            config.PASTA_RUNTIME / f"resultado_pgfn_{self.carimbo}.csv"
        )
        self._salvar_csv_resultados(csv_resultados)
        LOGGER.info("Resultado completo em: %s", csv_resultados)

        if not self.registro.itens:
            LOGGER.info("Todas as empresas foram processadas com sucesso.")
            return

        nome_base = f"{config.PREFIXO_RELATORIO_FALHAS}_{self.carimbo}"

        resumo = [
            f"Total de empresas na base: {self.total}",
            f"Processadas com sucesso:   {self.sucessos}",
            f"Não processadas:           {len(self.registro)}",
        ]

        pdf_runtime = self.registro.gerar_pdf(
            config.PASTA_RUNTIME / f"{nome_base}.pdf", resumo=resumo
        )
        csv_runtime = self.registro.gerar_csv(
            config.PASTA_RUNTIME / f"{nome_base}.csv"
        )

        LOGGER.info("Relatório de empresas não processadas: %s", pdf_runtime)
        LOGGER.info("Planilha de empresas não processadas: %s", csv_runtime)

        caminhos = [str(pdf_runtime), str(csv_runtime)]

        if config.SALVAR_RELATORIO_FALHAS_NA_PASTA_DESTINO and pdf_runtime:
            copia = self.pasta_destino / f"{nome_base}.pdf"

            try:
                copia.write_bytes(Path(pdf_runtime).read_bytes())
                LOGGER.info("Cópia do relatório de falhas em: %s", copia)
                caminhos.append(str(copia))
            except OSError as exc:
                LOGGER.warning(
                    "Não foi possível copiar o relatório de falhas: %s", exc
                )

        utils.emitir_evento("relatorio_falhas", arquivos=caminhos)

    def _salvar_csv_resultados(self, destino: Path) -> None:
        import csv as _csv

        if not self.resultados:
            return

        with destino.open("w", newline="", encoding="utf-8-sig") as arquivo:
            escritor = _csv.DictWriter(
                arquivo,
                fieldnames=list(self.resultados[0].keys()),
                delimiter=";",
            )
            escritor.writeheader()
            escritor.writerows(self.resultados)


# =============================================================================
# LINHA DE COMANDO
# =============================================================================

def analisar_argumentos(argumentos: Optional[list[str]] = None) -> Any:
    analisador = argparse.ArgumentParser(
        description=(
            "Emite o Relatório Consolidado da Dívida Ativa da União "
            "(PGFN / Regularize) para todas as empresas da base."
        )
    )

    analisador.add_argument(
        "--pasta-destino",
        default=str(config.PASTA_DESTINO_PADRAO),
        help="Pasta onde os PDFs serão salvos.",
    )
    analisador.add_argument(
        "--planilha",
        default=config.PLANILHA_PADRAO,
        help="Planilha de clientes. Vazio usa a mesma do ecac_download.py.",
    )
    analisador.add_argument(
        "--limite",
        type=int,
        default=0,
        help="Processa apenas as N primeiras empresas (útil em testes).",
    )
    analisador.add_argument(
        "--cnpj",
        default="",
        help="Processa somente o CNPJ informado (útil em testes).",
    )
    analisador.add_argument(
        "--modo-navegador",
        default=config.MODO_NAVEGADOR,
        choices=["anexar", "selenium"],
        help="Como o Chrome será controlado.",
    )
    analisador.add_argument(
        "--pular-login",
        action="store_true",
        help="Não aguarda o login manual (use apenas em testes).",
    )
    analisador.add_argument(
        "--fechar-navegador",
        action="store_true",
        help="Fecha o Chrome ao final da execução.",
    )
    analisador.add_argument(
        "--testar-relatorio-falhas",
        action="store_true",
        help=(
            "Não abre o navegador: apenas gera um PDF de exemplo do "
            "relatório de empresas não processadas."
        ),
    )

    return analisador.parse_args(argumentos)


def gerar_relatorio_de_exemplo() -> int:
    """Teste isolado do relatório final, sem navegador e sem portal."""
    carimbo = datetime.now().strftime("%Y%m%d_%H%M%S")
    utils.configurar_log(config.PASTA_LOGS, carimbo)

    registro = utils.RegistroFalhas()
    registro.registrar(
        "00000000000100", "EMPRESA A", ETAPA_REGULARIZE, "Cadastro necessário"
    )
    registro.registrar(
        "11111111000111",
        "EMPRESA B",
        ETAPA_CONSULTA,
        "Página do Regularize indisponível",
    )
    registro.registrar(
        "22222222000122",
        "EMPRESA C",
        ETAPA_RELATORIO,
        "Elemento não encontrado",
    )

    destino = config.PASTA_RUNTIME / f"exemplo_nao_processadas_{carimbo}.pdf"
    caminho = registro.gerar_pdf(destino, resumo=["Execução de teste."])
    registro.gerar_csv(destino.with_suffix(".csv"))

    LOGGER.info("Relatório de exemplo gerado em: %s", caminho)
    return 0


def main(argumentos: Optional[list[str]] = None) -> int:
    opcoes = analisar_argumentos(argumentos)

    if opcoes.testar_relatorio_falhas:
        return gerar_relatorio_de_exemplo()

    automacao = AutomacaoPGFN(
        pasta_destino=Path(opcoes.pasta_destino),
        planilha=opcoes.planilha,
        limite=opcoes.limite,
        apenas_cnpj=opcoes.cnpj,
        pular_login_manual=opcoes.pular_login,
        fechar_navegador=opcoes.fechar_navegador,
        modo_navegador=opcoes.modo_navegador,
    )

    return automacao.executar()


if __name__ == "__main__":
    raise SystemExit(main())
