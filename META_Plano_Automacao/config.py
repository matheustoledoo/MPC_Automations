#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Configurações da automação PGFN / Regularize.

Este arquivo concentra tudo o que costuma mudar com o tempo:

- URLs do portal;
- tempos de espera e timeouts;
- pastas padrão;
- configurações do navegador;
- localizadores (seletores) de cada etapa do fluxo.

Nada aqui importa Selenium. Os localizadores são descritos como tuplas
simples e traduzidos em `pgfn_utils.py`. Assim a configuração pode ser
lida, revisada e ajustada por quem não conhece a biblioteca.

FORMATO DOS LOCALIZADORES
-------------------------
Cada etapa recebe uma LISTA de localizadores, testados na ordem em que
aparecem. O primeiro que encontrar o elemento vence.

    ("id",     "natTodosCheck")
    ("css",    "input[formcontrolname='todas']")
    ("xpath",  "//*[@id='btn255']/a")
    ("texto",  "Consultar Dívida Ativa")
    ("texto",  "Imprimir", "button,a,input")

O tipo "texto" procura o texto visível ignorando maiúsculas, minúsculas e
acentos. É o mais estável quando o HTML muda. Os Full XPaths do documento
de especificação ficam sempre por último, apenas como rede de segurança.
"""

from __future__ import annotations

import os
from pathlib import Path


# =============================================================================
# PASTAS
# =============================================================================

LOCAL_APP_DATA = Path(os.environ.get("LOCALAPPDATA", str(Path.home())))

# Arquivos técnicos (logs, CSV de falhas, perfil do navegador).
# A pasta escolhida pelo usuário recebe somente os PDFs dos relatórios.
PASTA_RUNTIME = LOCAL_APP_DATA / "PGFNRegularize" / "runtime"
PASTA_LOGS = PASTA_RUNTIME / "logs"

# Perfil próprio do Chrome usado pela automação.
# O login manual do Gov.br fica salvo aqui entre execuções.
PASTA_PERFIL_CHROME = LOCAL_APP_DATA / "PGFNRegularize" / "chrome-profile"

# Pasta de downloads do navegador controlado (usada apenas como plano B
# quando a impressão direta em PDF não estiver disponível).
PASTA_DOWNLOAD_TEMPORARIA = PASTA_RUNTIME / "downloads"

PASTA_DESTINO_PADRAO = Path(
    os.environ.get(
        "PGFN_OUTPUT_DIR",
        str(Path.home() / "Desktop" / "Relatorios_PGFN"),
    )
)

# Planilha de clientes. Quando vazio, a automação usa a mesma resolução do
# `ecac_download.py` (variável ECAC_EXCEL_PATH e caminhos padrão).
PLANILHA_PADRAO = os.environ.get("PGFN_EXCEL_PATH", "").strip()


# =============================================================================
# URLS
# =============================================================================

URL_LOGIN = (
    "https://servicos.receitafederal.gov.br/servico/"
    "pendencias/#/analise-pendencias"
)

# Página usada para representar a próxima empresa.
URL_PORTAL = URL_LOGIN

# Página inicial do e-CAC, onde ficam os botões de serviço.
# A automação tenta os endereços na ordem e usa o primeiro que exibir o
# menu de serviços (botão "Dívida Ativa da União").
URL_ECAC = "https://cav.receita.fazenda.gov.br/ecac/"
URLS_ECAC_ALTERNATIVAS = (
    "https://cav.receita.fazenda.gov.br/autenticacao/login",
    "https://cav.receita.fazenda.gov.br/ecac/Aplicacao.aspx",
)

# Trechos que identificam cada área do fluxo pela URL.
FRAGMENTO_URL_ECAC = "receita.fazenda.gov.br"
FRAGMENTO_URL_REGULARIZE = "consultadebitos"
FRAGMENTO_URL_CADASTRO = "/cadastro"


# =============================================================================
# NAVEGADOR
# =============================================================================

# "anexar"  -> a automação abre o Chrome com depuração remota e o Selenium
#              se conecta à janela aberta. O navegador continua vivo mesmo
#              que a automação seja encerrada.
# "selenium"-> o próprio Selenium abre e controla o Chrome.
MODO_NAVEGADOR = os.environ.get("PGFN_MODO_NAVEGADOR", "anexar").strip()

PORTA_DEPURACAO = int(os.environ.get("PGFN_CHROME_PORT", "9333"))

# Caminho manual do Chrome. Vazio = detecção automática pelo ecac_download.
CAMINHO_CHROME = os.environ.get("PGFN_CHROME_PATH", "").strip()

ARGUMENTOS_CHROME = [
    "--no-first-run",
    "--no-default-browser-check",
    "--start-maximized",
    "--disable-features=Translate",
]

TIMEOUT_ABERTURA_CHROME = 60


# =============================================================================
# TEMPOS DE ESPERA
# =============================================================================

# Espera padrão por elementos.
TIMEOUT_PADRAO = 30

# Espera curta usada para modais e elementos opcionais.
TIMEOUT_OPCIONAL = 6

# Verificação instantânea de modais em telas onde eles raramente aparecem.
# Multiplicado por várias telas e centenas de empresas, cada segundo aqui
# custa muito tempo de execução.
TIMEOUT_MODAL_RAPIDO = 1.5

# Intervalo entre as verificações das esperas inteligentes.
INTERVALO_POLLING = 0.4

# Carregamento de página (document.readyState).
TIMEOUT_CARREGAMENTO_PAGINA = 45

# Abertura da nova aba do Regularize (link com target="_blank").
TIMEOUT_NOVA_ABA = 30

# Documento: aguardar aproximadamente 25 segundos após "Consultar Dívida Ativa".
TIMEOUT_CONSULTA_DIVIDA = 25

# Documento: aguardar aproximadamente 20 segundos após "Gerar Relatório".
TIMEOUT_GERACAO_RELATORIO = 20

# Espera de segurança depois de clicar em "Imprimir".
ESPERA_APOS_IMPRIMIR = 3.0

# Espera obrigatória do portal antes de clicar em "Representar".
# O mesmo valor usado pelo ecac_download.py, que evita a mensagem
# pedindo para aguardar antes de uma nova representação.
ESPERA_ANTES_REPRESENTAR = 30

# Tempo máximo aguardando o portal concluir a representação.
TIMEOUT_APOS_REPRESENTAR = 60

# Pausa entre empresas.
ESPERA_ENTRE_EMPRESAS = 4.0

# Tempo máximo aguardando uma ação manual do usuário (login, primeira
# representação ou recuperação de uma empresa).
TIMEOUT_ACAO_MANUAL = 900


# =============================================================================
# COMPORTAMENTO
# =============================================================================

# Quantas vezes uma mesma empresa é tentada antes de ser registrada
# como não processada.
TENTATIVAS_POR_EMPRESA = 2

# Quando a representação automática falhar, pedir ajuda ao usuário em vez
# de descartar a empresa imediatamente.
PERMITIR_REPRESENTACAO_MANUAL = True

# Interromper a empresa quando o CNPJ representado não confere com o
# CNPJ esperado. Com False, apenas registra um aviso e continua.
EXIGIR_CONFERENCIA_DO_CNPJ = True

# Copiar o relatório de empresas não processadas para a pasta escolhida
# pelo usuário, além de salvá-lo na pasta de runtime.
SALVAR_RELATORIO_FALHAS_NA_PASTA_DESTINO = True

PREFIXO_RELATORIO_FALHAS = "_EMPRESAS_NAO_PROCESSADAS"

# Tamanho mínimo aceitável para considerar o PDF válido.
TAMANHO_MINIMO_PDF_BYTES = 1024


# =============================================================================
# IMPRESSÃO EM PDF
# =============================================================================

# Parâmetros de Page.printToPDF (Chrome DevTools Protocol). As medidas
# são em polegadas; o padrão equivale a uma folha A4 retrato.
OPCOES_IMPRESSAO = {
    "printBackground": True,
    "preferCSSPageSize": False,
    "paperWidth": 8.27,
    "paperHeight": 11.69,
    "marginTop": 0.4,
    "marginBottom": 0.4,
    "marginLeft": 0.4,
    "marginRight": 0.4,
    "scale": 1.0,
}


# =============================================================================
# LOCALIZADORES — e-CAC
# =============================================================================

# Painel/menu que abre a troca de perfil (representação).
LOC_ABRIR_PAINEL_REPRESENTACAO = [
    ("css", "#perfil-acesso, #btnAlterarPerfil, #alterar-perfil"),
    ("texto", "Alterar perfil de acesso", "a,button,span,div"),
    ("texto", "Alterar perfil", "a,button,span,div"),
    ("texto", "Trocar perfil", "a,button,span,div"),
    ("texto", "Perfil de acesso", "a,button,span,div"),
]

LOC_CAMPO_CNPJ = [
    (
        "css",
        "input#cnpj, input#txtCNPJ, input[name*='cnpj' i], "
        "input[id*='cnpj' i], input[formcontrolname*='cnpj' i], "
        "input[placeholder*='CNPJ' i]",
    ),
    (
        "css",
        "input[name*='documento' i], input[id*='documento' i], "
        "input[placeholder*='documento' i]",
    ),
]

LOC_SELECT_PERFIL = [
    (
        "css",
        "select#perfil, select[name*='perfil' i], select[id*='perfil' i], "
        "select[formcontrolname*='perfil' i]",
    ),
    ("css", "select"),
]

LOC_ABRIR_LISTA_PERFIL = [
    ("css", "#perfil, [id*='perfil' i][role='combobox']"),
    ("texto", "Selecione o perfil", "div,span,button,a,label"),
    ("texto", "Procurador", "div,span,button,a,label"),
]

LOC_OPCAO_PROCURADOR = [
    ("xpath", "//option[contains(normalize-space(.), 'Procurador')]"),
    ("texto", "Procurador de Pessoa Jurídica", "li,span,div,a,label,option"),
    ("texto", "Procurador", "li,span,div,a,label,option"),
]

LOC_BOTAO_REPRESENTAR = [
    ("xpath", "//button[contains(normalize-space(.), 'Representar')]"),
    ("xpath", "//input[@type='submit' and contains(@value, 'Representar')]"),
    ("texto", "Representar", "button,a,input"),
]

# "Dívida Ativa da União" no menu de serviços do e-CAC.
LOC_DIVIDA_ATIVA = [
    ("xpath", '//*[@id="btn255"]/a'),
    ("texto", "Dívida Ativa da União", "a,button,span"),
    ("xpath", "/html/body/div[3]/div[2]/ul/li[5]/a"),
]

# "PGFN - Todos os serviços do Regularize" (abre em nova aba).
LOC_PGFN_REGULARIZE = [
    ("xpath", '//*[@id="containerServicos255"]/div[3]/ul/li/a'),
    ("xpath", "//a[contains(@href, 'PGFN/consultaDebitos')]"),
    ("texto", "Todos os serviços do Regularize", "a,span,button"),
    ("texto", "Regularize", "a,span,button"),
    (
        "xpath",
        "/html/body/div[3]/div[2]/div[3]/div[5]/div[2]/div[3]/ul/li/a",
    ),
]


# =============================================================================
# LOCALIZADORES — REGULARIZE
# =============================================================================

LOC_BOTAO_CIENTE = [
    ("xpath", "//modal-container//button[contains(normalize-space(.), 'Ciente')]"),
    ("texto", "Ciente", "button,a"),
    ("xpath", "/html/body/modal-container/div[2]/div/div[2]/div/button"),
]

LOC_BOTAO_FECHAR = [
    ("xpath", "//modal-container//button[contains(normalize-space(.), 'Fechar')]"),
    ("texto", "Fechar", "button,a"),
    ("xpath", "/html/body/modal-container/div[2]/div/div[2]/div/button[2]"),
]

LOC_CONSULTAR_DIVIDA_ATIVA = [
    (
        "xpath",
        "//app-servico//a[.//*[contains(normalize-space(.), "
        "'Consultar Dívida Ativa')]]",
    ),
    (
        "xpath",
        "//span[contains(normalize-space(.), 'Consultar Dívida Ativa')]"
        "/ancestor::a[1]",
    ),
    ("texto", "Consultar Dívida Ativa", "a,button,span,div"),
    (
        "xpath",
        "/html/body/app-root/div/div[2]/app-pagina-principal/main/div/"
        "app-servicos/section[2]/div/div[1]/app-servico/a/div/div/div",
    ),
]

LOC_RELATORIO_CONSOLIDADO = [
    ("xpath", "//button[contains(@routerlink, 'relatorio')]"),
    ("xpath", "//button[contains(normalize-space(.), 'Relatório Consolidado')]"),
    ("texto", "Relatório Consolidado", "button,a"),
    (
        "xpath",
        "/html/body/app-root/div/div[2]/app-consulta-divida/main/div/div/"
        "fieldset[1]/button",
    ),
]

LOC_TODAS_NATUREZAS = [
    ("id", "natTodosCheck"),
    ("css", "input#natTodosCheck[formcontrolname='todas']"),
    ("xpath", '//*[@id="natTodosCheck"]'),
]

LOC_TODAS_SITUACOES = [
    ("id", "sitTodosCheck"),
    ("css", "input#sitTodosCheck[formcontrolname='todas']"),
    ("xpath", '//*[@id="sitTodosCheck"]'),
]

LOC_GERAR_RELATORIO = [
    ("xpath", "//button[contains(normalize-space(.), 'Gerar Relatório')]"),
    ("texto", "Gerar Relatório", "button,a,input"),
    (
        "xpath",
        "/html/body/app-root/div/div[2]/app-relatorio-consolidado/div/table/"
        "tbody/tr/td/div/fieldset/form/button[1]",
    ),
]

LOC_IMPRIMIR = [
    ("xpath", "//button[contains(normalize-space(.), 'Imprimir')]"),
    ("texto", "Imprimir", "button,a,input"),
    (
        "xpath",
        "/html/body/app-root/div/div[2]/app-relatorio-consolidado/div/table/"
        "tbody/tr/td/div/div[3]/button[2]",
    ),
]

# Área do relatório já renderizado. Serve para confirmar que a geração
# realmente aconteceu antes de imprimir.
LOC_CONTEUDO_RELATORIO = [
    ("css", "app-relatorio-consolidado table tbody tr td div div"),
    ("css", "app-relatorio-consolidado"),
]

# Elementos de carregamento que precisam sumir antes do próximo clique.
SELETORES_CARREGANDO = (
    "ngx-spinner .la-ball-spin-clockwise, .loading-overlay, "
    ".spinner-border, #loading, .cdk-overlay-backdrop-showing"
)


# =============================================================================
# DETECÇÃO DE EMPRESA SEM CADASTRO NO REGULARIZE
# =============================================================================

# Basta um destes trechos aparecer na página para a empresa ser
# registrada como "cadastro necessário" e o fluxo seguir para a próxima.
TEXTOS_CADASTRO_NECESSARIO = (
    "realize seu cadastro",
    "efetue seu cadastro",
    "criar cadastro",
    "cadastre-se",
    "faca seu cadastro",
    "nao possui cadastro",
    "cadastro nao encontrado",
    "primeiro acesso ao regularize",
    "complete seu cadastro",
    "atualize seu cadastro",
)

TEXTOS_INDISPONIVEL = (
    "servico indisponivel",
    "sistema indisponivel",
    "em manutencao",
    "tente novamente mais tarde",
    "erro interno",
    "503 service unavailable",
    "502 bad gateway",
)

TEXTOS_SEM_DEBITOS = (
    "nao existem debitos",
    "nao ha debitos",
    "nenhum debito encontrado",
    "nao possui debitos inscritos",
)


# =============================================================================
# NOMES DE ARQUIVO
# =============================================================================

# Modelo do nome final: "12345678000100 - EMPRESA EXEMPLO LTDA.pdf"
MODELO_NOME_ARQUIVO = "{cnpj} - {nome}.pdf"

# Usar o CNPJ apenas com números (True) ou formatado (False).
NOME_ARQUIVO_COM_CNPJ_SEM_PONTUACAO = True
