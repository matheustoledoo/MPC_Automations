META PLANO - CENTRAL DE AUTOMACOES
=================================

O projeto possui tres automacoes:

- as duas primeiras ficam integradas na tela do INICIAR.bat;
- a terceira (PGFN / Regularize) tem tela propria, no INICIAR_PGFN.bat.

1) BAIXAR RELATORIOS DA RECEITA FEDERAL
   - Usa a planilha de clientes.
   - Executa o fluxo ja desenvolvido no portal da Receita.
   - Mantem o Chrome normal, sem Selenium.
   - Salva na pasta escolhida SOMENTE os PDFs finais renomeados.
   - Mantem 10 posicoes de teste do botao Baixar Relatorio.
   - Mantem espera de 30 s antes de Representar e 45 s depois de Representar.

2) IDENTIFICAR EMPRESAS INAPTAS / INAPTOS
   - Le todos os PDFs da pasta selecionada.
   - Procura o campo Situacao / Situacao Cadastral.
   - Aceita maiusculas/minusculas e acentos diferentes.
   - Considera apenas INAPTA ou INAPTO logo apos o campo, ignorando texto posterior.
   - Nao inclui empresas ATIVAS no Excel.
   - Cria um Excel formatado com os principais dados cadastrais das empresas inaptas.

3) EMITIR SITUACAO FISCAL PGFN / REGULARIZE
   - Usa a MESMA planilha de clientes da automacao 1.
   - Reaproveita do ecac_download.py a leitura da base, a validacao e a
     formatacao de CNPJ, o saneamento de nomes de arquivo e a localizacao
     do Chrome.
   - Para cada empresa: representa, entra no e-CAC, abre Divida Ativa da
     Uniao, abre o PGFN / Regularize, consulta a divida ativa, abre o
     Relatorio Consolidado, marca todas as naturezas e todas as situacoes,
     gera o relatorio, clica em Imprimir e salva o PDF.
   - O PDF e salvo direto com o nome final: "CNPJ - NOME DA EMPRESA.pdf".
   - Empresa sem cadastro no Regularize, pagina indisponivel ou elemento
     que nao apareceu NAO param o lote: a empresa e registrada e o robo
     segue para a proxima.
   - No fim gera um PDF e um CSV com as empresas nao processadas
     (CNPJ, empresa, motivo e etapa).

COMO USAR
---------
1. Coloque todos os arquivos desta pasta juntos. Nao separe app.py, ecac_download.py,
   pdf_inaptos.py, pgfn_interface.py, pgfn_automation.py, pgfn_utils.py, config.py
   e a pasta assets.
2. Clique duas vezes em INICIAR.bat.
3. Na tela:
   - Selecione a planilha de clientes.
   - Selecione a pasta de PDFs (ex.: Desktop\Saidas).
   - Escolha onde salvar Relatorio_Empresas_Inaptas.xlsx.
4. Ative/desative os dois switches conforme o que deseja executar:
   - apenas downloads;
   - apenas analise;
   - os dois, um depois do outro.
5. Clique em INICIAR AUTOMACAO.
6. Durante a automacao da Receita, use o botao CONTINUAR / ENTER da tela quando
   as instrucoes pedirem para voltar ao terminal e pressionar Enter.
7. Os cliques manuais de aprendizado e as marcacoes com F8 continuam iguais.

COMO USAR A AUTOMACAO 3 (PGFN / REGULARIZE)
-------------------------------------------
1. Clique duas vezes em INICIAR_PGFN.bat.
2. Selecione a planilha de clientes e a pasta de destino dos PDFs.
3. Clique em INICIAR AUTOMACAO. O programa abre o Chrome sozinho.
4. Faca o login manualmente pelo Gov.br na janela que abriu.
5. Represente MANUALMENTE a primeira empresa da lista (o console mostra
   qual e o CNPJ).
6. Volte a tela e clique em CONTINUAR / ENTER.
7. O robo assume o controle e processa as demais empresas sozinho.
8. Acompanhe empresa atual, CNPJ, progresso, status e os contadores.

O login, o certificado e o CAPTCHA nunca sao automatizados.
Se a representacao automatica de alguma empresa falhar, o programa pede
que voce represente aquela empresa e clique novamente em CONTINUAR.

Arquivos tecnicos da automacao 3 (log, CSV, relatorio de falhas e o
perfil do Chrome usado por ela):
%LOCALAPPDATA%\PGFNRegularize

Ajustes finos (tempos de espera, seletores das telas, nome do arquivo e
modo do navegador) ficam todos no arquivo config.py.

IMPORTANTE SOBRE A PASTA DE PDFS
--------------------------------
A automacao nao cria logs, CSVs, pastas de diagnostico ou arquivos Excel dentro
na pasta selecionada para os PDFs. Ela fica reservada exclusivamente aos PDFs.
O Excel de empresas inaptas deve ser salvo FORA dessa pasta.

ARQUIVOS DO PROJETO
-------------------
app.py              tela das automacoes 1 e 2
ecac_download.py    automacao 1 (Receita Federal, por mouse e teclado)
pdf_inaptos.py      automacao 2 (analise dos PDFs)
pgfn_interface.py   tela da automacao 3 (PGFN / Regularize)
pgfn_automation.py  fluxo da automacao 3
pgfn_utils.py       esperas, cliques seguros, abas, logs e PDFs
config.py           tempos, seletores e pastas da automacao 3

ARQUIVOS TECNICOS
-----------------
Logs e controles da automacao 1 ficam em:
%LOCALAPPDATA%\ReceitaFederalECAC\runtime

Configuracoes da interface ficam em:
%LOCALAPPDATA%\METAPlanoAutomacao\config.json

DEPENDENCIAS
------------
O projeto usa Python 3, Tkinter (incluso normalmente no Python Windows), openpyxl,
pypdf, pynput e selenium. A automacao principal, o analisador e a automacao da
PGFN tentam instalar dependencias Python ausentes automaticamente.

A automacao 3 controla o Chrome pelo Selenium. O driver correto e baixado
automaticamente pelo proprio Selenium na primeira execucao, entao a maquina
precisa de acesso a internet nesse momento.
