META PLANO - CENTRAL DE AUTOMACOES
=================================

O projeto possui duas automacoes integradas em uma unica tela:

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

COMO USAR
---------
1. Coloque todos os arquivos desta pasta juntos. Nao separe app.py, ecac_download.py,
   pdf_inaptos.py e a pasta assets.
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

IMPORTANTE SOBRE A PASTA DE PDFS
--------------------------------
A automacao nao cria logs, CSVs, pastas de diagnostico ou arquivos Excel dentro
na pasta selecionada para os PDFs. Ela fica reservada exclusivamente aos PDFs.
O Excel de empresas inaptas deve ser salvo FORA dessa pasta.

ARQUIVOS TECNICOS
-----------------
Logs e controles da automacao 1 ficam em:
%LOCALAPPDATA%\ReceitaFederalECAC\runtime

Configuracoes da interface ficam em:
%LOCALAPPDATA%\METAPlanoAutomacao\config.json

DEPENDENCIAS
------------
O projeto usa Python 3, Tkinter (incluso normalmente no Python Windows), openpyxl,
pypdf e pynput. A automacao principal e o analisador tentam instalar dependencias
Python ausentes automaticamente.
