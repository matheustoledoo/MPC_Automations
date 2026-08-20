META PLANO - CENTRAL DE AUTOMACOES
=================================

O projeto possui tres automacoes integradas em uma unica tela:

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

3) RELATORIOS PGFN / REGULARIZE (DIVIDA ATIVA DA UNIAO)
   - Usa a MESMA planilha de clientes da automacao 1.
   - Reaproveita integralmente a representacao do eCAC (CNPJ, Procurador,
     espera de 30 s, Representar, espera de 45 s, fechar menu lateral).
   - Depois de representar, segue o fluxo:
     eCAC -> Divida Ativa da Uniao -> PGFN Todos os servicos do Regularize ->
     Ciente -> (Fechar aviso, quando aparece) -> Consultar Divida Ativa ->
     espera de 25 s -> Relatorio consolidado -> Natureza "Todos" ->
     Situacao "Todos" -> Gerar relatorio -> espera de 20 s ->
     scroll ate o fim da pagina -> Imprimir -> Salvar como PDF.
   - Mantem 5 posicoes de teste do botao Imprimir, marcadas com F8.
   - Salva o PDF direto na pasta escolhida, com o nome da empresa e o CNPJ.
   - Empresas sem cadastro no Regularize (tela/endereco de cadastro) nao
     travam a execucao: sao puladas, registradas e listadas em um PDF final
     na propria pasta dos relatorios PGFN.
   - Tambem sem Selenium, sem ChromeDriver e sem PyAutoGUI.

COMO USAR
---------
1. Coloque todos os arquivos desta pasta juntos. Nao separe app.py, ecac_download.py,
   pdf_inaptos.py e a pasta assets.
2. Clique duas vezes em INICIAR.bat. Nao separe pgfn_regularize.py dos demais.
3. Na tela:
   - Selecione a planilha de clientes.
   - Selecione a pasta de PDFs (ex.: Desktop\Saidas).
   - Escolha onde salvar Relatorio_Empresas_Inaptas.xlsx.
   - Selecione a pasta dos relatorios PGFN (ex.: Desktop\Saidas PGFN).
     Ela precisa ser diferente da pasta de PDFs do eCAC.
4. Ative/desative os switches conforme o que deseja executar:
   - apenas downloads;
   - apenas analise;
   - apenas PGFN / Regularize;
   - qualquer combinacao. As etapas ligadas rodam na ordem 1, 2 e 3.
5. Clique em INICIAR AUTOMACAO.
6. Durante as automacoes 1 e 3, use o botao CONTINUAR / ENTER da tela quando
   as instrucoes pedirem para voltar ao terminal e pressionar Enter.
7. Os cliques manuais de aprendizado e as marcacoes com F8 continuam iguais.
8. Na automacao 3 o aprendizado acontece na PRIMEIRA empresa: voce faz o fluxo
   clique a clique e o robo grava tudo. Da segunda empresa em diante ele repete
   sozinho.

IMPORTANTE SOBRE A PASTA DE PDFS
--------------------------------
A automacao nao cria logs, CSVs, pastas de diagnostico ou arquivos Excel dentro
na pasta selecionada para os PDFs. Ela fica reservada exclusivamente aos PDFs.
O Excel de empresas inaptas deve ser salvo FORA dessa pasta.
Os relatorios do PGFN / Regularize tambem tem pasta propria, para nao se
misturarem com os relatorios do eCAC analisados pela automacao 2.

ARQUIVOS TECNICOS
-----------------
Logs e controles da automacao 1 ficam em:
%LOCALAPPDATA%\ReceitaFederalECAC\runtime

Logs e controles da automacao 3 ficam em:
%LOCALAPPDATA%\ReceitaFederalPGFN\runtime

Configuracoes da interface ficam em:
%LOCALAPPDATA%\METAPlanoAutomacao\config.json

DEPENDENCIAS
------------
O projeto usa Python 3, Tkinter (incluso normalmente no Python Windows), openpyxl,
pypdf e pynput. A automacao principal e o analisador tentam instalar dependencias
Python ausentes automaticamente.
