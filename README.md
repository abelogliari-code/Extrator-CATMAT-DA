<h3>Extrator de CATMATs Pro</h3>

BPS · DESID · Ministério da Saúde

Ferramenta para extrair registros de preços do Portal de Compras Governamentais
(dadosabertos.compras.gov.br) e consolidá-los com o histórico do DW (SIASG).
São três abas:

<strong>1. Extração por CATMAT</strong> — a partir de uma lista pronta de códigos (.xlsx ou .csv),
busca todos os registros de preços, repara os CSVs corrompidos da API e grava em Excel
ou CSV. O seletor "Buscar por" alterna entre CATMAT (`codigoItemCatalogo`) e PDM
(`codigoPdm`), espelhando o parâmetro `tipo` da API de Pesquisa de Preço.

<strong>2. Extração por Classes</strong> — descobre os PDMs de uma ou mais Classes e, a partir
deles, os CATMATs para extração. Marcando "Extrair preços direto por PDM", o programa pula
a expansão PDM → CATMAT e consulta a Pesquisa de Preço com `tipo=codigoPdm`, o que reduz
drasticamente o número de requisições.

<strong>3. Consolidação DW + DA</strong> — junta o histórico do DW (SIASG) com o do DA
(Dados Abertos) em uma planilha por Classe, com as abas `dw-XXXX` e `da-XXXX`, removendo do
DA todo registro que já exista no DW — o DW é a fonte preferencial, por trazer mais
informação. A chave é o identificador do item da compra, com 22 dígitos:

- DW → coluna "Identif Item Compra" (já vem com 22 dígitos)
- DA → `idCompra` (zeros à esquerda até 17) + `numeroItemCompra` (até 5)

Além das planilhas por classe são gerados o `Relatorio_Consolidacao.xlsx` (contagens por
classe), o `linhas_em_quarentena.csv` (linhas corrompidas na origem, preservadas na íntegra
e fora das planilhas) e, opcionalmente, o `duplicatas_removidas.csv` para auditoria.

O mesmo motor continua disponível em linha de comando:

```
python consolidar_dw_da.py -e PASTA_DE_ENTRADA -s PASTA_DE_SAIDA --salvar-duplicatas
```

O arquivo `consolidar_dw_da.py` precisa ficar ao lado do `ExtratorCatmat.py`. Sem ele, as
duas primeiras abas seguem funcionando normalmente e a terceira informa o que falta.

<h3>Requisitos Obrigatórios - <strong>PRIORIDADE</strong></h3>

- Consertar os csvs "corrompidos"; --- OK (+/- não detecta todos as falhas dos CSVs)
- Criação do Instalador; --- OK (Inno Setup)
- Extração por PDM e Classe; --- OK
- Avaliar mudar interface para PySide6/PyQt6 ou até uma mini interface web (Flask + HTML) --- Será necessário? é uma solução provisória.
- Checkbox se quer ou não salvar os corrompidos; --- OK
- Excluir colunas nomeUnidadeMedida e percentualMaiorDesconto se não houver dados; --- OK
- Consolidar DW + DA removendo as duplicatas; --- OK (aba 3, mesmo motor do consolidar_dw_da.py)
- Gerar arquivo PDF oficial para licitações (modelo BPS-Legado) --- Será necessário?

<h3>Requisitos Desejáveis</h3>

- Criar o texto padrão para ficar na interface do programa quando abrir; --- OK
- Barra de progresso da consolidação por linha lida, e não por arquivo; --- Não Iniciado

<h3>Requisitos Possíveis</h3>

- Manter da tabela em "Verde Claro, Estilo de Tabela Média 28" no excel final; --- Não Iniciado
