# Inventario

Ferramenta que extrai dados de estoque de um relatório em PDF exportado do ERP (e-Millennium) e gera uma planilha Excel, um log de conferência e uma verificação automática contra os totais impressos no próprio PDF.

## Como usar

1. Exporte o relatório "Quantidade de Itens em Estoque" do ERP em PDF e salve como:

   ```
   Inventario\Inventario\61 MENDES.pdf
   ```

   Esse é o caminho fixo lido pelo script (`caminho_pdf` em `Inventario.py`). Se o nome do arquivo ou da filial mudar, atualize esse caminho no código.

2. Execute:

   ```
   python Inventario.py
   ```

3. O script gera, na mesma pasta do PDF (`Inventario\Inventario\`):
   - `61_MENDES_inventario.xlsx` — planilha com todos os produtos extraídos (Código, Nome do Produto, Cor/Variação, Qtde em Estoque, Qtde Reservada, Qtde Disponível, Página).
   - `61_MENDES_reservas.log` — log detalhado, listando por página os produtos com quantidade reservada > 0, mais um resumo geral.

4. No terminal, é exibido um resumo (totais, primeiros/últimos produtos, lista de reservados) seguido de uma **verificação automática**: o script lê as linhas "Qtde Total em Estoque" e "Qtde Total Reservada" que o próprio PDF imprime na última página e compara com a soma calculada a partir dos produtos extraídos. Se a diferença ficar em 0 (ou muito próxima disso), a extração está correta para aquele relatório.

## Como funciona a extração

O PDF não tem uma tabela estruturada — é texto solto por página, então o script percorre linha a linha e classifica cada uma:

- **Cabeçalho de produto** (`<código> - <nome> ... U * <UN|MT|RL|...>`): identificado pela função `extrair_cabecalho_produto`. O código pode ter poucos dígitos, letras (`105C`) ou pontos (`50.0261`), então a detecção não depende da quantidade de dígitos — ela procura o "terminador" da linha (`* UN`, a palavra isolada `U`, ou um número seguido de 2-4 letras, para cobrir os casos em que a extração do PDF corta a unidade no fim da linha).
- **Cor/Variação** (`<código de 3 dígitos> - <nome>`): guardada como contexto do produto atual.
- **Qtde em Estoque** / **Qtde Reservada**: linhas de totais por produto, aceitam valores decimais no formato brasileiro (vírgula), já que alguns itens medidos (metros de tecido, por exemplo) vêm com casas decimais.
- **Qtde Total em Estoque** / **Qtde Total Reservada**: linhas de totais gerais impressas ao final do relatório, usadas só para a verificação automática.

Cada produto é fechado (gravado na lista de resultados) assim que o próximo cabeçalho aparece, ou ao final do PDF.

## Limitações conhecidas

- Em raríssimos casos (bem menos de 1% das linhas), a extração de texto do PDF corrompe ou corta o valor de estoque/reserva de um produto específico (ex.: um número decimal muito longo dividido entre duas linhas). Esses produtos podem ficar de fora do resultado. Se a verificação automática não bater exatamente, normalmente é por causa disso — não é um bug de lógica, é uma limitação da extração de texto do PDF de origem.
- O caminho do PDF de entrada e os nomes de saída (`61_MENDES_...`) estão fixos no código, pensados para o relatório da filial "61 MENDES". Para outra filial/relatório, ajuste `caminho_pdf` em `Inventario.py`.
