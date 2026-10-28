# Guia passo a passo: como a planilha é feita

Este guia explica o `gerar_planilha.py` por partes, para quem está começando em Python. Leia com o código aberto ao lado.

## 1. A biblioteca openpyxl

O Python sozinho não sabe criar arquivos do Excel. Quem faz isso é a biblioteca **openpyxl**, instalada pelo `requirements.txt`. As ideias principais:

```python
livro = Workbook()                  # um arquivo .xlsx novo, em branco
aba = livro.create_sheet("Nome")    # uma aba nova
aba["A1"] = "Olá"                   # escreve na célula A1
aba.cell(5, 2, 100)                 # escreve 100 na linha 5, coluna 2 (B5)
livro.save("arquivo.xlsx")          # grava no disco
```

## 2. As configurações ficam no topo

Logo no começo do arquivo estão as **categorias e os limites do orçamento**:

```python
DESPESAS = {
    "Moradia": 1500,
    "Mercado": 900,
    ...
}
```

Para mudar a sua planilha, comece por aqui: troque os nomes, os valores, apague ou acrescente linhas. As cores também ficam todas juntas (`VERDE`, `VERMELHO`...), então dá para mudar o visual inteiro num lugar só.

## 3. Uma função para cada aba

O programa é dividido em funções, e cada uma monta uma aba: `aba_lancamentos`, `aba_orcamento`, `aba_resumo`, `aba_categorias` e `aba_como_usar`. A função `main`, no final, chama todas na ordem e salva o arquivo.

Funções pequenas que se repetem, como `titulo` e `cabecalho`, evitam copiar o mesmo código em todas as abas.

## 4. Menus de escolha (validação de dados)

Na aba Lançamentos, o Tipo só aceita "Receita" ou "Despesa":

```python
tipo = DataValidation(type="list", formula1='"Receita,Despesa"')
tipo.add("C5:C1004")
```

O menu da Categoria vem da aba Categorias (`formula1="=Categorias!$A$5:$A$..."`). Por isso, quando você cria uma categoria nova lá, ela aparece no menu.

O mesmo recurso recusa valores errados: data que não é data, valor zero ou negativo.

## 5. As fórmulas fazem as contas

O Python **escreve fórmulas**, e quem calcula é o Excel. Por isso a planilha continua funcionando depois, sem Python.

A fórmula mais importante é a **SOMASES** (em inglês, `SUMIFS`): soma os valores que atendem a várias condições ao mesmo tempo. No Orçamento, o gasto de uma categoria no mês é:

> somar a coluna Valor dos lançamentos **onde** o Tipo é "Despesa" **e** a Categoria é esta **e** a data está dentro do mês escolhido.

No código, a função `soma_do_mes` monta essa fórmula para o Resumo mensal.

Observação: no código as fórmulas estão em inglês (`SUMIFS`, `IF`, `DATE`). É assim que o arquivo do Excel guarda as fórmulas por dentro. Ao abrir no Excel em português, você vê `SOMASES`, `SE` e `DATA`.

## 6. Cores que mudam sozinhas (formatação condicional)

O alerta do orçamento é uma **formatação condicional**: uma regra que pinta a célula conforme o valor dela.

```python
FormulaRule(formula=['$F6="Estourou"'], fill=preencher(VERMELHO_CLARO))
```

Isso quer dizer: "se a coluna F desta linha disser Estourou, pinte a linha de vermelho". A coluna F (Situação) é outra fórmula:

```
=SE(Gasto > Limite; "Estourou"; SE(Gasto >= 80% do Limite; "Atenção"; "Dentro"))
```

Quer que o amarelo apareça em 70% em vez de 80%? Troque o `0.8` por `0.7` na fórmula da Situação.

## 7. Gráficos

Os gráficos apontam para um pedaço da planilha (`Reference`), então se atualizam sozinhos quando os números mudam:

```python
grafico = BarChart()
grafico.add_data(Reference(aba, min_col=2, max_col=3, min_row=5, max_row=17), titles_from_data=True)
```

## 8. O cuidado com os seus dados

No começo da `main`, o programa confere se a planilha já existe. Se existir, **não faz nada**, porque ela pode ter meses de lançamentos seus. Para gerar outra por cima, é preciso pedir: `python gerar_planilha.py --substituir`.

É um hábito bom para qualquer programa: nunca apagar dado do usuário sem ele pedir.

## Exercícios para praticar

1. Acrescente a categoria "Pets" com limite de R$ 150.
2. Mude a cor do cabeçalho para verde-escuro.
3. Mude o alerta amarelo para 70% do limite.
4. Desafio: crie uma coluna "Forma de pagamento" na aba Lançamentos, com menu (Pix, Débito, Crédito, Dinheiro).
