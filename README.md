# Comparador de Planilhas Judiciais

Aplicação web em Python + Streamlit para comparar duas planilhas Excel
contendo processos judiciais e gerar uma nova planilha com anotações
transferidas da situação anterior para a situação atual.

## Funcionalidades

- Upload de duas planilhas `.xlsx`.
- Escolha da aba de cada arquivo.
- Seleção manual das colunas de CNJ, responsável e anotações.
- Normalização do CNJ para comparação, removendo pontos, hífens, espaços e
  outros caracteres não numéricos.
- Identificação de processos excluídos, novos e mantidos.
- Transferência automática das anotações da Planilha 1 para os processos
  mantidos na Planilha 2.
- Preservação de todas as linhas e da ordem original da Planilha 2.
- Criação de novas colunas quando a Planilha 2 já possui uma coluna com o
  mesmo nome da anotação selecionada.
- Aviso de CNJs duplicados, com continuidade do processamento.
- Concatenação de responsáveis e anotações duplicadas por CNJ.
- Dashboard com cards, gráfico de barras, gráfico de pizza e tabelas
  filtráveis.
- Exportação em Excel dos excluídos, novos e arquivo completo final.

## Estrutura

```text
app.py
comparison.py
excel_utils.py
dashboard.py
charts.py
export.py
requirements.txt
README.md
```

## Instalação

No Windows PowerShell, dentro da pasta do projeto:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Se o comando `python` não estiver disponível, use o Python Launcher:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install --upgrade pip
py -m pip install -r requirements.txt
```

## Execução

```powershell
streamlit run app.py
```

Depois disso, abra o endereço local exibido pelo Streamlit, normalmente:

```text
http://localhost:8501
```

## Como usar

1. Envie a Planilha 1, que representa a situação anterior.
2. Envie a Planilha 2, que representa a situação atual.
3. Na barra lateral, escolha a aba de cada arquivo.
4. Escolha a coluna CNJ e a coluna responsável em cada planilha.
5. Escolha, na Planilha 1, as colunas que devem ser tratadas como anotações.
6. Clique em `Comparar planilhas`.
7. Revise os cards, gráficos e tabelas.
8. Baixe os arquivos desejados.

## Regras de comparação

A comparação considera exclusivamente o CNJ normalizado. Assim, estes dois
valores são considerados o mesmo processo:

```text
0000000-00.0000.0.00.0000
00000000000000000000
```

Campos como classe, movimentação, vara, assunto ou qualquer outra coluna não
interferem na identificação de novos, excluídos ou mantidos.

## Regras para duplicidades

Quando há CNJs duplicados, a aplicação informa o usuário e continua:

- os processos são contados por CNJ normalizado único;
- responsáveis duplicados são concatenados, preservando valores únicos;
- anotações duplicadas da Planilha 1 são concatenadas por coluna;
- o arquivo final preserva todas as linhas da Planilha 2, inclusive eventuais
  duplicidades.

## Arquivo final

O Excel completo gerado contém:

- todas as linhas da Planilha 2;
- todas as colunas originais da Planilha 2;
- colunas de anotações copiadas da Planilha 1;
- linhas na mesma ordem da Planilha 2;
- células de anotação vazias para processos novos.

Se uma coluna de anotação da Planilha 1 já existir na Planilha 2, a aplicação
cria uma nova coluna com o sufixo `- Planilha 1`.

## Observações

Planilhas com dezenas de milhares de linhas são processadas com operações
vetorizadas do Pandas. A exportação usa `openpyxl` e aplica formatação simples,
filtro automático e congelamento da primeira linha.
