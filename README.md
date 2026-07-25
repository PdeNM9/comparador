# Produtividade por Servidor

Aplicação web em Python + Streamlit para comparar uma lista antiga de
processos distribuídos em abas por servidor com uma planilha atual de processos
em 120 dias.

## Regra de produtividade

Um processo conta como produtivo para o servidor quando:

- estava em uma aba do arquivo `lista 01.26.xlsx`;
- não aparece mais no arquivo atual de `120 dias`.

A coluna `Situação` da lista antiga não entra no cálculo da produtividade. Ela
é preservada apenas como informação histórica.

## Formato esperado

### Arquivo dos servidores

Workbook `.xlsx` com uma aba por servidor. O nome da aba é usado
automaticamente como o nome do servidor.

Cada aba deve conter:

- `Número do Processo`
- `Situação`
- `Observação`

Colunas extras, como `Unnamed: 3`, são preservadas no comparativo e nas
exportações.

### Arquivo atual

Workbook `.xlsx` com a planilha atual de processos em 120 dias, contendo:

- `Descrição Classe`
- `Número Processo`
- `Quantidade de Dias na Situação Atual Processo`
- `Situação Atual`
- `Última Tarefa PJE`

## Saídas do app

O dashboard mostra:

- total de servidores;
- total de processos da lista antiga;
- total de processos atuais;
- quantidade de processos produtivos;
- quantidade de processos novos;
- ranking de produtividade por servidor;
- tabelas filtráveis de resumo, produtivos, novos, atual enriquecida e lista
  antiga comparada.

## Exportações

O relatório completo em Excel possui estas abas:

- `Resumo por servidor`
- `Atual enriquecida`
- `Produtivos - saíram`
- `Novos`
- `Lista antiga comparada`

A planilha atual enriquecida mantém todas as colunas originais do arquivo atual
e adiciona:

- `Servidor`
- `Situação anterior`
- `Observação anterior`
- colunas extras anteriores, quando existirem;
- `Status comparativo`, com `Permaneceu`, `Novo` ou `CNJ vazio`.

Os processos com status `Saiu` aparecem na aba `Produtivos - saíram`, pois não
existem na planilha atual.

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

O endereço local padrão é:

```text
http://localhost:8501
```

## Uso

1. Envie o arquivo dos servidores ou marque a opção para usar os `.xlsx`
   detectados na pasta do projeto.
2. Envie o arquivo atual de 120 dias.
3. Se o arquivo atual tiver mais de uma aba, escolha a aba correta na barra
   lateral.
4. Clique em `Processar produtividade`.
5. Revise o dashboard e baixe as exportações.

## Observações técnicas

- A comparação usa apenas o número CNJ normalizado, removendo pontuação,
  hífens, espaços e outros caracteres não numéricos.
- As operações principais são vetorizadas com Pandas.
- A exportação usa `openpyxl`, com filtros, congelamento de cabeçalho e ajuste
  básico de largura das colunas.
