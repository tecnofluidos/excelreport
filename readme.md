# Documentação da Classe TExcelReport

## Índice
- [Introdução](#introdução)
- [Instalação](#instalação)
- [Configuração](#configuração)
  - [Arquivo config.yaml](#arquivo-configyaml)
    - [Styles](#styles)
    - [Cells](#cells)
    - [Dataset](#dataset)
- [Uso da Classe TExcelReport](#uso-da-classe-texcelreport)
  - [Como utilizar](#como-utilizar-a-classe-texcelreport-para-gerar-um-relatório)
- [API da Classe TExcelReport](#api-da-classe-texcelreport)
  - [Métodos](#métodos)
- [Considerações Finais](#considerações-finais)

## Introdução

A classe `TExcelReport` foi desenvolvida para facilitar a criação de relatórios em Excel a partir de um conjunto de dados e um arquivo de configuração YAML. Esta documentação visa fornecer uma visão detalhada de como configurar e utilizar essa classe para gerar relatórios de forma eficiente.

## Instalação

Para utilizar a classe `TExcelReport`, você precisa ter o Python e as bibliotecas necessárias instaladas. Você pode instalar as bibliotecas necessárias usando `pip`:

```bash
pip install pyyaml openpyxl
```
## Configuração

### Arquivo config.yaml
O arquivo config.yaml contém a configuração necessária para a criação do relatório. <br>
Ele deve ser estruturado conforme o exemplo abaixo:

#### Styles
A sessão <b>styles</b> permite a criação de estilos nomeados (Named Styles) que podem ser utilizados em outros elementos do relatório através do parâmetro <b>use</b> na chave <b>style</b> do elemento.<br>
```yaml
styles:
  name_of_style:
    style:
      name : 'name_of_style'
      font:
        name: 'Calibri'
        size: 15
        bold: True
      border:
        style: 'thin'
        color: 'D9D9D9'
        side: ['bottom']   
      fill:
        patternType: 'darkDown'
        start_color: '262626'
        end_color: '262626'           
      alignment:
        horizontal: 'left'
        vertical: 'center' 
```
Todos os componentes de estilo podem ser usados na definição de um estilo nomeado, incluindo propriedades como fonte, borda, preenchimento e alinhamento. No entanto, componentes de formatação numérica devem ser aplicados diretamente ao estilo da célula (ou coluna), imediatamente após a chave <b>use</b>, conforme exemplificado abaixo

```yaml
            ...  
              style:
                use: 'master_data'            
                format: 'dd/mm/yyyy;@' 
```

#### Cells
A Seção Cells permite a configuração de celulas específicas para uso em conjunto com a função <b>setCell</b>

```yaml
cells:
  A1:
    merge: 'A1:H1'
    style:
      use: 'named_style_1'   
```

No exemplo acima, verificamos o uso do estilo nomeado <b>named_style_1</b> para a célula <b>A1</b>, além da aplicação da <b>mesclagem</b> de células para <b>A1:H1</b>.<br><br>
Além do uso de estilos nomeados, a seção cells permite também a aplicação direta de um estilo específico para uma determinada célula. Este estilo atuará exclusivamente nesta célula e não poderá ser utilizado em outra parte do relatório, a menos que seja completamente replicado.

```yaml
cells:
  ...
  B1:
    style:
        font:
            name: 'Calibri'
            size: 10    
            color: 'FFFFFF'    
        fill:
            patternType: 'darkDown'
            start_color: '404040'
            end_color: '404040'
        alignment:
            horizontal: 'center'
            vertical: 'center'  
```

Neste exemplo, será aplicado para a célula <b>B1</b> uma definição exclusiva de fonte, preenchimento e alinhamento.

#### Dataset
A seção <b>dataset</b> é responsável pela impressão dos dados do relatório. Estes dados são obtidos transferindo uma lista de objetos na criação da classe (init)

```python
dataset = [
    {"docnum": "1001", "serial": "A123", "instnum": 1, "docdate": date(2023, 1, 1), "duedate": date(2023, 1, 1), "doctotal": 1000.00, "status": "Pago", "cardname": "Cliente A", "cardcode": "C001", "phone": "(11) 256-0873"},
    # ... mais objetos podem ser adicionados à lista
]
yaml_file = "config.yaml"
report = TExcelReport('Por Cliente', dataset, yaml_file)
```
##### Parametrização da seção dataset
- start<br>
  - Indica a coordenada incial de impressão do relatorio
- break_fields<br>
  - Seção responsável pela "quebra" em grupos durante a impressão (group header, data, group footer)
  - <b>fields</b>: Informe a lista de campos (ou campo) que será analizado para gerar a quebra de grupo. O dataset deverá estar ordenado por este campo.
  - <b>data</b>: Dados adicionais a serem impressos no cabeçalho do grupo (group header) antes da impressão das colunas. Estes dados devem estar disponíveis no dataset, podendo ser aplicados estilos e formatações sobre eles. É importante notar que devemos informar apenas as colunas onde a informação será impressa, pois a linha é calculada pelo gerador.<br>

Neste exemplo acima, a impressão dos dados começa na célula A6. <br>
A quebra de grupo é baseada no campo cardcode, e o valor deste campo será impresso na coluna A com o estilo group_header aplicado.<br>

```yaml
  start: "A6"
  break_fields:
    fields: ['cardcode']
    data: 
      A:
        field: 'cardcode'
        style:
          use: 'group_header'                        

```
- columns<br>
  Nesta seção definimos a impressão dos dados do dataset, sendo que cada coluna deve ser associada a um unico campo do dataset.

```yaml
  columns:  
    A:            
      width: 13
      header: 
        value: 'Título'
        style: 
          use: 'column_header'
      data:
        field: 'docnum'
        style:
          use: 'master_data'       
      footer:
        value: '/*'              
        merge: 'A%:F%'
        style: 
          use: 'column_header'

    # ... mais colunas podem ser adicionados à lista

        formula: 'SUM'

```
  - Coluna (A)
    - width : define a largura da coluna
    - header: 
      - value: define um texto fixo como título ou rótulo da coluna. 
      - style: definição de estilo de formatação ou o uso de um estilo nomeado
    - data:
      - field: associa a coluna a um campo do dataset
      - style: definição de estilo de formatação ou o uso de um estilo nomeado
    - footer 
      - value: define um texto fixo para o rodapé da coluna. Use <b>value: '/*'</b> para identificar um conteudo vazio.
      - formula: define a função totalizadora da coluna, podendo ser aplicada <b>SUM</b>,<b>AVG</b>,<b>COUNT</b>,<b>MIN</b> e <b>MAX</b>.
      - style: definição de estilo de formatação ou o uso de um estilo nomeado<br><br>

  Outras configurações podem ser obtidas extrapolando o arquivo <b>config.yaml</b> do modelo.

## Uso da Classe TExcelReport

### Como utilizar a classe TExcelReport para gerar um relatório:

```python
from datetime import date
from excelreport import TExcelReport

dataset = [
    {"docnum": "1001", "serial": "A123", "instnum": 1, "docdate": date(2023, 1, 1), "duedate":date(2023, 1, 1),"doctotal": 1000.00, "status": "Pago", "cardname": "Cliente A","cardcode":"C001","phone":"(11) 256-0873"},
    {"docnum": "1002", "serial": "A124", "instnum": 1, "docdate": date(2023, 1, 2), "duedate":date(2023, 1, 1),"doctotal": 1500.00, "status": "Pendente", "cardname": "Cliente A","cardcode":"C001","phone":"(11) 256-0873"},
    # ... mais objetos podem ser adicionados à lista
]

dataset.sort(key=lambda x: x['cardcode'])

yaml_file = "config.yaml"

report = TExcelReport('Por Cliente',dataset, yaml_file)
date_of_issue = date.today().strftime('%d/%m/%Y')
report.setCell("A1",f"Títulos em Aberto até {date_of_issue}")
report.setCell("A2",f"Emissão")
report.setCell("B2",f"{date_of_issue}")
report.build()
report.save("relatorio.xlsx")
```
      
## API da Classe TExcelReport

### Métodos
Abaixo está a descrição dos principais métodos da classe TExcelReport:<br>

<b>TExcelReport(sheetname, data, config_file)</b><br>
Inicializa a instância da classe TExcelReport.
 - Parâmetros:
    - <b>sheetname (str)</b>: Nome da planilha.
    - <b>data (list)</b>: Lista de dicionários contendo os dados do relatório.
    - <b>config_file (str)</b>: Caminho para o arquivo de configuração YAML.<br>

<b>setCell(cell, value, style)</b><br>
Define o valor de uma célula específica.
  - Parâmetros:
    - <b>cell (str)</b>: Identificação da célula (ex.: "A1").
    - <b>value (str)</b>: Valor a ser atribuído à célula.
    - <b>style (list)</b>: dicionario de estilo (opcional)<br>

<b>build()</b><br>
Constrói o relatório conforme a configuração e os dados fornecidos.<br>
<b>save(filename)</b><br>
Salva o relatório em um arquivo Excel.
  - Parâmetros:
    - <b>filename (str)</b>: Nome do arquivo a ser salvo (ex.: "relatorio.xlsx").

## Considerações Finais
A classe TExcelReport proporciona uma maneira flexível e configurável de gerar relatórios em Excel.<br>
Utilizando um arquivo YAML para a configuração, é possível customizar o layout e o estilo do relatório facilmente. Esta documentação visa fornecer todas as informações necessárias para configurar e utilizar a classe de forma eficiente.<br>

Para quaisquer dúvidas, problemas ou expansão, consulte a documentação oficial das bibliotecas utilizadas.<br>

```
Essa documentação em formato Markdown deve cobrir todos os aspectos necessários para entender e utilizar a classe `TExcelReport` e o arquivo de configuração `config.yaml`. Você pode adicionar mais detalhes conforme necessário para adaptar às necessidades específicas dos usuários.
```