> **ATENÇÃO**: ESTE DOCUMENTO AINDA ESTÁ EM DESENVOLVIMENTO E NÃO ESTÁ COMPLETO. 

# Customizando o código 👨🏽‍💻

É possível alterar os parâmetros como as datas padrão, a quantidade de colunas e também a lista do menu suspenso em "Tipo de publicação". Para isso basta alterar algumas seções específicas do código. As alterações realizadas no código não têm efeito retroativo.

> Atenção: alterar o código pode ocasionar em erro caso não feito corretamente, atente-se as instruções. Caso algo der errado, basta reverter as alterações.

### Alterando as datas padrão (fácil)

> 🟢 Um conhecimento básico de JavaScript auxiliará no processo.

1. Procure pelo seguinte pedaço de código dentro da função [NOME]:
    ```JavaScript
    //DATAS BASE DO CALENDÁRIO
    let datasBase = [["01/01","Ano Novo", "Felicitar",""], ["08/03","Dia Internacional da Mulher", "Felicitar", ""], ["23/03", "Aniversário de Curitiba", "Felicitar", ""], ["01/04", "Dia da Mentira", "Felicitar",""], ["17/04","Páscoa","Felicitar",""],["01/05", "Dia Mundial do Trabalho", "Felicitar", ""], ["08/05", "Dia das Mães", "Felicitar", ""], ["14/08", "Dia dos Pais", "Felicitar", ""], ["07/09", "Dia da Independência", "Felicitar", ""], ["15/09", "Dia do cliente", "Felicitar", ""], ["25/12", "Natal", "Felicitar", ""]]
    ```
    - Note que a variável ```datasBase``` é um vetor composto por vetores com 4 valores (```["valor", "valor", "valor", "valor"]```).
      -  O primeiro valor refere-se à data, no estilo DD/MM.
      - O segundo valor refere-se ao nome da Data.
      - O terceiro valor refere-se a que tipo de item da lista suspensa ele pertece.
      - O quarto valor refere-se as observações adicionais.
    - Para adicionar uma nova data, ao final da lista, insira a nova data conforme especificado anteriormente.

### Alterando os itens da lista do menu suspenso (médio)

> 🟢 Um conhecimento básico de JavaScript e Apps Script, é recomendado.

1. Procure pelo seguinte pedaço de código dentro da função [NOME]:
    ```JavaScript
    planilha.getRange(2,3,planilha.getMaxRows()-1,1).setDataValidation(SpreadsheetApp.newDataValidation()
      .setAllowInvalid(true)
      .requireValueInList(["Autopromoção", "Felicitar", "Gancho para artigo de blog", "Promoção de produto"], true)
      .build());
    ```
2. Altere o trecho de ```.requireValueInList``` inserindo um novo valor, separado por vírgulas e entre aspas duplas:
    ```JavaScript
    //antigo
    .requireValueInList(["Autopromoção", "Felicitar", "Gancho para artigo de blog", "Promoção de produto"], true)
    
    //novo
    .requireValueInList(["Autopromoção", "Felicitar", "Gancho para artigo de blog", "Promoção de produto", "NOVO VALOR 1", "NOVO VALOR 2"], true)
    ```
3. Agora vá em busca do seguinte código:
    ```JavaScript
    //FORMATAÇÃO CONDICIONAL
    abaAtiva = SpreadsheetApp.getActiveSheet();
    let celulas = abaAtiva.getRange(`C2:C${datasBase.length+1}`)

    let regra1 = SpreadsheetApp.newConditionalFormatRule()
      .whenCellEmpty()
      .setBackground("#F4CCCC")
      .setRanges([celulas])
      .build();

    let regra2 = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("Felicitar")
      .setBackground("#fff2cd")
      .setRanges([celulas])
      .build();

    let regra3 = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("Promoção")
      .setBackground("#dad2e9")
      .setRanges([celulas])
      .build();

    let regra4 = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("Gancho")
      .setBackground("#b7e1cd")
      .setRanges([celulas])
      .build();

    let regra5 = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("autopromoção")
      .setBackground("#c9daf8")
      .setRanges([celulas])
      .build();
    
    let regras = abaAtiva.getConditionalFormatRules();
    regras.push(regra1, regra2, regra3, regra4, regra5);
    abaAtiva.setConditionalFormatRules(regras)
    ```
1. Insira uma nova regra para cada um dos novos valores adicionados. Copie e cole este modelo no código:
    ```JavaScript
    let regraX = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("NOVO VALOR 1")
      .setBackground("#FF0000")
      .setRanges([celulas])
      .build();
    ```
    - Insira o código após a regra 5, como no exemplo abaixo:
        ```JavaScript
        [...]

        let regra5 = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("autopromoção")
          .setBackground("#c9daf8")
          .setRanges([celulas])
          .build();

        let regraX = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("NOVO VALOR 1")
          .setBackground("#FF0000")
          .setRanges([celulas])
          .build();
        
        let regras = abaAtiva.getConditionalFormatRules();
        
        [...]
        ```
1. Altere os trechos ```.whenTextContains("NOVO VALOR 1")``` e ```.setBackground("#FF0000")``` pelo nome do novo valor inserido nas etapas anteriores e o código HEX da cor que deseja, respectivamente.
    - Cada novo valor inserido nas etapas anteriores necessita que uma nova regra seja adicionada.
1. Salve o código.

Pronto, agora todas as novas datas e calendários criados virão acompanhados dessa mudança!

### Alterando a quandidade de colunas (difícil)

> ⛔ Atenção: Para executar essa alteração, é recomendado ter conhecimento básico de JavaScript e Apps Script. Conhecimentos em Excel VBA não se aplicam.

1. Procure pelo seguinte pedaço de código dentro da função [NOME]:
    ```JavaScript
    //DADOS DA PLANILHA
    const CABECALHO = planilha.getRange(1,1)
    const NOME_DA_DATA = planilha.getRange(1,2)
    const TIPO_DE_PUBLICACAO = planilha.getRange(1,3)
    const OBSERVACOES = planilha.getRange(1,4)
    const TODA_PLANILHA = planilha.getRange(1,1,planilha.getMaxRows(),planilha.getMaxColumns())
    const CORPO_PLANILHA = planilha.getRange(2,1,planilha.getMaxRows(),planilha.getMaxColumns())

    CABECALHO.setValue("Data")
      planilha.setColumnWidth(1, 126)
    NOME_DA_DATA.setValue("Nome")
      planilha.setColumnWidth(2, 252)
    TIPO_DE_PUBLICACAO.setValue("Tipo de publicação")
      planilha.setColumnWidth(3, 126)
    OBSERVACOES.setValue("Observações")
      planilha.setColumnWidth(4, 252)
    ```