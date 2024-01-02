//ATIVAÇÃO
function onOpen(e){
  adicionarMenu()
}

function adicionarLinha(){adicionarLinha()}
function criarNovoCalendario(){criarNovoCalendario()}
function organizarDatas(){organizarDatas()}
function abrirSobre(){abrirSobre()}

//FUNÇÕES
function adicionarMenu(){
  SpreadsheetApp.getUi()
  .createMenu("Calendário")
    .addItem("Adicionar data", "adicionarLinha")
    .addItem("Criar novo calendário de datas", "criarNovoCalendario")
    .addItem("Organizar datas do calendário", "organizarDatas")
    .addSeparator()
    .addItem("Sobre o script", "abrirSobre")
  .addToUi()
}

let abaAtiva = SpreadsheetApp.getActiveSheet()
let planilhaAtual = SpreadsheetApp.getActive()
 
function criarNovoCalendario(){
  let teste = Browser.msgBox("Criação de novos calendários", "Crie calendários de datas comemorativas para os clientes com apenas um clique!\\n\\nO calendário virá com algumas datas comemorativas padrão, mas você poderá adicionar ou remover elas sem problema.", Browser.Buttons.OK_CANCEL)
  if(teste != "ok"){return SpreadsheetApp.getUi().alert("Nada foi feito.")}

  let nomeCliente = SpreadsheetApp.getUi().prompt("Insira o nome do novo cliente:")

  //DATAS BASE DO CALENDÁRIO
  let datasBase = [["01/01","Ano Novo", "Felicitar",""], ["08/03","Dia Internacional da Mulher", "Felicitar", ""], ["23/03", "Aniversário de Curitiba", "Felicitar", ""], ["01/04", "Dia da Mentira", "Felicitar",""], ["17/04","Páscoa","Felicitar",""],["01/05", "Dia Mundial do Trabalho", "Felicitar", ""], ["08/05", "Dia das Mães", "Felicitar", ""], ["14/08", "Dia dos Pais", "Felicitar", ""], ["07/09", "Dia da Independência", "Felicitar", ""], ["15/09", "Dia do cliente", "Felicitar", ""], ["25/12", "Natal", "Felicitar", ""]]

  //CRIAÇÃO DA PLANILHA
  planilhaAtual.toast(`Criando calendário ${nomeCliente.getResponseText()}...`)
  SpreadsheetApp.getActiveSpreadsheet().insertSheet(nomeCliente.getResponseText())
  let novaPlanilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeCliente.getResponseText())
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(novaPlanilha)
  let planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()

  planilha.deleteRows(datasBase.length+1,(planilha.getMaxRows() - datasBase.length))
  planilha.deleteColumns(4, planilha.getMaxColumns()-4)
  planilha.setRowHeightsForced(1, planilha.getMaxRows(),42)

  //ESTILIZAÇÃO
  const ESTILO_CABECALHO = SpreadsheetApp.newTextStyle()
  .setFontSize(10)
  .setFontFamily("Lato")
  .setBold(true)
  .build();

  const ESTILO_CORPO = SpreadsheetApp.newTextStyle()
  .setFontFamily("Lato")
  .setFontSize(10)
  .setBold(false)
  .build()

  planilha.getRange(1,1,1,4).setTextStyle(ESTILO_CABECALHO)
  planilha.getRange(1,1,planilha.getMaxRows(),1).setTextStyle(ESTILO_CABECALHO)
  planilha.getRange(2,2,planilha.getMaxRows(),planilha.getMaxColumns()-1).setTextStyle(ESTILO_CORPO)
  planilha.getRange(1,1,planilha.getMaxRows(), planilha.getMaxColumns()).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY)
  planilha.getRange(1,1,planilha.getMaxRows(), planilha.getMaxColumns()).createFilter()
  planilha.getRange(1,1,planilha.getMaxRows(), planilha.getMaxColumns()).setBorder(true, true, true, true, true, true, "#9f9f9f", SpreadsheetApp.BorderStyle.SOLID)
  planilha.getRange(1,1,1,4).setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID)

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

  TODA_PLANILHA
    .setVerticalAlignment("middle")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    .setHorizontalAlignment("center")

  planilha.getRange(2,3,planilha.getMaxRows()-1,1).setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .requireValueInList(["Autopromoção", "Felicitar", "Gancho para artigo de blog", "Promoção de produto"], true)
  .build());

  //INSERÇÃO DAS DATAS BASE
  planilhaAtual.toast(`Adicionando datas base...`)
  planilha.getRange(2,1,datasBase.length,4).setValues(datasBase)

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
  
  planilhaAtual.setFrozenRows(1)
}

function abrirSobre(){
  Browser.msgBox("Sobre o script", `
  Este script foi criado por André Moura Santos (contato@andremourasantos.com.br) e disponibilizado, sob licença de domínio público (MIT), no GitHub.
  \\n\\n
  Para reportar erros ou obter ajuda, acesse o repositório no GitHub: https://github.com/andremourasantos/calendario-de-datas-comemorativas
  `, Browser.Buttons.OK)
}

function organizarDatas(){
  planilhaAtual.getRange("A1").activate();
  planilhaAtual.getActiveSheet().getFilter().sort(1, true);
  planilhaAtual.toast(`As datas da planilha ${planilhaAtual.getSheetName()} foram organizadas.`)
}

function adicionarLinha(){
  abaAtiva.insertRowBefore(2)
  abaAtiva.getRange(2,1,1,abaAtiva.getMaxColumns()).setBorder(false, true, true, true, true, true, "#9f9f9f", SpreadsheetApp.BorderStyle.SOLID)
  abaAtiva.getRange(1,1,1,abaAtiva.getMaxColumns()).setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID)
  abaAtiva.setActiveRange(abaAtiva.getRange(2,1))
}