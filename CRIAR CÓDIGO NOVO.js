// --- FUNÇÃO PRINCIPAL (COM NOMES EXATOS E PASTA PROF.) ---
function criarNovaEstrutura() {
  console.log("🟢 INÍCIO: Começando a execução...");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let abaPainel = null;
  const abas = ss.getSheets();
  for (let i = 0; i < abas.length; i++) {
    if (abas[i].getName().trim().toUpperCase() === "PAINEL DE CONTROLE") {
      abaPainel = abas[i];
      break;
    }
  }
  
  if (!abaPainel) abaPainel = ss.getActiveSheet();
  
  // 1. Identificar a próxima linha livre verificando a coluna D (CÓDIGO)
  const dadosCodigo = abaPainel.getRange("D1:D").getValues();
  let proximaLinha = 4;
  for (let i = 3; i < dadosCodigo.length; i++) { 
    if (dadosCodigo[i][0] !== "") {
      proximaLinha = i + 2; 
    }
  }
  
  let proximoCodigo = String.fromCharCode(65 + (proximaLinha - 4));
  console.log("🔢 PASSO 1: A próxima linha é a " + proximaLinha + ". Criando código '" + proximoCodigo + "'.");
  
  const linhaModelo = 4; 
  
  function pegarLinkDaCelula(letraColuna) {
    const celula = abaPainel.getRange(letraColuna + linhaModelo);
    let link = celula.getRichTextValue().getLinkUrl(); 
    if (!link) link = celula.getValue(); 
    return link;
  }
  
  // Pega os links originais
  const urlPastaBase = pegarLinkDaCelula("E");
  const idPastaBase = extrairIdDoUrl(urlPastaBase);
  
  const urlFormulario = pegarLinkDaCelula("F"); 
  const idFormulario = extrairIdDoUrl(urlFormulario);
  
  const urlPlanilhaCadastro = pegarLinkDaCelula("H"); 
  const idPlanilhaCadastro = extrairIdDoUrl(urlPlanilhaCadastro);
  
  const urlPlanilhaTurmas = pegarLinkDaCelula("M");
  const idPlanilhaTurmas = extrairIdDoUrl(urlPlanilhaTurmas);
  
  if (!idPastaBase || !idFormulario || !idPlanilhaCadastro || !idPlanilhaTurmas) {
    exibirMensagem("Erro: Falta um dos links do modelo na linha 4 (Colunas E, F, H ou M).");
    return;
  }

  try {
    console.log("📝 Lendo os nomes exatos para criação...");
    
    // Forçando o nome da pasta principal para "PASTA PROF." e buscando os outros da linha 2
    let nomePasta = "PASTA PROF."; 
    let nomeForm  = abaPainel.getRange("F2").getValue().toString().trim() || "FORMULÁRIO DE INGLÊS";
    let nomeDoc   = abaPainel.getRange("G2").getValue().toString().trim() || "DOCUMENTOS";
    let nomeCad   = abaPainel.getRange("H2").getValue().toString().trim() || "PLANILHA DE CADASTRO DE INGLÊS";
    let nomeTur   = abaPainel.getRange("M2").getValue().toString().trim() || "PLANILHA DE TURMAS E ALUNOS";
    
    console.log("📁 PASSO 3: Criando novas pastas...");
    const pastaBaseDrive = DriveApp.getFolderById(idPastaBase);
    const pastaPai = pastaBaseDrive.getParents().next(); 
    
    // Cria pastas com "NOME - LETRA"
    const novaPastaPrincipal = pastaPai.createFolder(nomePasta + " - " + proximoCodigo);
    const pastaDocumentos = novaPastaPrincipal.createFolder(nomeDoc + " - " + proximoCodigo);
    
    Utilities.sleep(2000); 
    
    console.log("🔄 PASSO 4: Copiando Formulário e Planilhas...");
    
    // Copiando o Formulário
    let formOriginal = DriveApp.getFileById(idFormulario);
    let arquivoNovoFormulario = formOriginal.makeCopy(nomeForm + " - " + proximoCodigo, novaPastaPrincipal);
    const idNovoFormulario = arquivoNovoFormulario.getId();
    
    // Copiando Planilha de Cadastro
    let planCadOriginal = SpreadsheetApp.openById(idPlanilhaCadastro);
    let novaPlanilhaAppCad = planCadOriginal.copy(nomeCad + " - " + proximoCodigo);
    let arquivoNovaCadastro = DriveApp.getFileById(novaPlanilhaAppCad.getId());
    arquivoNovaCadastro.moveTo(novaPastaPrincipal);
    const idNovoCadastro = arquivoNovaCadastro.getId();
    
    // Copiando Planilha de Turmas
    let planTurOriginal = SpreadsheetApp.openById(idPlanilhaTurmas);
    let novaPlanilhaAppTur = planTurOriginal.copy(nomeTur + " - " + proximoCodigo);
    let arquivoNovaTurmas = DriveApp.getFileById(novaPlanilhaAppTur.getId());
    arquivoNovaTurmas.moveTo(novaPastaPrincipal);
    const idNovoTurmas = arquivoNovaTurmas.getId();
    
    console.log("📝 PASSO 5: Injetando os novos links na linha " + proximaLinha + "...");
    
    abaPainel.getRange(proximaLinha, 4).setValue(proximoCodigo); 
    abaPainel.getRange(proximaLinha, 5).setValue(novaPastaPrincipal.getUrl()); 
    abaPainel.getRange(proximaLinha, 7).setValue(pastaDocumentos.getUrl()); 
    
    copiarLinkFormatado(abaPainel, linhaModelo, proximaLinha, 6, idFormulario, idNovoFormulario);
    
    for (let col = 8; col <= 12; col++) { 
      copiarLinkFormatado(abaPainel, linhaModelo, proximaLinha, col, idPlanilhaCadastro, idNovoCadastro);
    }
    
    for (let col = 13; col <= 24; col++) { 
      copiarLinkFormatado(abaPainel, linhaModelo, proximaLinha, col, idPlanilhaTurmas, idNovoTurmas);
    }
    
    SpreadsheetApp.flush();
    console.log("🎉 SUCESSO TOTAL!");
    exibirMensagem("✅ SUCESSO! A estrutura foi criada perfeitamente!");
    
  } catch (erro) {
    console.log("❌ ERRO: " + erro.message);
    exibirMensagem("⚠️ Ocorreu um erro: " + erro.message);
  }
}

function extrairIdDoUrl(url) {
  if (!url) return "";
  let texto = url.toString();
  let match = texto.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) return match[1];
  let fallbackMatch = texto.match(/([a-zA-Z0-9-_]{25,44})/);
  return fallbackMatch ? fallbackMatch[0] : "";
}

function copiarLinkFormatado(aba, linhaOrigem, linhaDestino, coluna, idAntigo, idNovo) {
  let celulaOrigem = aba.getRange(linhaOrigem, coluna);
  let celulaDestino = aba.getRange(linhaDestino, coluna);
  let richText = celulaOrigem.getRichTextValue();
  
  if (richText && richText.getLinkUrl()) {
     let linkOrigem = richText.getLinkUrl();
     let texto = richText.getText();
     let novoLink = linkOrigem.replace(idAntigo, idNovo); 
     let novoRichText = SpreadsheetApp.newRichTextValue().setText(texto).setLinkUrl(novoLink).build();
     celulaDestino.setRichTextValue(novoRichText);
  } else {
     let valor = celulaOrigem.getValue().toString();
     celulaDestino.setValue(valor.replace(idAntigo, idNovo));
  }
}

function exibirMensagem(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { }
}