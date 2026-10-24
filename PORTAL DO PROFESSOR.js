//// ========================================================
// 🛡️ SISTEMA DE SEGURANÇA: TOKENS, COFRE E RESOLUÇÃO SEGURA
// ========================================================
var CHAVE_SECRETA = "OnEnglish_Seguranca_Total_2026"; 

// CORREÇÃO FALHA 8: Sanitização contra Formula/CSV Injection
function sanitizarEntrada(texto) {
  if (!texto) return texto;
  var str = String(texto).trim();
  if (/^[=+\-@]/.test(str)) { return "'" + str; }
  return str;
}

function gerarHashSenha(senha) {
  if (!senha) return "";
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, senha, Utilities.Charset.UTF_8);
  var txtHash = '';
  for (var i = 0; i < rawHash.length; i++) {
    var hashVal = rawHash[i];
    if (hashVal < 0) { hashVal += 256; }
    if (hashVal.toString(16).length == 1) { txtHash += '0'; }
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}
function gerarSenhaAleatoria() {
    var chars = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789@#$*";
    var senha = "";
    for (var i = 0; i < 8; i++) { senha += chars.charAt(Math.floor(Math.random() * chars.length)); }
    return senha;
}
function criarTokenSessao(email, tipo, idUsuario) {
  var dados = { email: email, tipo: tipo, idUsuario: idUsuario, expiracao: Date.now() + (1000 * 60 * 60 * 24) };
  var payload = Utilities.base64Encode(JSON.stringify(dados));
  var assinatura = Utilities.base64Encode(Utilities.computeHmacSha256Signature(payload, CHAVE_SECRETA));
  return payload + "." + assinatura;
}
function validarTokenSessao(token) {
  if (!token) return false;
  var partes = token.split(".");
  if (partes.length !== 2) return false;
  var payload = partes[0];
  var assinaturaRecebida = partes[1];
  var assinaturaReal = Utilities.base64Encode(Utilities.computeHmacSha256Signature(payload, CHAVE_SECRETA));
  if (assinaturaRecebida !== assinaturaReal) return false;
  try {
    var dados = JSON.parse(Utilities.newBlob(Utilities.base64Decode(payload)).getDataAsString());
    if (Date.now() > dados.expiracao) return false;
    return dados;
  } catch(e) { return false; }
}
function obterUrlsSeguras(cracha) {
  if (!cracha || !cracha.email) throw new Error("Acesso Negado: Sessão inválida.");
  var cache = CacheService.getScriptCache();
  var cacheKey = "urls_" + cracha.tipo + "_" + cracha.email.replace(/[^a-zA-Z0-9]/g, "_");
  var urlsCacheadas = cache.get(cacheKey);
  if (urlsCacheadas) return JSON.parse(urlsCacheadas);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var res = { urlTurmas: "", urlCadastro: "" };
  if (cracha.tipo === 'professor') {
    var abaProf = ss.getSheetByName('ACESSO PROFESSOR');
    if (abaProf) {
      var dadosProf = abaProf.getDataRange().getValues();
      for (var i = 1; i < dadosProf.length; i++) {
        var emailsP = String(dadosProf[i][3]).toLowerCase().split(/[\s,;\n]+/).filter(Boolean);
        if (emailsP.includes(cracha.email)) { res.urlCadastro = dadosProf[i][6] || ""; res.urlTurmas = dadosProf[i][7] || ""; break; }
      }
    }
  } else if (cracha.tipo === 'aluno') {
    var abaAluno = ss.getSheetByName('ACESSO ALUNO');
    if (abaAluno) {
      var dadosAluno = abaAluno.getDataRange().getValues();
      for (var r = 5; r < dadosAluno.length; r++) {
        for (var c = 1; c < dadosAluno[r].length; c += 4) {
          var emailsA = String(dadosAluno[r][c]).toLowerCase().split(/[\s,;\n]+/).filter(Boolean);
          if (emailsA.includes(cracha.email)) { res.urlCadastro = dadosAluno[2][c] || ""; res.urlTurmas = dadosAluno[3][c] || ""; break; }
        }
        if (res.urlTurmas) break;
      }
    }
  }
  if (!res.urlTurmas && !res.urlCadastro) throw new Error("Configuração de banco não localizada.");
  cache.put(cacheKey, JSON.stringify(res), 21600); 
  return res;
}
function registrarAcesso(email, status, tipo, localizacao) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaLog = ss.getSheetByName('LOG_ACESSO'); 
    if (abaLog) { abaLog.appendRow([Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"), email, status, tipo, localizacao]); }
  } catch (e) {}
}
function buscarSenhaCofre(idBase) {
    var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('COFRE_SENHAS');
    if(!aba) return null;
    var ultimaLinha = aba.getLastRow();
    if (ultimaLinha < 1) return null;
    var dados = aba.getRange(1, 1, ultimaLinha, 2).getValues(); 
    var idBuscado = String(idBase).trim().toLowerCase();
    for(var i = 0; i < dados.length; i++) { if(String(dados[i][0]).trim().toLowerCase() === idBuscado) return String(dados[i][1]).trim(); }
    return null; 
}
function salvarSenhaCofre(idBase, novoHash) {
    var lock = LockService.getScriptLock();
    lock.waitLock(10000); 
    try {
        var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('COFRE_SENHAS');
        if(!aba) return;
        var ultimaLinha = aba.getLastRow();
        var atualizou = false; var linhasParaDeletar = []; var idBuscado = String(idBase).trim().toLowerCase();
        if (ultimaLinha > 0) {
            var dados = aba.getRange(1, 1, ultimaLinha, 1).getValues();
            for(var i = 0; i < dados.length; i++) {
                if(String(dados[i][0]).trim().toLowerCase() === idBuscado) {
                    if (!atualizou) { aba.getRange(i + 1, 2).setValue(novoHash); atualizou = true; } 
                    else { linhasParaDeletar.push(i + 1); }
                }
            }
        }
        if (!atualizou) aba.appendRow([idBase, novoHash]);
        else if (linhasParaDeletar.length > 0) { for (var j = linhasParaDeletar.length - 1; j >= 0; j--) aba.deleteRow(linhasParaDeletar[j]); }
        SpreadsheetApp.flush();
    } finally { lock.releaseLock(); }
}
function recuperarSenha(emailBuscado) {
    try {
        var emailLimpo = String(emailBuscado).toLowerCase().trim();
        if (!emailLimpo || !emailLimpo.includes("@")) return {sucesso: false, erro: "E-mail inválido."};
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var senhaProvisoria = gerarSenhaAleatoria();
        var hashProvisorio = gerarHashSenha(senhaProvisoria);
        
        var abaProf = ss.getSheetByName('ACESSO PROFESSOR');
        if (abaProf) {
            var dadosProf = abaProf.getDataRange().getValues();
            for (var i = 1; i < dadosProf.length; i++) {
                var emailsP = String(dadosProf[i][3]).toLowerCase().split(/[\s,;\n]+/).filter(Boolean);
                if (emailsP.includes(emailLimpo)) {
                    salvarSenhaCofre(String(dadosProf[i][5]).trim(), hashProvisorio);
                    enviarEmailRecuperacao(emailLimpo, senhaProvisoria);
                    registrarAcesso(emailLimpo, "Reset Automático", "Professor", "Sistema");
                    return {sucesso: true, msg: "Nova senha provisória enviada para seu e-mail!"};
                }
            }
        }
        var abaAluno = ss.getSheetByName('ACESSO ALUNO');
        if (abaAluno) {
            var dadosAluno = abaAluno.getDataRange().getValues();
            for (var r = 5; r < dadosAluno.length; r++) {
                for (var c = 1; c < dadosAluno[r].length; c += 4) {
                    var emailsA = String(dadosAluno[r][c]).toLowerCase().split(/[\s,;\n]+/).filter(Boolean);
                    if (emailsA.includes(emailLimpo)) {
                        var idBase = String(dadosAluno[r][c + 2]).trim();
                        if(idBase) {
                            salvarSenhaCofre(idBase, hashProvisorio);
                            enviarEmailRecuperacao(emailLimpo, senhaProvisoria);
                            registrarAcesso(emailLimpo, "Reset Automático", "Aluno", "Sistema");
                            return {sucesso: true, msg: "Nova senha provisória enviada para seu e-mail!"};
                        }
                    }
                }
            }
        }
        return {sucesso: false, erro: "E-mail não encontrado no sistema."};
    } catch (e) { return {sucesso: false, erro: "Erro ao recuperar senha."}; }
}
function enviarEmailRecuperacao(emailDestino, senhaProvisoria) {
    var remetenteOficial = "onenglishcourse@gmail.com";
    var assunto = "ON ENGLISH - Recuperação de Senha";
    var mensagem = "Sua senha foi resetada com sucesso pelo sistema.\n\nSua nova senha provisória é: " + senhaProvisoria + "\n\nPor favor, faça o login e altere-a imediatamente no seu painel acessando a aba 'Editar Cadastro'.";
    try {
        var aliases = GmailApp.getAliases();
        var emailLogado = Session.getActiveUser().getEmail();
        var opcoes = { name: "ON ENGLISH" };
        if (aliases.includes(remetenteOficial) && emailLogado !== remetenteOficial) opcoes.from = remetenteOficial;
        GmailApp.sendEmail(emailDestino, assunto, mensagem, opcoes);
    } catch(e) {}
}

function doGet(e) {
  // CORREÇÃO FALHA 9: Mantida permissão de Iframe conforme exigido para o Google Sites funcionar.
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ON ENGLISH - Portal de Acesso')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) 
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function verificarLoginUnificado(emailBuscado, senhaBuscada, localizacao) {
  var local = localizacao || "Desconhecida"; 
  try {
    var emailLimpo = String(emailBuscado || "").toLowerCase().trim();
    var senhaLimpa = String(senhaBuscada || "").trim();
    if (!emailLimpo || !senhaLimpa) return JSON.stringify({ sucesso: false, erro: "Por favor, preencha o e-mail e a senha." });
    
    var cache = CacheService.getScriptCache();
    var chaveHash = emailLimpo.replace(/[^a-zA-Z0-9]/g, "_");
    var cacheBloqueioKey = "bloq_" + chaveHash;
    var cacheTentativasKey = "tent_" + chaveHash;
    if (cache.get(cacheBloqueioKey)) {
      registrarAcesso(emailLimpo, "BLOQUEIO RATE LIMIT", "Bloqueado", local);
      return JSON.stringify({ sucesso: false, erro: "Conta temporariamente bloqueada. Tente novamente em 15 minutos." });
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var senhaDigitadaHash = gerarHashSenha(senhaLimpa); 
    var emailEncontradoEmAlgumLugar = false;
   
    var abaProf = ss.getSheetByName('ACESSO PROFESSOR');
    if (abaProf) {
      var dadosProf = abaProf.getDataRange().getValues();
      for (var i = 1; i < dadosProf.length; i++) {
        var emailsP = String(dadosProf[i][3]).toLowerCase().split(/[\s,;\n]+/).filter(Boolean);
        if (emailsP.includes(emailLimpo)) {
          emailEncontradoEmAlgumLugar = true;
          var codigoProf = String(dadosProf[i][5]).trim();
          var hashNoCofreProf = buscarSenhaCofre(codigoProf);
          var loginAprovado = false;
          if (hashNoCofreProf) {
              if (senhaDigitadaHash === hashNoCofreProf) loginAprovado = true;
          } else {
              var senhaP_legada = dadosProf[i][10] ? String(dadosProf[i][10]).trim() : "";
              if (senhaP_legada === senhaLimpa || senhaP_legada === senhaDigitadaHash) { loginAprovado = true; salvarSenhaCofre(codigoProf, senhaDigitadaHash); }
          }
          if (loginAprovado) {
            cache.remove(cacheTentativasKey); cache.remove(cacheBloqueioKey);
            registrarAcesso(emailLimpo, "Sucesso", "Professor", local);
            var nasc = dadosProf[i][11];
            return JSON.stringify({
              sucesso: true, token: criarTokenSessao(emailLimpo, 'professor', codigoProf), tipo: 'professor', linhaProfessor: i + 1,
              nomeProf: dadosProf[i][1] || "Professor", fotoProf: dadosProf[i][2] || "", emailProf: dadosProf[i][3] || "Não cadastrado", pixProf: dadosProf[i][4] || "Aguardando...", codigoProf: codigoProf, siteInscricao: dadosProf[i][8] || "#", nascimentoProf: (nasc instanceof Date) ? Utilities.formatDate(nasc, Session.getScriptTimeZone(), "dd/MM/yyyy") : (nasc ? String(nasc).trim() : "Não informado")
            });
          }
        }
      }
    }
   
    var abaAluno = ss.getSheetByName('ACESSO ALUNO');
    if (abaAluno) {
      var dadosAluno = abaAluno.getDataRange().getValues();
      var rEncontrado = -1, cEncontrado = -1;
      for (var r = 5; r < dadosAluno.length; r++) {
        for (var c = 1; c < dadosAluno[r].length; c += 4) {
          var emailsA = String(dadosAluno[r][c]).toLowerCase().split(/[\s,;\n]+/).filter(Boolean);
          if (emailsA.includes(emailLimpo)) {
            emailEncontradoEmAlgumLugar = true;
            var idBase = dadosAluno[r][c + 2] ? String(dadosAluno[r][c + 2]).trim() : "";
            var hashNoCofre = buscarSenhaCofre(idBase);
            var loginAprovadoAluno = false; var ehPrimeiroAcesso = false;
            if (hashNoCofre) {
                if (senhaDigitadaHash === hashNoCofre) { loginAprovadoAluno = true; if (senhaLimpa === "1111") ehPrimeiroAcesso = true; }
            } else {
                if (senhaLimpa === "1111") { loginAprovadoAluno = true; ehPrimeiroAcesso = true; salvarSenhaCofre(idBase, senhaDigitadaHash); } 
                else if (String(dadosAluno[r][c + 3]).trim() === senhaLimpa || String(dadosAluno[r][c + 3]).trim() === senhaDigitadaHash) { loginAprovadoAluno = true; salvarSenhaCofre(idBase, senhaDigitadaHash); }
            }
            if (loginAprovadoAluno) { rEncontrado = r; cEncontrado = c; break; }
          }
        }
        if (rEncontrado !== -1) break;
      }
      
      if (rEncontrado !== -1) {
        cache.remove(cacheTentativasKey); cache.remove(cacheBloqueioKey);
        var matricula = dadosAluno[rEncontrado][cEncontrado + 1] || "";
        var idBaseConfirmado = dadosAluno[rEncontrado][cEncontrado + 2] || "";
        var urls = obterUrlsSeguras({ email: emailLimpo, tipo: 'aluno' });
        if (!urls.urlTurmas) return JSON.stringify({ sucesso: false, erro: "Link da planilha de turmas não configurado." });
        registrarAcesso(emailLimpo, "Sucesso", "Aluno", local);
        var baseAlunos = SpreadsheetApp.openByUrl(urls.urlTurmas).getSheetByName("ALUNOS").getDataRange().getDisplayValues();
        var perfis = [];
        for (var i = 1; i < baseAlunos.length; i++) {
           var emailsPlan = String(baseAlunos[i][5]).toLowerCase().split(/[\s,;\n]+/).filter(Boolean);
           var matPlan = String(baseAlunos[i][1]).trim();
           if (emailsPlan.includes(emailLimpo) || (matPlan !== "" && matPlan === String(matricula).trim())) {
               perfis.push({ idAluno: baseAlunos[i][0] || idBaseConfirmado, nome: baseAlunos[i][6] || "Aluno", turma: baseAlunos[i][7] || "", horario: baseAlunos[i][8] || "" });
           }
        }
        if(perfis.length === 0) perfis.push({ idAluno: idBaseConfirmado, nome: "Aluno", turma: "", horario: "" });
        return JSON.stringify({ sucesso: true, token: criarTokenSessao(emailLimpo, 'aluno', idBaseConfirmado || matricula), tipo: 'aluno', email: emailLimpo, matricula: matricula, perfis: perfis, primeiroAcesso: ehPrimeiroAcesso });
      }
    }
    
    var tentativas = parseInt(cache.get(cacheTentativasKey) || "0") + 1;
    Utilities.sleep(1000);
    if (tentativas >= 5) {
      cache.put(cacheBloqueioKey, "bloqueado", 900); cache.remove(cacheTentativasKey);
      return JSON.stringify({ sucesso: false, erro: "Conta bloqueada após 5 tentativas. Tente novamente em 15 minutos." });
    } else {
      cache.put(cacheTentativasKey, String(tentativas), 900);
      return JSON.stringify({ sucesso: false, erro: emailEncontradoEmAlgumLugar ? "Senha incorreta." : "E-mail não cadastrado no sistema." });
    }
  } catch (e) { return JSON.stringify({ sucesso: false, erro: "Erro interno no servidor." }); } // CORREÇÃO FALHA 10
}

function alterarSenhaPrimeiroAcesso(token, novaSenha) {
    var cracha = validarTokenSessao(token);
    if (!cracha) return { sucesso: false, erro: "Sessão inválida." };
    if (!novaSenha || novaSenha.trim().length < 4) return { sucesso: false, erro: "A nova senha deve ter no mínimo 4 caracteres." };
    salvarSenhaCofre(cracha.idUsuario, gerarHashSenha(novaSenha.trim()));
    try {
        var abaAluno = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ACESSO ALUNO');
        if (abaAluno) {
            var dadosAluno = abaAluno.getDataRange().getValues();
            for (var r = 5; r < dadosAluno.length; r++) {
                for (var c = 1; c < dadosAluno[r].length; c += 4) {
                    if (String(dadosAluno[r][c + 2]).trim() === String(cracha.idUsuario).trim()) { abaAluno.getRange(r + 1, c + 4).setValue("[PROTEGIDO]"); break; }
                }
            }
        }
    } catch(e) {}
    return { sucesso: true };
}

function salvarEdicaoCadastroAluno(token, idAluno, senhaAtual, novaSenha, dadosForm) {
    var cracha = validarTokenSessao(token);
    if(!cracha) return {sucesso: false, erro: "Acesso Negado."};
    if (cracha.tipo === 'aluno' && String(idAluno).trim() !== String(cracha.idUsuario).trim()) return { sucesso: false, erro: "Você só pode editar seu próprio perfil." };
    
    var urls = obterUrlsSeguras(cracha);
    if (novaSenha && novaSenha.trim() !== "") {
        var hashAtual = gerarHashSenha(senhaAtual.trim());
        var hashNoCofre = buscarSenhaCofre(idAluno);
        var senhaValida = false;
        if (hashNoCofre) { if (hashNoCofre === hashAtual) senhaValida = true; } 
        else if (senhaAtual.trim() === "1111") senhaValida = true;
        if (!senhaValida) return {sucesso: false, erro: "Senha atual incorreta!"};
        salvarSenhaCofre(idAluno, gerarHashSenha(novaSenha.trim())); 
    }
    
    if (dadosForm && dadosForm.linha) {
        try {
            var abas = SpreadsheetApp.openByUrl(urls.urlCadastro).getSheets();
            var abaCad = null;
            for (var i = 0; i < abas.length; i++) {
                var nomeAba = abas[i].getName().toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").trim();
                if (nomeAba.indexOf("formulario") !== -1 && nomeAba.indexOf("cadastro") !== -1) { abaCad = abas[i]; break; }
            }
            if (abaCad) {
                var l = parseInt(dadosForm.linha);
                if (cracha.tipo === 'aluno') {
                    if (String(abaCad.getRange(l, 10).getDisplayValue()).toLowerCase().trim() !== cracha.email && String(abaCad.getRange(l, 2).getDisplayValue()).trim() !== String(cracha.idUsuario).trim()) return { sucesso: false, erro: "Acesso Negado: A linha solicitada não pertence ao seu usuário." };
                }
                // CORREÇÃO FALHA 8: Input sanitizado para prevenir Formula Injection
                abaCad.getRange(l, 6).setValue(sanitizarEntrada(dadosForm.nomeResp)); 
                abaCad.getRange(l, 7).setValue(sanitizarEntrada(dadosForm.nomeAluno)); 
                abaCad.getRange(l, 8).setValue(sanitizarEntrada(dadosForm.aniversario)); 
                abaCad.getRange(l, 9).setValue(sanitizarEntrada(dadosForm.cpf)); 
                abaCad.getRange(l, 10).setValue(sanitizarEntrada(dadosForm.email)); 
                abaCad.getRange(l, 11).setValue(sanitizarEntrada(dadosForm.telefone)); 
                abaCad.getRange(l, 12).setValue(sanitizarEntrada(dadosForm.cep)); 
                abaCad.getRange(l, 13).setValue(sanitizarEntrada(dadosForm.endereco)); 
                abaCad.getRange(l, 14).setValue(sanitizarEntrada(dadosForm.paisEstado)); 
                abaCad.getRange(l, 15).setValue(sanitizarEntrada(dadosForm.instagram)); 
                abaCad.getRange(l, 16).setValue(sanitizarEntrada(dadosForm.tiktok));
            }
        } catch(e) { return {sucesso: false, erro: "Erro interno no servidor."}; } // CORREÇÃO FALHA 10
    }
    return {sucesso: true};
}

function getDadosPainelAluno(token, turma, matricula, idAlunoEspecifico) {
    var cracha = validarTokenSessao(token);
    if(!cracha) throw new Error("Acesso Negado.");
    var urls = obterUrlsSeguras(cracha);
    var resp = { acad: null, fin: null, crono: null, turmaInfo: { nivel: "-", porcentagem: "-", desempenhoTurma: "-", linkAulas: "-", aulasGravadas: "-" } };
    try {
        var ssTurmas = SpreadsheetApp.openByUrl(urls.urlTurmas);
        if (turma && turma.toUpperCase().indexOf("TURMA") !== -1) {
            var abaTurma = ssTurmas.getSheetByName(turma);
            if (abaTurma) {
                resp.turmaInfo.nivel = abaTurma.getRange("D4").getDisplayValue();
                resp.turmaInfo.porcentagem = abaTurma.getRange("E4").getDisplayValue();
                var abaProf = ssTurmas.getSheetByName("PROFESSOR");
                if (abaProf) {
                    var profDados = abaProf.getDataRange().getDisplayValues();
                    for(var r = 0; r < profDados.length; r++) {
                        if(profDados[r][5] && String(profDados[r][5]).trim().toUpperCase() === turma.trim().toUpperCase()) {
                            resp.turmaInfo.desempenhoTurma = profDados[r].length > 13 ? profDados[r][13] : "-";
                            resp.turmaInfo.linkAulas = profDados[r].length > 14 ? profDados[r][14] : "-";
                            resp.turmaInfo.aulasGravadas = profDados[r].length > 15 ? profDados[r][15] : "-";
                            break;
                        }
                    }
                }
                var rangeAulas = abaTurma.getRange("B12:E61");
                var aulasValores = rangeAulas.getDisplayValues();
                var aulasRichText = rangeAulas.getRichTextValues();
                var cronograma = [];
                for (var i = 0; i < aulasValores.length; i++) {
                    if (aulasValores[i].join("").trim() === "") continue;
                    var linhaAula = [];
                    for (var j = 0; j < 4; j++) {
                       var texto = aulasValores[i][j] || "";
                       var link = aulasRichText[i][j] ? aulasRichText[i][j].getLinkUrl() : null;
                       linhaAula.push(link ? "<a href='" + link + "' target='_blank' style='color:#3498db; text-decoration:none; font-weight:bold;'>" + texto + "</a>" : texto);
                    }
                    cronograma.push(linhaAula);
                }
                resp.crono = cronograma;
                var alunosMatriz = abaTurma.getRange("H3:AZ61").getDisplayValues();
                var colAluno = -1;
                for (var c = 0; c < alunosMatriz[0].length; c++) {
                    var idLinha4 = String(alunosMatriz[1][c]).trim();
                    if (idLinha4 !== "" && idLinha4 === String(idAlunoEspecifico).trim()) { colAluno = c; break; }
                }
                if (colAluno === -1) {
                    for (var c = 0; c < alunosMatriz[0].length; c++) {
                        if (String(alunosMatriz[1][c]).trim() === String(matricula).trim()) { colAluno = c; break; }
                    }
                }
                if (colAluno !== -1) {
                    var palavras = [];
                    for (var p = 9; p <= 58; p++) { if (alunosMatriz[p][colAluno] && String(alunosMatriz[p][colAluno]).trim() !== "") palavras.push(alunosMatriz[p][colAluno]); }
                    var corridaTurma = [];
                    for (var c2 = 0; c2 < alunosMatriz[0].length; c2++) {
                        if (alunosMatriz[2][c2] && String(alunosMatriz[2][c2]).trim() !== "") {
                            corridaTurma.push({ nota: parseFloat(alunosMatriz[6][c2]) || 0, cor: alunosMatriz[8][c2] || "", isMe: (c2 === colAluno) });
                        }
                    }
                    resp.acad = { frequencia: alunosMatriz[5][colAluno] || "-", nota: alunosMatriz[6][colAluno] || "-", relatorio: alunosMatriz[7][colAluno] || "-", mensagem: alunosMatriz[4][colAluno] || "-", palavras: palavras, cor: alunosMatriz[8][colAluno] || "", corridaTurma: corridaTurma };
                }
            }
        }
    } catch(e) { resp.erroAcad = "Erro interno ao carregar dados acadêmicos."; }
   
    try {
        var ssCadastro = SpreadsheetApp.openByUrl(urls.urlCadastro);
        var abaMens = ssCadastro.getSheetByName("MENSALIDADES");
        var abaCabecalho = ssCadastro.getSheetByName('DADOS FINANCEIRO') || ssCadastro.getSheetByName('DADOS GERAIS');
        var pixProf = abaCabecalho ? (abaCabecalho.getRange("C6").getDisplayValue() || "Não informado") : "Não informado";
        if (abaMens) {
            var dadosMens = abaMens.getDataRange().getDisplayValues();
            for (var i = 2; i < dadosMens.length; i++) {
                if (String(dadosMens[i][2]).trim() === String(matricula).trim() || String(dadosMens[i][1]).trim() === String(idAlunoEspecifico).trim()) {
                    resp.fin = { vencimento: dadosMens[i][9]||"-", valor: dadosMens[i][10]||"-", statusAtual: dadosMens[i][11]||"-", proxMes: dadosMens[i][12]||"-", meses: [ dadosMens[i][13], dadosMens[i][14], dadosMens[i][15], dadosMens[i][16], dadosMens[i][17], dadosMens[i][18], dadosMens[i][19], dadosMens[i][20], dadosMens[i][21], dadosMens[i][22], dadosMens[i][23], dadosMens[i][24] ], pix: pixProf }; break;
                }
            }
        }
    } catch(e) { resp.erroFin = "Erro interno ao carregar mensalidades."; } // CORREÇÃO FALHA 10
    return JSON.stringify(resp);
}

function obterDadosFinanceiros(token) {
    var cracha = validarTokenSessao(token);
    if(!cracha || cracha.tipo !== 'professor') throw new Error("Acesso Negado.");
    var urls = obterUrlsSeguras(cracha);
    try {
        var ss = SpreadsheetApp.openByUrl(urls.urlCadastro);
        var abaMens = ss.getSheetByName('MENSALIDADES');
        var abaCont = ss.getSheetByName('CONTATOS');
        if (!abaMens) throw new Error("Aba MENSALIDADES não encontrada.");
       
        var abaCabecalho = ss.getSheetByName('DADOS FINANCEIRO') || ss.getSheetByName('DADOS GERAIS');
        var infoCabecalho = { codigo: "-", professor: "-", email: "-", pix: "-", totalAlunos: "0", valorEsperado: "R$ 0,00", valorAtrasado: "R$ 0,00", valorArrecadado: "R$ 0,00" };
        if (abaCabecalho) {
            try {
                var cabecalhoDados = abaCabecalho.getRange("C2:P13").getDisplayValues();
                infoCabecalho.codigo = cabecalhoDados[0][0] || "-"; infoCabecalho.professor = cabecalhoDados[1][0] || "-"; infoCabecalho.email = cabecalhoDados[3][0] || "-"; infoCabecalho.pix = cabecalhoDados[4][0] || "-"; infoCabecalho.totalAlunos = cabecalhoDados[2][4] || cabecalhoDados[3][4] || "0"; infoCabecalho.valorEsperado = cabecalhoDados[2][6] || cabecalhoDados[3][6] || "R$ 0,00"; infoCabecalho.valorAtrasado = cabecalhoDados[2][8] || cabecalhoDados[3][8] || "R$ 0,00"; infoCabecalho.valorArrecadado = cabecalhoDados[0][12] || cabecalhoDados[3][12] || cabecalhoDados[2][10] || "R$ 0,00";
            } catch(e) { }
        }
       
        var dadosMens = abaMens.getDataRange().getDisplayValues();
        var dadosCont = abaCont ? abaCont.getDataRange().getDisplayValues() : [];
        var alunos = [];
        for (var i = 2; i < dadosMens.length; i++) {
            if(dadosMens[i][3] || dadosMens[i][4]) {
                var matriculaStr = String(dadosMens[i][2] || "").toLowerCase().trim();
                var emailContato = "", telefoneContato = "", msgPadrao = "", msgPersonalizada = "";
                if (abaCont && matriculaStr !== "") {
                    for(var j = 1; j < dadosCont.length; j++) {
                        if(dadosCont[j].join(" | ").toLowerCase().includes(matriculaStr)) {
                            telefoneContato = dadosCont[j][6] || ""; emailContato = dadosCont[j][7] || ""; msgPadrao = dadosCont[j][9] || ""; msgPersonalizada = dadosCont[j][11] || ""; break;
                        }
                    }
                }
                alunos.push({ linhaPlanilha: i + 1, numero: dadosMens[i][1] || "", matricula: dadosMens[i][2] || "", nomeResp: dadosMens[i][3] || "", nomeAluno: dadosMens[i][4] || "", vencimento: dadosMens[i][9] || "", valor: dadosMens[i][10] || "", statusAtual: dadosMens[i][11] || "", proximoMes: dadosMens[i][12] || "", meses: [ dadosMens[i][13], dadosMens[i][14], dadosMens[i][15], dadosMens[i][16], dadosMens[i][17], dadosMens[i][18], dadosMens[i][19], dadosMens[i][20], dadosMens[i][21], dadosMens[i][22], dadosMens[i][23], dadosMens[i][24] ], emailContato: emailContato, telefoneContato: telefoneContato, msgPadrao: msgPadrao, msgPersonalizada: msgPersonalizada });
            }
        }
        return JSON.stringify({ cabecalho: infoCabecalho, alunos: alunos });
    } catch(e) { throw new Error("Erro interno no servidor."); } // CORREÇÃO FALHA 10
}

function enviarEmailsAlunos(token, listaEnvio) {
    var cracha = validarTokenSessao(token);
    if(!cracha || cracha.tipo !== 'professor') throw new Error("Acesso Negado.");
    if(!listaEnvio || !Array.isArray(listaEnvio) || listaEnvio.length === 0) return JSON.stringify({ sucessos: 0, erros: ["Nenhum e-mail para processar."] });
    
    var urls = obterUrlsSeguras(cracha);
    var ssCadastro = SpreadsheetApp.openByUrl(urls.urlCadastro);
    var emailsAutorizados = {};
    var abaCont = ssCadastro.getSheetByName('CONTATOS');
    if (abaCont) {
        var dadosCont = abaCont.getDataRange().getDisplayValues();
        for (var j = 1; j < dadosCont.length; j++) {
            var eCont = String(dadosCont[j][7] || "").toLowerCase().trim();
            if (eCont && eCont.includes("@")) { eCont.split(/[\s;,]+/).forEach(function(item) { emailsAutorizados[item.trim()] = true; }); }
        }
    }
    
    var resultados = { sucessos: 0, erros: [] };
    var remetenteOficial = "onenglishcourse@gmail.com";
    try {
        var aliases = GmailApp.getAliases();
        var emailLogado = Session.getActiveUser().getEmail();
        for(var i = 0; i < listaEnvio.length; i++) {
            try {
                var item = listaEnvio[i];
                var emailDestinoBruto = String(item.email || "").toLowerCase().trim();
                var destinatarios = emailDestinoBruto.split(/[\s;,]+/).map(function(e){ return e.trim(); }).filter(Boolean);
                var destinatariosValidos = destinatarios.filter(function(e) { return emailsAutorizados[e] === true; });
                if (destinatariosValidos.length === 0) { resultados.erros.push("Destinatário não autorizado: " + emailDestinoBruto); continue; }
                
                var textoSeguro = String(item.mensagem || "").replace(/<script/gi, "").replace(/<iframe/gi, "");
                var textoPuro = textoSeguro.replace(/<br\s*[\/]?>/gi, "\n").replace(/<[^>]+>/g, "");
                var assuntoOficial = String(item.assunto || "Atualização").indexOf("ON ENGLISH") === -1 ? "ON ENGLISH - " + item.assunto : item.assunto;
                
                var opcoes = { htmlBody: textoSeguro, name: "ON ENGLISH" };
                if (aliases.includes(remetenteOficial) && emailLogado !== remetenteOficial) opcoes.from = remetenteOficial;
                GmailApp.sendEmail(destinatariosValidos.join(","), assuntoOficial, textoPuro, opcoes);
                resultados.sucessos++;
                if (i < listaEnvio.length - 1) Utilities.sleep(1500);
            } catch(errEmail) { resultados.erros.push("Erro interno ao enviar."); } // CORREÇÃO FALHA 10
        }
    } catch (eGeral) { resultados.erros.push("Erro no serviço de envio."); }
    return JSON.stringify(resultados);
}

function salvarEdicaoCadastroCompleto(token, linha, dadosForm) {
    var cracha = validarTokenSessao(token);
    if(!cracha || cracha.tipo !== 'professor') return { sucesso: false, erro: "Acesso Negado." };
    var urls = obterUrlsSeguras(cracha);
    try {
        var ss = SpreadsheetApp.openByUrl(urls.urlCadastro);
        var abas = ss.getSheets();
        var aba = null;
        for (var i = 0; i < abas.length; i++) { if (abas[i].getName().toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").includes("cadastro")) { aba = abas[i]; break; } }
        if (!aba) return { sucesso: false, erro: "Aba de Formulário não encontrada." };
        var l = parseInt(linha);
        var valoresAtuais = aba.getRange(l, 2, 1, 15).getDisplayValues()[0];
        var colunasParaSalvar = [ { id: "valorMatricula", col: 4 }, { id: "qtdAlunos", col: 5 }, { id: "nomeResp", col: 6 }, { id: "nomeAluno", col: 7 }, { id: "aniversario", col: 8 }, { id: "cpf", col: 9 }, { id: "email", col: 10 }, { id: "telefone", col: 11 }, { id: "cep", col: 12 }, { id: "endereco", col: 13 }, { id: "paisEstado", col: 14 }, { id: "instagram", col: 15 }, { id: "tiktok", col: 16 } ];
        for (var c = 0; c < colunasParaSalvar.length; c++) {
            var config = colunasParaSalvar[c];
            if (dadosForm[config.id] !== undefined && String(dadosForm[config.id]).trim() !== String(valoresAtuais[config.col - 2]).trim()) {
                aba.getRange(l, config.col).setValue(sanitizarEntrada(dadosForm[config.id])); // CORREÇÃO FALHA 8
            }
        }
        return { sucesso: true };
    } catch(e) { return { sucesso: false, erro: "Erro interno no servidor." }; }
}

function salvarCorAluno(token, nomeTurma, idUnico, matricula, cor) {
    var cracha = validarTokenSessao(token);
    if(!cracha) return false;
    if (cracha.tipo === 'aluno' && String(idUnico).trim() !== String(cracha.idUsuario).trim() && String(matricula).trim() !== String(cracha.idUsuario).trim()) return false;
    var urls = obterUrlsSeguras(cracha);
    try {
        var aba = SpreadsheetApp.openByUrl(urls.urlTurmas).getSheetByName(nomeTurma);
        var rangeLinha4 = aba.getRange("H4:AZ4").getDisplayValues()[0];
        var colIndex = -1;
        for (var i = 0; i < rangeLinha4.length; i++) { if (String(rangeLinha4[i]).trim() === String(idUnico).trim()) { colIndex = i + 8; break; } }
        if (colIndex !== -1) { aba.getRange(11, colIndex).setValue(sanitizarEntrada(cor)); return true; }
        return false;
    } catch(e) { return false; }
}

function salvarProgressoJogoAluno(token, nomeTurma, idUnico, matricula, pontosGanhos, frasesUpdate, palavrasRemover) {
    var cracha = validarTokenSessao(token);
    if(!cracha) return false;
    if (cracha.tipo === 'aluno' && String(idUnico).trim() !== String(cracha.idUsuario).trim() && String(matricula).trim() !== String(cracha.idUsuario).trim()) return false;
    
    // CORREÇÃO FALHA 11: Validação Server-side contra spoofing de pontos via console
    var maxPontosValidos = (frasesUpdate ? frasesUpdate.length : 0) + (palavrasRemover ? palavrasRemover.length : 0);
    if (maxPontosValidos < 1) maxPontosValidos = 50; 
    var pontosSeguros = Math.min(Number(pontosGanhos) || 0, maxPontosValidos);
    
    var urls = obterUrlsSeguras(cracha);
    try {
        var aba = SpreadsheetApp.openByUrl(urls.urlTurmas).getSheetByName(nomeTurma);
        var rangeLinha4 = aba.getRange("H4:AZ4").getDisplayValues()[0];
        var colIndex = -1;
        for (var i = 0; i < rangeLinha4.length; i++) { if (String(rangeLinha4[i]).trim() === String(idUnico).trim()) { colIndex = i + 8; break; } }
        if (colIndex !== -1) {
            var celulaNota = aba.getRange(9, colIndex);
            celulaNota.setValue((parseFloat(celulaNota.getValue()) || 0) + pontosSeguros); // Usando limite seguro
            var palavrasValores = aba.getRange(12, colIndex, 50, 1).getValues();
            var listaOrganizada = [];
            for (var j = 0; j < palavrasValores.length; j++) {
                var palavraAtual = String(palavrasValores[j][0]).trim();
                if (palavraAtual === "") continue;
                if (palavrasRemover && palavrasRemover.indexOf(palavraAtual) !== -1) continue;
                var foiAtualizada = false;
                if (frasesUpdate && frasesUpdate.length > 0) {
                    for (var f = 0; f < frasesUpdate.length; f++) { if (String(frasesUpdate[f].original).trim() === palavraAtual) { listaOrganizada.push(frasesUpdate[f].novo); foiAtualizada = true; break; } }
                }
                if (!foiAtualizada) listaOrganizada.push(palavraAtual);
            }
            var arraySalvar = [];
            for (var p = 0; p < 50; p++) { arraySalvar.push([p < listaOrganizada.length ? sanitizarEntrada(listaOrganizada[p]) : ""]); }
            aba.getRange(12, colIndex, 50, 1).setValues(arraySalvar);
            return true;
        }
        return false;
    } catch(e) { return false; }
}

function salvarPerfilProfessor(token, linha, email, pix, foto, nascimento, senhaAtual, novaSenha) {
    var cracha = validarTokenSessao(token);
    if(!cracha || cracha.tipo !== 'professor') return { sucesso: false, erro: "Acesso Negado." };
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ACESSO PROFESSOR');
        var dadosProf = ss.getDataRange().getValues();
        var codigoProf = String(cracha.idUsuario).trim();
        var linhaReal = -1;
        for (var i = 1; i < dadosProf.length; i++) { if (String(dadosProf[i][5]).trim() === codigoProf) { linhaReal = i + 1; break; } }
        if (linhaReal === -1) return { sucesso: false, erro: "Registro não localizado." };
        if (novaSenha && novaSenha.trim() !== "") {
            var hashAtual = gerarHashSenha(senhaAtual.trim());
            var hashNoCofre = buscarSenhaCofre(codigoProf);
            var senhaValida = false;
            if (hashNoCofre) { if (hashNoCofre === hashAtual) senhaValida = true; } 
            else if (String(dadosProf[linhaReal - 1][10] || "").trim() === senhaAtual.trim() || String(dadosProf[linhaReal - 1][10] || "").trim() === hashAtual) senhaValida = true;
            if (!senhaValida) return { sucesso: false, erro: "Senha atual incorreta!" };
            salvarSenhaCofre(codigoProf, gerarHashSenha(novaSenha.trim()));
        }
        ss.getRange(linhaReal, 3).setValue(sanitizarEntrada(foto)); ss.getRange(linhaReal, 4).setValue(sanitizarEntrada(email)); ss.getRange(linhaReal, 5).setValue(sanitizarEntrada(pix)); ss.getRange(linhaReal, 12).setValue(sanitizarEntrada(nascimento));
        return { sucesso: true };
    } catch(e) { return { sucesso: false, erro: "Erro interno." }; }
}

function getDadosProfessor(token) {
    var cracha = validarTokenSessao(token);
    if(!cracha || cracha.tipo !== 'professor') throw new Error("Acesso Negado.");
    var urls = obterUrlsSeguras(cracha);
    try {
        var aba = SpreadsheetApp.openByUrl(urls.urlTurmas).getSheetByName("PROFESSOR");
        var ultimaLinha = aba.getLastRow() < 3 ? 10 : aba.getLastRow();
        return JSON.stringify({ identificacao: aba.getRange("B1:E7").getDisplayValues(), configTurmas: aba.getRange("F1:P" + ultimaLinha).getDisplayValues() });
    } catch(e) { throw new Error("Erro interno."); }
}

function getDadosAlunos(token) {
    var cracha = validarTokenSessao(token);
    if(!cracha || cracha.tipo !== 'professor') throw new Error("Acesso Negado.");
    var urls = obterUrlsSeguras(cracha);
    try { return JSON.stringify(SpreadsheetApp.openByUrl(urls.urlTurmas).getSheetByName("ALUNOS").getDataRange().getDisplayValues()); }
    catch(e) { throw new Error("Erro interno."); }
}

function getListaTurmas(token) {
    var cracha = validarTokenSessao(token);
    if(!cracha || cracha.tipo !== 'professor') throw new Error("Acesso Negado.");
    var urls = obterUrlsSeguras(cracha);
    try {
        var abas = SpreadsheetApp.openByUrl(urls.urlTurmas).getSheets(); var turmas = [];
        for (var i = 0; i < abas.length; i++) { var nome = abas[i].getName(); if (nome.toUpperCase().indexOf("TURMA") !== -1 && nome.toUpperCase() !== "TURMAS") turmas.push(nome); }
        return JSON.stringify(turmas);
    } catch(e) { throw new Error("Erro interno."); }
}

function getDadosDaTurma(token, nomeTurma) {
    var cracha = validarTokenSessao(token);
    if(!cracha || cracha.tipo !== 'professor') throw new Error("Acesso Negado.");
    var urls = obterUrlsSeguras(cracha);
    try {
        var ss = SpreadsheetApp.openByUrl(urls.urlTurmas);
        var aba = ss.getSheetByName(nomeTurma);
        if (!aba) throw new Error("Turma não encontrada.");
       
        var cabecalho = { status: aba.getRange("B4").getDisplayValue(), horario: aba.getRange("C4").getDisplayValue(), nivel: aba.getRange("D4").getDisplayValue(), porcentagem: aba.getRange("E4").getDisplayValue() };
        var profDados = ss.getSheetByName("PROFESSOR").getDataRange().getDisplayValues();
        var metricasProf = { numAlunos: "-", desempenho: "-", linkAulas: "-", aulasGravadas: "-" };
        for(var r = 0; r < profDados.length; r++) {
            if(profDados[r][5] && profDados[r][5].toString().trim().toUpperCase() === nomeTurma.trim().toUpperCase()) {
                metricasProf.numAlunos = profDados[r].length > 12 ? profDados[r][12] : "-"; metricasProf.desempenho = profDados[r].length > 13 ? profDados[r][13] : "-"; metricasProf.linkAulas = profDados[r].length > 14 ? profDados[r][14] : "-"; metricasProf.aulasGravadas = profDados[r].length > 15 ? profDados[r][15] : "-"; break;
            }
        }
       
        var rangeAulas = aba.getRange("B12:E61");
        var aulasValores = rangeAulas.getDisplayValues();
        var aulasRichText = rangeAulas.getRichTextValues();
        var cronograma = [];
        for (var i = 0; i < aulasValores.length; i++) {
            if (aulasValores[i].join("").trim() === "") continue;
            var linhaAula = [];
            for (var j = 0; j < 4; j++) { var texto = aulasValores[i][j]; var link = aulasRichText[i][j] ? aulasRichText[i][j].getLinkUrl() : null; linhaAula.push(link ? "<a href='" + link + "' target='_blank' style='color:#3498db; text-decoration:none; font-weight:bold;'>" + texto + "</a>" : texto); }
            cronograma.push(linhaAula);
        }
       
        var alunosMatriz = aba.getRange("H3:AZ61").getDisplayValues();
        var listaAlunos = [];
        for (var c = 0; c < alunosMatriz[0].length; c++) {
            if (!alunosMatriz[2][c] || String(alunosMatriz[2][c]).trim() === "") continue;
            var palavras = []; for(var p = 9; p <= 58; p++) { if(alunosMatriz[p][c] && String(alunosMatriz[p][c]).trim() !== "") palavras.push(alunosMatriz[p][c]); }
            listaAlunos.push({ numero: alunosMatriz[0][c], matricula: alunosMatriz[1][c], nome: alunosMatriz[2][c], responsavel: alunosMatriz[3][c], mensagem: alunosMatriz[4][c], frequencia: alunosMatriz[5][c], nota: alunosMatriz[6][c], relatorio: alunosMatriz[7][c], cor: alunosMatriz[8][c] || "", palavras: palavras });
        }
        return JSON.stringify({ cabecalho: cabecalho, metricasProf: metricasProf, cronograma: cronograma, alunos: listaAlunos });
    } catch (e) { throw new Error("Erro interno ao carregar turma."); } // CORREÇÃO FALHA 10
}

function salvarTurmasProf(token, matrizFront) { 
    var cracha = validarTokenSessao(token);
    if(!cracha || cracha.tipo !== 'professor') throw new Error("Acesso Negado.");
    var urls = obterUrlsSeguras(cracha);
    var aba = SpreadsheetApp.openByUrl(urls.urlTurmas).getSheetByName("PROFESSOR"); var colunaNomesTurmas = aba.getRange("F1:F" + aba.getLastRow()).getValues(); 
    for(var i = 0; i < matrizFront.length; i++) { 
        var nomeTurma = matrizFront[i][0]; if(!nomeTurma || nomeTurma.trim() === "" || nomeTurma.toUpperCase().indexOf("TURMA") === -1) continue; 
        var linhaCorreta = -1; for(var r = 0; r < colunaNomesTurmas.length; r++) { if(colunaNomesTurmas[r][0] == nomeTurma) { linhaCorreta = r + 1; break; } } 
        if(linhaCorreta !== -1) { 
            aba.getRange(linhaCorreta, 7, 1, 6).setValues([[ sanitizarEntrada(matrizFront[i][1]), sanitizarEntrada(matrizFront[i][2]), sanitizarEntrada(matrizFront[i][3]), sanitizarEntrada(matrizFront[i][4]), sanitizarEntrada(matrizFront[i][5]), sanitizarEntrada(matrizFront[i][6]) ]]); 
            if (matrizFront[i].length > 10) { aba.getRange(linhaCorreta, 13, 1, 4).setValues([[ sanitizarEntrada(matrizFront[i][7]), sanitizarEntrada(matrizFront[i][8]), sanitizarEntrada(matrizFront[i][9]), sanitizarEntrada(matrizFront[i][10]) ]]); } 
        } 
    } 
}
function salvarMetricasUnicasTurma(token, nomeTurma, numAlunos, desempenho, linkAulas, aulasGravadas) { 
    var cracha = validarTokenSessao(token);
    if(!cracha || cracha.tipo !== 'professor') return false;
    var urls = obterUrlsSeguras(cracha);
    var aba = SpreadsheetApp.openByUrl(urls.urlTurmas).getSheetByName("PROFESSOR"); var colunaNomesTurmas = aba.getRange("F1:F" + aba.getLastRow()).getValues(); 
    for(var r = 0; r < colunaNomesTurmas.length; r++) { 
        if(colunaNomesTurmas[r][0] && colunaNomesTurmas[r][0].toString().trim().toUpperCase() === nomeTurma.trim().toUpperCase()) { 
            aba.getRange(r + 1, 13, 1, 4).setValues([[ sanitizarEntrada(numAlunos), sanitizarEntrada(desempenho), sanitizarEntrada(linkAulas), sanitizarEntrada(aulasGravadas) ]]); 
            return true; 
        } 
    } return false; 
}
function salvarAlunosProf(token, matrizFront) { 
    var cracha = validarTokenSessao(token);
    if(!cracha || cracha.tipo !== 'professor') throw new Error("Acesso Negado.");
    var urls = obterUrlsSeguras(cracha);
    var aba = SpreadsheetApp.openByUrl(urls.urlTurmas).getSheetByName("ALUNOS"); var colunaIDs = aba.getRange("A1:A" + aba.getLastRow()).getValues(); 
    for(var i = 0; i < matrizFront.length; i++) { 
        var idUnico = matrizFront[i][0]; if(!idUnico || idUnico.trim() === "" || idUnico.toUpperCase() === "ID ÚNICO") continue; 
        var linhaCorreta = -1; for(var r = 0; r < colunaIDs.length; r++) { if(colunaIDs[r][0] == idUnico) { linhaCorreta = r + 1; break; } } 
        if(linhaCorreta !== -1) aba.getRange(linhaCorreta, 7, 1, 2).setValues([[ sanitizarEntrada(matrizFront[i][6]), sanitizarEntrada(matrizFront[i][7]) ]]); 
    } 
}
function salvarStatusCronograma(token, nomeTurma, matrizStatus) { 
    var cracha = validarTokenSessao(token);
    if(!cracha || cracha.tipo !== 'professor') return false;
    var urls = obterUrlsSeguras(cracha);
    var aba = SpreadsheetApp.openByUrl(urls.urlTurmas).getSheetByName(nomeTurma); 
    if(matrizStatus && matrizStatus.length > 0) aba.getRange(12, 5, matrizStatus.length, 1).setValues(matrizStatus); 
    return true; 
}
function salvarDadosAlunoTurma(token, nomeTurma, idUnico, freq, nota, relatorio, msg, palavras) {
    var cracha = validarTokenSessao(token);
    if(!cracha || cracha.tipo !== 'professor') throw new Error("Acesso Negado.");
    var urls = obterUrlsSeguras(cracha);
    try {
        var aba = SpreadsheetApp.openByUrl(urls.urlTurmas).getSheetByName(nomeTurma);
        var rangeLinha4 = aba.getRange("H4:AZ4").getDisplayValues()[0];
        var colIndex = -1;
        for (var i = 0; i < rangeLinha4.length; i++) { if (String(rangeLinha4[i]).trim() === String(idUnico).trim()) { colIndex = i + 8; break; } }
        if (colIndex !== -1) {
            aba.getRange(7, colIndex).setValue(sanitizarEntrada(msg)); aba.getRange(8, colIndex).setValue(sanitizarEntrada(freq)); aba.getRange(9, colIndex).setValue(sanitizarEntrada(nota)); aba.getRange(10, colIndex).setValue(sanitizarEntrada(relatorio));
            aba.getRange(12, colIndex, 50, 1).clearContent();
            if (palavras && palavras.length > 0) {
                var palavrasArray = [];
                for (var p = 0; p < 50; p++) { palavrasArray.push([p < palavras.length ? sanitizarEntrada(palavras[p]) : ""]); }
                aba.getRange(12, colIndex, 50, 1).setValues(palavrasArray);
            }
            return true;
        }
        return false;
    } catch(e) { throw new Error("Erro interno."); }
}
function atualizarVariosAlunos(token, listaAtualizacoes) {
    var cracha = validarTokenSessao(token);
    if(!cracha || cracha.tipo !== 'professor') throw new Error("Acesso Negado.");
    var urls = obterUrlsSeguras(cracha);
    var aba = SpreadsheetApp.openByUrl(urls.urlCadastro).getSheetByName('MENSALIDADES');
    for(var i = 0; i < listaAtualizacoes.length; i++) {
        var item = listaAtualizacoes[i];
        aba.getRange(item.linha, 12).setValue(item.status); aba.getRange(item.linha, 13).setValue(item.proxMes); aba.getRange(item.linha, 14, 1, 12).setValues([item.meses]);
    }
    return "Salvo com sucesso!";
}
function obterDadosDosNiveis(token) {
  if(!validarTokenSessao(token)) return JSON.stringify({ sucesso: false, erro: "Acesso Negado." });
  try {
    var ssGeral = SpreadsheetApp.getActiveSpreadsheet();
    var abaGeral = ssGeral.getSheets()[0];
    var urlNiveis = abaGeral.getRange("B8").getValue() || abaGeral.getRange("B7").getValue();
    if (!urlNiveis || String(urlNiveis).indexOf("http") === -1) return JSON.stringify({ sucesso: false, erro: "Link não encontrado." });
    var urlImagem = abaGeral.getRange("B9").getValue() || abaGeral.getRange("B10").getValue();
    var bgImagem = "";
    if (urlImagem && String(urlImagem).indexOf("http") !== -1) {
       var idMatch = String(urlImagem).match(/(?:id=|\/d\/)([\w-]+)/);
       if (idMatch && idMatch[1]) bgImagem = 'https://lh3.googleusercontent.com/d/' + idMatch[1];
       else bgImagem = urlImagem;
    }
    var ssNiveis = SpreadsheetApp.openByUrl(urlNiveis);
    var resultado = { 'A1': [], 'A2': [], 'B1': [], 'B2': [], 'C1': [], 'C2': [] };
    ssNiveis.getSheets().forEach(function(aba) {
       var nome = aba.getName().toUpperCase();
       var nivelMatch = nome.match(/(A1|A2|B1|B2|C1|C2)/);
       if (nivelMatch) {
          var nivelKey = nivelMatch[1];
          var ultimaLinha = aba.getLastRow();
          if (ultimaLinha >= 5) {
             var range = aba.getRange(5, 3, ultimaLinha - 4, 1);
             var textos = range.getDisplayValues(); var richTexts = range.getRichTextValues(); var cores = range.getFontColors();
             for (var i = 0; i < textos.length; i++) {
                var texto = textos[i][0].trim();
                if (!texto) continue;
                resultado[nivelKey].push({ texto: texto, link: richTexts[i][0].getLinkUrl() || null, cor: cores[i][0] || "#ffffff" });
             }
          }
       }
    });
    return JSON.stringify({ sucesso: true, dados: resultado, imagemFundo: bgImagem });
  } catch(e) { return JSON.stringify({ sucesso: false, erro: "Erro interno no servidor." }); } // CORREÇÃO FALHA 10
}
function getDadosCompletosCadastroProf(token, matricula) {
    var cracha = validarTokenSessao(token);
    if(!cracha || cracha.tipo !== 'professor') return JSON.stringify({ sucesso: false, erro: "Acesso Negado." });
    var urls = obterUrlsSeguras(cracha);
    try {
        var ss = SpreadsheetApp.openByUrl(urls.urlCadastro);
        var abas = ss.getSheets(); var aba = null;
        for (var i = 0; i < abas.length; i++) { if (abas[i].getName().toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").includes("cadastro")) { aba = abas[i]; break; } }
        if (!aba) return JSON.stringify({ sucesso: false, erro: "Aba não encontrada." });
        var ultimaLinha = aba.getLastRow();
        if (ultimaLinha < 2) return JSON.stringify({ sucesso: false, erro: "Planilha vazia." });
        var dados = aba.getRange(1, 1, ultimaLinha, 18).getDisplayValues();
        for (var r = 1; r < dados.length; r++) {
            if (String(dados[r][1]).trim() === String(matricula).trim()) {
                var linkContrato = null;
                try { var richText = aba.getRange(r + 1, 18).getRichTextValue(); linkContrato = richText ? richText.getLinkUrl() : null; } catch(e_rt) {}
                var valorContrato = dados[r][17] || "";
                return JSON.stringify({
                    sucesso: true,
                    cadastro: {
                        linha: r + 1, matricula: dados[r][1] || "", dataMatricula: dados[r][2] || "", valorMatricula: dados[r][3] || "",
                        qtdAlunos: dados[r][4] || "", nomeResp: dados[r][5] || "", nomeAluno: dados[r][6] || "", aniversario: dados[r][7] || "",
                        cpf: dados[r][8] || "", email: dados[r][9] || "", telefone: dados[r][10] || "", cep: dados[r][11] || "",
                        endereco: dados[r][12] || "", paisEstado: dados[r][13] || "", instagram: dados[r][14] || "", tiktok: dados[r][15] || "",
                        comoConheceu: dados[r][16] || "", contratoTexto: valorContrato, contratoLink: (linkContrato) ? linkContrato : (valorContrato.startsWith("http") ? valorContrato : "")
                    }
                });
            }
        }
        return JSON.stringify({ sucesso: false, erro: "Matrícula não encontrada." });
    } catch(e) { return JSON.stringify({ sucesso: false, erro: "Erro interno no servidor." }); }
}