// FREQUÊNCIA — 4ª Formação de Gestores e Coordenadores (27/11/2026)
// SME Tupanatinga-PE · página de confirmação de presença
const ABA = 'LISTA';
const CAB = ['Nº', 'ESCOLA', 'NOME', 'FUNÇÃO', 'RESPOSTA', 'MOTIVO', 'DATA DA RESPOSTA', 'ASSINATURA'];

// Rodar uma vez: prepara a aba LISTA com o cabeçalho.
function preparar() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(ABA);
  if (!sh) { sh = ss.getSheets()[0]; sh.setName(ABA); }
  sh.getRange(1, 1, 1, CAB.length).setValues([CAB])
    .setFontWeight('bold').setBackground('#282C62').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, 1, 45); sh.setColumnWidths(2, 1, 260); sh.setColumnWidths(3, 1, 280);
  sh.setColumnWidths(4, 1, 170); sh.setColumnWidths(5, 1, 100); sh.setColumnWidths(6, 1, 220);
  sh.setColumnWidths(7, 1, 150); sh.setColumnWidths(8, 1, 220);
}

function norm_(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, ' ').trim().toUpperCase();
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// A página pede a lista de escolas e nomes (só escola, nome e função).
function doGet() {
  const sh = SpreadsheetApp.getActive().getSheetByName(ABA);
  const v = sh.getDataRange().getValues();
  const lista = [];
  for (let i = 1; i < v.length; i++) {
    if (v[i][1] && v[i][2]) lista.push({ escola: String(v[i][1]), nome: String(v[i][2]), funcao: String(v[i][3] || '') });
  }
  return json_({ ok: true, lista: lista });
}

// A página envia a resposta: grava na linha da pessoa ou cria linha nova.
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const escola = String(d.escola || '').slice(0, 200);
    const nome = String(d.nome || '').slice(0, 200);
    const funcao = String(d.funcao || '').slice(0, 120);
    const resposta = d.resposta === 'VAI' ? 'VAI' : 'NÃO VAI';
    const motivo = resposta === 'VAI' ? '' : String(d.motivo || '').slice(0, 300);
    if (!escola || !nome) return json_({ ok: false, erro: 'Faltou escola ou nome.' });

    const sh = SpreadsheetApp.getActive().getSheetByName(ABA);
    const v = sh.getDataRange().getValues();
    let linha = -1;
    for (let i = 1; i < v.length; i++) {
      if (norm_(v[i][1]) === norm_(escola) && norm_(v[i][2]) === norm_(nome)) { linha = i + 1; break; }
    }
    if (linha === -1) {
      linha = sh.getLastRow() + 1;
      sh.getRange(linha, 1, 1, 4).setValues([[linha - 1, escola, nome, funcao]]);
      sh.getRange(linha, 1, 1, 8).setBackground('#FFF9C4'); // amarelo = nome que não estava na lista
    } else if (funcao) {
      sh.getRange(linha, 4).setValue(funcao);
    }
    sh.getRange(linha, 5, 1, 3).setValues([[resposta, motivo, new Date()]]);
    sh.getRange(linha, 7).setNumberFormat('dd/MM/yyyy HH:mm');
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}
