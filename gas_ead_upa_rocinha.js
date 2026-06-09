// ============================================================
// Google Apps Script — EAD UPA Rocinha
// Aba: "Tarefas" | Cabeçalho na linha 4 | Dados a partir linha 5
// Colunas: B=Nome C=CPF D=Conselho E=Setor F=Curso G=Status
//          H=Tempo I=DataInício J=DataConclusão K=CertAssinado
//          L=AssinaturaColabBase64 M=% N=CorCurso O=Fases P=CódCert
// ============================================================

const SHEET_NAME = 'Tarefas';
const DATA_START  = 5;   // primeira linha de dados (linha 4 = cabeçalho)

const C = {
  nome:        2,   // B
  cpf:         3,   // C
  conselho:    4,   // D
  setor:       5,   // E
  curso:       6,   // F
  status:      7,   // G
  tempo:       8,   // H
  dataInicio:  9,   // I
  dataConcl:   10,  // J
  certStatus:  11,  // K
  sigColab:    12,  // L  ← assinatura base64
  pct:         13,  // M
  corCurso:    14,  // N
  fases:       15,  // O
  certCode:    16   // P
};

// ── Helpers CORS ───────────────────────────────────────────────
function corsOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET: buscar pendentes ──────────────────────────────────────
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'getPending') return getPendingCerts();
  return corsOutput({status:'ok', msg:'ready'});
}

// ── POST: roteador principal ───────────────────────────────────
// O HTML envia Content-Type: text/plain para evitar preflight CORS
// O GAS lê normalmente via e.postData.contents
function doPost(e) {
  try {
    const raw = e.postData ? e.postData.contents : '{}';
    const p   = JSON.parse(raw);
    const action = p.action || '';

    if (action === 'savePendingCert')   return corsOutput(savePendingCert(p));
    if (action === 'updateCertificate') return corsOutput(markAsSigned(p));
    if (action === 'updateStatus')      return corsOutput(updateStatus(p));
    return corsOutput(registerInitial(p));

  } catch(err) {
    return corsOutput({status:'error', msg: err.toString()});
  }
}

// ── Registro inicial ───────────────────────────────────────────
function registerInitial(p) {
  getSheet().appendRow(buildRow(p, false));
  return {status:'ok', msg:'registered'};
}

// ── Salva pendente + assinatura base64 na col L ────────────────
function savePendingCert(p) {
  const sheet   = getSheet();
  const rowIdx  = findRow(sheet, p.cpf, p.curso);

  if (rowIdx > 0) {
    sheet.getRange(rowIdx, C.status).setValue('Concluído');
    sheet.getRange(rowIdx, C.tempo).setValue(p.tempoTotal  || '—');
    sheet.getRange(rowIdx, C.dataConcl).setValue(p.dataConclusao);
    sheet.getRange(rowIdx, C.certStatus).setValue('PENDENTE');
    sheet.getRange(rowIdx, C.sigColab).setValue(p.sigColabData);
    sheet.getRange(rowIdx, C.pct).setValue(p.pct);
    sheet.getRange(rowIdx, C.corCurso).setValue(p.courseColor);
    sheet.getRange(rowIdx, C.fases).setValue(p.phases);
    sheet.getRange(rowIdx, C.certCode).setValue(p.certCode);
  } else {
    getSheet().appendRow(buildRow(p, true));
  }
  return {status:'ok', msg:'saved'};
}

// ── Retorna pendentes (col K = PENDENTE e col L preenchida) ────
function getPendingCerts() {
  const sheet   = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) return corsOutput({status:'ok', pendentes:[]});

  const numRows = lastRow - DATA_START + 1;
  const data    = sheet.getRange(DATA_START, 1, numRows, 17).getValues();
  const pendentes = [];

  data.forEach(row => {
    const kVal = String(row[C.certStatus - 1]);
    const lVal = String(row[C.sigColab   - 1]);
    if (kVal === 'PENDENTE' && lVal && lVal.startsWith('data:image')) {
      pendentes.push({
        nome:          row[C.nome       - 1],
        cpf:           row[C.cpf        - 1],
        council:       row[C.conselho   - 1],
        cat:           row[C.setor      - 1],
        curso:         row[C.curso      - 1],
        duracao:       row[C.tempo      - 1],
        dataInicio:    row[C.dataInicio - 1],
        dataConclusao: row[C.dataConcl  - 1],
        certCode:      row[C.certCode   - 1],
        sigColabData:  row[C.sigColab   - 1],
        pct:           row[C.pct        - 1],
        courseColor:   row[C.corCurso   - 1],
        phases:        row[C.fases      - 1]
      });
    }
  });

  return corsOutput({status:'ok', pendentes: pendentes});
}

// ── Marca ASSINADO e limpa col L ───────────────────────────────
function markAsSigned(p) {
  const sheet  = getSheet();
  const rowIdx = findRowByCode(sheet, p.certCode) || findRow(sheet, p.cpf, p.curso);
  if (rowIdx > 0) {
    sheet.getRange(rowIdx, C.certStatus).setValue('ASSINADO');
    sheet.getRange(rowIdx, C.sigColab).setValue('');
  }
  return {status:'ok', msg:'signed'};
}

// ── updateStatus (compatibilidade) ────────────────────────────
function updateStatus(p) {
  const sheet  = getSheet();
  const rowIdx = findRow(sheet, p.cpf, p.curso);
  if (rowIdx > 0) {
    if (p.status)        sheet.getRange(rowIdx, C.status).setValue(p.status);
    if (p.tempoTotal)    sheet.getRange(rowIdx, C.tempo).setValue(p.tempoTotal);
    if (p.dataConclusao) sheet.getRange(rowIdx, C.dataConcl).setValue(p.dataConclusao);
    if (p.certAssinado)  sheet.getRange(rowIdx, C.certStatus).setValue(p.certAssinado);
  }
  return {status:'ok', msg:'updated'};
}

// ── Helpers ────────────────────────────────────────────────────
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}

function findRow(sheet, cpf, curso) {
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) return -1;
  const n = lastRow - DATA_START + 1;
  const cpfs   = sheet.getRange(DATA_START, C.cpf,  n, 1).getValues();
  const cursos = sheet.getRange(DATA_START, C.curso, n, 1).getValues();
  for (let i = 0; i < cpfs.length; i++) {
    if (String(cpfs[i][0]) === String(cpf) &&
        String(cursos[i][0]) === String(curso)) return DATA_START + i;
  }
  return -1;
}

function findRowByCode(sheet, certCode) {
  if (!certCode) return -1;
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) return -1;
  const n     = lastRow - DATA_START + 1;
  const codes = sheet.getRange(DATA_START, C.certCode, n, 1).getValues();
  for (let i = 0; i < codes.length; i++) {
    if (String(codes[i][0]) === String(certCode)) return DATA_START + i;
  }
  return -1;
}

function buildRow(p, isPending) {
  return [
    '',                          // A (vazio — dados começam em B)
    p.nome        || '',         // B
    p.cpf         || '',         // C
    p.conselho    || '—',        // D
    p.setor       || '',         // E
    p.curso       || '',         // F
    isPending ? 'Concluído' : (p.status || 'Em andamento'), // G
    p.tempoTotal  || '—',        // H
    p.dataInicio  || '',         // I
    isPending ? (p.dataConclusao || '') : '', // J
    isPending ? 'PENDENTE' : (p.certAssinado || 'NÃO'), // K
    isPending ? (p.sigColabData  || '') : '', // L
    isPending ? (p.pct           || '') : '', // M
    isPending ? (p.courseColor   || '') : '', // N
    isPending ? (p.phases        || '') : '', // O
    isPending ? (p.certCode      || '') : ''  // P
  ];
}
