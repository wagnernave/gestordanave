// ============================================================
//  GESTOR DA NAVE — Google Apps Script
//  Regras de escrita:
//    database1/database2 → apenas O-S (15-19)
//    novas_aulas          → apenas A,B,C,D,J,M,N,O,P,W (1,2,3,4,10,13,14,15,16,23)
// ============================================================

const PLANILHA_ID = '18dOTG1kjP-xfKQa1FbtBn-WbAZg3NYwXU2dDaELB7n0';
const SNAPSHOT_FOLDER_ID = '18ynJteC2nvWI3AlIjaGqUEWd29jBF1eg';

function doGet(e) {
  const params = e.parameter;
  let result;

  if (params.action === 'getData') {
    result = getData(params.sheet);
  } else if (params.action === 'updateStatus') {
    result = updateStatus(params);
  } else if (params.action === 'addAtividade') {
    result = addAtividade(params);
  } else if (params.action === 'getSnapshotData') {
    result = getSnapshotData(params);
  } else if (params.action === 'snapshot') {
    result = gerarSnapshot();
  } else if (params.action === 'syncFull') {
    const snap = gerarSnapshot();
    const limpo = limparSnapshots();
    result = { ok: true, snapshot: snap, limpeza: limpo };
  } else if (params.action === 'getNovasAtividadesDrive') {
    result = getNovasAtividadesDrive();
  } else if (params.action === 'salvarNovaAtividadeDrive') {
    result = salvarNovaAtividadeDrive(params);
  } else if (params.action === 'excluirNovaAtividadeDrive') {
    result = excluirNovaAtividadeDrive(params);
  } else if (params.action === 'getInstrutores') {
    result = getInstrutores();
  } else if (params.action === 'addInstrutor') {
    result = addInstrutor(params);
  } else if (params.action === 'updateInstrutor') {
    result = updateInstrutor(params);
  } else if (params.action === 'deleteInstrutor') {
    result = deleteInstrutor(params);
  } else {
    result = { error: 'Ação desconhecida' };
  }

  const output = JSON.stringify(result);

  if (params.callback) {
    return ContentService
      .createTextOutput(`${params.callback}(${output})`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  LEITURA — database1 / database2 (colunas A-S)
// ============================================================
function getData(sheetName) {
  const name = sheetName === 'database2' ? 'database2' : 'database1';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Aba "' + name + '" não encontrada.');

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return { sheet: name, rows: [] };

  const data = sheet.getRange(3, 1, lastRow - 2, 19).getValues();
  const rows = [];

  data.forEach(function(row, i) {
    const id = row[0];
    if (!id || String(id).trim() === '') return;

    rows.push({
      _row:          i + 3,
      id:            String(row[0]  || '').trim(),
      san:           String(row[1]  || '').trim(),
      nome:          String(row[2]  || '').trim(),
      tipo:          String(row[3]  || '').trim(),
      eixo:          String(row[4]  || '').trim(),
      vagas:         row[5] || 0,
      parceria:      String(row[6]  || '').trim(),
      turmaAberta:   String(row[7]  || '').trim(),
      publicoAlvo:   String(row[8]  || '').trim(),
      turno:         String(row[9]  || '').trim(),
      data:          formatDate(row[10]),
      horario:       String(row[11] || '').trim(),
      instrutores:   String(row[12] || '').trim(),
      local:         String(row[13] || '').trim(),
      status:        String(row[14] || '').trim() || null,
      novaData:      formatDate(row[15]) || null,
      novoHorario:   String(row[16] || '').trim() || null,
      novoInstrutor: String(row[17] || '').trim() || null,
      novoLocal:     String(row[18] || '').trim() || null,
    });
  });

  return { sheet: name, rows: rows };
}

// ============================================================
//  ESCRITA — database1 / database2: apenas O-S (15-19)
// ============================================================
function updateStatus(body) {
  const sheetName = body.sheet === 'database2' ? 'database2' : 'database1';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Aba "' + sheetName + '" não encontrada.');

  let row;
  if (body.row) {
    row = parseInt(body.row);
  } else if (body.id) {
    const data = sheet.getDataRange().getValues();
    row = -1;
    for (var i = 2; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(body.id).trim()) {
        row = i + 1;
        break;
      }
    }
    if (row < 0) throw new Error('Atividade ID "' + body.id + '" não encontrada na planilha.');
  } else {
    throw new Error('Forneça row ou id.');
  }

  const status = body.status || 'CONFIRMADA';

  sheet.getRange(row, 15).setValue(status); // O = STATUS

  if (status === 'REMARCADA') {
    sheet.getRange(row, 16).setValue(body.novaData      || '');
    sheet.getRange(row, 17).setValue(body.novoHorario   || '');
    sheet.getRange(row, 18).setValue(body.novoInstrutor || '');
    sheet.getRange(row, 19).setValue(body.novoLocal     || '');
  } else {
    sheet.getRange(row, 16, 1, 4).clearContent(); // limpa P-S
  }

  return { ok: true, sheet: sheetName, row: row, status: status };
}

// ============================================================
//  NOVA ATIVIDADE — adiciona linha em novas_aulas
//  Colunas editáveis: A(1)=eixo, B(2)=nome, C(3)=tipo, D(4)=data,
//    J(10)=horario, M(13)=turmaAberta, N(14)=parceria,
//    O(15)=instrutores, P(16)=local, W(23)=publicoAlvo
// ============================================================
function addAtividade(params) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('novas_aulas');
  if (!sheet) throw new Error('Aba "novas_aulas" não encontrada.');

  const lastRow = sheet.getLastRow();
  const nextRow = Math.max(lastRow + 1, 2);

  if (params.eixo)        sheet.getRange(nextRow, 1).setValue(params.eixo);         // A
  if (params.nome)        sheet.getRange(nextRow, 2).setValue(params.nome);         // B
  if (params.tipo)        sheet.getRange(nextRow, 3).setValue(params.tipo);         // C
  if (params.data)        sheet.getRange(nextRow, 4).setValue(params.data);         // D
  if (params.horario)     sheet.getRange(nextRow, 10).setValue(params.horario);     // J
  if (params.turmaAberta) sheet.getRange(nextRow, 13).setValue(params.turmaAberta); // M
  if (params.parceria)    sheet.getRange(nextRow, 14).setValue(params.parceria);    // N
  if (params.instrutores) sheet.getRange(nextRow, 15).setValue(params.instrutores); // O
  if (params.local)       sheet.getRange(nextRow, 16).setValue(params.local);       // P
  if (params.publicoAlvo) sheet.getRange(nextRow, 23).setValue(params.publicoAlvo); // W

  return { ok: true };
}

// ============================================================
//  EDIÇÃO — linha existente em novas_aulas
//  Só altera as colunas editáveis, uma a uma
// ============================================================
function updateNovaAula(params) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('novas_aulas');
  if (!sheet) throw new Error('Aba "novas_aulas" não encontrada.');

  const row = parseInt(params.row);
  if (!row || row < 2) throw new Error('Linha inválida: ' + params.row);

  if (params.hasOwnProperty('eixo'))        sheet.getRange(row, 1).setValue(params.eixo || '');
  if (params.hasOwnProperty('nome'))        sheet.getRange(row, 2).setValue(params.nome || '');
  if (params.hasOwnProperty('tipo'))        sheet.getRange(row, 3).setValue(params.tipo || '');
  if (params.hasOwnProperty('data'))        sheet.getRange(row, 4).setValue(params.data || '');
  if (params.hasOwnProperty('horario'))     sheet.getRange(row, 10).setValue(params.horario || '');
  if (params.hasOwnProperty('turmaAberta')) sheet.getRange(row, 13).setValue(params.turmaAberta || '');
  if (params.hasOwnProperty('parceria'))    sheet.getRange(row, 14).setValue(params.parceria || '');
  if (params.hasOwnProperty('instrutores')) sheet.getRange(row, 15).setValue(params.instrutores || '');
  if (params.hasOwnProperty('local'))       sheet.getRange(row, 16).setValue(params.local || '');
  if (params.hasOwnProperty('publicoAlvo')) sheet.getRange(row, 23).setValue(params.publicoAlvo || '');

  return { ok: true };
}

// ============================================================
//  NOVAS ATIVIDADES — JSON no Drive (pasta oculta .gestor-nave)
// ============================================================

const NA_FOLDER_NAME = '.gestor-nave'; // pasta oculta no Drive
const NA_FILE_NAME = 'novas_atividades.json';

function getNAFolder_() {
  var folders = DriveApp.getFoldersByName(NA_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(NA_FOLDER_NAME);
}

function lerNAJson_() {
  var folder = getNAFolder_();
  var files = folder.getFilesByName(NA_FILE_NAME);
  if (!files.hasNext()) return [];
  var content = files.next().getBlob().getDataAsString();
  var data = JSON.parse(content);
  return Array.isArray(data) ? data : [];
}

function escreverNAJson_(rows) {
  var folder = getNAFolder_();
  var files = folder.getFilesByName(NA_FILE_NAME);
  var json = JSON.stringify(rows, null, 2);
  if (files.hasNext()) {
    files.next().setContent(json);
  } else {
    folder.createFile(NA_FILE_NAME, json);
  }
}

function getNovasAtividadesDrive() {
  return { ok: true, rows: lerNAJson_() };
}

function salvarNovaAtividadeDrive(params) {
  var rows = lerNAJson_();
  var item = {
    id: params.id || String(Date.now()),
    eixo: params.eixo || '',
    nome: params.nome || '',
    tipo: params.tipo || '',
    data: params.data || '',
    horario: params.horario || '',
    turmaAberta: params.turmaAberta || '',
    parceria: params.parceria || '',
    instrutores: params.instrutores || '',
    local: params.local || '',
    publicoAlvo: params.publicoAlvo || '',
    cadastroEm: params.cadastroEm || new Date().toISOString()
  };

  if (params.id) {
    var idx = rows.findIndex(function(r) { return r.id === params.id; });
    if (idx >= 0) {
      rows[idx] = item;
    } else {
      rows.push(item);
    }
  } else {
    rows.push(item);
  }

  escreverNAJson_(rows);
  return { ok: true, id: item.id };
}

function excluirNovaAtividadeDrive(params) {
  if (!params.id) return { ok: false, error: 'id obrigatório' };
  var rows = lerNAJson_();
  var idx = rows.findIndex(function(r) { return r.id === params.id; });
  if (idx >= 0) rows.splice(idx, 1);
  escreverNAJson_(rows);
  return { ok: true };
}

// ============================================================
//  INSTRUTORES E AGENTES — armazenados no servidor (PropertiesService)
//  (fora do cache de snapshots, banco separado)
// ============================================================

var PROP_INSTRUTORES = 'gestor_instrutores';

function lerInstrutoresJson() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(PROP_INSTRUTORES);
    return raw ? JSON.parse(raw) : [];
  } catch (e) {
    return [];
  }
}

function escreverInstrutoresJson(dados) {
  PropertiesService.getScriptProperties().setProperty(PROP_INSTRUTORES, JSON.stringify(dados));
}

function getInstrutores() {
  var dados = lerInstrutoresJson();
  return { ok: true, dados: dados };
}

function addInstrutor(params) {
  var dados = lerInstrutoresJson();
  dados.push({
    id: params.id || String(Date.now()),
    nome: params.nome || '',
    nickname: params.nickname || '',
    tipo: params.tipo || '',
    nascimento: params.nascimento || '',
    idade: params.idade || '',
    telefone: params.telefone || '',
    email: params.email || '',
    endereco: params.endereco || '',
    obs: params.obs || '',
    cadastroEm: params.cadastroEm || new Date().toISOString(),
    ativo: params.hasOwnProperty('ativo') ? params.ativo : true,
  });
  escreverInstrutoresJson(dados);
  return { ok: true, total: dados.length };
}

function updateInstrutor(params) {
  var dados = lerInstrutoresJson();
  var idx = -1;
  for (var i = 0; i < dados.length; i++) {
    if (String(dados[i].id) === String(params.id)) { idx = i; break; }
  }
  if (idx < 0) throw new Error('Instrutor com id "' + params.id + '" não encontrado.');
  var campos = ['id','nome','nickname','tipo','nascimento','idade','telefone','email','endereco','obs','cadastroEm','ativo'];
  for (var j = 0; j < campos.length; j++) {
    var c = campos[j];
    if (params.hasOwnProperty(c)) dados[idx][c] = params[c];
  }
  escreverInstrutoresJson(dados);
  return { ok: true };
}

function deleteInstrutor(params) {
  var dados = lerInstrutoresJson();
  var encontrou = false;
  for (var i = dados.length - 1; i >= 0; i--) {
    if (String(dados[i].id) === String(params.id)) {
      dados.splice(i, 1);
      encontrou = true;
      break;
    }
  }
  if (!encontrou) throw new Error('Instrutor com id "' + params.id + '" não encontrado.');
  escreverInstrutoresJson(dados);
  return { ok: true, total: dados.length };
}

// ============================================================
//  CONFIG
// ============================================================
function getConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config');
  if (!sheet) throw new Error('Aba "config" não encontrada.');
  const v = sheet.getRange(2, 1, 1, 6).getValues()[0];
  return {
    usuario: String(v[0] || 'admin').trim(),
    senha:   String(v[1] || 'S3nh@987').trim(),
    acessar: String(v[2] || 'Conectar').trim(),
    dbUrl:   String(v[5] || '').trim(),
  };
}

function updateConfig(body) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config');
  if (!sheet) throw new Error('Aba "config" não encontrada.');
  const cur = sheet.getRange(2, 1, 1, 6).getValues()[0];
  if (body.usuario) cur[0] = body.usuario;
  if (body.senha)   cur[1] = body.senha;
  if (body.acessar) cur[2] = body.acessar;
  if (body.dbUrl)   cur[5] = body.dbUrl;
  sheet.getRange(2, 1, 1, 6).setValues([cur]);
  return { ok: true };
}

// ============================================================
//  UTILITÁRIOS
// ============================================================
function formatDate(value) {
  if (!value) return null;
  if (value instanceof Date) {
    var d = String(value.getDate()).padStart(2, '0');
    var m = String(value.getMonth() + 1).padStart(2, '0');
    var y = value.getFullYear();
    var h = value.getHours();
    var min = value.getMinutes();
    if (h === 0 && min === 0) {
      return d + '/' + m + '/' + y;
    }
    return d + '/' + m + '/' + y + ' ' + String(h).padStart(2, '0') + ':' + String(min).padStart(2, '0');
  }
  return String(value).trim() || null;
}

// ============================================================
//  SNAPSHOT — exporta todas as abas para CSV no Google Drive
//  Uso:       gerarSnapshot()                  (manual ou gatilho)
//  Uso (web): ?action=snapshot                 (via doGet)
//  Gatilho:   instalarTriggerSnapshots()       (uma vez no console)
// ============================================================

const SNAPSHOT_FOLDER_NAME = 'Snapshots Gestor da Nave';

function gerarSnapshot() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const pasta = obterPastaSnapshots();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
  const tabs = [
    { nome:'database1', colInicio:3, colFim:22, cabecalho:'A2:V2' },
    { nome:'database2', colInicio:3, colFim:22, cabecalho:'A2:V2' },
    { nome:'novas_aulas', colInicio:2, colFim:23, cabecalho:'A1:W1' },
  ];
  const resultados = [];

  tabs.forEach(function(t) {
    const sheet = ss.getSheetByName(t.nome);
    if (!sheet) {
      resultados.push({ aba: t.nome, erro: 'Aba não encontrada' });
      return;
    }
    var dados = sheet.getDataRange().getValues();
    var inicio = t.colInicio - 1; // converte pra índice 0-based
    var linhas = [];
    for (var i = inicio; i < dados.length; i++) {
      var row = dados[i];
      if (row.slice(0, t.colFim).some(function(c) { return c !== '' && c !== null && c !== undefined; })) {
        linhas.push(row.slice(0, t.colFim));
      }
    }
    if (linhas.length === 0) {
      resultados.push({ aba: t.nome, linhas: 0 });
      return;
    }
    var headers = sheet.getRange(t.cabecalho).getValues()[0].slice(0, t.colFim);
    var csv = gerarCSV(headers, linhas);
    var arquivo = pasta.createFile(t.nome + '_' + timestamp + '.csv', csv, MimeType.CSV);
    resultados.push({ aba: t.nome, linhas: linhas.length, arquivo: arquivo.getName() });
  });

  return { ok: true, timestamp: timestamp, resultados: resultados };
}

function obterPastaSnapshots() {
  return DriveApp.getFolderById(SNAPSHOT_FOLDER_ID);
}

function gerarCSV(headers, rows) {
  var saida = '';
  if (headers && headers.length) {
    saida += headers.join('|') + '\r\n';
  }
  if (rows && rows.length) {
    rows.forEach(function(row) {
      if (row && row.length) {
        saida += row.map(function(v) {
          var s;
          if (v instanceof Date) {
            s = formatDate(v) || '';
          } else {
            s = String(v || '');
          }
          return s.indexOf('|') >= 0 ? '"' + s.replace(/"/g, '""') + '"' : s;
        }).join('|') + '\r\n';
      }
    });
  }
  return saida;
}

function citarCSV(valor) {
  var v = String(valor || '');
  if (v.indexOf(',') >= 0 || v.indexOf('"') >= 0 || v.indexOf('\n') >= 0) {
    v = '"' + v.replace(/"/g, '""') + '"';
  }
  return v;
}

// ============================================================
//  LEITURA DE SNAPSHOT — lê o último CSV de uma aba+mês e retorna JSON
//  Uso (web): ?action=getSnapshotData&aba=database1&mes=2026-06
// ============================================================

function getSnapshotData(params) {
  const aba = params.aba || 'database1';
  const mes = params.mes || '';

  if (!mes) throw new Error('Parâmetro "mes" obrigatório (formato YYYY-MM)');

  const pasta = obterPastaSnapshots();
  const arquivos = pasta.getFilesByType(MimeType.CSV);
  const candidatos = [];

  while (arquivos.hasNext()) {
    var f = arquivos.next();
    var nome = f.getName();
    var regex = new RegExp('^' + aba + '_' + mes + '\\-\\d{2}_\\d{2}\\-\\d{2}\\.csv$');
    if (regex.test(nome)) {
      candidatos.push({ arquivo: f, nome: nome, data: f.getDateCreated() });
    }
  }

  if (candidatos.length === 0) {
    return { sheet: aba, mes: mes, rows: [] };
  }

  candidatos.sort(function(a, b) { return b.data - a.data; }); // newest first
  var ultimo = candidatos[0];
  var conteudo = ultimo.arquivo.getBlob().getDataAsString();

  var linhas = conteudo.split('\n');
  var dados = [];
  var cabecalho = true;

  for (var i = 0; i < linhas.length; i++) {
    var linha = linhas[i].trim();
    if (!linha) continue;
    if (cabecalho) { cabecalho = false; continue; } // skip header row
    dados.push(parsePipedLine(linha));
  }

  return { sheet: aba, mes: mes, rows: dados, arquivo: ultimo.nome };
}

function parsePipedLine(line) {
  var cols = [];
  var current = '';
  var inQuotes = false;

  for (var i = 0; i < line.length; i++) {
    var c = line[i];
    if (inQuotes) {
      if (c === '"') {
        if (i + 1 < line.length && line[i + 1] === '"') {
          current += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        current += c;
      }
    } else {
      if (c === '"') {
        inQuotes = true;
      } else if (c === '|') {
        cols.push(current);
        current = '';
      } else {
        current += c;
      }
    }
  }
  cols.push(current);
  return cols;
}

// ============================================================
//  LIMPEZA — elimina snapshots duplicados, mantém ao menos 1/mês
//  Uso:       limparSnapshots()            (manual ou gatilho mensal)
//  Lógica:    agrupa por aba+mês, compara conteúdo (MD5),
//             deleta consecutivos idênticos, preserva o último.
// ============================================================

function limparSnapshots() {
  const pasta = obterPastaSnapshots();
  const arquivos = pasta.getFilesByType(MimeType.CSV);
  var lista = [];

  while (arquivos.hasNext()) {
    var f = arquivos.next();
    var nome = f.getName();
    var partes = nome.match(/^(.+)_(\d{4}-\d{2})-\d{2}_\d{2}-\d{2}\.csv$/);
    if (!partes) continue;
    lista.push({
      arquivo: f,
      aba: partes[1],
      mes: partes[2],
      data: f.getDateCreated(),
    });
  }

  if (lista.length === 0) return { ok: true, removidos: 0, mantidos: 0 };

  var grupos = {};
  lista.forEach(function(item) {
    var chave = item.aba + '|' + item.mes;
    if (!grupos[chave]) grupos[chave] = [];
    grupos[chave].push(item);
  });

  var removidos = 0;
  var mantidos = 0;

  Object.keys(grupos).forEach(function(chave) {
    var items = grupos[chave];
    // Ordena do mais antigo pro mais novo
    items.sort(function(a, b) { return a.data - b.data; });
    // Mantém apenas o último (mais recente)
    var ultimo = items[items.length - 1];

    items.forEach(function(item) {
      if (item === ultimo) {
        mantidos++;
      } else {
        item.arquivo.setTrashed(true);
        removidos++;
      }
    });
  });

  return { ok: true, removidos: removidos, mantidos: mantidos };
}

function instalarTriggerSnapshots() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    var nome = t.getHandlerFunction();
    if (nome === 'gerarSnapshot' || nome === 'limparSnapshots') {
      ScriptApp.deleteTrigger(t);
    }
  });

  var horarios = [8, 12, 16, 20];
  horarios.forEach(function(h) {
    ScriptApp.newTrigger('gerarSnapshot')
      .timeBased()
      .everyDays(1)
      .atHour(h)
      .nearMinute(0)
      .create();
  });

  // Limpeza mensal no dia 1 de cada mês às 6h
  ScriptApp.newTrigger('limparSnapshots')
    .timeBased()
    .onMonthDay(1)
    .atHour(6)
    .create();

  return { ok: true, triggers: horarios.length + 1 };
}
