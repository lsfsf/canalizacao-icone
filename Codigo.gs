/**
 * @OnlyCurrentDoc
 *
 * Backend da app "Canalização • Luís & Paulo".
 * Os números do acerto vêm SEMPRE da aba "Acerto de Contas (Luís & Paulo)",
 * para a app e a folha nunca darem valores diferentes.
 */
const SHEET_TRABALHOS = 'Trabalhos';
const SHEET_DESPESAS = 'Despesas e Portagens';
const SHEET_FUNDO = 'Fundo de Caixa';
const SHEET_ACERTO = 'Acerto de Contas (Luís & Paulo)';
const SHEET_REGISTO = 'Registo de Alterações';
const SHEET_UTILIZADORES = 'Utilizadores';
const SHEET_FECHOS = 'Acertos Fechados';

const PRIMEIRA_LINHA_DADOS = 5;

// Ícone da app (assets/icone-512.png), publicado no repositório público lsfsf/canalizacao-icone.
// O Google Drive não serve: devolve uma página HTML e o Apps Script rejeita-o como "tipo não suportado".
const ICONE_URL = 'https://raw.githubusercontent.com/lsfsf/canalizacao-icone/main/icone-512.png';

// Células da aba "Acerto de Contas" (coluna B = base cobrada, C = base após cobrar tudo).
const CEL_ACERTO = {
  cobrado: 'B6', pendente: 'B7', faturacao: 'B8',
  dinheiroMao: 'B15', despLuis: 'B16', despPaulo: 'B17', sobra: 'B18',
  reserva: 'B19', aDividir: 'B20', metade: 'B21', totalLuis: 'B22', totalPaulo: 'B23',
  conferencia: 'D24',
  metadeTotal: 'C21', totalLuisTotal: 'C22', totalPauloTotal: 'C23'
};
// Células da aba "Fundo de Caixa".
const CEL_FUNDO = { inicial: 'C6', gastoFundo: 'C7', troco: 'C8', saldoFisico: 'C9' };

function doGet() {
  const nomesPossiveis = ['Index', 'index'];
  let html = null;
  for (let i = 0; i < nomesPossiveis.length; i++) {
    try {
      html = HtmlService.createHtmlOutputFromFile(nomesPossiveis[i]);
      break;
    } catch (err) {}
  }

  if (!html) {
    return HtmlService.createHtmlOutput('<h3>Erro: Cria o ficheiro HTML na coluna da esquerda com o nome Index</h3>');
  }

  html
    .setTitle('Canalização • Luís & Paulo')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  // O ícone é opcional: se o Apps Script não aceitar o URL, a app abre na mesma.
  try {
    if (ICONE_URL) html.setFaviconUrl(ICONE_URL);
  } catch (err) {
    console.warn('Ícone não aplicado: ' + err);
  }
  return html;
}

// A página (doGet) não contém dados: os dados só saem pelas duas funções abaixo, que exigem o código.
function apiLerDados(codigo) {
  const acesso = verificarAcesso(SpreadsheetApp.getActiveSpreadsheet(), codigo);
  return lerDadosCompletos(acesso);
}

function apiGuardarDados(payload) {
  const acesso = verificarAcesso(SpreadsheetApp.getActiveSpreadsheet(), payload && payload.codigo);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return guardarNaFolha(payload, acesso);
  } finally {
    lock.releaseLock();
  }
}

// ---------- Controlo de acesso ----------
// A aba "Utilizadores" (coluna A = nome, coluna B = código de acesso) diz quem pode usar a app.
// Vazia: a app funciona sem proteção (como antes). Com códigos: só entra quem tiver um código
// válido, e o nome de quem altera vem do código. Não exige que os sócios tenham acesso à folha:
// a implementação fica "Executar como: Eu" e o código é pedido uma vez em cada telemóvel.
// Os erros que começam por "CODIGO:" fazem a app pedir o código; os "ACESSO:" só mostram a mensagem.

const MAX_FALHAS_CODIGO = 8;      // códigos errados seguidos antes de bloquear
const BLOQUEIO_SEGUNDOS = 600;    // duração do bloqueio (10 minutos)

function lerUtilizadores(ss) {
  const sh = ss.getSheetByName(SHEET_UTILIZADORES);
  if (!sh) return [];
  const ult = sh.getLastRow();
  if (ult < 2) return [];
  return sh.getRange(2, 1, ult - 1, 2).getValues()
    .map(r => ({ nome: String(r[0]).trim(), codigo: String(r[1]).trim() }))
    .filter(u => u.nome || u.codigo);
}

function verificarAcesso(ss, codigo) {
  if (!ss) throw new Error('ACESSO: Não foi possível abrir a folha.');
  const linhas = lerUtilizadores(ss);
  if (!linhas.length) return { protegido: false, nome: '' };

  const validos = linhas.filter(u => u.nome && u.codigo.length >= 6);
  if (validos.length !== linhas.length) {
    throw new Error('ACESSO: A aba "Utilizadores" está mal preenchida: cada linha precisa de um nome e de um código com pelo menos 6 caracteres.');
  }
  if (new Set(validos.map(u => u.codigo)).size !== validos.length) {
    throw new Error('ACESSO: Dois sócios têm o mesmo código na aba "Utilizadores".');
  }

  const cache = CacheService.getScriptCache();
  const falhas = Number(cache.get('falhas_codigo') || 0);
  if (falhas >= MAX_FALHAS_CODIGO) {
    throw new Error('CODIGO: Demasiadas tentativas erradas. Tenta outra vez daqui a 10 minutos.');
  }
  const c = String(codigo == null ? '' : codigo).trim();
  const u = c ? validos.find(x => x.codigo === c) : null;
  if (!u) {
    if (c) cache.put('falhas_codigo', String(falhas + 1), BLOQUEIO_SEGUNDOS);
    throw new Error('CODIGO: ' + (c ? 'Código incorreto.' : 'Escreve o teu código de acesso.'));
  }
  return { protegido: true, nome: u.nome };
}

function encontrarAba(ss, nomesPossiveis) {
  const abas = ss.getSheets();
  for (let n of nomesPossiveis) {
    const exata = ss.getSheetByName(n);
    if (exata) return exata;
  }
  for (let aba of abas) {
    const nomeLimpo = aba.getName().toLowerCase();
    for (let n of nomesPossiveis) {
      if (nomeLimpo.indexOf(n.toLowerCase()) !== -1) return aba;
    }
  }
  return null;
}

function abaObrigatoria(ss, nomesPossiveis) {
  const sh = encontrarAba(ss, nomesPossiveis);
  if (!sh) throw new Error('Aba "' + nomesPossiveis[0] + '" não encontrada na folha.');
  return sh;
}

const ABAS = {
  trabalhos: [SHEET_TRABALHOS, 'Trabalhos_Obras', 'Obras'],
  despesas: [SHEET_DESPESAS, 'Despesas_e_Portagens', 'Despesas'],
  fundo: [SHEET_FUNDO, 'Fundo_Caixa_e_Acerto'],
  acerto: [SHEET_ACERTO, 'Acerto de Contas', 'Acerto']
};

function guardarNaFolha(payload, acesso) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  let resultado = null;
  let acaoTxt = '', detalhe = '';
  const quem = (acesso && acesso.protegido) ? acesso.nome : identificarUtilizador(payload.utilizador);

  if (payload.action === 'addTrabalho') {
    const sh = abaObrigatoria(ss, ABAS.trabalhos);
    const linha = obterLinhaLivre(sh, 10, [1]);
    const novoId = gerarProximoIdObra(sh);
    const temValor = payload.valor !== null && payload.valor !== undefined && payload.valor !== '';
    const estado = temValor ? (payload.estadoPagamento || 'Pendente') : 'Pendente';
    const dataIni = normalizarData(payload.dataInicio, hoje);
    const dataRec = estado === 'Recebido em Caixa' ? normalizarData(payload.dataRecebimento, dataIni) : '';

    sh.getRange(linha, 1, 1, 10).setValues([[
      novoId,
      dataIni,
      payload.nomeObra || '',
      payload.localidade || '',
      payload.nomeCliente || '',
      payload.descricao || '',
      temValor ? Number(payload.valor) : '',
      estado,
      dataRec,
      'Aberto'
    ]]);
    acaoTxt = 'Nova obra';
    detalhe = novoId + ' • ' + (payload.nomeObra || payload.localidade || payload.nomeCliente || '(sem nome)') +
      ' — ' + (temValor ? fmtEur(payload.valor) : 'sem valor') + ' (' + estado + ')' +
      (payload.descricao ? ' — ' + String(payload.descricao).replace(/\n/g, ' | ') : '') + sufixoData(dataIni, hoje);
  }
  else if (payload.action === 'addDespesa') {
    // Regra do negócio: portagens são sempre Via Verde do Luís, e Via Verde só existe para portagens.
    let categoria = payload.categoria || 'Material Stock';
    let quemPagou = payload.quemPagou || 'Fundo Caixa';
    if (categoria === 'Portagens') quemPagou = 'Luís - Via Verde';
    else if (quemPagou === 'Luís - Via Verde') throw new Error('"Luís - Via Verde" só pode ser usado em Portagens.');

    const dataDesp = normalizarData(payload.data, hoje);
    const sh = abaObrigatoria(ss, ABAS.despesas);
    const linha = obterLinhaLivre(sh, 7, [1, 4]);

    sh.getRange(linha, 1, 1, 7).setValues([[
      dataDesp,
      categoria,
      payload.descricao || '',
      Number(payload.valor || 0),
      quemPagou,
      'Aberto',
      payload.notas || ''
    ]]);
    acaoTxt = 'Nova despesa';
    detalhe = categoria + ' — ' + (payload.descricao || '(sem descrição)') + ' — ' + fmtEur(payload.valor) + ' (' + quemPagou + ')' + sufixoData(dataDesp, hoje);
  }
  else if (payload.action === 'marcarObraRecebida') {
    const sh = abaObrigatoria(ss, ABAS.trabalhos);
    const dados = sh.getDataRange().getValues();
    for (let i = PRIMEIRA_LINHA_DADOS - 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === String(payload.idObra).trim()) {
        sh.getRange(i + 1, 8).setValue('Recebido em Caixa');
        const dataRecebida = normalizarData(payload.dataRecebimento, hoje);
        sh.getRange(i + 1, 9).setValue(dataRecebida);
        acaoTxt = 'Obra recebida';
        detalhe = String(dados[i][0]) + ' • ' + (dados[i][2] || dados[i][3] || dados[i][4] || '') + ' — ' + fmtEur(dados[i][6]) + ' recebido em caixa' + sufixoData(dataRecebida, hoje);
        break;
      }
    }
  }
  else if (payload.action === 'setValorObra') {
    const sh = abaObrigatoria(ss, ABAS.trabalhos);
    const dados = sh.getDataRange().getValues();
    let encontrada = false;
    for (let i = PRIMEIRA_LINHA_DADOS - 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === String(payload.idObra).trim()) {
        if (String(dados[i][9]) === 'Fechado') throw new Error('A obra ' + payload.idObra + ' já está fechada.');
        const temValor = payload.valor !== null && payload.valor !== undefined && payload.valor !== '';
        if (!temValor && String(dados[i][7]) === 'Recebido em Caixa') {
          throw new Error('A obra ' + payload.idObra + ' já está recebida: precisa de valor.');
        }
        const antes = dados[i][6] === '' ? 'sem valor' : fmtEur(dados[i][6]);
        sh.getRange(i + 1, 7).setValue(temValor ? Number(payload.valor) : '');
        if (payload.descricao !== undefined) sh.getRange(i + 1, 6).setValue(String(payload.descricao));
        acaoTxt = 'Valor da obra alterado';
        detalhe = String(dados[i][0]) + ' • ' + (dados[i][2] || dados[i][3] || dados[i][4] || '') + ': ' + antes + ' → ' +
          (temValor ? fmtEur(payload.valor) : 'sem valor') + (payload.descricao !== undefined ? ' (valores por trabalho atualizados)' : '');
        encontrada = true;
        break;
      }
    }
    if (!encontrada) throw new Error('Obra ' + payload.idObra + ' não encontrada.');
  }
  else if (payload.action === 'setFundoInicial') {
    const cel = abaObrigatoria(ss, ABAS.fundo).getRange(CEL_FUNDO.inicial);
    const antes = fmtEur(cel.getValue());
    cel.setValue(Number(payload.valor || 0));
    acaoTxt = 'Fundo inicial alterado';
    detalhe = antes + ' → ' + fmtEur(payload.valor);
  }
  else if (payload.action === 'setReserva') {
    const cel = abaObrigatoria(ss, ABAS.acerto).getRange(CEL_ACERTO.reserva);
    const antes = fmtEur(cel.getValue());
    cel.setValue(Number(payload.valor || 0));
    acaoTxt = 'Valor que fica no envelope alterado';
    detalhe = antes + ' → ' + fmtEur(payload.valor);
  }
  else if (payload.action === 'fecharCiclo') {
    resultado = fecharCiclo(ss, hoje, quem);
    acaoTxt = 'Ciclo fechado';
    detalhe = 'Luís ' + fmtEur(resultado.luis) + ' • Paulo ' + fmtEur(resultado.paulo) + ' • fica no envelope ' + fmtEur(resultado.reserva) +
      ' • ' + resultado.obrasFechadas + ' obras fechadas, ' + resultado.despesasLiquidadas + ' despesas liquidadas';
  }
  else if (payload.action === 'editarDespesa') {
    const sh = abaObrigatoria(ss, ABAS.despesas);
    const fim = encontrarLinhaAntesDoTotal(sh, 1, 'Total') - 1;
    const linha = Number(payload.linha);
    if (!(linha >= PRIMEIRA_LINHA_DADOS && linha <= fim)) throw new Error('Despesa não encontrada. Atualiza a app e tenta de novo.');
    const atual = sh.getRange(linha, 1, 1, 7).getValues()[0];
    const disp = sh.getRange(linha, 1, 1, 7).getDisplayValues()[0];
    const esp = payload.esperado || {};
    // Confirma que a linha ainda é a despesa que a pessoa viu (alguém pode ter mexido entretanto).
    const igual = String(disp[0]) === String(esp.data) && String(disp[1]) === String(esp.categoria) &&
      String(disp[2]) === String(esp.descricao) && String(disp[4]) === String(esp.origem) &&
      Math.abs(Number(atual[3] || 0) - Number(esp.valor || 0)) < 0.005;
    if (!igual) throw new Error('Esta despesa mudou entretanto. Atualiza a app e tenta de novo.');
    if (String(atual[5]) !== 'Aberto') throw new Error('Esta despesa já foi liquidada e não pode ser alterada.');

    let categoria = String(disp[1]), descricao = String(disp[2]), quemPagou = String(disp[4]), valor = Number(atual[3] || 0);
    const antes = textoDespesa(categoria, descricao, valor, quemPagou);
    let dataNova = null;
    if (payload.anular) {
      descricao = /^ANULADA/.test(descricao) ? descricao : 'ANULADA — ' + (descricao || '(sem descrição)');
      valor = 0;
      acaoTxt = 'Despesa anulada';
    } else {
      categoria = payload.categoria || categoria;
      quemPagou = payload.quemPagou || quemPagou;
      descricao = String(payload.descricao || '').trim();
      valor = Number(payload.valor);
      if (isNaN(valor) || valor < 0) throw new Error('Valor inválido.');
      if (categoria === 'Portagens') {
        quemPagou = 'Luís - Via Verde';
        if (!descricao) descricao = 'Portagem Via Verde';
      } else if (quemPagou === 'Luís - Via Verde') {
        throw new Error('"Luís - Via Verde" só pode ser usado em Portagens.');
      }
      if (!descricao) throw new Error('Preenche a descrição da despesa.');
      if (payload.data) dataNova = normalizarData(payload.data, null);
      acaoTxt = 'Despesa corrigida';
    }
    sh.getRange(linha, 2, 1, 4).setValues([[categoria, descricao, valor, quemPagou]]);
    detalhe = antes + ' → ' + textoDespesa(categoria, descricao, valor, quemPagou);
    if (dataNova && dataNova !== String(disp[0])) {
      sh.getRange(linha, 1).setValue(dataNova);
      detalhe += ' — data ' + dataPT(String(disp[0])) + ' → ' + dataPT(dataNova);
    }
  }
  else {
    throw new Error('Ação desconhecida: ' + payload.action);
  }

  if (acaoTxt) registar(ss, quem, acaoTxt, detalhe);
  SpreadsheetApp.flush();
  return { status: 'ok', resultado: resultado, data: lerDadosCompletos(acesso) };
}

// ---------- Registo de alterações ----------

// Datas chegam da app como aaaa-mm-dd. Vazio devolve o valor por omissão; uma data impossível dá erro.
function normalizarData(txt, padrao) {
  if (txt === undefined || txt === null || txt === '') return padrao;
  const m = String(txt).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error('Data inválida.');
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  if (d.getFullYear() !== Number(m[1]) || d.getMonth() !== Number(m[2]) - 1 || d.getDate() !== Number(m[3])) throw new Error('Data inválida.');
  return String(txt);
}
function dataPT(iso) {
  const m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? m[3] + '/' + m[2] + '/' + m[1] : String(iso);
}
function sufixoData(d, hoje) { return d && d !== hoje ? ' — data ' + dataPT(d) : ''; }

function textoDespesa(categoria, descricao, valor, origem) {
  return categoria + ' — ' + (descricao || '(sem descrição)') + ' — ' + fmtEur(valor) + ' (' + origem + ')';
}

function fmtEur(v) {
  const n = Number(v);
  if (v === '' || v === null || v === undefined || isNaN(n)) return '—';
  return n.toFixed(2).replace('.', ',') + ' €';
}

// O nome vem da app (cada telemóvel escolhe "Luís" ou "Paulo"). Não é uma autenticação:
// serve só para saber quem fez cada alteração.
function identificarUtilizador(nome) {
  const n = String(nome || '').trim();
  if (n === 'Luís' || n === 'Paulo') return n;
  let email = '';
  try { email = Session.getActiveUser().getEmail(); } catch (e) {}
  return email || 'Desconhecido';
}

function registar(ss, quem, acao, detalhe) {
  try {
    let sh = ss.getSheetByName(SHEET_REGISTO);
    if (!sh) {
      sh = ss.insertSheet(SHEET_REGISTO);
      sh.getRange(1, 1, 1, 4).setValues([['Data e hora', 'Quem', 'Ação', 'Detalhe']]).setFontWeight('bold');
      sh.setFrozenRows(1);
      sh.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');
      sh.setColumnWidth(1, 150); sh.setColumnWidth(2, 80); sh.setColumnWidth(3, 220); sh.setColumnWidth(4, 620);
    }
    sh.appendRow([new Date(), quem, acao, detalhe]);
  } catch (err) {
    console.error('Não foi possível registar a alteração: ' + err); // nunca impede a gravação
  }
}

// Últimas alterações (mais recentes primeiro).
function lerRegisto(ss) {
  const sh = ss.getSheetByName(SHEET_REGISTO);
  if (!sh) return [];
  const ult = sh.getLastRow();
  if (ult < 2) return [];
  const n = Math.min(60, ult - 1);
  return sh.getRange(ult - n + 1, 1, n, 4).getDisplayValues().reverse()
    .map(r => ({ data: r[0], quem: r[1], acao: r[2], detalhe: r[3] }));
}

/**
 * Fecha o ciclo depois de o dinheiro ter sido dividido:
 * - obras "Recebido em Caixa" e "Aberto" passam a "Fechado" (as pendentes ficam abertas);
 * - despesas "Aberto" passam a "Liquidado";
 * - o fundo inicial do novo ciclo passa a ser a reserva deixada no envelope (Linha G).
 */
function fecharCiclo(ss, hoje, quem) {
  const shAcerto = abaObrigatoria(ss, ABAS.acerto);
  const shTrab = abaObrigatoria(ss, ABAS.trabalhos);
  const shDesp = abaObrigatoria(ss, ABAS.despesas);
  const shFundo = abaObrigatoria(ss, ABAS.fundo);

  const conferencia = String(shAcerto.getRange(CEL_ACERTO.conferencia).getDisplayValue());
  if (conferencia.indexOf('100%') !== 0) {
    throw new Error('A conferência do acerto não bate certo (' + conferencia + '). Ciclo não fechado.');
  }
  const sobra = Number(shAcerto.getRange(CEL_ACERTO.sobra).getValue()) || 0;
  const reserva = Number(shAcerto.getRange(CEL_ACERTO.reserva).getValue()) || 0;
  if (reserva < 0 || reserva > sobra) {
    throw new Error('A reserva para o envelope (' + reserva + ' €) não pode ser maior que o valor que sobra (' + sobra + ' €).');
  }
  const resumo = {
    data: hoje,
    luis: Number(shAcerto.getRange(CEL_ACERTO.totalLuis).getValue()) || 0,
    paulo: Number(shAcerto.getRange(CEL_ACERTO.totalPaulo).getValue()) || 0,
    reserva: reserva,
    obrasFechadas: 0,
    despesasLiquidadas: 0,
    obras: []
  };
  // Números do acerto tal como estão agora (depois de fechar ficam a zero).
  const foto = {
    cobrado: Number(shAcerto.getRange(CEL_ACERTO.cobrado).getValue()) || 0,
    despLuis: Number(shAcerto.getRange(CEL_ACERTO.despLuis).getValue()) || 0,
    despPaulo: Number(shAcerto.getRange(CEL_ACERTO.despPaulo).getValue()) || 0,
    mesa: Number(shAcerto.getRange(CEL_ACERTO.dinheiroMao).getValue()) || 0,
    gastoFundo: Number(shFundo.getRange(CEL_FUNDO.gastoFundo).getValue()) || 0
  };

  // 1.º passo (só leitura): que obras e despesas fecham.
  const fimTrab = encontrarLinhaAntesDoTotal(shTrab, 1, 'Total') - 1;
  let ciclos = [], nTrab = 0;
  if (fimTrab >= PRIMEIRA_LINHA_DADOS) {
    nTrab = fimTrab - PRIMEIRA_LINHA_DADOS + 1;
    const ids = shTrab.getRange(PRIMEIRA_LINHA_DADOS, 1, nTrab, 1).getValues();
    const estados = shTrab.getRange(PRIMEIRA_LINHA_DADOS, 8, nTrab, 1).getValues();
    ciclos = shTrab.getRange(PRIMEIRA_LINHA_DADOS, 10, nTrab, 1).getValues();
    for (let i = 0; i < nTrab; i++) {
      if (String(estados[i][0]) === 'Recebido em Caixa' && String(ciclos[i][0]) === 'Aberto') {
        ciclos[i][0] = 'Fechado';
        resumo.obrasFechadas++;
        resumo.obras.push(String(ids[i][0]));
      }
    }
  }
  const fimDesp = encontrarLinhaAntesDoTotal(shDesp, 1, 'Total') - 1;
  let nDesp = 0, vals = [], estadosDesp = [], notasDesp = [];
  if (fimDesp >= PRIMEIRA_LINHA_DADOS) {
    nDesp = fimDesp - PRIMEIRA_LINHA_DADOS + 1;
    vals = shDesp.getRange(PRIMEIRA_LINHA_DADOS, 1, nDesp, 7).getValues();
    estadosDesp = vals.map(r => [r[5]]);
    notasDesp = vals.map(r => [r[6]]);
    for (let i = 0; i < nDesp; i++) {
      if (String(estadosDesp[i][0]) === 'Aberto' && String(vals[i][0]).trim() !== '') {
        estadosDesp[i][0] = 'Liquidado';
        notasDesp[i][0] = (notasDesp[i][0] ? notasDesp[i][0] + ' | ' : '') + 'Acerto ' + hoje;
        resumo.despesasLiquidadas++;
      }
    }
  }

  // 2.º passo: guarda o resumo do acerto ANTES de alterar o que quer que seja
  // (se falhar, o ciclo não fecha).
  registarFecho(ss, quem, resumo, foto);

  // 3.º passo: fecha.
  if (nTrab) shTrab.getRange(PRIMEIRA_LINHA_DADOS, 10, nTrab, 1).setValues(ciclos);
  if (nDesp) {
    shDesp.getRange(PRIMEIRA_LINHA_DADOS, 6, nDesp, 1).setValues(estadosDesp);
    shDesp.getRange(PRIMEIRA_LINHA_DADOS, 7, nDesp, 1).setValues(notasDesp);
  }
  shFundo.getRange(CEL_FUNDO.inicial).setValue(reserva);
  return resumo;
}

// Uma linha por ciclo fechado, na aba "Acertos Fechados" (criada automaticamente).
function registarFecho(ss, quem, resumo, foto) {
  let sh = ss.getSheetByName(SHEET_FECHOS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_FECHOS);
    sh.getRange(1, 1, 1, 12).setValues([['Data e hora', 'Fechado por', 'Obras cobradas (€)', 'Gastos do fundo (€)',
      'Despesas do Luís (€)', 'Despesas do Paulo (€)', 'Dinheiro na mesa (€)', 'Ficou no envelope (€)',
      'Luís recebeu (€)', 'Paulo recebeu (€)', 'Obras fechadas', 'Despesas liquidadas']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
    sh.getRange('C:J').setNumberFormat('#,##0.00 "€"');
    sh.setColumnWidth(1, 130); sh.setColumnWidth(11, 260);
  }
  sh.appendRow([new Date(), quem, foto.cobrado, foto.gastoFundo, foto.despLuis, foto.despPaulo, foto.mesa,
    resumo.reserva, resumo.luis, resumo.paulo, resumo.obras.join(', '), resumo.despesasLiquidadas]);
}

// Últimos acertos fechados (mais recentes primeiro).
function lerAcertos(ss) {
  const sh = ss.getSheetByName(SHEET_FECHOS);
  if (!sh) return [];
  const ult = sh.getLastRow();
  if (ult < 2) return [];
  const n = Math.min(12, ult - 1);
  return sh.getRange(ult - n + 1, 1, n, 12).getDisplayValues().reverse().map(r => ({
    data: r[0], quem: r[1], cobrado: r[2], luis: r[8], paulo: r[9], envelope: r[7], obras: r[10], despesas: r[11]
  }));
}

// Lê todas as células do mapa numa só chamada (bloco A1 até à maior célula usada).
function lerNumeros(sheet, mapa) {
  const pos = {};
  let maxL = 1, maxC = 1;
  Object.keys(mapa).forEach(k => {
    const m = mapa[k].match(/^([A-Z])(\d+)$/);
    const c = m[1].charCodeAt(0) - 64, l = parseInt(m[2], 10);
    pos[k] = [l - 1, c - 1];
    maxL = Math.max(maxL, l); maxC = Math.max(maxC, c);
  });
  const vals = sheet.getRange(1, 1, maxL, maxC).getValues();
  const disp = sheet.getRange(1, 1, maxL, maxC).getDisplayValues();
  const out = {};
  Object.keys(pos).forEach(k => {
    const v = vals[pos[k][0]][pos[k][1]];
    const d = disp[pos[k][0]][pos[k][1]];
    out[k] = (typeof v === 'number') ? v : (k === 'conferencia' ? d : limparMoeda(d));
  });
  return out;
}

function lerDadosCompletos(acesso) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shTrab = encontrarAba(ss, ABAS.trabalhos);
  const shDesp = encontrarAba(ss, ABAS.despesas);
  const shFundo = encontrarAba(ss, ABAS.fundo);
  const shAcerto = encontrarAba(ss, ABAS.acerto);

  const trabalhos = [];
  if (shTrab) {
    const fim = encontrarLinhaAntesDoTotal(shTrab, 1, 'Total') - 1;
    if (fim >= PRIMEIRA_LINHA_DADOS) {
      const n = fim - PRIMEIRA_LINHA_DADOS + 1;
      const raw = shTrab.getRange(PRIMEIRA_LINHA_DADOS, 1, n, 10).getValues();
      const disp = shTrab.getRange(PRIMEIRA_LINHA_DADOS, 1, n, 10).getDisplayValues();
      for (let i = 0; i < n; i++) {
        const id = String(disp[i][0] || '').trim();
        if (!id) continue;
        trabalhos.push({
          id: id,
          dataInicio: disp[i][1],
          obra: disp[i][2],
          localidade: disp[i][3],
          cliente: disp[i][4],
          descricao: disp[i][5],
          valor: raw[i][6] === '' ? null : (Number(raw[i][6]) || 0),
          estadoPagamento: disp[i][7],
          dataRecebimento: disp[i][8],
          ciclo: disp[i][9] || 'Aberto'
        });
      }
    }
  }

  const despesas = [];
  if (shDesp) {
    const fim = encontrarLinhaAntesDoTotal(shDesp, 1, 'Total') - 1;
    if (fim >= PRIMEIRA_LINHA_DADOS) {
      const n = fim - PRIMEIRA_LINHA_DADOS + 1;
      const raw = shDesp.getRange(PRIMEIRA_LINHA_DADOS, 1, n, 7).getValues();
      const disp = shDesp.getRange(PRIMEIRA_LINHA_DADOS, 1, n, 7).getDisplayValues();
      for (let i = 0; i < n; i++) {
        if (String(disp[i][0] || '').trim() === '') continue;
        despesas.push({
          linha: PRIMEIRA_LINHA_DADOS + i,
          data: disp[i][0],
          categoria: disp[i][1],
          descricao: disp[i][2],
          valor: Number(raw[i][3]) || 0,
          origem: disp[i][4],
          estado: disp[i][5] || 'Aberto',
          notas: disp[i][6] || ''
        });
      }
    }
  }

  const fundo = shFundo ? lerNumeros(shFundo, CEL_FUNDO) : null;
  const acerto = shAcerto ? lerNumeros(shAcerto, CEL_ACERTO) : null;

  return {
    updatedAt: new Date().toISOString(),
    fundoInicial: fundo ? fundo.inicial : 0,
    fundo: fundo,
    acerto: acerto,
    registo: lerRegisto(ss),
    acertos: lerAcertos(ss),
    acesso: acesso || { protegido: false, nome: '' },
    trabalhos: trabalhos,
    despesas: despesas
  };
}

function encontrarLinhaAntesDoTotal(sheet, colIndex, textoTotal) {
  const vals = sheet.getDataRange().getDisplayValues();
  for (let i = PRIMEIRA_LINHA_DADOS - 1; i < vals.length; i++) {
    if (String(vals[i][colIndex - 1]).toLowerCase().indexOf(textoTotal.toLowerCase()) !== -1) {
      return i + 1;
    }
  }
  return sheet.getLastRow() + 1;
}

/**
 * Devolve a linha onde gravar um novo registo, sempre DENTRO do intervalo
 * usado pelas fórmulas de total (ex.: SUMIFS(D5:D10; ...)).
 * 1) Usa a primeira linha vazia entre a linha 5 e a linha "Total".
 * 2) Se não houver, insere uma linha antes da última linha de dados (o que
 *    alarga automaticamente os intervalos das fórmulas), sobe o último registo
 *    para a nova linha e devolve a última posição, mantendo a ordem cronológica.
 * colsChave: colunas (base 1) que têm de estar vazias para a linha contar como livre.
 */
function obterLinhaLivre(sheet, numCols, colsChave) {
  const linhaTotal = encontrarLinhaAntesDoTotal(sheet, 1, 'Total');
  const ultimaDados = linhaTotal - 1;

  if (ultimaDados >= PRIMEIRA_LINHA_DADOS) {
    const vals = sheet.getRange(PRIMEIRA_LINHA_DADOS, 1, ultimaDados - PRIMEIRA_LINHA_DADOS + 1, numCols).getValues();
    for (let i = 0; i < vals.length; i++) {
      const livre = colsChave.every(c => String(vals[i][c - 1]).trim() === '');
      if (livre) return PRIMEIRA_LINHA_DADOS + i;
    }
    sheet.insertRowBefore(ultimaDados);
    // O último registo passou para ultimaDados + 1; copia-o para a nova linha.
    sheet.getRange(ultimaDados + 1, 1, 1, numCols).copyTo(sheet.getRange(ultimaDados, 1, 1, numCols));
    return ultimaDados + 1;
  }

  sheet.insertRowBefore(linhaTotal);
  return linhaTotal;
}

function gerarProximoIdObra(sheet) {
  const linhaTotal = encontrarLinhaAntesDoTotal(sheet, 1, 'Total');
  const numLinhas = Math.max(1, linhaTotal - PRIMEIRA_LINHA_DADOS);
  const vals = sheet.getRange(PRIMEIRA_LINHA_DADOS, 1, numLinhas, 1).getDisplayValues();
  let maxNum = 0;
  vals.forEach(r => {
    const m = String(r[0]).match(/(\d+)/);
    if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
  });
  return 'OBR-' + String(maxNum + 1).padStart(3, '0');
}

function limparMoeda(str) {
  if (typeof str === 'number') return str;
  const limpo = String(str || '0').replace(/[^\d,.-]/g, '').replace(/\./g, '').replace(',', '.');
  return parseFloat(limpo) || 0;
}
