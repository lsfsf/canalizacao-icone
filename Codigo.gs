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

function apiLerDados() {
  return lerDadosCompletos();
}

function apiGuardarDados(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return guardarNaFolha(payload);
  } finally {
    lock.releaseLock();
  }
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

function guardarNaFolha(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  let resultado = null;

  if (payload.action === 'addTrabalho') {
    const sh = abaObrigatoria(ss, ABAS.trabalhos);
    const linha = obterLinhaLivre(sh, 10, [1]);
    const novoId = gerarProximoIdObra(sh);
    const temValor = payload.valor !== null && payload.valor !== undefined && payload.valor !== '';
    const estado = temValor ? (payload.estadoPagamento || 'Pendente') : 'Pendente';
    const dataRec = (estado === 'Recebido em Caixa' && !payload.dataRecebimento) ? hoje : (payload.dataRecebimento || '');

    sh.getRange(linha, 1, 1, 10).setValues([[
      novoId,
      payload.dataInicio || hoje,
      payload.nomeObra || '',
      payload.localidade || '',
      payload.nomeCliente || '',
      payload.descricao || '',
      temValor ? Number(payload.valor) : '',
      estado,
      dataRec,
      'Aberto'
    ]]);
  }
  else if (payload.action === 'addDespesa') {
    // Regra do negócio: portagens são sempre Via Verde do Luís, e Via Verde só existe para portagens.
    let categoria = payload.categoria || 'Material Stock';
    let quemPagou = payload.quemPagou || 'Fundo Caixa';
    if (categoria === 'Portagens') quemPagou = 'Luís - Via Verde';
    else if (quemPagou === 'Luís - Via Verde') throw new Error('"Luís - Via Verde" só pode ser usado em Portagens.');

    const sh = abaObrigatoria(ss, ABAS.despesas);
    const linha = obterLinhaLivre(sh, 7, [1, 4]);

    sh.getRange(linha, 1, 1, 7).setValues([[
      payload.data || hoje,
      categoria,
      payload.descricao || '',
      Number(payload.valor || 0),
      quemPagou,
      'Aberto',
      payload.notas || ''
    ]]);
  }
  else if (payload.action === 'marcarObraRecebida') {
    const sh = abaObrigatoria(ss, ABAS.trabalhos);
    const dados = sh.getDataRange().getValues();
    for (let i = PRIMEIRA_LINHA_DADOS - 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === String(payload.idObra).trim()) {
        sh.getRange(i + 1, 8).setValue('Recebido em Caixa');
        sh.getRange(i + 1, 9).setValue(payload.dataRecebimento || hoje);
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
        sh.getRange(i + 1, 7).setValue(temValor ? Number(payload.valor) : '');
        if (payload.descricao !== undefined) sh.getRange(i + 1, 6).setValue(String(payload.descricao));
        encontrada = true;
        break;
      }
    }
    if (!encontrada) throw new Error('Obra ' + payload.idObra + ' não encontrada.');
  }
  else if (payload.action === 'setFundoInicial') {
    abaObrigatoria(ss, ABAS.fundo).getRange(CEL_FUNDO.inicial).setValue(Number(payload.valor || 0));
  }
  else if (payload.action === 'setReserva') {
    abaObrigatoria(ss, ABAS.acerto).getRange(CEL_ACERTO.reserva).setValue(Number(payload.valor || 0));
  }
  else if (payload.action === 'fecharCiclo') {
    resultado = fecharCiclo(ss, hoje);
  }
  else {
    throw new Error('Ação desconhecida: ' + payload.action);
  }

  SpreadsheetApp.flush();
  return { status: 'ok', resultado: resultado, data: lerDadosCompletos() };
}

/**
 * Fecha o ciclo depois de o dinheiro ter sido dividido:
 * - obras "Recebido em Caixa" e "Aberto" passam a "Fechado" (as pendentes ficam abertas);
 * - despesas "Aberto" passam a "Liquidado";
 * - o fundo inicial do novo ciclo passa a ser a reserva deixada no envelope (Linha G).
 */
function fecharCiclo(ss, hoje) {
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
    despesasLiquidadas: 0
  };

  const fimTrab = encontrarLinhaAntesDoTotal(shTrab, 1, 'Total') - 1;
  if (fimTrab >= PRIMEIRA_LINHA_DADOS) {
    const n = fimTrab - PRIMEIRA_LINHA_DADOS + 1;
    const estados = shTrab.getRange(PRIMEIRA_LINHA_DADOS, 8, n, 1).getValues();
    const ciclos = shTrab.getRange(PRIMEIRA_LINHA_DADOS, 10, n, 1).getValues();
    for (let i = 0; i < n; i++) {
      if (String(estados[i][0]) === 'Recebido em Caixa' && String(ciclos[i][0]) === 'Aberto') {
        ciclos[i][0] = 'Fechado';
        resumo.obrasFechadas++;
      }
    }
    shTrab.getRange(PRIMEIRA_LINHA_DADOS, 10, n, 1).setValues(ciclos);
  }

  const fimDesp = encontrarLinhaAntesDoTotal(shDesp, 1, 'Total') - 1;
  if (fimDesp >= PRIMEIRA_LINHA_DADOS) {
    const n = fimDesp - PRIMEIRA_LINHA_DADOS + 1;
    const vals = shDesp.getRange(PRIMEIRA_LINHA_DADOS, 1, n, 7).getValues();
    const estados = vals.map(r => [r[5]]);
    const notas = vals.map(r => [r[6]]);
    for (let i = 0; i < n; i++) {
      if (String(estados[i][0]) === 'Aberto' && String(vals[i][0]).trim() !== '') {
        estados[i][0] = 'Liquidado';
        notas[i][0] = (notas[i][0] ? notas[i][0] + ' | ' : '') + 'Acerto ' + hoje;
        resumo.despesasLiquidadas++;
      }
    }
    shDesp.getRange(PRIMEIRA_LINHA_DADOS, 6, n, 1).setValues(estados);
    shDesp.getRange(PRIMEIRA_LINHA_DADOS, 7, n, 1).setValues(notas);
  }

  shFundo.getRange(CEL_FUNDO.inicial).setValue(reserva);
  return resumo;
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

function lerDadosCompletos() {
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
