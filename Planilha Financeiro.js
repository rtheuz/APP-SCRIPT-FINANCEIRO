// ======================================================
// ⚡ CODE GS OTIMIZADO (DESEMPENHO MÁXIMO)
// ======================================================
const ss = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_LD = ss.getSheetByName("LD");
const SHEET_CAD = ss.getSheetByName("CADASTRO");
const HEADER_ROW_LD = 4;
const CACHE_DURATION = 600; // 10 minutos

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ Financeiro')
    .addItem('🔄 Atualizar Saldos', 'AtualizarSaldosLD')
    .addItem('📥 Importar Projetos', 'importarDadosFixos')
    .addItem('🎨 Pintar Tudo', 'pintarTudoLD')
    .addItem('🔄 Atualizar Planilha LD', 'reaplicarDadosCadastroLD')
    .addSeparator()
    .addItem('↕ Classificar por Caixa', 'classificarCaixa')
    .addItem('↕ Classificar por Vencimento', 'classificarVencimento')
    .addSeparator()
    .addItem('📅 Ir para Última Linha Data Caixa', 'irParaUltimaLinhaDataCaixa')
    .addItem('🗑️ Limpar Cache', 'limparCacheCadastro')
    .addToUi();
}
// ======================================================
// 🔹 CÁLCULO DE SALDOS (Q e R)
// ======================================================
// ----------------------
// getLastDataRow_ (corrigida, retorna linha absoluta)
// OTIMIZAÇÃO: Lê ambas as colunas de uma vez em vez de separadamente
// ----------------------
function getLastDataRow_(sh) {
  const firstDataRow = HEADER_ROW_LD + 1;
  const lastRowSheet = sh.getLastRow();
  if (lastRowSheet < firstDataRow) return HEADER_ROW_LD;

  const numRows = lastRowSheet - HEADER_ROW_LD;
  
  // OTIMIZADO: Lê M e P de uma só vez (colunas 13-16)
  const dados = sh.getRange(firstDataRow, 13, numRows, 4).getValues();

  for (let i = numRows - 1; i >= 0; i--) {
    const c = dados[i][0]; // M (col 13, índice 0)
    const v = dados[i][3]; // P (col 16, índice 3)
    if ((c !== "" && c != null) || (v !== "" && v != null)) {
      return firstDataRow + i;
    }
  }
  return HEADER_ROW_LD;
}

// ----------------------
// Função utilitária: tenta extrair número mesmo se vier como string "1.234,56" ou "R$ 1.234,56"
// ----------------------
function parseBrasilNumber(raw) {
  if (raw === null || raw === undefined || raw === '') return 0;
  if (typeof raw === 'number') return raw;
  let s = String(raw).trim();
  // remove "R$", espaços e outros símbolos, mantém dígitos, '.' e ','
  s = s.replace(/[R$\s]/g, '');
  // se tem vírgula e ponto, assume formato pt-BR: remove pontos (milhar) e troca vírgula por ponto
  if (s.indexOf(',') > -1 && s.indexOf('.') > -1) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (s.indexOf(',') > -1 && s.indexOf('.') === -1) {
    // "1234,56" -> "1234.56"
    s = s.replace(',', '.');
  } else {
    // "1.234" pode ser 1234 ou 1.234 (provável milhar) -> remover quaisquer espaços
    s = s.replace(/\s/g, '');
  }
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}

// ----------------------
// atualizarSaldosLD (corrigida: Q/R/O2:P2 com futuros; O1:P1 SÓ realizado/Caixa)
// Modificado: O2:P2 usa fórmula SUBTOTAL para atualizar automaticamente com filtros
// ----------------------
function atualizarSaldosLD() {
  const sh = SHEET_LD;
  if (!sh) return;

  const headerRow = HEADER_ROW_LD;
  const lastRow = getLastDataRow_(sh);
  if (lastRow <= headerRow) return;

  const contaSelecionada = sh.getRange("Q4").getValue();
  if (!contaSelecionada) return;

  const startRow = headerRow + 1;
  const numRows = lastRow - headerRow; // linhas reais a ler

  // 1. LÊ COLUNAS M:N:O:P:Q (13..17) para o cálculo total (Q e R)
  const dados = sh.getRange(startRow, 13, numRows, 5).getValues();

  // 2. LÊ COLUNA D (4) para verificar se o lançamento é "realizado" (Data de Caixa)
  const datasCaixa = sh.getRange(startRow, 4, numRows, 1).getValues();

  let saldoGeral = 0; // Acumula todos os lançamentos (para R)
  let saldoContaTotal = 0; // Acumula todos os lançamentos (para Q)
  
  let saldoContaRealizado = 0; // Acumula APENAS os lançamentos com Data de Caixa (para O1:P1)
  
  const resultadosQ = [];
  const resultadosR = [];

  for (let i = 0; i < dados.length; i++) {
    const linha = dados[i];
    const conta = linha[0]; // M
    const rawValor = linha[3]; // P (pode ser número ou string)
    
    const dataCaixa = datasCaixa[i][0]; // D (Valor da coluna D da linha atual)
    const valor = parseBrasilNumber(rawValor);

    // ----------------------------------------------------------------------
    // CÁLCULO TOTAL (Para Q e R - soma todos os lançamentos)
    // ----------------------------------------------------------------------
    saldoGeral += valor;
    if (conta === contaSelecionada) saldoContaTotal += valor;

    resultadosQ.push([saldoContaTotal]); // Q
    resultadosR.push([saldoGeral]); // R

    // ----------------------------------------------------------------------
    // CÁLCULO REALIZADO (Para O1:P1 - APENAS se tiver Data de Caixa)
    // ----------------------------------------------------------------------
    if (dataCaixa && conta === contaSelecionada) {
        saldoContaRealizado += valor;
    }
  }

  // grava Q e R
  sh.getRange(startRow, 17, resultadosQ.length, 1).setValues(resultadosQ); // Q
  sh.getRange(startRow, 18, resultadosR.length, 1).setValues(resultadosR); // R

  // --------------------------------------------------
  // ✅ Exibe o Saldo da Conta Selecionada (O1:P1) - NÃO afetado por filtros
  // --------------------------------------------------
  const targetRangeO1P1 = sh.getRange("O1:P1");
  if (!targetRangeO1P1.isPartOfMerge()) targetRangeO1P1.merge();
  targetRangeO1P1
    .setNumberFormat('"R$" #,##0.00')
    .setHorizontalAlignment("center")
    // Usa o saldo REALIZADO (só lançamentos com Data de Caixa)
    .setValue(saldoContaRealizado); 

  // --------------------------------------------------
  // 🧮 Exibe o Saldo Geral (O2:P2) - USA FÓRMULA SUBTOTAL para atualizar automaticamente com filtros
  // --------------------------------------------------

  // Usa SUBTOTAL(109, P5:P) que soma apenas valores visíveis automaticamente
  // Função 109 = SOMA ignorando valores ocultos por filtros
  const targetRangeO2P2 = sh.getRange("O2:P2");
  if (!targetRangeO2P2.isPartOfMerge()) targetRangeO2P2.merge();
  targetRangeO2P2
    .setNumberFormat('"R$" #,##0.00')
    .setHorizontalAlignment("center");
  // Fórmula SUBTOTAL atualiza automaticamente quando filtros são aplicados
  // Deve ser aplicada apenas na célula O2 (primeira célula do merge)
  sh.getRange("O2").setFormula('=SUBTOTAL(109;P5:P)');

  // --------------------------------------------------
  // 💬 TOAST
  // --------------------------------------------------
  SpreadsheetApp.getActive().toast(
    `✅ Saldos atualizados — Conta Selecionada (Caixa): R$ ${saldoContaRealizado.toLocaleString('pt-BR', { minimumFractionDigits: 2 })} | Geral: Fórmula SUBTOTAL (atualiza automaticamente com filtros)`,
    "Saldos",
    5
  );
}

// ======================================================
// 🔹 FUNÇÕES DE CACHE PARA PERFORMANCE
// ======================================================

/**
 * Obtém dados do cadastro com cache para melhor performance
 * @param {string} tipo - 'projetos' ou 'contas'
 * @return {Array} Dados do cadastro
 */
function getCadastroCached_(tipo) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `cadastro_${tipo}`;
  const cached = cache.get(cacheKey);
  
  if (cached) {
    return JSON.parse(cached);
  }
  
  const cadastro = SHEET_CAD;
  if (!cadastro) return [];
  
  let dados;
  if (tipo === 'projetos') {
    const lastRow = cadastro.getLastRow();
    if (lastRow < 5) return [];
    dados = cadastro.getRange('AY5:BA' + lastRow).getValues();
  } else if (tipo === 'contas') {
    const lastRow = cadastro.getLastRow();
    if (lastRow < 5) return [];
    dados = cadastro.getRange('J5:L' + lastRow).getValues();
  } else {
    return [];
  }
  
  cache.put(cacheKey, JSON.stringify(dados), CACHE_DURATION);
  return dados;
}

/**
 * Cria índice Map para lookup O(1) em vez de O(n)
 * @param {Array} dados - Array 2D de dados
 * @return {Map} Índice de valores para linha
 */
function criarIndiceCadastro_(dados) {
  const indice = new Map();
  
  dados.forEach((row, idx) => {
    row.forEach(val => {
      if (val && val.toString().trim()) {
        const key = val.toString().trim().toLowerCase();
        // Armazena o primeiro match encontrado
        if (!indice.has(key)) {
          indice.set(key, idx);
        }
      }
    });
  });
  
  return indice;
}

/**
 * Limpa o cache de cadastros
 */
function limparCacheCadastro() {
  const cache = CacheService.getScriptCache();
  cache.remove('cadastro_projetos');
  cache.remove('cadastro_contas');
  SpreadsheetApp.getActive().toast('✅ Cache de cadastros limpo!', 'Cache', 3);
}

// ======================================================
// 🔹 ONEDIT PRINCIPAL (corrigido com formatação automática de datas)
// ======================================================
function onEdit(e) {
  try {
    const r = e.range;
    const sh = r.getSheet();
    if (sh.getName() !== 'LD') return;

    const row = r.getRow();
    const col = r.getColumn();
    const val = e.value;
    if (!val) return;

    const cadastro = SHEET_CAD;

    // ======================================================
// 🗓️ FORMATAÇÃO AUTOMÁTICA DE DATAS (colunas B, C, D)
// ======================================================
if (row >= 5 && col >= 2 && col <= 4) {
  const num = val.toString().replace(/\D/g, '');

  // Aceita entradas com 4,5,6,7 ou 8 dígitos:
  // - 4 -> DD MM (ex: 0110 => 01/10/ANOATUAL)
  // - 5 -> D M YY ou DD M YY (ex: 11025 => 1/10/25 ou 10125 => 10/1/25)
  // - 6 -> DD MM YY (ex: 011025 => 01/10/25)
  // - 7 -> D MM YYYY (ex: 1112025 => 1/11/2025)
  // - 8 -> DD MM YYYY (ex: 01112025 => 01/11/2025)
  
  if (num.length >= 4 && num.length <= 8) {
    let dia, mes, anoNum;
    
    // Lógica de parsing baseada no comprimento
    if (num.length === 4) {
      // DDMM - adiciona ano atual
      dia = num.slice(0, 2);
      mes = num.slice(2, 4);
      anoNum = new Date().getFullYear();
      
    } else if (num.length === 5) {
      // Pode ser DMMYY (ex: 11025) ou DDMYY (ex: 10125)
      // Vamos tentar DMMYY primeiro
      dia = num.slice(0, 1);
      mes = num.slice(1, 3);
      const anoStr = num.slice(3, 5);
      anoNum = Number(anoStr) < 80 ? 2000 + Number(anoStr) : 1900 + Number(anoStr);
      
      // Valida se o mês é válido, senão tenta DDMYY
      if (Number(mes) > 12) {
        dia = num.slice(0, 2);
        mes = num.slice(2, 3);
        anoNum = Number(anoStr) < 80 ? 2000 + Number(anoStr) : 1900 + Number(anoStr);
      }
      
    } else if (num.length === 6) {
      // DDMMYY
      dia = num.slice(0, 2);
      mes = num.slice(2, 4);
      const anoStr = num.slice(4, 6);
      anoNum = Number(anoStr) < 80 ? 2000 + Number(anoStr) : 1900 + Number(anoStr);
      
    } else if (num.length === 7) {
      // DMMYYYY
      dia = num.slice(0, 1);
      mes = num.slice(1, 3);
      anoNum = Number(num.slice(3, 7));
      
    } else if (num.length === 8) {
      // DDMMYYYY
      dia = num.slice(0, 2);
      mes = num.slice(2, 4);
      anoNum = Number(num.slice(4, 8));
    }

    // Valida dia e mês
    const diaNum = Number(dia);
    const mesNum = Number(mes);
    
    if (diaNum >= 1 && diaNum <= 31 && mesNum >= 1 && mesNum <= 12) {
      // Corrige fuso: cria data às 12h
      const data = new Date(anoNum, mesNum - 1, diaNum, 12, 0, 0);

      if (!isNaN(data.getTime())) {
        r.setValue(data);
        r.setNumberFormat('dd/mm/yy');
        return;
      }
    }
  }
}

    // ======================================================
    // 🔸 BLOCO PROJETO (G:H:I → AY:BA)
    // OTIMIZADO: Usa cache e índice para lookup O(1)
    // ======================================================
    if (col >= 7 && col <= 9) {
      const cadastroData = getCadastroCached_('projetos');
      if (cadastroData.length > 0) {
        const indice = criarIndiceCadastro_(cadastroData);
        const valKey = val.toString().trim().toLowerCase();
        const idx = indice.get(valKey);
        
        if (idx !== undefined) {
          const cadastroRow = cadastroData[idx];
          sh.getRange(row, 7, 1, 3).setValues([[cadastroRow[0], cadastroRow[1], cadastroRow[2]]]);
          // REMOVIDO: SpreadsheetApp.flush() - desnecessário e lento
        }
      }
    }

    // ======================================================
    // 🔸 BLOCO CONTA CONTÁBIL (J:K:L → J:L)
    // OTIMIZADO: Usa cache e índice para lookup O(1)
    // ======================================================
    if (col >= 10 && col <= 12) {
      const cadastroData = getCadastroCached_('contas');
      if (cadastroData.length > 0) {
        const indice = criarIndiceCadastro_(cadastroData);
        const valKey = val.toString().trim().toLowerCase();
        const idx = indice.get(valKey);
        
        if (idx !== undefined) {
          const cadastroRow = cadastroData[idx];
          sh.getRange(row, 10, 1, 3).setValues([[cadastroRow[0], cadastroRow[1], cadastroRow[2]]]);
          // REMOVIDO: SpreadsheetApp.flush() - desnecessário e lento
        }
      }
    }

    // ======================================================
    // 🔸 ATUALIZA SALDOS (se mudar Q4 ou colunas M/P)
    // ======================================================
    if (r.getA1Notation() === 'Q4') return atualizarSaldosLD();
    if (col === 13 || col === 16) atualizarSaldosLD();

    // ======================================================
    // 🔸 FORMATAÇÃO CONDICIONAL DINÂMICA
    // ======================================================
    aplicarFormatacaoLD(r);

  } catch (err) {
    Logger.log('Erro em onEdit: ' + err);
  }
}

function aplicarFormatacaoLD(r) {
  const sh = r.getSheet();
  const row = r.getRow();
  	
  // 🔹 Ignorar linhas 1 a HEADER_ROW_LD (cabeçalho e resumo)
  if (row <= HEADER_ROW_LD) return;

  // OTIMIZADO: Lê P e S de uma vez em vez de duas chamadas separadas
  const valores = sh.getRange(row, 16, 1, 4).getValues()[0]; // P, Q, R, S (cols 16-19)
  const valP = valores[0]; // P
  const valS = valores[3]; // S

  // --- 1) Define a cor base da linha (saldo negativo ou branca)
  let linhaCor = "white";
  if (valP < 0) linhaCor = "#b7e1cd";

  // Aplica a cor base na linha A:W
  sh.getRange(`A${row}:W${row}`).setBackground(linhaCor);

  // --- 2) Define cor especial (se aplicável)
  let bg = null;
  if (valS === "À pagar - CONFIRMAR VALOR !!") bg = "#ff60ff";
  else if (valS === "Pago - CHEQUE !!") bg = "#00ffff";
  else if (valS === "SALDO") bg = "#9900ff";
  else if (typeof valS === "string" && valS.includes("À pagar - CARTÃO DE CRÉDITO")) bg = "#46bdc6";
  else if (typeof valS === "string" && valS.includes("Agendado")) bg = "#6d9eeb";
  else if (valS === "Pago" && valP < 0) bg = "#43f643";
  else if (valS === "Pago" && valP > 0) bg = "#dd7e6b";

  // --- 3) Aplica a cor específica (bg) sobre as colunas D, F, P, Q, R, S
  if (bg) {
    // OTIMIZADO: Agrupa operações de coloração
    sh.getRange(row, 4).setBackground(bg); // D
    sh.getRange(row, 6).setBackground(bg); // F
    sh.getRange(row, 16, 1, 4).setBackground(bg); // P, Q, R, S de uma vez
  }
}


// ======================================================
// 🔹 BOTÃO MANUAL: Atualizar saldos LD
// ======================================================
function AtualizarSaldosLD() {
  atualizarSaldosLD();
}

// ======================================================
// 🔹 FUNÇÕES DE CLASSIFICAÇÃO (BOTÕES MANUAIS)
// ======================================================
function classificarCaixa() {
  const sh = SHEET_LD;
  const headerRow = HEADER_ROW_LD;
  const lastRow = sh.getLastRow();
  if (lastRow <= headerRow) return;

  const colDataCaixa = 4;   // D
  const colContaFin = 13;   // M
  const colMeioPag = 14;    // N
  const colValor = 16;      // P
  const colDataComp = 2;   // B
  const colDataVenc = 3;    // C

  const range = sh.getRange(headerRow + 1, 1, lastRow - headerRow, sh.getLastColumn());
  range.sort([
    { column: colDataCaixa, ascending: true },
    { column: colDataVenc, ascending: true },
    { column: colContaFin, ascending: true },
    { column: colMeioPag, ascending: true },
    { column: colValor, ascending: true },
    { column: colDataComp, ascending: true },
  ]);

  SpreadsheetApp.getActive().toast("✅ Classificação por CAIXA concluída", "Ordenação", 3);
}
// ======================================================
// 🔹 FUNÇÕES DE CLASSIFICAÇÃO (BOTÕES MANUAIS)
// ======================================================
function classificarVencimento() {
  const sh = SHEET_LD;
  const headerRow = HEADER_ROW_LD;
  const lastRow = sh.getLastRow();
  if (lastRow <= headerRow) return;

  const colDataCaixa = 4;   // D
  const colContaFin = 13;   // M
  const colMeioPag = 14;    // N
  const colValor = 16;      // P
  const colDataComp = 2;   // B
  const colDataVenc = 3;    // C

  const range = sh.getRange(headerRow + 1, 1, lastRow - headerRow, sh.getLastColumn());
  range.sort([
    { column: colDataVenc, ascending: true },
    { column: colDataCaixa, ascending: true },

  ]);

  SpreadsheetApp.getActive().toast("✅ Classificação por VENCIMENTO concluída", "Ordenação", 3);
}
// ======================================================
// 🔹 IR PARA ÚLTIMA LINHA PREENCHIDA (COLUNA "Data Caixa" = D)
// ======================================================
function irParaUltimaLinhaDataCaixa() {
  const sh = SHEET_LD;
  if (!sh) return;

  const colDataCaixa = 4; // Coluna D
  const lastRow = getLastDataRowByCol_(sh, colDataCaixa);
  if (lastRow <= HEADER_ROW_LD) {
    SpreadsheetApp.getActive().toast("⚠ Nenhum dado encontrado na coluna D", "Ir para linha", 3);
    return;
  }

  sh.activate(); // garante que está na aba LD
  sh.getRange(lastRow, colDataCaixa).activate();
  SpreadsheetApp.getActive().toast(`📅 Focado na última linha de Data Caixa (linha ${lastRow})`, "Navegação", 3);
}

// Função auxiliar (versão genérica)
function getLastDataRowByCol_(sh, col) {
  const values = sh.getRange(HEADER_ROW_LD + 1, col, sh.getLastRow() - HEADER_ROW_LD, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "" && values[i][0] != null) return i + HEADER_ROW_LD + 1;
  }
  return HEADER_ROW_LD;
}
// ======================================================
// 🚀 NOVA FUNÇÃO DE IMPORTAÇÃO (Substitui IMPORTRANGE)
// ======================================================

function importarDadosFixos() {
  const shCadastro = SHEET_CAD

  // --- DADOS DO CADASTRO DE CONTAS (Antigo A2) ---
  const sourceIdContas = "1JuoLc9KPTvm3B56tYVYuD4IWIsjUGyEN_Efd04PLoWs"; // Substitua pela ID real
  const sourceRangeContas = "CONTAS CONTÁBEIS!A:H"; // Exemplo: Range completo da aba

  // --- DADOS DOS PROJETOS (Antigo AK2) ---
  const sourceIdProjetos = "1UlVRpUOFwiAXrl9hkGe5wNvLjCDTcHeoOyv2XuTNmz4"; // Substitua pela ID real
  const sourceRangeProjetos = "LD!A:M"; // Exemplo: Range completo da aba

  if (!shCadastro) {
    Logger.log("Aba 'CADASTRO' não encontrada.");
    return;
  }

  // 1. Importa e fixa as Contas Contábeis (A2 em diante)
  try {
    const dadosContas = SpreadsheetApp.openById(sourceIdContas).getRange(sourceRangeContas).getValues();
    shCadastro.getRange("A2").clearContent(); // Limpa A2 (onde estava o IMPORTRANGE)
    shCadastro.getRange(2, 1, dadosContas.length, dadosContas[0].length).setValues(dadosContas);
    SpreadsheetApp.getActive().toast(`✅ Contas Contábeis importadas.`, "Importação", 3);
  } catch (e) {
    Logger.log("Erro ao importar Contas Contábeis: " + e);
  }

  // 2. Importa e fixa os Projetos (AK2 em diante)
  try {
    const dadosProjetos = SpreadsheetApp.openById(sourceIdProjetos).getRange(sourceRangeProjetos).getValues();
    shCadastro.getRange("AK2").clearContent(); // Limpa AK2 (onde estava o IMPORTRANGE)
    shCadastro.getRange(2, 37, dadosProjetos.length, dadosProjetos[0].length).setValues(dadosProjetos); // 37 é a coluna AK
    SpreadsheetApp.getActive().toast(`✅ Projetos importados.`, "Importação", 3);
  } catch (e) {
    Logger.log("Erro ao importar Projetos: " + e);
  }
  
  // NOVO: Limpa o cache após importar dados novos
  limparCacheCadastro();
}
function pintarTudoLD() {
  const sh = SHEET_LD;
  if (!sh) return;

  // --- OTIMIZAÇÃO: Usa a função utilitária para obter a última linha com dados
  const lastRow = getLastDataRow_(sh);
  const startRow = HEADER_ROW_LD + 1; // 5

  if (lastRow < startRow) {
    // Se não houver dados, limpa a formatação de todas as linhas de dados até o final da planilha.
    // Isso evita formatação de sobra se a planilha foi encolhida.
    sh.getRange(`${startRow}:${sh.getMaxRows()}`).setBackground(null);
    SpreadsheetApp.getActive().toast("⚠ Sem dados para formatar (apenas cabeçalho)", "LD", 3);
    return;
  }

  const numRows = lastRow - startRow + 1;

  const colS = 19; // Coluna S
  const colP = 16; // Coluna P

  // --- Intervalos específicos a pintar (ajusta o final deles para a nova lastRow)
  const intervalos = [
    `D${startRow}:D${lastRow}`, `F${startRow}:F${lastRow}`, `P${startRow}:S${lastRow}`,
    // Os intervalos fixos que você tinha, devem ser mantidos, pois são células específicas
    "B2182:C2183", "B2331:C2334", "C2574",
    "C2832:C2833", "C2880:C2882", "C2912",
    "W3006:W3007", "C3256:C3257", "B3287"
  ];

  // Se o lastRow for menor que os intervalos fixos, eles serão ignorados (o que é correto)

  // --- Leitura em lote
  const valoresS = sh.getRange(startRow, colS, numRows).getValues();
  const valoresP = sh.getRange(startRow, colP, numRows).getValues();

  // --- 1) Mapeia a cor da linha (saldo negativo)
  const corLinha = [];
  for (let i = 0; i < numRows; i++) {
    corLinha[i] = valoresP[i][0] < 0 ? "#b7e1cd" : "white";
  }

  // --- 2) Pinta toda a área A:W com a cor base (saldo negativo ou branca)
  const rangeTotal = sh.getRange(`A${startRow}:W${lastRow}`);
  const bgTotal = corLinha.map(cor => Array(23).fill(cor));
  rangeTotal.setBackgrounds(bgTotal);

  // Limpa a formatação abaixo da última linha de dados (IMPORTANTE)
  if (lastRow < sh.getLastRow()) {
    sh.getRange(`A${lastRow + 1}:W${sh.getLastRow()}`).setBackground(null);
  }

  // --- 3) Define cores específicas conforme coluna S
  const coresPorLinha = Array(numRows).fill(null);
  for (let i = 0; i < numRows; i++) {
    const S = valoresS[i][0];
    const P = valoresP[i][0];
    let cor = null;

    if (S === "À pagar - CONFIRMAR VALOR !!") cor = "#ff60ff";
    else if (S === "Pago - CHEQUE !!") cor = "#00ffff";
    else if (S === "SALDO") cor = "#9900ff";
    else if (typeof S === "string" && S.includes("À pagar - CARTÃO DE CRÉDITO")) cor = "#46bdc6";
    else if (typeof S === "string" && S.includes("Agendado")) cor = "#6d9eeb";
    else if (S === "Pago" && P < 0) cor = "#43f643";
    else if (S === "Pago" && P > 0) cor = "#dd7e6b";

    coresPorLinha[i] = cor; // se null, mantém cor da linha
  }

  // --- 4) Aplica cores específicas sobre os intervalos
  // A lógica de aplicação de cor foi mantida, mas ajustada para os novos intervalos
  intervalos.forEach(ref => {
    const range = sh.getRange(ref);
    const height = range.getNumRows();
    const width = range.getNumColumns();
    const firstRow = range.getRow(); // Pega a primeira linha real do intervalo
    const offset = firstRow - startRow; // Offset em relação ao array de dados (startRow=5)
    const bg = [];

    // Se o intervalo está fora do array de dados (ex: uma célula 3287 fora do lastRow), ignora.
    if (offset < 0 || offset >= numRows) {
      // Para intervalos fixos que podem estar abaixo do lastRow dinâmico,
      // garantimos que eles sejam pintados com a cor base (white) ou null.
      // Já limpamos o resto da planilha acima, e o código original não limpava os fixos,
      // então vamos manter o comportamento de limpeza apenas para o range principal.
      return;
    }

    for (let i = 0; i < height; i++) {
      const linhaDoArray = offset + i;

      // Checa se a linha do intervalo está dentro do array de dados lido
      if (linhaDoArray < 0 || linhaDoArray >= numRows) {
        // Isso pode ocorrer em intervalos fixos se eles caírem fora do range lido.
        // Para estes casos, vamos apenas sair do loop (ou pintar de branco/null)
        break; // Sai do loop para este intervalo
      }

      const cor = coresPorLinha[linhaDoArray];
      if (cor) {
        bg.push(Array(width).fill(cor));
      } else {
        // mantém a cor da linha (verde ou branca)
        const corBase = corLinha[linhaDoArray];
        bg.push(Array(width).fill(corBase));
      }
    }

    // Aplica a formatação apenas se houver cores definidas (o que exclui o break acima)
    if (bg.length > 0) {
      // Se o tamanho for menor (devido ao break), ajusta o range para evitar erro
      const rangeToSet = sh.getRange(firstRow, range.getColumn(), bg.length, width);
      rangeToSet.setBackgrounds(bg);
    }
  });

  SpreadsheetApp.getActive().toast("✅ Formatação aplicada com sucesso", "LD", 3);
  SpreadsheetApp.flush();
}
// ======================================================
// 🔹 FUNÇÃO MANUAL: Reaplicar dados do CADASTRO em Lote
// ======================================================
function reaplicarDadosCadastroLD() {
  const sh = SHEET_LD;
  const cadastro = SHEET_CAD;
  if (!sh || !cadastro) return;

  const headerRow = HEADER_ROW_LD;
  // Reusa a função otimizada para a última linha
  const lastRow = getLastDataRow_(sh); 
  
  if (lastRow <= headerRow) {
    SpreadsheetApp.getActive().toast("⚠ Sem dados para atualização.", "Atualização em Lote", 3);
    return;
  }

  const startRow = headerRow + 1;
  const numRows = lastRow - headerRow;

  // 1. Leitura em lote de todas as colunas de destino: G:L (6 colunas)
  const dadosLD_GL = sh.getRange(startRow, 7, numRows, 6).getValues(); // G (col 7) até L (col 12)

  // 2. Leitura dos dados de cadastro (Projetos e Contas Contábeis)
  const cadastroProj = cadastro.getRange('AY5:BA' + cadastro.getLastRow()).getValues(); // AY:AZ:BA
  const cadastroContas = cadastro.getRange('J5:L' + cadastro.getLastRow()).getValues(); // J:K:L

  const novosDadosLD_GL = [];

  for (let i = 0; i < numRows; i++) {
    let linha = dadosLD_GL[i]; // [G, H, I, J, K, L]

    const chaveProj = linha[0]; // G (usado como chave de lookup para Projeto)
    const chaveConta = linha[3]; // J (usado como chave de lookup para Conta Contábil)

    // A. Processar Projeto (G:H:I)
    if (chaveProj && chaveProj.toString().trim() !== "") {
      for (let j = 0; j < cadastroProj.length; j++) {
        // Assume que a primeira coluna do cadastroProj (AY) é a chave principal.
        if (cadastroProj[j][0] && cadastroProj[j][0].toString().trim() === chaveProj.toString().trim()) {
          linha[0] = cadastroProj[j][0]; // G
          linha[1] = cadastroProj[j][1]; // H
          linha[2] = cadastroProj[j][2]; // I
          break;
        }
      }
    }

    // B. Processar Conta Contábil (J:K:L)
    if (chaveConta && chaveConta.toString().trim() !== "") {
      for (let j = 0; j < cadastroContas.length; j++) {
        // Assume que a primeira coluna do cadastroContas (J) é a chave principal.
        if (cadastroContas[j][0] && cadastroContas[j][0].toString().trim() === chaveConta.toString().trim()) {
          linha[3] = cadastroContas[j][0]; // J
          linha[4] = cadastroContas[j][1]; // K
          linha[5] = cadastroContas[j][2]; // L
          break;
        }
      }
    }

    novosDadosLD_GL.push(linha);
  }

  // 3. Grava os novos valores em lote (G:L)
  if (novosDadosLD_GL.length > 0) {
    sh.getRange(startRow, 7, novosDadosLD_GL.length, 6).setValues(novosDadosLD_GL); // G:L
    SpreadsheetApp.getActive().toast("✅ Dados de PROJETO e CONTA em LD atualizados conforme CADASTRO.", "Atualização em Lote", 5);
  }
}
