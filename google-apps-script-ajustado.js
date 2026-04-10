// ============================================
// GOOGLE APPS SCRIPT - REGISTRO DE PONTO (CORRIGIDO)
// ============================================
const SPREADSHEET_ID = '1VPZ0PrC7EwDQv_Zy1FhjyDiw478H9i-G_S0KtZJ1vv4';

function doGet(e) {
  try {
    const params = e?.parameter || {};
    const acao = params.acao || 'registrar';
    
    console.log("📥 Evento GET - Ação:", acao);
    
    if (acao === 'consultar_colaboradores') {
      return consultarColaboradores(params);
    }
    
    if (acao === 'consultar') {
      return consultarRegistros(params);
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName("Registros");
    
    if (!sheet) {
      sheet = ss.insertSheet("Registros");
      criarCabecalho(sheet);
    }
    
    if (sheet.getLastRow() === 0) {
      criarCabecalho(sheet);
    }
    
    if (acao === 'resetar') {
      const dataHoje = params.data;
      const colaborador = params.colaborador;
      
      if (dataHoje && colaborador) {
        removerRegistrosDoDia(sheet, dataHoje, colaborador);
        
        const sheetResumo = ss.getSheetByName("Resumo");
        if (sheetResumo) {
          removerDoResumo(sheetResumo, dataHoje, colaborador);
        }
      }
      
      const callback = params.callback || 'callback';
      const resposta = JSON.stringify({status: 'success', message: 'Registros resetados!'});
      return ContentService
        .createTextOutput(callback + '(' + resposta + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    
    // Garantir que a data está no formato YYYY-MM-DD
    const dataParaSalvar = params.data || new Date().toISOString().split('T')[0];
    
    const novaLinha = [
      params.colaborador || 'NÃO INFORMADO',
      dataParaSalvar,
      params.marcacao1 || '',
      params.local1 || '',
      params.marcacao2 || '',
      params.local2 || '',
      params.marcacao3 || '',
      params.local3 || '',
      params.marcacao4 || '',
      params.local4 || '',
      params.tipo_marcacao || '',
      params.iso_timestamp || new Date().toISOString(),
      new Date().toLocaleString('pt-BR')
    ];
    
    sheet.appendRow(novaLinha);
    console.log("✅ Linha adicionada na planilha!");
    
    atualizarAbasResumoEColaboradores(ss, params);
    
    const callback = params.callback || 'callback';
    const resposta = JSON.stringify({status: 'success', message: 'Ponto registrado!'});
    
    return ContentService
      .createTextOutput(callback + '(' + resposta + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
      
  } catch(error) {
    console.error("❌ Erro:", error.toString());
    const callback = e?.parameter?.callback || 'callback';
    const resposta = JSON.stringify({status: 'error', message: error.toString()});
    return ContentService
      .createTextOutput(callback + '(' + resposta + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
}

// Função auxiliar para converter data para ISO (YYYY-MM-DD)
function converterDataParaIso(valor) {
  if (!valor && valor !== 0) return '';
  
  // Se já é uma data no formato YYYY-MM-DD
  if (typeof valor === 'string' && valor.match(/^\d{4}-\d{2}-\d{2}$/)) {
    return valor;
  }
  
  // Se é um objeto Date
  if (valor instanceof Date) {
    const ano = valor.getFullYear();
    const mes = String(valor.getMonth() + 1).padStart(2, '0');
    const dia = String(valor.getDate()).padStart(2, '0');
    return `${ano}-${mes}-${dia}`;
  }
  
  // Converter de string
  const texto = String(valor).trim();
  
  // Formato DD/MM/YYYY
  const ddmmyyyy = texto.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
  if (ddmmyyyy) {
    const dia = ddmmyyyy[1].padStart(2, '0');
    const mes = ddmmyyyy[2].padStart(2, '0');
    const ano = ddmmyyyy[3];
    return `${ano}-${mes}-${dia}`;
  }
  
  // Formato DDMMYYYY
  const compact = texto.match(/^(\d{2})(\d{2})(\d{4})$/);
  if (compact) {
    return `${compact[3]}-${compact[2]}-${compact[1]}`;
  }
  
  // Tentar parse como Date
  const parsed = new Date(texto);
  if (!isNaN(parsed.getTime())) {
    return parsed.toISOString().split('T')[0];
  }
  
  return texto;
}

function consultarColaboradores(params) {
  try {
    console.log("👥 Consultando colaboradores...");
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Colaboradores");
    
    if (!sheet) {
      console.log("⚠️ Aba Colaboradores não encontrada, retornando padrão");
      const callback = params.callback || 'callback';
      const resposta = JSON.stringify({status: 'success', colaboradores: [{nome: 'Funcionário Padrão'}]});
      return ContentService
        .createTextOutput(callback + '(' + resposta + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    
    const todasLinhas = sheet.getDataRange().getValues();
    const colaboradores = [];
    
    // Pular cabeçalho (índice 0)
    for (let i = 1; i < todasLinhas.length; i++) {
      if (todasLinhas[i][0] && String(todasLinhas[i][0]).trim() !== '') {
        colaboradores.push({ nome: String(todasLinhas[i][0]).trim() });
      }
    }
    
    // Remover duplicatas
    const nomesUnicos = [...new Set(colaboradores.map(c => c.nome))];
    const colaboradoresUnicos = nomesUnicos.map(nome => ({ nome }));
    
    if (colaboradoresUnicos.length === 0) {
      colaboradoresUnicos.push({ nome: 'Funcionário Padrão' });
    }
    
    console.log("👥 Colaboradores encontrados:", colaboradoresUnicos.length);
    
    const callback = params.callback || 'callback';
    const resposta = JSON.stringify({status: 'success', colaboradores: colaboradoresUnicos});
    
    return ContentService
      .createTextOutput(callback + '(' + resposta + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
      
  } catch(error) {
    console.error("❌ Erro ao consultar colaboradores:", error);
    const callback = params?.callback || 'callback';
    const resposta = JSON.stringify({status: 'error', colaboradores: [{nome: 'Funcionário Padrão'}]});
    return ContentService
      .createTextOutput(callback + '(' + resposta + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
}

// ============================================
// FUNÇÃO CONSULTAR REGISTROS - CORRIGIDA
// ============================================
function consultarRegistros(params) {
  try {
    console.log("🔍 Iniciando consulta de registros");
    
    if (!params) params = {};
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Registros");
    
    if (!sheet) {
      console.log("❌ Aba Registros não encontrada");
      const callback = params.callback || 'callback';
      const resposta = JSON.stringify({status: 'error', registros: [], message: 'Aba não encontrada'});
      return ContentService
        .createTextOutput(callback + '(' + resposta + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    
    const todasLinhas = sheet.getDataRange().getValues();
    console.log("📊 Total de linhas:", todasLinhas.length);
    
    if (todasLinhas.length <= 1) {
      console.log("📋 Planilha vazia");
      const callback = params.callback || 'callback';
      const resposta = JSON.stringify({status: 'success', registros: []});
      return ContentService
        .createTextOutput(callback + '(' + resposta + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    
    const dataInicio = params.dataInicio ? converterDataParaIso(params.dataInicio) : null;
    const dataFim = params.dataFim ? converterDataParaIso(params.dataFim) : null;
    const colaborador = params.colaborador ? String(params.colaborador).trim() : 'todos';
    const debug = params.debug === 'true';
    
    // Mapear cabeçalho
    const header = todasLinhas[0].map(c => String(c || '').trim().toLowerCase());
    
    function findHeaderIndex(names, fallback) {
      for (const name of names) {
        const idx = header.indexOf(name.toLowerCase());
        if (idx !== -1) return idx;
      }
      return fallback;
    }

    const idxColaborador = findHeaderIndex(['colaborador'], 0);
    const idxData = findHeaderIndex(['data'], 1);
    const idxMarcacao1 = findHeaderIndex(['1ª marcação', '1ª marcação (início)', '1º marcação', 'marcação 1'], 2);
    const idxLocal1 = findHeaderIndex(['local 1ª marcação', 'local 1ª', 'local 1'], 3);
    const idxMarcacao2 = findHeaderIndex(['2ª marcação', '2ª marcação (saída almoço)', '2º marcação', 'marcação 2'], 4);
    const idxLocal2 = findHeaderIndex(['local 2ª marcação', 'local 2ª', 'local 2'], 5);
    const idxMarcacao3 = findHeaderIndex(['3ª marcação', '3ª marcação (retorno almoço)', '3º marcação', 'marcação 3'], 6);
    const idxLocal3 = findHeaderIndex(['local 3ª marcação', 'local 3ª', 'local 3'], 7);
    const idxMarcacao4 = findHeaderIndex(['4ª marcação', '4ª marcação (fim expediente)', '4º marcação', 'marcação 4'], 8);
    const idxLocal4 = findHeaderIndex(['local 4ª marcação', 'local 4ª', 'local 4'], 9);

    console.log("📌 Índices:", {idxColaborador, idxData, idxMarcacao1, idxMarcacao2, idxMarcacao3, idxMarcacao4});
    console.log("🔎 Filtros - Início:", dataInicio, "Fim:", dataFim, "Colaborador:", colaborador);

    const registros = [];

    function formatarHora(valor) {
      if (!valor) return '';
      if (valor instanceof Date) {
        return valor.toLocaleTimeString('pt-BR', {hour: '2-digit', minute: '2-digit', second: '2-digit'});
      }
      if (typeof valor === 'string') return valor.trim();
      return String(valor);
    }

    for (let i = 1; i < todasLinhas.length; i++) {
      const linha = todasLinhas[i];
      
      // Pular linhas vazias
      if (!linha[idxColaborador] && !linha[idxData]) continue;
      
      const dataComparavel = converterDataParaIso(linha[idxData]);
      let dataFormatada = dataComparavel;
      
      // Formatar para exibição DD/MM/YYYY
      if (dataComparavel.match(/^(\d{4})-(\d{2})-(\d{2})$/)) {
        const [ano, mes, dia] = dataComparavel.split('-');
        dataFormatada = `${dia}/${mes}/${ano}`;
      }

      const reg = {
        colaborador: String(linha[idxColaborador] || '').trim(),
        data: dataFormatada,
        marcacao1: formatarHora(linha[idxMarcacao1]),
        local1: String(linha[idxLocal1] || '').trim(),
        marcacao2: formatarHora(linha[idxMarcacao2]),
        local2: String(linha[idxLocal2] || '').trim(),
        marcacao3: formatarHora(linha[idxMarcacao3]),
        local3: String(linha[idxLocal3] || '').trim(),
        marcacao4: formatarHora(linha[idxMarcacao4]),
        local4: String(linha[idxLocal4] || '').trim(),
        tipo_marcacao: String(linha[10] || '').trim()
      };

      // Filtrar por colaborador
      if (colaborador && colaborador.toLowerCase() !== 'todos' && 
          reg.colaborador.toLowerCase() !== colaborador.toLowerCase()) {
        if (debug) console.log(`⏭️ Colaborador não corresponde: "${reg.colaborador}" !== "${colaborador}"`);
        continue;
      }

      // Filtrar por data
      if (dataInicio && dataComparavel < dataInicio) {
        if (debug) console.log(`⏭️ Data antes do início: ${dataComparavel} < ${dataInicio}`);
        continue;
      }
      if (dataFim && dataComparavel > dataFim) {
        if (debug) console.log(`⏭️ Data após o fim: ${dataComparavel} > ${dataFim}`);
        continue;
      }

      if (debug) console.log(`✅ Registro aceito: ${reg.colaborador} - ${dataComparavel}`);
      registros.push(reg);
    }
    
    console.log("📋 Registros após filtro:", registros.length);
    
    // Agrupar por data + colaborador
    const agrupado = {};
    registros.forEach(reg => {
      const chave = reg.data + '_' + reg.colaborador;
      if (!agrupado[chave]) {
        agrupado[chave] = reg;
      } else {
        // Manter a primeira marcação de cada tipo
        if (reg.marcacao1 && !agrupado[chave].marcacao1) agrupado[chave].marcacao1 = reg.marcacao1;
        if (reg.marcacao2 && !agrupado[chave].marcacao2) agrupado[chave].marcacao2 = reg.marcacao2;
        if (reg.marcacao3 && !agrupado[chave].marcacao3) agrupado[chave].marcacao3 = reg.marcacao3;
        if (reg.marcacao4 && !agrupado[chave].marcacao4) agrupado[chave].marcacao4 = reg.marcacao4;
      }
    });
    
    const resultado = Object.values(agrupado);
    
    // Ordenar por data (mais recente primeiro)
    resultado.sort((a, b) => {
      const dataA = a.data.split('/').reverse().join('-');
      const dataB = b.data.split('/').reverse().join('-');
      return dataB.localeCompare(dataA);
    });
    
    console.log("✅ Registros agrupados:", resultado.length);
    
    const callback = params.callback || 'callback';
    const respostaObj = {
      status: 'success',
      registros: resultado
    };
    
    if (debug) {
      respostaObj.debugInfo = {
        totalRows: todasLinhas.length - 1,
        registrosEncontrados: resultado.length,
        idxColaborador,
        idxData,
        dataInicio,
        dataFim,
        colaborador
      };
    }
    
    const resposta = JSON.stringify(respostaObj);
    
    return ContentService
      .createTextOutput(callback + '(' + resposta + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
      
  } catch(error) {
    console.error("❌ Erro na consulta:", error.toString());
    const callback = params?.callback || 'callback';
    const resposta = JSON.stringify({status: 'error', registros: [], message: error.toString()});
    return ContentService
      .createTextOutput(callback + '(' + resposta + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
}

function removerRegistrosDoDia(sheet, data, colaborador) {
  const dados = sheet.getDataRange().getValues();
  const linhasParaDeletar = [];
  
  // Procurar de baixo para cima para não bagunçar os índices
  for (let i = dados.length - 1; i >= 1; i--) {
    let dataLinha = converterDataParaIso(dados[i][1]); // Índice 1 = coluna Data
    
    if (dataLinha === data && dados[i][0] === colaborador) { // Índice 0 = Colaborador
      linhasParaDeletar.push(i + 1); // +1 porque sheet.deleteRow usa índice 1-based
    }
  }
  
  // Ordenar decrescente para deletar de baixo para cima
  linhasParaDeletar.sort((a, b) => b - a);
  linhasParaDeletar.forEach(linha => sheet.deleteRow(linha));
  
  console.log(`🗑️ ${linhasParaDeletar.length} registros removidos dos Registros`);
}

function removerDoResumo(sheetResumo, data, colaborador) {
  const dados = sheetResumo.getDataRange().getValues();
  const linhasParaDeletar = [];
  
  // Estrutura do Resumo: Data (col 0), Colaborador (col 1), Entrada, Saída, Status
  for (let i = dados.length - 1; i >= 1; i--) {
    let dataLinha = converterDataParaIso(dados[i][0]); // Índice 0 = Data
    
    if (dataLinha === data && dados[i][1] === colaborador) { // Índice 1 = Colaborador
      linhasParaDeletar.push(i + 1);
    }
  }
  
  linhasParaDeletar.sort((a, b) => b - a);
  linhasParaDeletar.forEach(linha => sheetResumo.deleteRow(linha));
  
  console.log(`🗑️ ${linhasParaDeletar.length} registros removidos do Resumo`);
}

function atualizarAbasResumoEColaboradores(ss, dados) {
  try {
    // Atualizar aba Colaboradores
    let sheetColabs = ss.getSheetByName("Colaboradores");
    if (!sheetColabs) {
      sheetColabs = ss.insertSheet("Colaboradores");
      sheetColabs.getRange(1, 1, 1, 4)
        .setValues([['Nome', 'Data Cadastro', 'Total Registros', 'Último Registro']])
        .setFontWeight('bold').setBackground('#620000').setFontColor('#FFFFFF');
    }
    
    const dadosColabs = sheetColabs.getDataRange().getValues();
    const nomeColab = dados.colaborador;
    let encontrou = false;
    
    for (let i = 1; i < dadosColabs.length; i++) {
      if (dadosColabs[i][0] === nomeColab) {
        const totalAtual = parseInt(dadosColabs[i][2]) || 0;
        sheetColabs.getRange(i + 1, 3).setValue(totalAtual + 1);
        sheetColabs.getRange(i + 1, 4).setValue(new Date().toLocaleString('pt-BR'));
        encontrou = true;
        break;
      }
    }
    
    if (!encontrou) {
      sheetColabs.appendRow([
        nomeColab,
        new Date().toLocaleDateString('pt-BR'),
        1,
        new Date().toLocaleString('pt-BR')
      ]);
    }
    
    // Atualizar aba Resumo
    let sheetResumo = ss.getSheetByName("Resumo");
    if (!sheetResumo) {
      sheetResumo = ss.insertSheet("Resumo");
      sheetResumo.getRange(1, 1, 1, 5)
        .setValues([['Data', 'Colaborador', 'Entrada', 'Saída', 'Status']])
        .setFontWeight('bold').setBackground('#620000').setFontColor('#FFFFFF');
    }
    
    // Garantir formato YYYY-MM-DD
    const dataResumo = dados.data || new Date().toISOString().split('T')[0];
    
    // Remover registro existente do dia para este colaborador
    removerDoResumo(sheetResumo, dataResumo, nomeColab);
    
    const status = (dados.marcacao1 && dados.marcacao2 && dados.marcacao3 && dados.marcacao4) 
      ? '✅ Completo' : '⏳ Parcial';
    
    sheetResumo.appendRow([
      dataResumo,
      nomeColab,
      dados.marcacao1 || '—',
      dados.marcacao4 || '—',
      status
    ]);
    
    console.log("✅ Abas atualizadas!");
    
  } catch(error) {
    console.error("Erro ao atualizar abas:", error);
  }
}

function criarCabecalho(sheet) {
  const cabecalho = [
    'Colaborador', 'Data',
    '1ª Marcação', 'Local 1ª',
    '2ª Marcação', 'Local 2ª',
    '3ª Marcação', 'Local 3ª',
    '4ª Marcação', 'Local 4ª',
    'Tipo', 'Timestamp ISO', 'Registro'
  ];
  
  sheet.getRange(1, 1, 1, cabecalho.length)
    .setValues([cabecalho])
    .setFontWeight('bold')
    .setBackground('#620000')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');
  
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 13);
}

function testarConsulta() {
  const hoje = new Date();
  const dataHoje = hoje.toISOString().split('T')[0];
  
  const params = {
    acao: 'consultar',
    dataInicio: dataHoje,
    dataFim: dataHoje,
    colaborador: 'todos',
    callback: 'test',
    debug: 'true'
  };
  
  console.log("🧪 Testando consulta para data:", dataHoje);
  const resultado = consultarRegistros(params);
  console.log("Resultado:", resultado.getContent());
  return resultado.getContent();
}

function inicializarPlanilha() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    let sheet = ss.getSheetByName("Registros");
    if (!sheet) { 
      sheet = ss.insertSheet("Registros"); 
      criarCabecalho(sheet); 
    }
    
    if (!ss.getSheetByName("Resumo")) {
      const r = ss.insertSheet("Resumo");
      r.getRange(1,1,1,5).setValues([['Data','Colaborador','Entrada','Saída','Status']])
       .setFontWeight('bold').setBackground('#620000').setFontColor('#FFFFFF');
    }
    
    if (!ss.getSheetByName("Colaboradores")) {
      const c = ss.insertSheet("Colaboradores");
      c.getRange(1,1,1,4).setValues([['Nome','Data Cadastro','Total Registros','Último Registro']])
       .setFontWeight('bold').setBackground('#620000').setFontColor('#FFFFFF');
    }
    
    console.log("✅ Planilha inicializada!");
    return "Planilha configurada!";
  } catch(error) {
    return "Erro: " + error.toString();
  }
}