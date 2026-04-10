// ============================================
// GOOGLE APPS SCRIPT - REGISTRO DE PONTO (AJUSTADO)
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
    
    // AJUSTE: Salvar data sempre como YYYY-MM-DD para consistência
    const dataParaSalvar = params.data || new Date().toISOString().split('T')[0];
    
    const novaLinha = [
      params.colaborador || 'NÃO INFORMADO',
      dataParaSalvar, // YYYY-MM-DD
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
    
    for (let i = 1; i < todasLinhas.length; i++) {
      if (todasLinhas[i][0] && String(todasLinhas[i][0]).trim() !== '') {
        colaboradores.push({ nome: String(todasLinhas[i][0]).trim() });
      }
    }
    
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
// FUNÇÃO CONSULTAR REGISTROS - AJUSTADA PARA COMPARAR DATAS COMO STRING YYYY-MM-DD
// ============================================
function consultarRegistros(params) {
  try {
    console.log("🔍 Iniciando consulta de registros");
    
    // GARANTIR QUE PARAMS É UM OBJETO
    if (!params) {
      params = {};
    }
    
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
    
    // Garantir valores padrão
    function converterDataParaIso(valor) {
      if (!valor && valor !== 0) return '';
      if (valor instanceof Date) {
        const ano = valor.getFullYear();
        const mes = String(valor.getMonth() + 1).padStart(2, '0');
        const dia = String(valor.getDate()).padStart(2, '0');
        return `${ano}-${mes}-${dia}`;
      }
      const texto = String(valor).trim();
      if (texto.match(/^\d{4}-\d{2}-\d{2}$/)) {
        return texto;
      }
      const ddmmyyyy = texto.match(/^(\d{1,2})\s*[\/\-\.]\s*(\d{1,2})\s*[\/\-\.]\s*(\d{4})$/);
      if (ddmmyyyy) {
        const dia = ddmmyyyy[1].padStart(2, '0');
        const mes = ddmmyyyy[2].padStart(2, '0');
        const ano = ddmmyyyy[3];
        return `${ano}-${mes}-${dia}`;
      }
      const compact = texto.match(/^(\d{2})(\d{2})(\d{4})$/);
      if (compact) {
        return `${compact[3]}-${compact[2]}-${compact[1]}`;
      }
      const parsed = new Date(texto);
      if (!isNaN(parsed.getTime())) {
        return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, '0')}-${String(parsed.getDate()).padStart(2, '0')}`;
      }
      return texto;
    }
    const dataInicio = params.dataInicio ? converterDataParaIso(params.dataInicio) : null;
    const dataFim = params.dataFim ? converterDataParaIso(params.dataFim) : null;
    const colaborador = params.colaborador ? String(params.colaborador).trim() : 'todos';
    const debug = params.debug === 'true';
    
    const header = todasLinhas[0].map(c => String(c || '').trim().toLowerCase());
    function findHeaderIndex(names, fallback) {
      for (const name of names) {
        const idx = header.indexOf(name);
        if (idx !== -1) return idx;
      }
      return fallback;
    }

    const idxColaborador = findHeaderIndex(['colaborador'], 0);
    const idxData = findHeaderIndex(['data'], 1);
    const idxMarcacao1 = findHeaderIndex(['1ª marcação', '1ª marcação (início)', '1º marcação', 'marcação 1', 'primeira marcação'], 2);
    const idxLocal1 = findHeaderIndex(['local 1ª marcação', 'local 1ª', 'local 1', 'local primeira marcação'], 3);
    const idxMarcacao2 = findHeaderIndex(['2ª marcação', '2ª marcação (saída almoço)', '2º marcação', 'marcação 2', 'segunda marcação'], 4);
    const idxLocal2 = findHeaderIndex(['local 2ª marcação', 'local 2ª', 'local 2', 'local segunda marcação'], 5);
    const idxMarcacao3 = findHeaderIndex(['3ª marcação', '3ª marcação (retorno almoço)', '3º marcação', 'marcação 3', 'terceira marcação'], 6);
    const idxLocal3 = findHeaderIndex(['local 3ª marcação', 'local 3ª', 'local 3', 'local terceira marcação'], 7);
    const idxMarcacao4 = findHeaderIndex(['4ª marcação', '4ª marcação (fim expediente)', '4º marcação', 'marcação 4', 'quarta marcação'], 8);
    const idxLocal4 = findHeaderIndex(['local 4ª marcação', 'local 4ª', 'local 4', 'local quarta marcação'], 9);

    console.log("📋 Header detectado:", header);
    console.log("📌 Índices usados: colaborador=", idxColaborador, "data=", idxData, "marcação1=", idxMarcacao1, "marcação2=", idxMarcacao2, "marcação3=", idxMarcacao3, "marcação4=", idxMarcacao4);
    console.log("🔎 Filtros - Início:", dataInicio, "Fim:", dataFim, "Colaborador:", colaborador, "Debug:", debug);

    const registros = [];

    function formatarHora(valor) {
      if (!valor) return '';
      if (valor instanceof Date) {
        const horas = String(valor.getHours()).padStart(2, '0');
        const minutos = String(valor.getMinutes()).padStart(2, '0');
        const segundos = String(valor.getSeconds()).padStart(2, '0');
        return horas + ':' + minutos + ':' + segundos;
      } else if (typeof valor === 'string') {
        return valor.trim();
      }
      return String(valor);
    }

    for (let i = 1; i < todasLinhas.length; i++) {
      const linha = todasLinhas[i];

      if (debug) {
        console.log(`🔍 Processando linha ${i}:`, linha);
      }

      const dataComparavel = converterDataParaIso(linha[idxData]);
      let dataFormatada = dataComparavel;
      if (dataComparavel.match(/^(\d{4})-(\d{2})-(\d{2})$/)) {
        const [ano, mes, dia] = dataComparavel.split('-');
        dataFormatada = `${dia}/${mes}/${ano}`;
      }
      if (debug) {
        console.log(`📅 Data convertida: '${linha[idxData]}' -> '${dataComparavel}'`);
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
      if (colaborador && colaborador.toLowerCase() !== 'todos' && reg.colaborador !== colaborador) {
        if (debug) console.log(`⏭️ Colaborador não corresponde: "${reg.colaborador}" !== "${colaborador}"`);
        continue;
      }

      // Filtrar por data - COMPARAÇÃO SIMPLES COMO STRING YYYY-MM-DD
      if (dataInicio && dataComparavel < dataInicio) {
        if (debug) console.log(`⏭️ Data antes do início: ${dataComparavel} < ${dataInicio}`);
        continue;
      }
      if (dataFim && dataComparavel > dataFim) {
        if (debug) console.log(`⏭️ Data após o fim: ${dataComparavel} > ${dataFim}`);
        continue;
      }

      if (debug) console.log(`✅ Registro aceito: ${reg.colaborador} - ${dataComparavel}`);
      console.log(`✅ Aceito: ${reg.colaborador} - ${dataComparavel}`);
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
        if (reg.marcacao1) agrupado[chave].marcacao1 = reg.marcacao1;
        if (reg.marcacao2) agrupado[chave].marcacao2 = reg.marcacao2;
        if (reg.marcacao3) agrupado[chave].marcacao3 = reg.marcacao3;
        if (reg.marcacao4) agrupado[chave].marcacao4 = reg.marcacao4;
      }
    });
    
    const resultado = Object.values(agrupado);
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
  
  for (let i = dados.length - 1; i >= 1; i--) {
    let dataLinha = dados[i][1];
    if (dataLinha instanceof Date) {
      const ano = dataLinha.getFullYear();
      const mes = String(dataLinha.getMonth() + 1).padStart(2, '0');
      const dia = String(dataLinha.getDate()).padStart(2, '0');
      dataLinha = `${ano}-${mes}-${dia}`;
    }
    if (dataLinha === data && dados[i][0] === colaborador) {
      linhasParaDeletar.push(i + 1);
    }
  }
  
  linhasParaDeletar.forEach(linha => sheet.deleteRow(linha));
  console.log(`🗑️ ${linhasParaDeletar.length} registros removidos`);
}

function removerDoResumo(sheetResumo, data, colaborador) {
  const dados = sheetResumo.getDataRange().getValues();
  const linhasParaDeletar = [];
  
  for (let i = dados.length - 1; i >= 1; i--) {
    let dataLinha = dados[i][0];
    if (dataLinha instanceof Date) {
      const ano = dataLinha.getFullYear();
      const mes = String(dataLinha.getMonth() + 1).padStart(2, '0');
      const dia = String(dataLinha.getDate()).padStart(2, '0');
      dataLinha = `${ano}-${mes}-${dia}`;
    }
    if (dataLinha === data && dados[i][1] === colaborador) {
      linhasParaDeletar.push(i + 1);
    }
  }
  
  linhasParaDeletar.forEach(linha => sheetResumo.deleteRow(linha));
  console.log(`🗑️ ${linhasParaDeletar.length} registros removidos do Resumo`);
}

function atualizarAbasResumoEColaboradores(ss, dados) {
  try {
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
    
    let sheetResumo = ss.getSheetByName("Resumo");
    if (!sheetResumo) {
      sheetResumo = ss.insertSheet("Resumo");
      sheetResumo.getRange(1, 1, 1, 5)
        .setValues([['Data', 'Colaborador', 'Entrada', 'Saída', 'Status']])
        .setFontWeight('bold').setBackground('#620000').setFontColor('#FFFFFF');
    }
    
    // Ajustar data para YYYY-MM-DD no resumo também
    const dataResumo = dados.data || new Date().toISOString().split('T')[0];
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
  const dia = String(hoje.getDate()).padStart(2, '0');
  const mes = String(hoje.getMonth() + 1).padStart(2, '0');
  const ano = hoje.getFullYear();
  const dataHoje = ano + '-' + mes + '-' + dia;
  
  const params = {
    acao: 'consultar',
    dataInicio: dataHoje,
    dataFim: dataHoje,
    colaborador: 'todos',
    callback: 'test'
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
    if (!sheet) { sheet = ss.insertSheet("Registros"); criarCabecalho(sheet); }
    
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