/* ==========================================================
   CONFIGURAÇÃO INICIAL
   ========================================================== */
var NOME_ABA_USUARIOS = "Usuarios";
var NOME_ABA_MAQUINAS = "Máquinas";
var NOME_ABA_NOTIFICACOES = "Notificacoes";

/* ==========================================================
   FUNÇÃO PRINCIPAL DO WEB APP (GET)
   ========================================================== */
function doGet(e) {
 
  // ROTA DE IMPRESSÃO
  if (e.parameter.page === 'print') {
    Logger.log("Requisição para página de impressão recebida.");

    var filtroStatus = e.parameter.filtro || 'pendentes';
    var filtroMaquina = e.parameter.filtroMaquina || 'todas';
    var filtroNome = e.parameter.filtroNome || 'Pendentes';

    if (filtroMaquina !== 'todas') {
      filtroNome += " (Máquina: " + filtroMaquina + ")";
    }

    // Se o filtro for "realizados", busca do histórico
    var dados;
    if (filtroStatus === 'realizados') {
      dados = buscarDadosHistorico(filtroMaquina);
    } else {
      dados = buscarDadosManutencaoComFiltro(filtroStatus, filtroMaquina);
    }

    var tabelaHtml = formatarDadosParaImpressao(dados, filtroStatus); 
    
    var template = HtmlService.createTemplateFromFile('Print');
    template.dataAsHtmlTable = tabelaHtml;
    template.filtroNome = filtroNome; 
    
    return template.evaluate()
        .setTitle('Imprimir Relatório - MARFIM TÊXTIL')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  
  // ROTA PRINCIPAL (APP)
  } else {
    Logger.log("Web App acessado. Servindo Index.html");
    return HtmlService.createHtmlOutputFromFile('Index.html')
        .setTitle("MARFIM TÊXTIL - Controle de Manutenção")
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

/* ==========================================================
   FUNÇÕES AUXILIARES (getAppUrl, getListaDeMaquinas)
   ========================================================== */
function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

function getListaDeMaquinas() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaMaquinas = ss.getSheetByName(NOME_ABA_MAQUINAS);
    var range = abaMaquinas.getRange(2, 1, abaMaquinas.getLastRow() - 1, 1);
    var nomes = range.getValues();
    
    var listaLimpa = nomes.map(function(linha) {
      return linha[0];
    }).filter(function(nome) {
      return nome !== "";
    });
    
    Logger.log("Retornando " + listaLimpa.length + " nomes de máquinas para o filtro.");
    return listaLimpa;
    
  } catch (e) {
    Logger.log("Erro ao buscar lista de máquinas: " + e.message);
    return []; 
  }
}

/* ==========================================================
   FUNÇÃO: Formatar Relatório (ATUALIZADA: Sem Coluna "Status")
   ========================================================== */
function formatarDadosParaImpressao(dados, filtro) {
  var html = "";

  // Cabeçalhos da Tabela
  html += "<thead><tr>";
  html += '<th style="width: 50px; text-align: center;">OK</th>';
  html += "<th>Máquina</th>";
  html += "<th>Intervalo</th>";
  html += "<th>Itens</th>";

  if (filtro === "realizados") {
    html += "<th>Realizado Por</th>";
    html += "<th>Data</th>";
    html += '<th style="width: 120px;">Assinatura</th>'; // Nova coluna de assinatura
  } else if (filtro === "todos") {
    html += "<th>Próxima Manutenção</th>";
    html += "<th>Realizado Por</th>";
    html += "<th>Data da Confirmação</th>";
  } else {
    html += "<th>Próxima Manutenção</th>";
  }
  html += "</tr></thead>";

  // Corpo da Tabela
  html += "<tbody>";

  if (dados.length === 0) {
    var colspan = (filtro === "realizados") ? "7" : "7";
    html += '<tr style="text-align: center;"><td colspan="' + colspan + '">Nenhum dado encontrado para este filtro.</td></tr>';
  }

  dados.forEach(function(item) {
    html += "<tr>";
    html += '<td style="text-align: center;"><div style="width:20px; height:20px; border:1px solid #333; margin:auto;"></div></td>';
    html += "<td>" + item.maquina + "</td>";
    html += "<td>" + item.intervalo + "</td>";
    html += "<td>" + (item.itens || 'N/A') + "</td>";

    if (filtro === "realizados") {
      html += "<td>" + (item.realizadoPor || 'N/A') + "</td>";
      html += "<td>" + (item.dataConfirmacaoFormatada || 'N/A') + "</td>";
      // Espaço para assinatura - célula vazia com altura para assinatura
      html += '<td style="height: 40px; border: 1px solid #333;">&nbsp;</td>';

    } else if (filtro === "todos") {
      if (item.tipo === "Realizado") {
        html += "<td> - </td>";
        html += "<td>" + (item.realizadoPor || 'N/A') + "</td>";
        html += "<td>" + (item.dataConfirmacaoFormatada || 'N/A') + "</td>";
      } else { // é Pendente
        html += "<td>" + (item.proximaData || 'N/A') + "</td>";
        html += "<td> - </td>";
        html += "<td> - </td>";
      }

    } else {
      // Filtros de Pendentes
      html += "<td>" + (item.proximaData || 'N/A') + "</td>";
    }
    html += "</tr>";
  });

  html += "</tbody>";
  return html;
}

/* ==========================================================
   FUNÇÃO DE LOGIN
   ========================================================== */
function verificarLogin(usuario, senha) { 
  Logger.log("Tentativa de login para o usuário: " + usuario);
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaUsuarios = ss.getSheetByName(NOME_ABA_USUARIOS);
    var dados = abaUsuarios.getRange(2, 1, abaUsuarios.getLastRow() - 1, 4).getValues(); 
    
    for (var i = 0; i < dados.length; i++) {
      var linha = dados[i];
      if (String(linha[0]).trim() === String(usuario).trim() && String(linha[1]).trim() === String(senha).trim()) {
        Logger.log("Usuário encontrado: " + linha[2] + ", Função: " + linha[3]);
        return {
          status: "sucesso",
          nome: linha[2],
          funcao: linha[3]
        };
      }
    }
    
    Logger.log("Falha no login: usuário ou senha inválidos para " + usuario);
    return { status: "erro", mensagem: "Usuário ou senha inválidos." };
    
  } catch (e) {
    Logger.log("Erro na função verificarLogin: " + e.message);
    return { status: "erro", mensagem: "Erro no servidor: " + e.message };
  }
}

/* ==========================================================
   FUNÇÕES DE AÇÃO (ATUALIZADAS PARA USAR 'IDENTIFICADOR' E RETORNAR DADOS)
   ========================================================== */

function registrarManutencao(identificador, nomeUsuario, filtroStatusAtual, filtroMaquinaAtual) {
  Logger.log("--- INÍCIO REGISTRARMANUTENCAO ---");
  Logger.log("Número da linha recebido: '" + identificador + "' para usuário: '" + nomeUsuario + "'");
  Logger.log("Filtros atuais: Status=" + filtroStatusAtual + ", Maquina=" + filtroMaquinaAtual);

  if (!identificador) {
    Logger.log("ERRO FATAL: Número da linha é nulo ou vazio.");
    return { status: "erro", mensagem: "Número da linha não foi enviado." };
  }

  // --- NOVA LÓGICA: Usa o número da linha diretamente ---
  var linhaEncontrada = parseInt(identificador, 10);

  if (isNaN(linhaEncontrada) || linhaEncontrada < 2) {
    Logger.log("ERRO FATAL: Número da linha inválido: '" + identificador + "'");
    return { status: "erro", mensagem: "Número da linha inválido." };
  }

  Logger.log("Linha a ser atualizada: " + linhaEncontrada);
  // --- FIM DA NOVA LÓGICA ---

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaMaquinas = ss.getSheetByName(NOME_ABA_MAQUINAS);
    if (!abaMaquinas) {
      Logger.log("ERRO FATAL: Aba '" + NOME_ABA_MAQUINAS + "' não encontrada.");
      return { status: "erro", mensagem: "Aba da planilha de máquinas não foi encontrada." };
    }

    // Validar se a linha existe
    if (linhaEncontrada > abaMaquinas.getLastRow()) {
      Logger.log("ERRO: Linha " + linhaEncontrada + " não existe na planilha.");
      return { status: "erro", mensagem: "Linha " + linhaEncontrada + " não existe na planilha." };
    }
    
    // A partir daqui, a lógica é a mesma, pois já temos a 'linhaEncontrada'
    Logger.log("Linha " + linhaEncontrada + ": Lendo intervalo da Coluna B.");
    var intervaloDias = abaMaquinas.getRange(linhaEncontrada, 2).getValue();
    if (typeof intervaloDias !== 'number' || intervaloDias <= 0) {
      Logger.log("ERRO: Intervalo de dias na Coluna B é inválido: " + intervaloDias);
      return { status: "erro", mensagem: "Intervalo de dias (Coluna B) é inválido para a máquina: " + nomeMaquinaParaBuscar };
    }
    
    var dataConfirmacao = new Date();
    var novaProximaManutencao = new Date(dataConfirmacao);
    novaProximaManutencao.setDate(novaProximaManutencao.getDate() + intervaloDias);
    
    Logger.log("Linha " + linhaEncontrada + ": Fazendo backup da Coluna E -> Coluna I.");
    var dataAntiga = abaMaquinas.getRange(linhaEncontrada, 5).getValue(); // Col E
    abaMaquinas.getRange(linhaEncontrada, 9).setValue(dataAntiga);       // Col I

    Logger.log("Linha " + linhaEncontrada + ": Gravando novos dados (Colunas D, E, F, G, H).");
    abaMaquinas.getRange(linhaEncontrada, 4).setValue(dataConfirmacao);         // Col D
    abaMaquinas.getRange(linhaEncontrada, 5).setValue(novaProximaManutencao); // Col E
    abaMaquinas.getRange(linhaEncontrada, 6).setValue("Realizado");          // Col F
    abaMaquinas.getRange(linhaEncontrada, 7).setValue(nomeUsuario);          // Col G
    abaMaquinas.getRange(linhaEncontrada, 8).setValue(dataConfirmacao);      // Col H
    
    Logger.log("Linha " + linhaEncontrada + ": Forçando salvamento (flush)...");
    SpreadsheetApp.flush();

    // Aguarda um momento para garantir que o flush completou
    Utilities.sleep(100);

    Logger.log("--- SUCESSO! Manutenção registrada. ---");

    Logger.log("VERIFICAÇÃO: Lendo novamente a linha " + linhaEncontrada + " para confirmar gravação:");
    var statusGravado = abaMaquinas.getRange(linhaEncontrada, 6).getValue();
    Logger.log("VERIFICAÇÃO: Status na planilha (Col F) = '" + statusGravado + "'");

    Logger.log("Buscando dados atualizados para o cliente...");
    var dadosAtualizados = buscarDadosManutencaoComFiltro(filtroStatusAtual, filtroMaquinaAtual);

    Logger.log("VERIFICAÇÃO: Dados retornados contêm " + dadosAtualizados.length + " itens.");

    return {
      status: "sucesso",
      dados: dadosAtualizados
    };
    
  } catch (e) {
    Logger.log("--- ERRO GRAVE NO BLOCO 'TRY' ---");
    Logger.log("ERRO: " + e.message);
    Logger.log("STACK: " + e.stack);
    return { status: "erro", mensagem: "Erro inesperado no servidor: " + e.message };
  }
}

// ---------------------------------------------------------------------------------

function desfazerManutencao(identificador, filtroStatusAtual, filtroMaquinaAtual) {
  Logger.log("--- INÍCIO DESFAZERMANUTENCAO ---");
  Logger.log("Número da linha recebido: '" + identificador + "'");
  Logger.log("Filtros atuais: Status=" + filtroStatusAtual + ", Maquina=" + filtroMaquinaAtual);

  if (!identificador) {
    Logger.log("ERRO FATAL: Número da linha é nulo ou vazio.");
    return { status: "erro", mensagem: "Número da linha não foi enviado." };
  }

  // --- NOVA LÓGICA: Usa o número da linha diretamente ---
  var linhaEncontrada = parseInt(identificador, 10);

  if (isNaN(linhaEncontrada) || linhaEncontrada < 2) {
    Logger.log("ERRO FATAL: Número da linha inválido: '" + identificador + "'");
    return { status: "erro", mensagem: "Número da linha inválido." };
  }

  Logger.log("Linha a ser restaurada: " + linhaEncontrada);
  // --- FIM DA NOVA LÓGICA ---

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaMaquinas = ss.getSheetByName(NOME_ABA_MAQUINAS);
    if (!abaMaquinas) {
      Logger.log("ERRO FATAL: Aba '" + NOME_ABA_MAQUINAS + "' não encontrada.");
      return { status: "erro", mensagem: "Aba da planilha de máquinas não foi encontrada." };
    }

    // Validar se a linha existe
    if (linhaEncontrada > abaMaquinas.getLastRow()) {
      Logger.log("ERRO: Linha " + linhaEncontrada + " não existe na planilha.");
      return { status: "erro", mensagem: "Linha " + linhaEncontrada + " não existe na planilha." };
    }
    
    Logger.log("Linha " + linhaEncontrada + ": Lendo backup da Coluna I.");
    var dataAntiga = abaMaquinas.getRange(linhaEncontrada, 9).getValue(); // Col I
    
    if (!dataAntiga || !(dataAntiga instanceof Date)) {
      Logger.log("Aviso: Não há data de backup válida na Coluna I. Restaurando como vazio.");
      dataAntiga = "";
    }
    
    Logger.log("Linha " + linhaEncontrada + ": Restaurando data antiga para Coluna E.");
    abaMaquinas.getRange(linhaEncontrada, 5).setValue(dataAntiga); // Col E
    
    Logger.log("Linha " + linhaEncontrada + ": Limpando colunas D, F, G, H, I.");
    abaMaquinas.getRange(linhaEncontrada, 4).clearContent(); // Col D
    abaMaquinas.getRange(linhaEncontrada, 6, 1, 4).clearContent(); // Limpa F, G, H, I

    Logger.log("Linha " + linhaEncontrada + ": Forçando salvamento (flush)...");
    SpreadsheetApp.flush();

    // Aguarda um momento para garantir que o flush completou
    Utilities.sleep(100);

    Logger.log("--- SUCESSO! Manutenção desfeita. ---");

    Logger.log("Buscando dados atualizados para o cliente...");
    var dadosAtualizados = buscarDadosManutencaoComFiltro(filtroStatusAtual, filtroMaquinaAtual);
    
    return { 
      status: "sucesso", 
      dados: dadosAtualizados
    };
    
  } catch (e) {
    Logger.log("--- ERRO GRAVE NO BLOCO 'TRY' ---");
    Logger.log("ERRO: " + e.message);
    Logger.log("STACK: " + e.stack);
    return { status: "erro", mensagem: "Erro inesperado no servidor: " + e.message };
  }
}

/* ==========================================================
   FUNÇÃO DE BUSCA (ATUALIZADA PARA CRIAR 'IDENTIFICADOR')
   ========================================================== */
function buscarDadosManutencaoComFiltro(filtroStatus, filtroMaquina) { 
  Logger.log("--- DEBUG: Iniciando busca com filtroStatus: '" + filtroStatus + "', filtroMaquina: '" + filtroMaquina + "' ---"); 
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaMaquinas = ss.getSheetByName(NOME_ABA_MAQUINAS);
    
    var dados = abaMaquinas.getRange(2, 1, abaMaquinas.getLastRow() - 1, 9).getValues(); 
    Logger.log("DEBUG: Lidas " + dados.length + " linhas da planilha.");
    
    var listaCompleta = [];
    var hoje = new Date();
    hoje.setHours(0, 0, 0, 0); 

    for (var i = 0; i < dados.length; i++) {
      var linha = dados[i];
      var nomeMaquina = linha[0];
      
      if (nomeMaquina === "") {
        continue;
      }
      
      var statusPlanilha = linha[5]; // Col F
      var realizadoPor = linha[6]; // Col G
      
      var item = {
        maquina: nomeMaquina,
        intervalo: linha[1] + " dias",
        itens: linha[2] || "Nenhum item cadastrado",
        tipo: "",
        status: "",
        statusTexto: "",
        proximaData: "",
        dataConf: null,
        diasRestantes: 0
      };

      // --- ATUALIZAÇÃO: Usa o NÚMERO DA LINHA como identificador único ---
      // Isso é mais confiável que usar Máquina|Item porque:
      // 1. Cada linha tem um número único
      // 2. Não há problemas com itens vazios ou duplicados
      // 3. A busca é direta e rápida
      item.identificador = String(i + 2); // +2 porque começamos da linha 2
      item.numeroLinha = i + 2; // Guardamos também como número
      // --- FIM DA ATUALIZAÇÃO ---

      // Verifica a COLUNA F (Status)
      var statusLimpo = String(statusPlanilha).trim().toLowerCase();

      // Log específico para linhas problemáticas
      var linhaAtual = i + 2;
      if (linhaAtual === 149 || linhaAtual === 151 || linhaAtual === 152) {
        Logger.log(">>> LINHA PROBLEMÁTICA " + linhaAtual + " <<<");
        Logger.log("    Máquina: " + nomeMaquina);
        Logger.log("    Item (Col C): '" + (linha[2] || "") + "'");
        Logger.log("    Status (Col F): '" + statusPlanilha + "'");
        Logger.log("    Identificador que será criado: '" + (nomeMaquina + "|" + (linha[2] || "")) + "'");
      }

      Logger.log("DEBUG - Linha " + (i+2) + " | Máquina: " + nomeMaquina + " | Status da planilha: '" + statusPlanilha + "' | Status limpo: '" + statusLimpo + "'");

      if (statusLimpo === "realizado") {

        item.tipo = "Realizado";
        item.status = "Realizado";
        item.statusTexto = "MANUTENÇÃO REALIZADA";
        item.realizadoPor = realizadoPor || "Não informado";
        item.diasRestantes = 99999;
        Logger.log("DEBUG - Linha " + (i+2) + " marcada como REALIZADO"); 
        
        var dataConfirmacao = linha[7]; // Col H
        if (dataConfirmacao && dataConfirmacao instanceof Date && !isNaN(new Date(dataConfirmacao))) {
          item.dataConf = new Date(dataConfirmacao);
          item.dataConfirmacaoFormatada = item.dataConf.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit', year: 'numeric' }) + ' às ' + item.dataConf.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
        } else {
          item.dataConf = null; 
          item.dataConfirmacaoFormatada = "Data não registrada";
        }
      
      } else {
        // Se a Coluna F NÃO é "Realizado", é "Pendente"
        Logger.log("DEBUG - Linha " + (i+2) + " NÃO é realizado, processando como PENDENTE");
        item.tipo = "Pendente";

        var proximaManutencaoValor = linha[4]; // Col E
        
        if (!proximaManutencaoValor) {
            continue;
        }

        var proximaManutencao = new Date(proximaManutencaoValor); 
        
        if (isNaN(proximaManutencao)) {
            continue; 
        }
        
        proximaManutencao.setHours(0, 0, 0, 0); 
        item.proximaData = proximaManutencao.toLocaleDateString('pt-BR', {timeZone: 'UTC'});
        
        var diffTime = proximaManutencao.getTime() - hoje.getTime();
        var diffDias = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        item.diasRestantes = diffDias; 
        
        if (diffDias <= 0) { 
          item.status = "Vencido";
          item.statusTexto = (diffDias === 0) ? "VENCE HOJE" : "VENCIDO HÁ " + Math.abs(diffDias) + " DIAS";
        } else if (diffDias <= 2) { 
          item.status = "Alerta";
          item.statusTexto = "VENCE EM " + diffDias + " DIAS";
        } else { 
          item.status = "Em Dia";
          item.statusTexto = "VENCE EM " + diffDias + " DIAS";
        }
      }
      
      listaCompleta.push(item);
    }
    
    Logger.log("DEBUG: Processamento do loop concluído. " + listaCompleta.length + " itens na listaCompleta.");
    
    // --- LÓGICA DE FILTRO (STATUS) ---
    var listaFiltrada = [];
    var limiteDias = 7;

    Logger.log("DEBUG FILTRO: Aplicando filtro '" + filtroStatus + "' na lista de " + listaCompleta.length + " itens.");

    if (filtroStatus === "pendentes") {
      listaFiltrada = listaCompleta.filter(m => m.tipo === "Pendente" && m.diasRestantes <= limiteDias);
      Logger.log("DEBUG FILTRO: Filtro 'pendentes' - itens com tipo='Pendente' e diasRestantes <= 7");
    } else if (filtroStatus === "vencidos") {
      listaFiltrada = listaCompleta.filter(m => m.status === "Vencido");
      Logger.log("DEBUG FILTRO: Filtro 'vencidos' - itens com status='Vencido'");
    } else if (filtroStatus === "prazo") {
      listaFiltrada = listaCompleta.filter(m => (m.status === "Em Dia" || m.status === "Alerta"));
      Logger.log("DEBUG FILTRO: Filtro 'prazo' - itens com status='Em Dia' ou 'Alerta'");
    } else if (filtroStatus === "realizados") {
      listaFiltrada = listaCompleta.filter(m => m.tipo === "Realizado");
      Logger.log("DEBUG FILTRO: Filtro 'realizados' - itens com tipo='Realizado'");
    } else { // "todos"
      listaFiltrada = listaCompleta;
      Logger.log("DEBUG FILTRO: Filtro 'todos' - sem filtro aplicado");
    }

    Logger.log("DEBUG: Filtro '" + filtroStatus + "' aplicado. " + listaFiltrada.length + " itens restantes.");

    // Log detalhado dos itens filtrados
    listaFiltrada.forEach(function(item, index) {
      Logger.log("DEBUG FILTRADO[" + index + "]: " + item.maquina + " | Item COMPLETO: " + item.itens + " | Tipo: " + item.tipo + " | Status: " + item.status);
      Logger.log("DEBUG FILTRADO[" + index + "] - Identificador: '" + item.identificador + "'");
    });

    // --- LÓGICA DE FILTRO (MÁQUINA) ---
    if (filtroMaquina && filtroMaquina !== "todas") {
      var antesDoFiltroMaquina = listaFiltrada.length;
      listaFiltrada = listaFiltrada.filter(function(m) {
        // Compara os nomes "limpos"
        return String(m.maquina).trim() === String(filtroMaquina).trim();
      });
      Logger.log("DEBUG: Filtro de MÁQUINA '" + filtroMaquina + "' aplicado. " + antesDoFiltroMaquina + " → " + listaFiltrada.length + " itens.");
    }

    // --- LÓGICA DE ORDENAÇÃO ---
    listaFiltrada.sort(function(a, b) {
      if (a.tipo === "Pendente" && b.tipo === "Realizado") { return -1; }
      if (a.tipo === "Realizado" && b.tipo === "Pendente") { return 1; }
      if (a.tipo === "Pendente") { return a.diasRestantes - b.diasRestantes; }
      if (a.tipo === "Realizado") {
        var aTime = a.dataConf ? a.dataConf.getTime() : 0;
        var bTime = b.dataConf ? b.dataConf.getTime() : 0;
        return bTime - aTime;
      }
      return 0;
    });
    
    // --- LIMPEZA PARA SERIALIZAÇÃO ---
    listaFiltrada.forEach(function(item) {
      delete item.dataConf; 
    });

    Logger.log("DEBUG: Busca concluída. Retornando " + listaFiltrada.length + " itens.");
    return listaFiltrada;

  } catch (e) {
    Logger.log("--- ERRO GRAVE ---");
    Logger.log("ERRO GRAVE em buscarDadosManutencaoComFiltro: " + e.message);
    Logger.log("Stack: " + e.stack);
    Logger.log("--- FIM DO ERRO ---");
    return []; 
  }
}

/* ==========================================================
   FUNÇÃO DE ARQUIVAMENTO AUTOMÁTICO (DIARIAMENTE)
   ========================================================== */
function arquivarManutencoesRealizadas() {
  Logger.log("=== INÍCIO DO ARQUIVAMENTO AUTOMÁTICO ===");

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaMaquinas = ss.getSheetByName(NOME_ABA_MAQUINAS);
    var abaHistorico = ss.getSheetByName("Historico");

    // Cria a aba Historico se não existir
    if (!abaHistorico) {
      Logger.log("Aba 'Historico' não existe. Criando...");
      abaHistorico = ss.insertSheet("Historico");

      // Cria cabeçalho
      abaHistorico.getRange(1, 1, 1, 10).setValues([[
        "NumeroLinha", "Máquina", "Intervalo", "Itens",
        "DataConfirmacao", "ProximaManutencao", "Status",
        "RealizadoPor", "DataRealizacao", "DataArquivamento"
      ]]);
      abaHistorico.getRange(1, 1, 1, 10).setFontWeight("bold");
      Logger.log("Aba 'Historico' criada com sucesso.");
    }

    // Lê todas as linhas da aba Máquinas
    var ultimaLinha = abaMaquinas.getLastRow();
    var dados = abaMaquinas.getRange(2, 1, ultimaLinha - 1, 9).getValues();

    // Lê o histórico existente
    var ultimaLinhaHistorico = abaHistorico.getLastRow();
    var historicoExistente = [];
    if (ultimaLinhaHistorico > 1) {
      historicoExistente = abaHistorico.getRange(2, 1, ultimaLinhaHistorico - 1, 10).getValues();
    }

    var linhasArquivadas = 0;
    var linhasIgnoradas = 0;
    var novasLinhasHistorico = []; // Array para operação em lote

    // Percorre todas as linhas da aba Máquinas
    for (var i = 0; i < dados.length; i++) {
      var linha = dados[i];
      var numeroLinha = i + 2; // +2 porque começamos da linha 2

      var maquina = linha[0];
      var intervalo = linha[1];
      var itens = linha[2];
      var dataConfirmacao = linha[3]; // Col D
      var proximaManutencao = linha[4]; // Col E
      var status = linha[5]; // Col F
      var realizadoPor = linha[6]; // Col G
      var dataRealizacao = linha[7]; // Col H

      // Verifica se o status é "Realizado"
      if (String(status).trim().toLowerCase() !== "realizado") {
        continue; // Pula linhas que não estão realizadas
      }

      // Verifica se já existe no histórico
      var jaExiste = false;

      for (var h = 0; h < historicoExistente.length; h++) {
        var linhaHistorico = historicoExistente[h];
        var numeroLinhaHistorico = linhaHistorico[0];
        var proximaManutencaoHistorico = linhaHistorico[5];

        // Se o número da linha é o mesmo E a data da ProximaManutencao é a mesma
        if (numeroLinhaHistorico === numeroLinha &&
            proximaManutencaoHistorico instanceof Date &&
            proximaManutencao instanceof Date &&
            proximaManutencaoHistorico.getTime() === proximaManutencao.getTime()) {
          jaExiste = true;
          linhasIgnoradas++;
          break;
        }
      }

      // Se não existe, prepara para adicionar ao histórico
      if (!jaExiste) {
        var novaLinhaHistorico = [
          numeroLinha,
          maquina,
          intervalo,
          itens || "Nenhum item cadastrado",
          dataConfirmacao,
          proximaManutencao,
          status,
          realizadoPor,
          dataRealizacao,
          new Date() // Data do arquivamento
        ];

        novasLinhasHistorico.push(novaLinhaHistorico);
        linhasArquivadas++;
      }
    }

    // Operação em lote: Adiciona todas as linhas ao histórico de uma vez
    if (novasLinhasHistorico.length > 0) {
      Logger.log("Arquivando " + novasLinhasHistorico.length + " linhas no histórico...");
      var proximaLinhaHistorico = abaHistorico.getLastRow() + 1;
      abaHistorico.getRange(proximaLinhaHistorico, 1, novasLinhasHistorico.length, 10)
        .setValues(novasLinhasHistorico);
      Logger.log("Linhas adicionadas ao histórico com sucesso.");

      // NÃO limpa mais os dados da aba Máquinas
      // Os dados permanecem e serão sobrescritos na próxima manutenção
      // O sistema detecta novos registros pela mudança na ProximaManutencao
    }

    SpreadsheetApp.flush();

    Logger.log("=== ARQUIVAMENTO CONCLUÍDO ===");
    Logger.log("Linhas arquivadas: " + linhasArquivadas);
    Logger.log("Linhas ignoradas (duplicadas): " + linhasIgnoradas);

    return {
      sucesso: true,
      linhasArquivadas: linhasArquivadas,
      linhasIgnoradas: linhasIgnoradas
    };

  } catch (e) {
    Logger.log("ERRO no arquivamento: " + e.message);
    Logger.log("Stack: " + e.stack);
    return {
      sucesso: false,
      erro: e.message
    };
  }
}

/* ==========================================================
   FUNÇÃO PARA CONFIGURAR O TRIGGER AUTOMÁTICO
   ========================================================== */
function configurarTriggerArquivamento() {
  // Remove triggers antigos para evitar duplicatas
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'arquivarManutencoesRealizadas') {
      ScriptApp.deleteTrigger(triggers[i]);
      Logger.log("Trigger antigo removido.");
    }
  }

  // Cria novo trigger para executar diariamente
  ScriptApp.newTrigger('arquivarManutencoesRealizadas')
    .timeBased()
    .everyDays(1)
    .atHour(2) // Executa às 2h da manhã
    .create();

  Logger.log("Trigger configurado: arquivarManutencoesRealizadas será executado diariamente às 2h.");
  return "Trigger configurado com sucesso!";
}

/* ==========================================================
   FUNÇÃO PARA BUSCAR DADOS DO HISTÓRICO (PARA RELATÓRIO)
   ========================================================== */
function buscarDadosHistorico(filtroMaquina) {
  Logger.log("--- BUSCAR DADOS HISTÓRICO ---");
  Logger.log("Filtro de máquina: '" + filtroMaquina + "'");

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaHistorico = ss.getSheetByName("Historico");

    if (!abaHistorico) {
      Logger.log("AVISO: Aba 'Historico' não existe. Retornando lista vazia.");
      return [];
    }

    var ultimaLinha = abaHistorico.getLastRow();
    if (ultimaLinha <= 1) {
      Logger.log("AVISO: Aba 'Historico' está vazia. Retornando lista vazia.");
      return [];
    }

    // Lê todos os dados do histórico (pula cabeçalho)
    var dados = abaHistorico.getRange(2, 1, ultimaLinha - 1, 10).getValues();
    Logger.log("DEBUG: Lidas " + dados.length + " linhas do histórico.");

    var lista = [];

    for (var i = 0; i < dados.length; i++) {
      var linha = dados[i];

      // Estrutura da aba Historico:
      // 0: NumeroLinha, 1: Máquina, 2: Intervalo, 3: Itens,
      // 4: DataConfirmacao, 5: ProximaManutencao, 6: Status,
      // 7: RealizadoPor, 8: DataRealizacao, 9: DataArquivamento

      var maquina = linha[1];
      if (!maquina || maquina === "") {
        continue; // Pula linhas vazias
      }

      var item = {
        numeroLinha: linha[0],
        maquina: maquina,
        intervalo: linha[2] + " dias",
        itens: linha[3] || "Nenhum item cadastrado",
        realizadoPor: linha[7] || "Não informado",
        dataConfirmacaoFormatada: "",
        tipo: "Realizado",
        status: "Realizado",
        statusTexto: "MANUTENÇÃO REALIZADA (ARQUIVADO)",
        arquivado: true  // Marca como dado do histórico
      };

      // Formata a data de realização
      var dataRealizacao = linha[8];
      if (dataRealizacao && dataRealizacao instanceof Date && !isNaN(new Date(dataRealizacao))) {
        item.dataConfirmacaoFormatada = dataRealizacao.toLocaleDateString('pt-BR', {
          day: '2-digit',
          month: '2-digit',
          year: 'numeric'
        }) + ' às ' + dataRealizacao.toLocaleTimeString('pt-BR', {
          hour: '2-digit',
          minute: '2-digit'
        });
      } else {
        item.dataConfirmacaoFormatada = "Data não registrada";
      }

      lista.push(item);
    }

    // Aplica filtro de máquina se necessário
    if (filtroMaquina && filtroMaquina !== "todas") {
      var antesDoFiltro = lista.length;
      lista = lista.filter(function(m) {
        return String(m.maquina).trim() === String(filtroMaquina).trim();
      });
      Logger.log("DEBUG: Filtro de máquina '" + filtroMaquina + "' aplicado. " + antesDoFiltro + " → " + lista.length + " itens.");
    }

    // Ordena por data de realização (mais recente primeiro)
    lista.sort(function(a, b) {
      // Como já temos a string formatada, vamos ordenar por linha (mais recente = maior número)
      return b.numeroLinha - a.numeroLinha;
    });

    Logger.log("DEBUG: Busca do histórico concluída. Retornando " + lista.length + " itens.");
    return lista;

  } catch (e) {
    Logger.log("ERRO GRAVE em buscarDadosHistorico: " + e.message);
    Logger.log("Stack: " + e.stack);
    return [];
  }
}

/* ==========================================================
   FUNÇÃO DE DISPARADOR DE E-MAILS
   ========================================================== */
function enviarEmailsNotificacao() {
  Logger.log("Iniciando verificação de notificações...");
  
  var maquinasPendentes = buscarDadosManutencaoComFiltro('pendentes', 'todas'); 
  
  if (maquinasPendentes.length === 0) {
    Logger.log("Nenhuma manutenção pendente. E-mail não enviado.");
    return;
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaNotificacoes = ss.getSheetByName(NOME_ABA_NOTIFICACOES);
  
  var dadosEmails = abaNotificacoes.getRange(2, 1, abaNotificacoes.getLastRow() - 1, 1).getValues();
  var listaDeEmails = dadosEmails
    .map(function(linha) { return linha[0]; })
    .filter(function(email) { return email !== ""; })
    .join(",");

  var textoPadrao = abaNotificacoes.getRange("B2").getValue();
  
  if (listaDeEmails === "") {
    Logger.log("Nenhum e-mail cadastrado. E-mail não enviado.");
    return;
  }
  
  Logger.log("Enviando alertas para: " + listaDeEmails);

  var assunto = "ALERTA DE MANUTENÇÃO - MARFIM TÊXTIL";
  var corpoEmail = ""; 
  
  if (textoPadrao && textoPadrao.trim() !== "") {
    var textoPadraoHtml = textoPadrao.replace(/\n/g, '<br>');
    corpoEmail += "<p>" + textoPadraoHtml + "</p>";
    corpoEmail += "<hr>";
  }
  
  corpoEmail += "<h3>Alerta de Manutenção:</h3>" +
                "<p>As seguintes manutenções estão vencendo ou já estão vencidas, informar ao supervisor de produção para verificação junto a manutenção:</p>" +
                "<ul>";
  
  maquinasPendentes.forEach(function(maquina) {
    corpoEmail += "<li>" +
                  "<strong>Máquina:</strong> " + maquina.maquina + " | " +
                  "<strong>Status:</strong> " + maquina.statusTexto + "<br>" +
                  "<small><i>Itens Necessários: " + maquina.itens + "</i></small>" + 
                  "</li>";
  });
  
  corpoEmail += "</ul>";
  corpoEmail += "<br><p>Atenciosamente,<br>" +
                "Controle de Prazos e Rotinas Marfim</p>";

  try {
    MailApp.sendEmail({
      to: listaDeEmails,
      subject: assunto,
      htmlBody: corpoEmail
    });
    Logger.log("E-mail de notificação enviado com sucesso.");
  } catch (e) {
    Logger.log("Falha ao enviar e-mail: " + e.message);
  }
}
