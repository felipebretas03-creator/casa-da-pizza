function doGet(e) {
  // 1. Lógica NOVA do Painel de Pedidos (Busca os pedidos na planilha)
  if (e && e.parameter && e.parameter.action == "listar_pedidos") {
    // Usando o mesmo ID da planilha que você já usava!
    var planilha = SpreadsheetApp.openById('1I0j7_UPKeQS2wD9y41tWX690sxYuB1QqIXSD8W4gRSQ');
    var aba = planilha.getSheetByName("Pedidos");
    
    // Se a aba Pedidos não existir, cria ela automaticamente
    if (!aba) {
      aba = planilha.insertSheet("Pedidos");
      aba.appendRow(["ID", "Data", "Cliente", "Telefone", "Total", "Itens", "Endereco", "Pagamento", "Status"]);
    }

    var dados = aba.getDataRange().getValues();
    var cabecalho = dados[0];
    var pedidos = [];
    
    // Lê da linha 2 em diante
    for (var i = 1; i < dados.length; i++) {
      var row = dados[i];
      var pedido = {};
      for (var j = 0; j < cabecalho.length; j++) {
        pedido[cabecalho[j]] = row[j];
      }
      pedidos.push(pedido);
    }
    
    // Inverte para o pedido mais recente aparecer no topo
    pedidos.reverse();
    
    return ContentService.createTextOutput(JSON.stringify(pedidos))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (e && e.parameter && e.parameter.action == "listar_caixa_mes") {
    var planilha = SpreadsheetApp.openById('1I0j7_UPKeQS2wD9y41tWX690sxYuB1QqIXSD8W4gRSQ');
    var aba = planilha.getSheetByName("caixa do mes");
    if (!aba) {
      return ContentService.createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    var dados = aba.getDataRange().getDisplayValues();
    var caixas = [];
    
    if (dados.length > 1) {
      var cabecalhos = dados[0];
      for (var i = 1; i < dados.length; i++) {
        var row = dados[i];
        var obj = {};
        for (var j = 0; j < cabecalhos.length; j++) {
          obj[cabecalhos[j]] = row[j];
        }
        caixas.push(obj);
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify(caixas))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // 2. Lógica ANTIGA do seu sistema de abrir e fechar a loja
  return HtmlService.createHtmlOutputFromFile('index');
}


// Sua função original intocada
function alterarStatus(status) {
  const planilha = SpreadsheetApp.openById('1I0j7_UPKeQS2wD9y41tWX690sxYuB1QqIXSD8W4gRSQ');
  const aba = planilha.getSheetByName('controle');
  aba.getRange('A2').setValue(status);
}


// Função NOVA para receber os pedidos e atualizar os status pelo site
function doPost(e) {
  var planilha = SpreadsheetApp.openById('1I0j7_UPKeQS2wD9y41tWX690sxYuB1QqIXSD8W4gRSQ');
  var aba = planilha.getSheetByName("Pedidos");
  
  if (!aba) {
    aba = planilha.insertSheet("Pedidos");
    aba.appendRow(["ID", "Data", "Cliente", "Telefone", "Total", "Itens", "Endereco", "Pagamento", "Status"]);
  }

  try {
    var payload = JSON.parse(e.postData.contents);
    // 0. Alterar o status da loja (Abrir/Fechar)
    if (payload.action == "alterar_status_loja") {
      alterarStatus(payload.status);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 1. Receber um NOVO pedido do site
    if (payload.action == "novo_pedido") {
      function getMaxIdFromSheet(sheet) {
        var mId = 0;
        if (sheet && sheet.getLastRow() > 1) {
          var ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
          for (var i = 0; i < ids.length; i++) {
            var val = String(ids[i][0]).replace(/\D/g, ''); // Extrai apenas números
            var num = parseInt(val, 10);
            if (!isNaN(num) && num > mId) {
              mId = num;
            }
          }
        }
        return mId;
      }

      var maxId = getMaxIdFromSheet(aba);
      if (maxId === 0) {
        // Se a aba Pedidos estiver vazia, tenta pegar do histórico para não perder a sequência
        var historicoAba = planilha.getSheetByName("historico de pedidos");
        maxId = getMaxIdFromSheet(historicoAba);
      }
      
      var idNum = maxId + 1;
      var id = idNum.toString().padStart(4, '0');
      
      var data = new Date().toLocaleString("pt-BR", {timeZone: "America/Sao_Paulo"});
      
      aba.appendRow([
        "'" + id, 
        data, 
        payload.cliente, 
        payload.telefone, 
        payload.total, 
        payload.itens, 
        payload.endereco, 
        payload.pagamento, 
        "Recebido"
      ]);
      
      return ContentService.createTextOutput(JSON.stringify({"sucesso": true, "id": id}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    // 2. Atualizar o STATUS de um pedido (pelo painel)
    if (payload.action === "atualizar_status") {
      var idPedido = payload.id;
      var novoStatus = payload.status;
      var dados = aba.getDataRange().getValues();
      
      for (var i = 1; i < dados.length; i++) {
        if (dados[i][0] == idPedido) {
          aba.getRange(i + 1, 9).setValue(novoStatus); // Atualiza o status na coluna 9 (I)
          return ContentService.createTextOutput(JSON.stringify({"sucesso": true}))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({"erro": "Pedido não encontrado"}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // 3. Fechar o Caixa do Dia
    if (payload.action === "fechar_caixa") {
      var caixaMes = planilha.getSheetByName("caixa do mes");
      if (!caixaMes) {
        caixaMes = planilha.insertSheet("caixa do mes");
        caixaMes.appendRow(["Data de Fechamento", "Quantidade Entregues", "Total Arrecadado"]);
      }
      
      var historico = planilha.getSheetByName("historico de pedidos");
      if (!historico) {
        historico = planilha.insertSheet("historico de pedidos");
        historico.appendRow(["ID", "Data", "Cliente", "Telefone", "Total", "Itens", "Endereco", "Pagamento", "Status"]);
        historico.hideSheet(); // Oculta a aba
      }

      var pedidos = aba.getDataRange().getValues();
      var qtdEntregues = 0;
      var totalArrecadado = 0;
      var pedidosParaMover = [];
      
      // Começa do 1 para pular o cabeçalho
      for (var i = 1; i < pedidos.length; i++) {
        var row = pedidos[i];
        var status = row[8]; // Coluna Status é a 9ª (índice 8)
        
        if (status === "Entregue") {
          qtdEntregues++;
          var valorStr = String(row[4]).replace(/[^\d,.-]/g, '').replace(',', '.'); // Total é a 5ª (índice 4)
          var valor = parseFloat(valorStr);
          if (!isNaN(valor)) {
            totalArrecadado += valor;
          }
        }
        pedidosParaMover.push(row);
      }
      
      // Mover para o histórico
      if (pedidosParaMover.length > 0) {
        historico.getRange(historico.getLastRow() + 1, 1, pedidosParaMover.length, pedidosParaMover[0].length).setValues(pedidosParaMover);
      }
      
      // Limpar a aba Pedidos atual (mantendo o cabeçalho)
      if (aba.getLastRow() > 1) {
        aba.getRange(2, 1, aba.getLastRow() - 1, aba.getLastColumn()).clearContent();
      }

      // Registrar o caixa
      var dataAtual = new Date().toLocaleString("pt-BR", {timeZone: "America/Sao_Paulo"});
      caixaMes.appendRow([dataAtual, qtdEntregues, "R$ " + totalArrecadado.toFixed(2).replace('.', ',')]);
      
      return ContentService.createTextOutput(JSON.stringify({
        success: true, 
        qtd: qtdEntregues, 
        total: totalArrecadado
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({"erro": err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
