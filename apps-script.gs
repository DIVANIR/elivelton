// ================================================================
//  Google Apps Script — Recebe inscrições e salva no Google Sheets
//
//  COMO CONFIGURAR:
//  1. Abra sua planilha no Google Sheets
//  2. Clique em "Extensões" > "Apps Script"
//  3. Apague o código existente e cole TODO este arquivo
//  4. Clique em "Salvar" (ícone de disquete)
//  5. Clique em "Implantar" > "Nova implantação"
//  6. Em "Tipo", selecione "App da Web"
//  7. Em "Executar como", selecione "Eu"
//  8. Em "Quem tem acesso", selecione "Qualquer pessoa"
//  9. Clique em "Implantar" e autorize quando solicitado
// 10. Copie a URL gerada e cole em CONFIG.APPS_SCRIPT_URL no HTML
// ================================================================

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Cria cabeçalho se a planilha estiver vazia
    if (sheet.getLastRow() === 0) {
      const headers = [
        'Data/Hora', 'Nome', 'E-mail', 'Telefone',
        'CPF', 'Instituição/Empresa', 'Evento', 'Valor (R$)', 'Status Pgto'
      ];
      sheet.appendRow(headers);

      // Formata cabeçalho
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#4F46E5');
      headerRange.setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);

      // Ajusta largura das colunas
      sheet.setColumnWidth(1, 160); // Data/Hora
      sheet.setColumnWidth(2, 220); // Nome
      sheet.setColumnWidth(3, 220); // E-mail
      sheet.setColumnWidth(4, 140); // Telefone
      sheet.setColumnWidth(5, 130); // CPF
      sheet.setColumnWidth(6, 200); // Instituição
      sheet.setColumnWidth(7, 180); // Evento
      sheet.setColumnWidth(8, 100); // Valor
      sheet.setColumnWidth(9, 130); // Status
    }

    const dados = JSON.parse(e.postData.contents);

    sheet.appendRow([
      dados.data_inscricao || new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' }),
      dados.nome        || '',
      dados.email       || '',
      dados.telefone    || '',
      dados.cpf         || '',
      dados.instituicao || '',
      dados.evento      || '',
      dados.valor       || '',
      'Aguardando pagamento',   // Status inicial
    ]);

    // Alterna cor das linhas (zebra)
    const lastRow = sheet.getLastRow();
    if (lastRow % 2 === 0) {
      sheet.getRange(lastRow, 1, 1, 9).setBackground('#F3F4F6');
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', linha: lastRow }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Teste manual: abre o editor, selecione a função "testar" e clique em Executar
function testar() {
  const mockEvent = {
    postData: {
      contents: JSON.stringify({
        nome:          'João da Silva',
        email:         'joao@email.com',
        telefone:      '(11) 99999-8888',
        cpf:           '000.000.000-00',
        instituicao:   'UFMG',
        evento:        'Evento Teste',
        valor:         50.00,
        data_inscricao: new Date().toLocaleString('pt-BR'),
      })
    }
  };
  const resultado = doPost(mockEvent);
  Logger.log(resultado.getContent());
}
