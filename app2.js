function verificarCpfsRepetidosRelatorios() {
  const pastaRelatoriosId = '1LuRx7EfpDNh8_e6nfiV47EdZ4-mlS-lj'; // AUXILIO ESTUDANTIL
  const arquivosConvertidos = [];
  const mapaCpfRelatorios = new Map();

  // Processa arquivos da pasta de relatórios
  const pastaRelatorios = DriveApp.getFolderById(pastaRelatoriosId);
  const arquivos = pastaRelatorios.getFiles();

  while (arquivos.hasNext()) {
    const arquivo = arquivos.next();
    const nome = arquivo.getName();

    // Apenas Relatorio1 até Relatorio5
    if (/Relatorio[1-5]\.xls$/i.test(nome) || /Relatorio[1-5]/i.test(nome)) {
      let planilha;
      if (arquivo.getMimeType() !== MimeType.GOOGLE_SHEETS) {
        // Converte .xls para Google Sheets
        const blob = arquivo.getBlob();
        const planilhaConvertida = Drive.Files.insert(
          { title: nome + ' (convertido)', parents: [{ id: pastaRelatoriosId }] },
          blob,
          { convert: true }
        );
        arquivosConvertidos.push(planilhaConvertida.id);
        planilha = SpreadsheetApp.openById(planilhaConvertida.id);
      } else {
        planilha = SpreadsheetApp.open(arquivo);
      }

      const aba = planilha.getSheets()[0];
      const dados = aba.getDataRange().getValues();
      const headerRel = dados[0];
      const idxCPF = headerRel.findIndex(col => col.toString().toUpperCase().includes("CPF"));
      if (idxCPF === -1) continue;

      for (let i = 1; i < dados.length; i++) {
        const cpf = normalizarCpf(dados[i][idxCPF]);
        if (!cpf) continue;

        if (!mapaCpfRelatorios.has(cpf)) {
          mapaCpfRelatorios.set(cpf, new Set());
        }
        mapaCpfRelatorios.get(cpf).add(nome);
      }
    }
  }

  // Cria nova planilha com os CPFs repetidos
  const novaPlanilha = SpreadsheetApp.create("Resultado CPFs Repetidos");
  const abaNova = novaPlanilha.getActiveSheet();
  abaNova.appendRow(["CPF", "Relatórios Encontrados"]);

  for (let [cpf, relatorios] of mapaCpfRelatorios.entries()) {
    if (relatorios.size > 1) {
      abaNova.appendRow([cpf, Array.from(relatorios).join(", ")]);
    }
  }

  // Move planilha criada para a pasta de relatórios
  const novaPlanilhaFile = DriveApp.getFileById(novaPlanilha.getId());
  pastaRelatorios.addFile(novaPlanilhaFile);
  DriveApp.getRootFolder().removeFile(novaPlanilhaFile);

  // Exclui arquivos convertidos temporários
  for (const id of arquivosConvertidos) {
    try {
      DriveApp.getFileById(id).setTrashed(true);
      Logger.log("Arquivo convertido excluído: " + id);
    } catch (e) {
      Logger.log("Erro ao excluir arquivo convertido: " + id + " → " + e);
    }
  }

  Logger.log("Execução concluída. Planilha com CPFs repetidos criada.");
}

function normalizarCpf(cpf) {
  if (!cpf) return '';
  return cpf.toString().replace(/[^\d]/g, '').padStart(11, '0');
}

