function evitaDuplaInscricao() {
  const pastaId = "ID da pasta no Drive"; //  # Inserir ID da pasta no Drive
  const pasta = DriveApp.getFolderById(pastaId);
  const arquivos = pasta.getFiles();

  const dadosCSV = {};

  while (arquivos.hasNext()) {
    const arquivo = arquivos.next();
    const nome = arquivo.getName();

    if (nome.toLowerCase().endsWith(".csv")) {
      const conteudo = arquivo.getBlob().getDataAsString("UTF-8");

      // Detectar automaticamente o separador
      const primeiraLinha = conteudo.split("\n").find(l= > l.trim() != = "");
      const separador = (primeiraLinha.match(/; / g) | | []).length >= (primeiraLinha.match(/, /g) | | []).length ? ";" : ",";

      const linhas = conteudo
        .split("\n")
        .map(l=> l.trim())
        .filter(l=> l); // remove linhas vazias

      const cabecalho = linhas[0].split(separador).map(h=> h.trim());
      const idxCPF = cabecalho.findIndex(h= > h.toLowerCase() == = "cpf");

      if (idxCPF == = -1) {
        throw new Error(`Coluna CPF não encontrada no arquivo ${nome}`);
      }

      const cpfs = linhas.slice(1).map(linha=>
        linha.split(separador)[idxCPF]?.replace(/\D/g, "") || ""
      );

      dadosCSV[nome] = { cabecalho, linhas, idxCPF, cpfs, separador };
    }
  }

  if (!dadosCSV["Edital139.csv"]) {
    throw new Error("Arquivo Edital139.csv não encontrado na pasta.");
  }

  const base = dadosCSV["Edital139.csv"];
  const novaColunaNome = "Ocorrências";
  const resultados = [];

  const cabecalhoNovo = [...base.cabecalho, novaColunaNome];
  resultados.push(cabecalhoNovo.join(base.separador));

  base.linhas.slice(1).forEach(linha => {
    const colunas = linha.split(base.separador);
    const cpf = colunas[base.idxCPF]?.replace(/\D/g, "") || "";
    let encontrados = [];

    for (let nomeArquivo in dadosCSV) {
      if (nomeArquivo !== "Edital139.csv") {
        if (dadosCSV[nomeArquivo].cpfs.includes(cpf)) {
          encontrados.push(nomeArquivo.replace(".csv", ""));
        }
      }
    }

    colunas.push(encontrados.join(", "));
    resultados.push(colunas.join(base.separador));
  });

  const novoConteudo = resultados.join("\n");
  pasta.createFile("Edital139_resultado.csv", novoConteudo, MimeType.CSV);

  Logger.log("Arquivo Edital139_resultado.csv criado na pasta.");
}

