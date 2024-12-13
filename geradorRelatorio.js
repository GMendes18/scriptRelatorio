const fs = require("fs");
const path = require("path");
const csv = require("csv-parser");
const _ = require("lodash");
const ExcelJS = require("exceljs");

// Caminho para o arquivo CSV
const arquivoCSV = path.resolve(__dirname, "dados.csv");

// Caminho para o arquivo Excel existente
const arquivoExcelExistente = path.resolve(__dirname, "teste3.xlsx"); // Substitua se necessário

// Função para contar ocorrências de uma determinada coluna
function contarOcorrencias(dados, coluna) {
  const contagem = _.countBy(dados, coluna);
  // Converter para um array de objetos
  const resultado = Object.keys(contagem).map((chave) => ({
    [coluna]: chave,
    QUANTIDADE: contagem[chave],
  }));
  // Ordenar em ordem decrescente
  return _.orderBy(resultado, ["QUANTIDADE"], ["desc"]);
}

// Função para salvar dados em um arquivo Excel existente
async function editarExcel(estabelecimentos, categorias) {
  const workbook = new ExcelJS.Workbook();
  try {
    console.log(
      `Carregando o arquivo Excel existente: ${arquivoExcelExistente}`
    );
    await workbook.xlsx.readFile(arquivoExcelExistente);
    console.log("Arquivo Excel carregado com sucesso.");

    // Selecionar a planilha específica
    let worksheet = workbook.getWorksheet("Relatório"); // Substitua 'Relatório' pelo nome da planilha desejada
    if (!worksheet) {
      console.warn(
        'Planilha "Relatório" não encontrada. Usando a primeira planilha.'
      );
      worksheet = workbook.worksheets[0];
    } else {
      console.log('Planilha "Relatório" encontrada e selecionada.');
    }

    // Inserir Estabelecimentos a partir da célula A6
    const inicioEstabelecimentos = { row: 6, column: 1 }; // A7
    worksheet.getCell(`A${inicioEstabelecimentos.row}`).value =
      "ESTABELECIMENTO";
    worksheet.getCell(`B${inicioEstabelecimentos.row}`).value = "QUANTIDADE";

    // Inserir os dados de Estabelecimentos
    estabelecimentos.forEach((item, index) => {
      const row = inicioEstabelecimentos.row + 1 + index;
      worksheet.getCell(`A${row}`).value = item.ESTABELECIMENTO;
      worksheet.getCell(`B${row}`).value = item.QUANTIDADE;
    });

    // Inserir Categorias a partir da célula D20
    const inicioCategorias = { row: 20, column: 4 }; // D21
    worksheet.getCell(`D${inicioCategorias.row}`).value = "CATEGORIA";
    worksheet.getCell(`E${inicioCategorias.row}`).value = "QUANTIDADE";

    // Inserir os dados de Categorias
    categorias.forEach((item, index) => {
      const row = inicioCategorias.row + 1 + index;
      worksheet.getCell(`D${row}`).value = item.CATEGORIA;
      worksheet.getCell(`E${row}`).value = item.QUANTIDADE;
    });

    // Salvar o arquivo Excel com as alterações
    const caminhoExcelEditado = path.resolve(
      __dirname,
      "relatorio_editado.xlsx"
    ); // Nome do arquivo editado
    await workbook.xlsx.writeFile(caminhoExcelEditado);
    console.log(`\nArquivo Excel editado e salvo em: ${caminhoExcelEditado}`);
  } catch (erro) {
    console.error("Erro ao editar o arquivo Excel:", erro);
  }
}

// Função principal
function gerarRelatorios() {
  const dados = [];

  console.log(`Iniciando a leitura do arquivo CSV: ${arquivoCSV}`);
  fs.createReadStream(arquivoCSV)
    .pipe(csv())
    .on("data", (linha) => {
      // Filtrar linhas com valores ausentes em ESTABELECIMENTO ou CATEGORIA
      if (linha.ESTABELECIMENTO && linha.CATEGORIA) {
        dados.push(linha);
      }
    })
    .on("end", async () => {
      console.log("Leitura do arquivo CSV concluída.");

      // Contar e ordenar estabelecimentos
      const estabelecimentos = contarOcorrencias(dados, "ESTABELECIMENTO");
      console.log(
        "\nEstabelecimentos únicos em ordem decrescente de frequência:"
      );
      console.table(estabelecimentos);

      // Contar e ordenar categorias
      const categorias = contarOcorrencias(dados, "CATEGORIA");
      console.log("\nCategorias únicas em ordem decrescente de frequência:");
      console.table(categorias);

      // Salvar os relatórios em arquivos CSV (opcional)
      const caminhoEstabelecimentos = path.resolve(
        __dirname,
        "relatorio_estabelecimentos.csv"
      );
      salvarCSV(estabelecimentos, caminhoEstabelecimentos, [
        "ESTABELECIMENTO",
        "QUANTIDADE",
      ]);

      const caminhoCategorias = path.resolve(
        __dirname,
        "relatorio_categorias.csv"
      );
      salvarCSV(categorias, caminhoCategorias, ["CATEGORIA", "QUANTIDADE"]);

      console.log("\nRelatórios CSV gerados com sucesso!");

      // Editar o arquivo Excel existente com os dados contados
      await editarExcel(estabelecimentos, categorias);

      console.log("\nRelatório Excel editado com sucesso!");
    })
    .on("error", (erro) => {
      console.error("Erro ao ler o arquivo CSV:", erro);
    });
}

// Função para salvar dados em um arquivo CSV (opcional)
function salvarCSV(dados, caminho, colunas) {
  const cabecalho = colunas.join(",") + "\n";
  const linhas = dados
    .map((obj) => colunas.map((col) => `"${obj[col]}"`).join(","))
    .join("\n");
  fs.writeFileSync(caminho, cabecalho + linhas, "utf8");
  console.log(`Arquivo CSV salvo em: ${caminho}`);
}

// Executar a função principal
gerarRelatorios();
