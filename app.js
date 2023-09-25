const ExcelJS = require('exceljs');
const officegen = require('officegen');
const fs = require('fs');

async function gerarCertificados() {
  // Carregue o arquivo ODS usando a biblioteca exceljs
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('Sheet1.xlsx');

  // Obtenha a primeira planilha do arquivo ODS (suponha que seja a única)
  const worksheet = workbook.getWorksheet(1);

  // Crie um objeto Word (docx)
  const docx = officegen('docx');
  const docxStream = fs.createWriteStream('certificados.docx');
  docxStream.on('error', (err) => {
    console.log('Erro ao criar o arquivo DOCX:', err);
  });

  // Crie um certificado para cada linha do arquivo ODS
  worksheet.eachRow({ includeEmpty: true }, (row, rowIndex) => {
    if (rowIndex === 1) return; // Ignorar a primeira linha (cabeçalho)

    const nome = row.getCell(1).value; // Suponha que o nome esteja na primeira coluna

    // Crie um documento de parágrafo para cada certificado
    const paragraph = docx.createP();

    // Substitua os marcadores de posição no modelo do certificado
    const modeloCertificado = `
    O Instituto SENAI de Tecnologia em Mecatrônica confere o presente atestado a

    ${nome}

    por sua participação no Workshop XXXXXXXXXXXXXXX, realizado nas dependências XXXXXXXXXXXXXXXXXXXXXXXXXXXX, no dia XX de XXXXX de 202X, com duração de XX horas.

    Caxias do Sul, XX de XXXXX de 202X.`;

    paragraph.addText(modeloCertificado);

    // Salte uma linha após cada certificado
    docx.createP();
  });

  // Finalize e salve o documento Word
  docx.generate(docxStream);
  console.log('Certificados gerados com sucesso!');
}

gerarCertificados();
