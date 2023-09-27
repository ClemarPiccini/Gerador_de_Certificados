const fs = require('fs');
const PDF = require('pdfmake/build/pdfmake'); // Importe o pdfmake
const PDF_Fonts = require('pdfmake/build/vfs_fonts'); // Importe os vfs_fonts

async function gerarCertificados() {
  // Carregue o arquivo ODS usando a biblioteca exceljs
  const ExcelJS = require('exceljs');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('Sheet1.xlsx');

  // Obtenha a primeira planilha do arquivo ODS (suponha que seja a única)
  const worksheet = workbook.getWorksheet(1);

  // Crie um array para armazenar todos os certificados
  const certificados = [];

  // Defina um layout personalizado para a página no modo paisagem
  const pageLayout = {
    pageSize: 'A4',
    pageOrientation: 'landscape',
    pageMargins: [40, 40, 40, 40],
  };

  // Crie um certificado para cada linha do arquivo ODS
  worksheet.eachRow({ includeEmpty: true }, (row, rowIndex) => {
    if (rowIndex === 1) return; // Ignorar a primeira linha (cabeçalho)

    const nome = row.getCell(1).value; // Suponha que o nome esteja na primeira coluna

    // Defina o conteúdo do certificado
    const modeloCertificado = [
      { text: 'O Instituto SENAI de Tecnologia em Mecatrônica confere o presente atestado a', style: 'header' },
      { text: nome, style: 'nome' },
      { text: 'por sua participação no Workshop XXXXXXXXXXXXXXX,', style: 'paragrafo' },
      { text: 'realizado nas dependências XXXXXXXXXXXXXXXXXXXXXXXXXXXX,', style: 'paragrafo' },
      { text: 'no dia XX de XXXXX de 202X, com duração de XX horas.', style: 'paragrafo' },
      { text: 'Caxias do Sul, XX de XXXXX de 202X.', style: 'paragrafo' },
      { text: '', pageBreak: 'after' }, // Quebra de página após cada certificado
    ];

    certificados.push(modeloCertificado);
  });

  // Defina os estilos para o PDF
  const styles = {
    header: {
      fontSize: 14,
      bold: true,
    },
    nome: {
      fontSize: 16,
      bold: true,
      margin: [0, 10, 0, 10],
    },
    paragrafo: {
      fontSize: 12,
      margin: [0, 5, 0, 5],
    },
  };

  // Defina o documento PDF
  const docDefinition = {
    pageOrientation: 'landscape',
    content: certificados,
    styles: styles,
    pageMargins: [40, 40, 40, 40],
  };

  // Defina o vfs para os fonts
  PDF.vfs = PDF_Fonts.pdfMake.vfs;

  // Crie o PDF usando pdfmake
  const pdfDocGenerator = PDF.createPdf(docDefinition);

  // Gere o PDF e salve no sistema de arquivos
  pdfDocGenerator.getBuffer((buffer) => {
    fs.writeFileSync('certificados.pdf', buffer);
    console.log('Certificados gerados com sucesso!');
  });
}

gerarCertificados().catch((err) => {
  console.error(err);
});
