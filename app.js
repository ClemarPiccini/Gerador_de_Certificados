const fs = require('fs');
const PDF = require('pdfmake/build/pdfmake');
const PDF_Fonts = require('pdfmake/build/vfs_fonts');
const ExcelJS = require('exceljs');
const path = require('path');
const { promisify } = require('util');

const readFileAsync = promisify(fs.readFile);

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
      { text: 'por sua participação no Workshop XXXXXXXXXXXXXXX, realizado nas dependências', style: 'paragrafo' },
      { text: ' XXXXXXXXXXXXXXXXXXXXXXXXXXXX, no dia XX de XXXXX de 202X, com', style: 'paragrafo' },
      { text: ' duração de XX horas.', style: 'paragrafo' },
      { text: 'Caxias do Sul, XX de XXXXX de 202X.', style: 'data' },
      { text: '', pageBreak: 'after' }, // Quebra de página após cada certificado
    ];

    certificados.push(modeloCertificado);
  });

  const styles = {
    header: {
      fontSize: 18,
      alignment: 'center', // Centralize o texto
      margin: [0, 130, 0, 0], // Margem superior maior para o cabeçalho
    },
    nome: {
      fontSize: 22,
      bold: true,
      alignment: 'center', // Centralize o texto
      margin: [0, 40, 0, 40], // Margem superior maior para o nome
    },
    paragrafo: {
      fontSize: 18,
      margin: [0, 5, 0, 5],
      alignment: 'center', // Centralize o texto
    },
    data: {
      fontSize: 18,
      margin: [0, 20, 0, 5],
      alignment: 'right', // Centralize o texto
    },
  };
  
  const headerImage = await readFileAsync(path.join(__dirname, 'Cabecalho.png'), 'base64');
  const footerImage = await readFileAsync(path.join(__dirname, 'Rodape.png'), 'base64');

  const docDefinition = {
    pageOrientation: 'landscape',
    content: certificados,
    styles: styles,
    pageMargins: [40, 40, 40, 40],
    header: {
      image: `data:image/png;base64,${headerImage}`,
      width: 842, // Largura da página A4 no modo paisagem
      margin: [0, 0, 0, 0], // Remova as margens do cabeçalho
    },
    footer: {
      image: `data:image/png;base64,${footerImage}`,
      width: 300, // Largura da página A4 no modo paisagem
      margin: [300, -20, 0, 0], // Remova as margens do rodapé
    },
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
