const express = require('express');
const app = express();
const PDF = require('pdfmake/build/pdfmake');
const PDF_Fonts = require('pdfmake/build/vfs_fonts');
const { promisify } = require('util');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const os = require('os');
const { v4: uuidv4 } = require('uuid');
const multer = require('multer');
const cors = require('cors');

const port = 3000;

const readFileAsync = promisify(fs.readFile);
const writeFileAsync = promisify(fs.writeFile);

app.use(express.json());
app.use(cors('*'));

const upload = multer({ dest: 'uploads/' }); 
app.post('/gerarCertificados', upload.single('arquivoExcel'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'Nenhum arquivo Excel foi enviado.' });
    }

    const excelFilePath = req.file.path;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelFilePath);
    const worksheet = workbook.getWorksheet(1);
    const certificados = [];
    const pageLayout = {
      pageSize: 'A4',
      pageOrientation: 'landscape',
      pageMargins: [40, 40, 40, 40],
    };

    worksheet.eachRow({ includeEmpty: true }, (row, rowIndex) => {
      if (rowIndex === 1) return;

      const nome = row.getCell(1).value;
      const workshop = row.getCell(2).value;
      const dependencias = row.getCell(3).value;
      const dia = row.getCell(4).value;
      const mes = row.getCell(5).value;
      const ano = row.getCell(6).value;
      const duracao = row.getCell(7).value;
      const dia2 = row.getCell(8).value;
      const mes2 = row.getCell(9).value;
      const ano2 = row.getCell(10).value;
      const nomeGestor = row.getCell(11).value;
      const funcao = row.getCell(12).value;

      const modeloCertificado = [
        { text: 'O Instituto SENAI de Tecnologia em Mecatrônica confere o presente atestado a', style: 'header' },
        { text: nome, style: 'nome' },
        { text: `por sua participação no Workshop ${workshop}, realizado nas dependências`, style: 'paragrafo' },
        { text: ` ${dependencias}, no dia ${dia} de ${mes} de 202${ano}, com`, style: 'paragrafo' },
        { text: ` duração de ${duracao} horas.`, style: 'paragrafo' },
        { text: `Caxias do Sul, ${dia2} de ${mes2} de 202${ano2}.`, style: 'data' },
        { text: `______________________________________`, style: 'assinatura' },
        { text: `${nomeGestor}`, style: 'paragrafo' },
        { text: `${funcao}`, style: 'paragrafo' },
        { text: '', pageBreak: 'after' },
      ];
      certificados.push(modeloCertificado);
    });

    const styles = {
      header: {
        fontSize: 18,
        bold: true,
        alignment: 'center',
        margin: [0, 90, 0, 0],
      },
      nome: {
        fontSize: 22,
        bold: true,
        alignment: 'center',
        margin: [0, 30, 0, 30],
      },
      paragrafo: {
        fontSize: 18,
        bold: true,
        margin: [0, 3, 0, 3],
        alignment: 'center',
      },
      data: {
        fontSize: 18,
        bold: true,
        margin: [0, 20, 0, 30],
        alignment: 'center',
      },
      assinatura: {
        fontSize: 18,
        margin: [0, 5, 0, 5],
        alignment: 'center',
      }
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
        width: 842,
        margin: [0, 0, 0, 0],
      },
      footer: {
        image: `data:image/png;base64,${footerImage}`,
        width: 300,
        margin: [300, -20, 0, 0],
      },
    };

    PDF.vfs = PDF_Fonts.pdfMake.vfs;
    const pdfDocGenerator = PDF.createPdf(docDefinition);
    const pdfBuffer = await new Promise((resolve, reject) => {
      pdfDocGenerator.getBuffer((buffer) => {
        resolve(buffer);
      });
    });

    // Gerar um nome de arquivo exclusivo usando uuid
    const tempPdfFileName = `${uuidv4()}.pdf`;
    const tempPdfFilePath = path.join(os.tmpdir(), tempPdfFileName);
    
    // Salvar o PDF em um arquivo temporário
    await writeFileAsync(tempPdfFilePath, pdfBuffer);

    // Enviar o PDF como um anexo de arquivo na resposta HTTP
    res.setHeader('Content-Disposition', `attachment; filename=${tempPdfFileName}`);
    res.sendFile(tempPdfFilePath, {}, (err) => {
      // Remover o arquivo temporário após o envio
      fs.unlink(tempPdfFilePath, (unlinkError) => {
        if (unlinkError) {
          console.error('Erro ao remover o arquivo temporário:', unlinkError);
        }
      });
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Ocorreu um erro ao gerar os certificados.' });
  }
});

app.listen(port, () => {
  console.log(`API de certificados está ouvindo na porta ${port}`);
});