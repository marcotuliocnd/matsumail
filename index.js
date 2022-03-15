const fs = require('fs')
const { Workbook } = require('exceljs');
const nodemailer = require('nodemailer')
const readline = require('readline')

const getUserInput = (message = '') => {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });
  return new Promise(resolve => {
    rl.question(message, (response) => {
      rl.close()
      resolve((response || 's').toLowerCase())
    })
  })
}

const getTemplateText = async () => {
  const data = await fs.readFileSync('./template.txt')
  return await data.toString()
}

const escapeRegExp = (string) => {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
}

const replaceAll = (str, find, replace) => {
  return str.replace(new RegExp(escapeRegExp(find), 'g'), replace);
}

const getConfig = async () => {
  const data = await fs.readFileSync('./config.json')
  const string = await data.toString()
  return JSON.parse(string)
}

const extractSpreadsheet = async (file) => {
  const workbook = new Workbook();
  const workbookFile = await workbook.xlsx.readFile(file);
  const worksheet = workbookFile.getWorksheet(1);
  const data = [];
  const columnsNames = [];

  worksheet.getRow(1).eachCell((cell, colNumber) => {
    columnsNames[colNumber] = cell.text.toLowerCase();
  });

  worksheet.eachRow((row) => {
    if (row.number === 1) {
      return;
    }

    const contact = {};

    row.eachCell((cell) => {
      const { col, value } = cell;
      contact[columnsNames[col]] = value;
    });
  
    data.push(contact);
  });

  return data;
}

const main = async () => {
  const config = await getConfig()
  const template = await getTemplateText()
  const contacts = await extractSpreadsheet('./BASE_DADOS_EMAIL.xlsx')
  console.log('='.repeat(10))
  console.log('Bem vindo ao Matsumail!')
  console.log('='.repeat(10))

  const transporter = nodemailer.createTransport({
    host: 'smtp.dreamhost.com',
    port: 587,
    secure: false,
    auth: {
      user: config.email,
      pass: config.senha,
    },
  });

  console.log(`> Conectado ao ${config.email}`)
  
  for (const contact of contacts) {
    try {
      const message = replaceAll(replaceAll(template, '[NOME]', contact.nome), '[EMPRESA]', contact.empresa)

      console.log('='.repeat(20))
      console.log(`Email: ${contact.email}`)
      console.log(`Assunto: ${config.assunto}`)
      console.log(`Mensagem: ${message}`)
      console.log('='.repeat(20))
      let response = ''
      while (response !== 's' && response !== 'n') {
        response = ''
        response = await getUserInput('Deseja confirmar o envio dessa mensagem? [S/n]: ')
      }

      if (response === 's') {
        await transporter.sendMail({
          from: config.email,
          to: contact.email,
          subject: config.assunto,
          text: message,
          cc: config.teste ? '' : config.supervisor,
          bcc: config.email
        })
        console.log(`> E-mail enviado com sucesso`)
      } else {
        console.log(`> E-mail n enviado`)
      }
    } catch (error) {
      console.error(`> Ocorreu um erro ao enviar e-mail`)
    }
  }
}

main()