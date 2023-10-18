require('dotenv').config()
const puppeteer = require('puppeteer-core');
const sql = require('mssql');

const login = process.env.C6_USER_000169
const password = process.env.C6_GLOBAL_PASSWD

async function main() {
  const config = {
    user: 'GFT',
    password: '%4ffbUBf8F7g',
    server: 'AZ-SQL',
    database: 'MIS_GFT',
    options: {
      encrypt: false,
      trustServerCertificate: false,
    },
  }

  const browser = await puppeteer.launch({
    executablePath: 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
    headless: true,
  });

  try {
    const page = await browser.newPage();
    await page.setDefaultNavigationTimeout(0);
    await page.goto('https://c6.c6consig.com.br/WebAutorizador/');

    await page.waitForTimeout(2000)

    await page.type('#EUsuario_CAMPO', login)
    await page.waitForTimeout(500)
    await page.type('#ESenha_CAMPO', password)
    await page.waitForTimeout(1000)
    await page.click('#lnkEntrar');

    page.on('dialog', async (dialog) => {
      await dialog.accept();
    });
    await page.waitForTimeout(2000);
    const [a] = await page.$x("//a[contains(., 'Esteira')]");
    if (a) {
      await a.hover();
    }
    await page.waitForTimeout(1000);

    await page.click('#WFP2010_PWTCNPROP');

    await page.waitForTimeout(5000);
    await page.click('#btnFiltro_txt');
    await page.waitForTimeout(3000);
    await page.select('select#ctl00_Cph_cboTipoPeriodo_CAMPO', 'DA');
    await page.waitForTimeout(1000);

    // Calcular a data atual menos 3 dias
    const dataAtual = new Date();
    dataAtual.setDate(dataAtual.getDate() - 3);
    const dataFormatada = dataAtual.toLocaleDateString('pt-BR'); // Formato 'dd/MM/yyyy'
    const dataInsert = dataAtual.toISOString().slice(0, 19).replace('T', ' ');

    // Inserir a data calculada no campo
    await page.evaluate((dataFormatada) => {
      document.querySelector('#ctl00_Cph_FIFaixaDatas_edit1_CAMPO').value = dataFormatada;
    }, dataFormatada);

    await page.waitForTimeout(500);
    await page.keyboard.press('PageDown');
    await page.waitForTimeout(1000);

    await page.click('#btnConfirmar_txt');

    await page.waitForTimeout(3000);

    let hasNextPage = true;
    let countInserted = 0;

    // Inicializa a conexão com o SQL Server
    await sql.connect(config);

    // Usar um conjunto para evitar duplicatas
    const uniqueRecords = new Set();

    while (hasNextPage) {
      const tableData = await page.evaluate(() => {
        const tableRows = Array.from(document.querySelectorAll('table tr'));

        const data = [];

        // ignorar o cabeçalho
        for (let i = 1; i < tableRows.length; i++) {
          const row = tableRows[i];
          const columns = Array.from(row.querySelectorAll('td'));

          if (columns.length >= 15) {
            const rowData = columns.map(column => column.textContent.trim());

            // Verifique se a linha não contém apenas botões
            if (!rowData.some(cell => cell.includes('Atualizar') ||
              cell.includes('Filtro') ||
              cell.includes('Bloquear / Desbloquear') ||
              cell.includes('Cancelar') ||
              cell.includes('Voltar'))) {
              data.push(rowData);
            }
          }
        }

        // Filtra as linhas em branco
        const filteredData = data.filter(row => row.length > 0);

        return filteredData;
      });

      for (const rowData of tableData) {
        const atividade = rowData[6].substring(0, 255);
        const proposta = rowData[0].substring(0, 255); // Use um campo exclusivo como chave

        if (atividade === 'ANALISE CORBAN' && !uniqueRecords.has(proposta)) {

          // Verificar se a proposta já existe no banco de dados
          const checkIfExistsQuery = `SELECT COUNT(*) AS count FROM Tb_Consulta_Manual_C6 WHERE PROPOSTA = '${proposta}'`;
          const checkIfExistsResult = await sql.query(checkIfExistsQuery);
          const request = new sql.Request();
          request.query(`
            DELETE FROM tb_consulta_manual_c6 WHERE CORRESPONDENTE = 'R2 PROMOTORA DE VEND'
          `);
          if (checkIfExistsResult.recordset[0].count === 0) {
            request.input('PROPOSTA', sql.VarChar(255), rowData[0].substring(0, 255));
            request.input('CPF', sql.VarChar(255), rowData[1].substring(0, 255));
            request.input('CLIENTE', sql.VarChar(255), rowData[2].substring(0, 255));
            request.input('MODALIDADE', sql.VarChar(255), rowData[3].substring(0, 255));
            request.input('CONVENIO', sql.VarChar(255), rowData[4].substring(0, 255));
            request.input('SITUACAO', sql.VarChar(255), rowData[5].substring(0, 255));
            request.input('ATIVIDADE', sql.VarChar(255), rowData[6].substring(0, 255));
            request.input('FORMALIZACAO', sql.VarChar(255), rowData[7].substring(0, 255));
            request.input('STATUS_FORMALIZACAO', sql.VarChar(255), rowData[8].substring(0, 255));
            request.input('DATA_ATIVIDADE', sql.VarChar(255), rowData[9].substring(0, 255));
            request.input('HORA_ATIVIDADE', sql.VarChar(255), rowData[10].substring(0, 255));
            request.input('CORRESPONDENTE', sql.VarChar(255), rowData[11].substring(0, 255));
            request.input('VLR_PARCELA', sql.VarChar(255), rowData[12].substring(0, 255));
            request.input('VALOR_SOLICITADO', sql.VarChar(255), rowData[13].substring(0, 255));
            request.input('USUARIO', sql.VarChar(255), rowData[14].substring(0, 255));
            request.query(`
              INSERT INTO Tb_Consulta_Manual_C6 
              (PROPOSTA, CPF, CLIENTE, MODALIDADE, CONVENIO, SITUACAO, ATIVIDADE, FORMALIZACAO, STATUS_FORMALIZACAO, DATA_ATIVIDADE, HORA_ATIVIDADE, CORRESPONDENTE, VLR_PARCELA, VALOR_SOLICITADO, USUARIO)
              VALUES 
              (LEFT('${rowData[0]}', 255), 
                LEFT('${rowData[1]}', 255), 
                LEFT('${rowData[2]}', 255), 
                LEFT('${rowData[3]}', 255), 
                LEFT('${rowData[4]}', 255), 
                LEFT('${rowData[5]}', 255), 
                LEFT('${rowData[6]}', 255),
                LEFT('${rowData[7]}', 255),
                LEFT('${rowData[8]}', 255),
                LEFT('${rowData[9]}', 255),
                LEFT('${rowData[10]}', 255), 
                LEFT('${rowData[11]}', 255), 
                LEFT('${rowData[12]}', 255),
                LEFT('${rowData[13]}', 255), 
                LEFT('${rowData[14]}', 255)
                )
            `);
            request.query(`
              INSERT INTO Tb_Log_Minerador_C6 VALUES (getDate(),'${rowData[0]}','${rowData[11]}')
            `);


            // Executa a consulta e aguarda a conclusão
            const result = await request.query();

            if (result.rowsAffected[0] > 0) {
              countInserted++;
              uniqueRecords.add(proposta); 
            }
          }
        }
      }

      // Verifica se o botão "Próximo" está disponível e clicável
      const [nextPageButton] = await page.$x("//a[contains(., 'Próximo >>')]");
      console.log("botão", nextPageButton)
      if (nextPageButton) {
        const isDisabled =  await page.evaluate((button) => {
          return button.getAttribute('disabled') === 'disabled';
          }, nextPageButton);
        console.log(isDisabled)
        if (isDisabled) {
          hasNextPage = false;
          await browser.close();
        } else {
          await nextPageButton.click();
          await page.waitForTimeout(1000);
        }

      } else {
        // Se o botão "Próximo" não for encontrado, encerra o loop
        hasNextPage = false;
      }

    }

    console.log(`Inseridos ${countInserted} registros.`);

    await sql.close();
  } catch (error) {
    console.error('Ocorreu um erro:', error);
  } 
 
}

main();