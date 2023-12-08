require("dotenv").config();
const xlsx = require("xlsx");
const puppeteer = require("puppeteer");

const keysFile = xlsx.readFile("./Extrema_Ranfs.xlsx", {
  cellText: true,
  //cellDates: true,
});

const worksheetName = "Planilha1";
const worksheet = keysFile.Sheets[worksheetName];

const jsonData = xlsx.utils.sheet_to_json(worksheet, {
  blankrows: true,
  defval: "",
  header: 1,
  rawNumbers: false,
});

const dataArr = [...jsonData].splice(1, jsonData.length - 1);

const testArr = [dataArr[0], dataArr[1]];

const parseDate = (date) => {
  const currentDate = new Date(date);
  console.log(date)

  const month = String(date.split("/")[0]).padStart(2, "0");
  const day = String(currentDate.getDate()).padStart(2, "0");
  const year = currentDate.getFullYear();

  return { day, month, year };
};

(async () => {
  const browser = await puppeteer.launch({
    headless: false,
  });
  const page = await browser.newPage();
  await page.goto("https://extremamg.webiss.com.br/autenticacao/entrar");


  //Function select
  async function selectCompet(dateNote) {
 
    await page.click('[id="mes-competencia"]');

    await page.waitForTimeout(3000);

    const monthNumber =
      dateNote.month[0] == 0 ? dateNote.month[1] : dateNote.month;

    console.log({ monthNumber, dateNote });

    const select = await page.$("select[name='MesCompetencia']");
    console.log({ select });

    select.select(monthNumber);
  }
  


  //- Acessa a página de login
  await page.type('[name="Login"]', process.env.LOGIN);
  await page.type('[name="Senha"]', process.env.PASSWORD);
  await page.click('[id="botao-logar"]');

  //- Acessar os campos
  await page.click('[id="recurso-issqn"]');
  await page.click('[id="recurso-issqn-ranfs"]');
  await page.waitForTimeout(2000);
  await page.click('[id="recurso-issqn-ranfs-prestador"]');
  await page.waitForTimeout(2000);

  for (let i = 0; i < testArr.length; i++) {
    // Trabalhar item
    const item = testArr[i];
    const cnpj = item[5];
    const numberNote = item[0];
    const dateNote = parseDate(item[1]);
    const description = item[8];
    const serviceAmount = Number(item[2]).toFixed(2);

    //Agregar a planilha com o cnpj para pegar a informação se ja foi escriturado

    await page.type('[name="DocumentoTomador"]', cnpj);

    await page.waitForTimeout(2000);
    await page.click('[id="buscar-por-filtro"]');
    await page.waitForTimeout(2000);

    const informacoes = await page.evaluate(() => {
      const linhas = document.querySelectorAll('tr[role="row"].odd'); // Seletor para todas as linhas que correspondem ao padrão
      const armazem = linhas[0];
      
      if (armazem) {
        const numberNote = armazem.querySelector("td:nth-child(6)");
        return numberNote.textContent;
      }
      return;
    });
    

    if (!informacoes) {

      await page.click('[id="criar-ranfs"]')
      await page.waitForTimeout(2000);

      const elements = await page.$$("a");

      elements.forEach(async (element) => {
        const textContent = await element.evaluate((e) => e.textContent);
        if (textContent === "OK") {
          element.click();
        }
      });

      await page.waitForTimeout(2000);
      await page.type('[id="numero"]', numberNote);
      await page.waitForTimeout(3000);
      await page.type(
        '[id="data-emissao"]',
        `${dateNote.day}/${dateNote.month}/${dateNote.year}`
      );
      await page.waitForTimeout(2000);
      await page.click('[id="btnProximo"]')
      await page.waitForTimeout(2000);
      await page.type('[id="tomador-numero-documento"]', cnpj);
      await page.waitForTimeout(3000);
      await page.click('[id="btn-buscar-tomador"]')
      await page.waitForTimeout(3000);
      await page.click('[id="btnProximo"]')
      await page.waitForTimeout(2000);


      selectCompet(dateNote);

      await page.waitForTimeout(2000);
      await page.click('[id="exigibilidade-iss"]')
      const select = await page.$("select[name='ExigibilidadeDeISS']");
      select.select('Exigivel');

      await page.waitForTimeout(2000);
      await page.click('[id="lista-de-servicos-prestador"]')
      const selectservice = await page.$("select[name='AtividadeNoMunicipio.Id']")
      selectservice.select('198')

      await page.waitForTimeout(2000);
      await page.click('[id="CnaeAtividade_Id"]')
      const selectcnae = await page.$("select[name='CnaeAtividade.Id']")
      selectcnae.select('885') 
      await page.waitForTimeout(2000);

      await page.type('[name="Discriminacao"]', description);
      await page.waitForTimeout(2000);
      await page.click('[id="btnProximo"]');
      await page.waitForTimeout(2000);

      const valueDelet = await page.$('[id="valores-servico"]');
      await valueDelet.click({ clickCount: 3 });
      await valueDelet.press("Backspace");
      await page.type('[id="valores-servico"]', serviceAmount);
      await page.waitForTimeout(2000);

      await page.waitForTimeout(3000);
      await page.click('[id="recurso-issqn-ranfs-prestador"]');

      continue;
    } 

    //buscar dados antigo da NFs Caso tenha
    await page.waitForTimeout(2000);
    await page.click('[id="criar-ranfs"]');
    await page.waitForTimeout(2000);

    const elements = await page.$$("a");

    elements.forEach(async (element) => {
      const textContent = await element.evaluate((e) => e.textContent);
      if (textContent === "OK") {
        element.click();
      }
    });

    await page.waitForTimeout(2000);
    await page.type('[id="numero-anterior"]', informacoes);
    await page.waitForTimeout(2000);
    await page.click('[id="botao-carregar-dados-anteriores"]');
    await page.waitForTimeout(2000);

    const campdelet = await page.$('[id="numero"]');
    await campdelet.click({ clickCount: 3 });
    await campdelet.press("Backspace");
    await page.type('[id="numero"]', numberNote);
    await page.waitForTimeout(3000);

    const datedelet = await page.$('[id="data-emissao"]');
    await datedelet.click({ clickCount: 3 });
    await datedelet.press("Backspace");

    await page.type(
      '[id="data-emissao"]',
      `${dateNote.day}/${dateNote.month}/${dateNote.year}`
    );

    await page.waitForTimeout(3000);
    await page.click('[id="btnProximo"]');
    await page.waitForTimeout(3000);
    await page.click('[id="btnProximo"]');

    await page.waitForTimeout(3000);

    //Seleção competencia

    selectCompet(dateNote);

    await page.waitForTimeout(3000);

    //Discriminação da Atividade
    const bodydelet = await page.$('[name="Discriminacao"]');
    await bodydelet.click({ clickCount: 3 });
    await bodydelet.press("Backspace");

    await page.type('[name="Discriminacao"]', description);

    await page.click('[id="btnProximo"]');

    await page.waitForTimeout(3000);

    //Valor da Nota
    const valueDelet = await page.$('[id="valores-servico"]');
    await valueDelet.click({ clickCount: 3 });
    await valueDelet.press("Backspace");

    await page.type('[id="valores-servico"]', serviceAmount);

    await page.waitForTimeout(3000);

    //const selectConfirm = await page.$("a[class='btn.btn-default.btn-large.finish.emitir-ranfs.glyphicons.circle_ok.page-loading']");
    //await selectConfirm.click()

    await page.waitForTimeout(3000);
    await page.click('[id="recurso-issqn-ranfs-prestador"]');
  }
})();
