const puppeteer = require('puppeteer');
const xl = require('excel4node');
const fs = require('fs');
const url = 'https://www.asme.org/about-asme/honors-awards/honors-policy/list-of-society-awards';
const filename = 'asme';

async function scrapeProduct(url) {
    let data = [];
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    const award = await browser.newPage();
    await page.goto(url);

    await page.waitForSelector('tbody');
    const [el] = await page.$x('//*[@id="generalPage"]/div/div[2]/div[2]/section/div[2]/table/tbody');
    const children = await el.getProperty('children');
    const len = await (await children.getProperty('length')).jsonValue();
    for (let i = 2; i <= len; i++) {
        console.log(i);
        const path = '//*[@id="generalPage"]/div/div[2]/div[2]/section/div[2]/table/tbody/tr[' + i + ']/td[1]/p/a';
        let [element] = await page.$x(path);
        if (element === undefined)[element] = await page.$x('//*[@id="generalPage"]/div/div[2]/div[2]/section/div[2]/table/tbody/tr[' + i + ']/td[1]/a');
        //*[@id="generalPage"]/div/div[2]/div[2]/section/div[2]/table/tbody/tr[2]/td[1]/a
        const title = await (await element.getProperty('textContent')).jsonValue();
        const aUrl = await (await element.getProperty('href')).jsonValue();
        await award.goto(aUrl);
        const [p] = await award.$x('//*[@id="generalPage"]/div/div[2]/div[2]/section/div[2]/p[1]');
        const eligibility = await (await p.getProperty('textContent')).jsonValue();
        let [tr] = await award.$x('//*[@id="generalPage"]/div/div[2]/div[2]/section/div[2]/table[1]/tbody/tr[4]/td[2]');
        if (tr === undefined)[tr] = await award.$x('//*[@id="generalPage"]/div/div[2]/div[2]/section/div[2]/div/table[1]/tbody/tr[4]/td[2]');
        const deadline = await (await tr.getProperty('textContent')).jsonValue();
        const cell1 = '=HYPERLINK("' + aUrl + '","' + title + '")';
        const jsonc = {
            award: cell1,
            source: "ASME",
            eligibility: eligibility,
            deadline: deadline
        }
        data[i - 2] = jsonc;
    }
    fs.writeFile('./json/' + filename + '.json', JSON.stringify(data), 'utf8', err => {
        if (err) console.log('Some error occured: ' + err);
        else console.log('Saved!');
    });
    console.log({ data });
    await browser.close();
    writeFile(filename);
}

function writeFile(filename) {
    const wb = new xl.Workbook();
    const ws = wb.addWorksheet(filename);

    const rawFile = fs.readFileSync('./json/' + filename + '.json');
    const raw = JSON.parse(rawFile);
    let files = [];
    for (each in raw) {
        files.push(raw[each]);
    }
    const obj = files.map(e => {
        return e;
    });

    let rowIndex = 1;
    obj.forEach(record => {
        let columnIndex = 1;
        Object.keys(record).forEach(columnName => {
            ws.cell(rowIndex, columnIndex++)
                .string(record[columnName])
        });
        rowIndex++;
    });
    wb.write('./xlsx/' + filename + '.xlsx');
}

scrapeProduct(url);