const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');

async function getMarks() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Marks');
  worksheet.columns = [
    { header: 'Roll No', key: 'rollNo' },
    { header: 'Name', key: 'name' },
    { header: 'Marks', key: 'marks' },
  ];

  const browser = await puppeteer.launch({ headless: false });
  const page = await browser.newPage();

  await page.goto('https://upmsp.edu.in/ResultIntermediate.aspx');

  // Wait for the city dropdown to load
  await page.waitForSelector('button[data-id="ctl00_cphBody_ddl_districtCode"]');

  // Click on the city dropdown
  await page.click('button[data-id="ctl00_cphBody_ddl_districtCode"]');

  // Wait for the city options to appear
  await page.waitForSelector('ul.dropdown-menu.show');

  // Select PRAYAGRAJ - 55
  await page.click('ul.dropdown-menu.show li:nth-child(61) a'); // 61 is the nth city in drop down

  // Wait for the year dropdown to load
  await page.waitForSelector('#ctl00_cphBody_ddl_ExamYear');

  // Select 2023
  await page.select('#ctl00_cphBody_ddl_ExamYear', '2023');

  // Loop through the roll numbers and get the marks
  for (let rollNo = 22xxxxxxxx; rollNo <= 22xxxxxxxx; rollNo++) {
    // Type the roll number in the input box
    await page.type('input[name="ctl00$cphBody$txt_RollNumber"]', rollNo.toString());

    // Click on the Submit button
    await Promise.all([
      page.waitForNavigation(),
      page.click('#ctl00_cphBody_btnSubmit')
    ]);

    // Wait for the marks to load
    await page.waitForSelector('#ctl00_cphBody_lbl_C_NAME');

    // Get the name and marks obtained
    const name = await page.$eval('#ctl00_cphBody_lbl_C_NAME', el => el.textContent.trim());
    const marks = await page.$eval('#ctl00_cphBody_lbl_MRK_OBT', el => el.textContent.trim());

    console.log(`Roll No: ${rollNo}, Name: ${name}, Marks: ${marks}`);

    // Add row to the worksheet
    worksheet.addRow({ rollNo, name, marks });

    // Go back to the search page
    await page.click('a.card-link.d-print-none');

    // Wait for the city dropdown to load again
    await page.waitForSelector('button[data-id="ctl00_cphBody_ddl_districtCode"]');

    // Click on the city dropdown
    await page.click('button[data-id="ctl00_cphBody_ddl_districtCode"]');

    // Wait for the city options to appear again
    await page.waitForSelector('ul.dropdown-menu.show');

    // Select PRAYAGRAJ - 55 again
    await page.click('ul.dropdown-menu.show li:nth-child(61) a');

    // Clear the roll number input box
    await page.$eval('input[name="ctl00$cphBody$txt_RollNumber"]', el => el.value = '');
  }

  // Write data to a file
  await workbook.xlsx.writeFile('marks.xlsx');

  await browser.close();
}

getMarks();

