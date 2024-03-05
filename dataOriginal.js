import { dataCompany } from "./floridaData.js";
import puppeteer from "puppeteer";
import * as fuzzball from "fuzzball";
import ExcelJS from "exceljs";

async function handleCardsNavigation() {
    const browser = await puppeteer.launch({
        headless: false,
        slowMo: 10,
    });

    const page = await browser.newPage();

    try {
        const mainPageUrl = `https://search.sunbiz.org/Inquiry/CorporationSearch/SearchResults/EntityName/‎/Page1?searchNameOrder=‎`;
        await page.goto(mainPageUrl);

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("Companies");

        worksheet.columns = [
            { header: "Company Name", key: "companyName" },
            { header: "Address", key: "companyAddress" },
            { header: "Owner", key: "companyOwner" },
            { header: "Phone Number", key: "phoneNumber" },
            { header: "Email", key: "email" },
            { header : "Document Number", key : "documentNumber"}
        ];

        for (const company of dataCompany) {
            await page.click('#SearchTerm');
            await page.type('#SearchTerm', company.Company.replace(/,/g, ''));
            await page.click('#maincontent > div:nth-child(1) > div.navigationBarForm > div > form > input.center-element');

            await page.waitForSelector('#search-results > table > tbody > tr:nth-child(1)', { timeout: 60000 });

            const tableData = await page.evaluate(() => {
                const rows = document.querySelectorAll("#search-results > table > tbody > tr");
                return Array.from(rows, row => {
                    const columns = row.querySelectorAll('td');
                    return Array.from(columns, column => column.textContent.trim().replace(/[,\.]/g, ''));
                });
            });

            const filteredData = tableData.filter(row => {
                const status = row[2]; 
                if (status.toLowerCase() === 'active') {
                    const similarity = fuzzball.partial_ratio(company.Company.toLowerCase(), row[0].toLowerCase());
                    return similarity >= 80; 
                }
                return false; 
            });

            if (filteredData.length > 0) {
                const companyLink = await page.evaluate(() => {
                    const firstRow = document.querySelector("td");
                    const linkElement = firstRow.querySelector('a');
                    return linkElement ? linkElement.href : null;
                });

                if (companyLink) {
                    await page.goto(companyLink);

                    const companyInfo = await page.evaluate(() => {
                        const companyNameElement = document.evaluate('/html/body/div[1]/div[1]/div[2]/div/div[2]/div[1]/p[2]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE,null).singleNodeValue;
                        const companyAddressElement = document.evaluate('/html/body/div[1]/div[1]/div[2]/div/div[2]/div[3]/span[2]/div', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE,null).singleNodeValue;
                        const companyOwnerElement = document.evaluate('/html/body/div[1]/div[1]/div[2]/div/div[2]/div[5]/span[2]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE,null).singleNodeValue;
                        
                        const companyDocumentElement = document.evaluate('/html/body/div[1]/div[1]/div[2]/div/div[2]/div[2]/span[2]/div/span[1]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE,null).singleNodeValue;
                    
                        const companyName = companyNameElement ? companyNameElement.textContent.trim() : null;
                        const companyAddress = companyAddressElement ? companyAddressElement.textContent.trim() : null;
                        const companyOwner = companyOwnerElement ? companyOwnerElement.textContent.trim() : null;
                        const companyNumberCode = companyDocumentElement ? companyDocumentElement.textContent.trim() : null;
                    
                        return { companyName, companyAddress, companyOwner, companyNumberCode };
                    });

                    const matchingData = dataCompany.find(item => item.Company.toLowerCase() === company.Company.toLowerCase());
                    if (matchingData) {
                        companyInfo.phoneNumber = matchingData["Phone Number"];
                        companyInfo.email = matchingData.Email;
                    }

                    console.log(companyInfo);

                    worksheet.addRow({
                        companyName: companyInfo.companyName,
                        companyAddress: companyInfo.companyAddress,
                        companyOwner: companyInfo.companyOwner,
                        phoneNumber: companyInfo.phoneNumber,
                        email: companyInfo.email,
                        documentNumber : companyInfo.companyNumberCode
                    });

                    await page.goBack(); 
                }
            }

            await page.$eval('#SearchTerm', input => input.value = '');
        }

        await workbook.xlsx.writeFile("companies.xlsx");
        console.log("Excel file generated successfully.");
    } catch (error) {
        // Handle 404 error here

        // Redirect to a URL
        const errorPageUrl = `https://search.sunbiz.org/Inquiry/CorporationSearch/SearchResults/EntityName/
        ‎/Page1?searchNameOrder=‎`;
        return await page.goto(errorPageUrl);
    } finally {
        await browser.close();
    }
}

handleCardsNavigation();
