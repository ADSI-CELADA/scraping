import puppeteer from "puppeteer";
import ExcelJS from "exceljs";

async function handleCardsNavigation() {
    const browser = await puppeteer.launch({
        headless: false,
        slowMo: 5,
    });

    const page = await browser.newPage();

    let countPage = 1;
    const maxPages = 79;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Data');

    worksheet.columns = [
        { header: 'Title', key: 'title' },
        { header: 'Address', key: 'address' },
        { header: 'Email', key: 'email' },
        { header: 'Phone Number', key: 'phoneNumber' },
    ];

    try {
        let hasNextPage = true;

        while (hasNextPage && countPage <= maxPages) {
            const mainPageUrl = `https://www.yellowpages.com/search?search_terms=Electricians&geo_location_terms=AL&page=${countPage}`;

            await page.setCacheEnabled(false);
            await navigateWithTimeout(page, mainPageUrl, 60000); // Setting timeout to 60 seconds

            const links = await page.evaluate(() => {
                const cardLinks = Array.from(document.querySelectorAll(".info-section.info-primary h2 a"));
                return cardLinks.map(link => link.href);
            });

            for (const link of links) {
                await navigateWithTimeout(page, link, 60000); // Setting timeout to 60 seconds
            
                const cardData = await page.evaluate(() => {
                    let title = document.querySelector("#main-header > article > div > h1");
                    let address = document.querySelector("#default-ctas > a.directions.small-btn > span");
                    let phoneNumber = document.querySelector("#default-ctas > a.phone.dockable > strong");
                    let emailElement = document.querySelector("a.email-business");
            
                    title = title ? title.innerText : null;
                    address = address ? address.innerText : null;
                    phoneNumber = phoneNumber ? phoneNumber.innerText : null;
            
                    let email = null;
                    if (emailElement) {
                        email = emailElement.getAttribute('href').replace('mailto:', '');
                    }
            
                    return { title, address, email, phoneNumber };
                });
            
                if (cardData.email !== null && cardData.email.trim() !== '') {
                    worksheet.addRow(cardData);
                }
            }

            countPage++;
        }

        await workbook.xlsx.writeFile('electrician-alabana.xlsx');

    } catch (error) {
        console.error('Error during navigation:', error);
    } finally {
        await browser.close();
    }
}

async function navigateWithTimeout(page, url, timeout) {
    try {
        await page.goto(url, { timeout });
    } catch (error) {
        console.error(`Navigation to ${url} timed out after ${timeout} milliseconds.`);
        return error;
    }
}

handleCardsNavigation();
