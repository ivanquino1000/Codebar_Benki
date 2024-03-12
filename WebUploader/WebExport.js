//const { default: puppeteer, registerCustomQueryHandler } = require("puppeteer");
const { chromium, devices } = require('playwright'); // Install npx install chromium 
const os = require('os');
const fs = require('fs');
const { Console } = require("console");
const { errorMonitor } = require("events");
const { urlencoded } = require("express");
// Export Method 
// Allowed Extensions From:
//      - .arca.digital 
//      - Odoo - Local IP AdddesS
//      - Non Specified Won't be executed 
const URL_Address_Local_Path =  `${__dirname}/WebExportUrls.txt`;
const Platform_Downloads_Path = `${os.homedir}\\Downloads\\`


// ************** ON DOMAIN CHANGE *********
// Change the URLs text file 
// Credentials in Login
// Buttons Remain the same



// Operation Result Set To Undefined
// file write Won't be reflected on debug mode dkw
fs.writeFile('WebExportResult.txt', 'Undefined', (err) => {
    if (err) {
        console.error('Error writing file:', err);
    } else {
        console.log('Result File Reset successfully.');
    }
});

const ClientCredentialsList=  {
    Soraza: {
        User:"admin.soraza@jbsistemas.com",
        Password:"@Misti456@"
    },
    Larico: {
        User:"admin.larico@jbsistemas.com",
        Password:"@Misti456@"
    },
    Odoo:{
        User:"12345@gmail.com",
        Password:"abcde"
    }
}

main();

async function main() {
    try {
        await ExportFromWeb();
    } catch (error) {
        console.error('Error:', error);
    }
}

async function ExportFromWeb(){
    //${__dirname} = C:\Users\ivan\Desktop\Codebar_Benki\WebExporter
    // Current Directory

    const ExportItemsPath = `${__dirname}/../`;
    Web_Url_List = getUrlList(URL_Address_Local_Path);
    
    Web_Url_List.forEach(url => {

        const urlRegex = /^(https?:\/\/)?((\d{1,3}\.){3}\d{1,3}|([a-zA-Z0-9-]+\.)+[a-zA-Z]{2,})(\/\S*)?$/;
        const match = url.match(urlRegex);
        
        var urlExtension = ""
        let urlObject ={}

        if (match) {
            const protocol = match[1] || 'http://';
            const domain = match[2];
            const path = match[5] || '/';
            
            urlObject = {
                completeUrl: protocol + domain + path,
                domain: domain,
                path: path,
                protocol: protocol
            };

        }else{
            return;
        }

        // Domain Last Word Criteria
        const parts = urlObject.domain.split('.');
        urlExtension = parts[parts.length - 1]; // Get the last part after splitting by '.'

        switch(urlExtension){
            case 'digital':
                ExportArcaDigital(urlObject)

                break;
            case 'odoo':
                
                break;
            default:
                console.log(`Invalid Url Extension: ${urlObject.completeUrl}` )
                return;
        }
        
    });
    //console.log(ExportItemsPath)


}

// Custom Export Method Depending on URL Extension
async function ExportArcaDigital(url){
    const browser = await chromium.launch({headless:false})
    const page = await browser.newPage();
    
    //Ensure Items Url Path to be loaded
    try{
        await page.goto(url.completeUrl)
    
        //Log in Required
        if (page.url().includes('login')){
            await loginArcaDgital(page,url)
        }
        let loginRetries = 0 ;
        while (loginRetries<4){
            try{
                if(!page.url().includes('items') ){
                    console.log(`URL Items Page Failed Retriying: ${loginRetries}`)
                    await loginArcaDgital(page,url);
                    await page.reload()
                }else{
                    console.log(`Urls Items Page Load Succesfully`);
                    break;
                }
                loginRetries++

            }catch (error) {
                console.error('Error during Export:', error.message);
                fs.writeFileSync('WebExportResult.txt', "Failed");
                await browser.close();
            } 
        }
    }catch (error) {
        console.error('Error during Export:', error.message);
        await browser.close();
    } finally {
        //await browser.close();
    }

    // Place the download from the Items page
    try {
        await start_Download_ArcaDigital (page,url)
        console.error('SUCCESS OPERATION');
        fs.writeFileSync('WebExportResult.txt', "Success");
        
    }catch(error){
        console.error('FAILED OPERATION: Error At Data Extracting :', error.message);
        fs.writeFileSync('WebExportResult.txt', "Failed");
    }finally{
        browser.close();
        
    }
    
}

async function start_Download_ArcaDigital (page,url){
    // Click Export Button
    await page.click('button.btn.btn-custom.btn-sm.mt-2.mr-2.dropdown-toggle');

    // Wait for the dropdown menu to appear
    await page.waitForSelector('a.dropdown-item.text-1');

    // Click on the first element in the dropdown menu - Listado
    await page.click('a.dropdown-item.text-1');
    
    // Click on Time Period Selector 
    const timeRangeLocator = await page.$$('.el-input__inner');
    await timeRangeLocator[2].click({timeout:3000}); 
    

    // Click the ALL option
    //const periodOption = await  page.$$('.el-select-dropdown__item');
    
    const timeRangeOption = await  page.getByText('Todos')//$$('.el-select-dropdown__item');
    await timeRangeOption.click();

   
    // Click on Procced Button
    
    const ProccessButton = await  page.$$('.el-button.el-button--primary.el-button--small');
 
    // Start waiting for download before clicking. Note no await.
    const downloadPromise = page.waitForEvent('download');

    await ProccessButton[2].click(); 
    //await page.getByText('Download file').click();
    const download = await downloadPromise;

    // Wait for the download process to complete and save the downloaded file somewhere.
    await download.saveAs(Platform_Downloads_Path + download.suggestedFilename());
}

async function loginArcaDgital(page,url){
    await page.goto(url.protocol + url.domain + "/login")
    await page.type('#email', ClientCredentialsList.Soraza.User)
    await page.type('#password', ClientCredentialsList.Soraza.Password)
    await page.click('.btn-signin')
    return;
}

// Returns a list with the URL Addresses
function getUrlList(ListPath){
    try {
        const data = fs.readFileSync(ListPath, 'utf-8');
        var urlList = data.trim().split('\n');
        if (urlList.length === 0 ){
            console.log('No elements Found')
            throw new Error('Empty URL List')
        }
        return urlList

    }catch(error){
        console.log('Error Reading Urls List from Text File', error)
        return []
    }
}