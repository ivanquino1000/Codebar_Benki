//const { default: puppeteer, registerCustomQueryHandler } = require("puppeteer");
const { chromium, devices } = require('playwright'); // Install npx install chromium 
const notifier = require('node-notifier');
const os = require('os');
const fs = require('node:fs/promises');
const { resourceLimits } = require('worker_threads');
// Export Method 
// Allowed Extensions From:
//      - .arca.digital 
//      - Odoo - Local IP AdddesS
//      - Non Specified Won't be executed 
const URL_Address_Local_Path =  `${__dirname}/WebExportUrls.json`;
const Platform_Downloads_Path = `${os.homedir}\\Downloads\\`

const Export_Result_Path =   `${__dirname}/WebExportResult.txt`


var ClientsCredentials = {}

const SucessNotification = {
        title: 'Descarga de Archivos',
        message: 'EXITOSO',
        timeout: 300000
}
const FailedNotification = {
    title: 'Descarga de Archivos',
    message: 'FALLIDA',
    timeout: 300000
}

// ************** ON DOMAIN CHANGE *********
// Change the URLs text file 
// Change the Clients Object
// Credentials in Login
// Buttons Remain the same


// Operation Result Set To Undefined
async function updateResultFile(result) {
    try {
      await fs.writeFile(Export_Result_Path, result);
    } catch (err) {
      console.log(err);
    }
}
  


// Operation Result Set To Undefined
// file write Won't be reflected on debug mode dkw
// fs.writeFile('WebExportResult.txt', 'Undefined', (err) => {
//     if (err) {
//         console.error('Error writing file:', err);
//     } else {
//         console.log('Result File Reset successfully.');
//     }
// });


async function getClientsData() {
    try {
        const data = await fs.readFile(URL_Address_Local_Path, 'utf8');
        try {
            // Parse the JSON content and return the parsed data
            return JSON.parse(data);
        } catch (parseError) {
            throw new Error(`Error Parsing Client Json: ${parseError}`);
        }
    } catch (readError) {
        throw new Error(`Error Reading Client Json: ${readError}`);
    }
}


main();

async function main() {
    
    await updateResultFile("Undefined");

    try {
        ClientsCredentials = await getClientsData();
        await ExportFromWeb();
    } catch (error) {
        console.error('Error:', error);
    }
}

async function ExportFromWeb(){
    //${__dirname} = C:\Users\ivan\Desktop\Codebar_Benki\WebExporter
    // Current Directory
    const browser = await chromium.launch({headless:false})
    let exportResult = "Undefined"
    let ResultsLogger = ""
    for (const client of ClientsCredentials.clients){
        switch(client.WebAppType){
            case 'ArcaDigital':
                exportResult = await ExportArcaDigital(client,browser);
                ResultsLogger += `${exportResult}`;
                break;
            case 'Odoo':
                exportResult = await exportOdoo(client,browser)
                ResultsLogger += `${exportResult}`;
                break;
            default:
                console.log(`${client.name} = Invalid WebApp: ${client.Url}`)
                break;
        }
    }
    console.log("ResultsLogger: ",ResultsLogger)
    if (ResultsLogger === "Success"){
        notifier.notify(SucessNotification)
    
    }else{
        notifier.notify(FailedNotification)
    }

    updateResultFile(ResultsLogger)
    browser.close();
}

// Custom Export Method Depending on URL Extension
async function ExportArcaDigital(client,browser){
    const page = await browser.newPage();
    
    //Ensure Items Url Path to be loaded
    try{

        await loginArcaDgital(page,client)
        await page.goto(client.Url)

        console.log("<> Items Page Loaded ")
    
        //Log in - Items Redirection
        if (page.url().includes('login')){
            await loginArcaDgital(page,client)
            console.log("<> First Login Completed ")
        }
        let loginRetries = 0 ;
        while (loginRetries<4){
            try{
                // Error Loading site or LogIn Redirection
                if(!page.url().includes('items') ){
                    console.log(`URL Items Page Failed Retriying: ${loginRetries}`)
                    
                    //  LogIn Redirection - Login
                    if (page.url().includes('login')){
                        await loginArcaDgital(page,client)
                    }
                    //  Error Redirection - Reload
                    await page.reload()
                }else{
                    console.log(`Urls Items Page Load Succesfully`);
                    break;
                }
                loginRetries++

            }catch (error) {
                console.error('Error during Export:', error.message);
                return "Failed";
            } 
        }
    }catch (error) {
        console.error('Error during Export:', error.message);
        return "Failed";
    } 

    // Place the download from the Items page
    try {
        await start_Download_ArcaDigital (page)
        console.log('SUCCESS OPERATION');
        return "Success";
        
    }catch(error){
        console.error('FAILED OPERATION: Error At Data Extracting :', error.message);
        return "Failed";
    }finally{
        return "Success";
        
    }
    
}

async function start_Download_ArcaDigital (page){
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
    const downloadPromise = page.waitForEvent('download',{timeout:300000 });

    await ProccessButton[2].click(); 
    //await page.getByText('Download file').click();
    const download = await downloadPromise;

    // Wait for the download process to complete and save the downloaded file somewhere.
    await download.saveAs(Platform_Downloads_Path + download.suggestedFilename());
}

function UrlFactory(url){
    const urlRegex = /^(https?:\/\/)?((\d{1,3}\.){3}\d{1,3}|([a-zA-Z0-9-]+\.)+[a-zA-Z]{2,})(\/\S*)?$/;
    const match = url.match(urlRegex);
    

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
        return urlObject

    }else{
        return;
    }
}

async function loginArcaDgital(page,client){
    let urlObject = UrlFactory(client.Url)
    
    await page.goto(urlObject.protocol + urlObject.domain + "/login")
    
    console.log("<> Login Page Loaded")
    
    await page.type('#email', client.User)
    await page.type('#password', client.Password)
    await page.click('.btn-signin')
    return;

}