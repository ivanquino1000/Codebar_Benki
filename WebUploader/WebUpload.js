const process = require('node:process');
const fs = require('node:fs/promises');
const { chromium, devices } = require('playwright'); // Install npx install chromium 
const os = require('os');
//onst fs = require('fs');
const path = require('path');
const { Console } = require("console");
const { errorMonitor } = require("events");
const { urlencoded } = require("express");
const { rejects } = require('assert');

// Export Method 
// Allowed Extensions From:
//      - .arca.digital 
//      - Odoo - Local IP AdddesS
//      - Non Specified Won't be executed 



const URL_Address_Local_Path =  `${__dirname}/WebUploadUrls.json`;

const Upload_Result_Path =   `${__dirname}/WebUploadResult.txt`

const Local_UploadFile_Path_ArcaDigital =  path.resolve(__dirname, '..', 'Config', 'items.xlsx');
const Local_UploadFile_Path_Odoo =  path.resolve(__dirname, '..', 'Config', 'Odoo_Items.xls');

var ClientsCredentials = {}

// *******************  ClientsCredentials FORMAT   *********************
// {
//     "clients": [
//       {
//         "name": "Soraza",
//         "Url": "https://soraza.arca.digital/items",
//         "WebAppType": "ArcaDigital",
//         "User": "admin.soraza@jbsistemas.com",
//         "Password": "@Misti456@"
//       }
//     ]
//   }
// **********************************************************************

//  RESULTS
//      Undefined
//      CompanyName=Success
//      CompanyName=Failed


// Operation Result Set To Undefined
async function updateResultFile(result) {
    try {
      await fs.writeFile(Upload_Result_Path, result);
    } catch (err) {
      console.log(err);
    }
}
  


main();

async function main() {

    await updateResultFile("Undefined");

    try {

        ClientsCredentials = await getClientsData();
        await UploadWebApp();

    } catch (error) {
        console.error('Error:', error);
    }
}

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



async function UploadWebApp(){
    // Current Directory
    const browser = await chromium.launch({headless:false})
    let uploadResult = "Undefined"
    let ResultsLogger = ""
    for (const client of ClientsCredentials.clients){
        switch(client.WebAppType){
            case 'ArcaDigital':
                uploadResult = await uploadArcaDigital(client,browser);
                ResultsLogger += `${client.name}=${uploadResult}\n`;
                break;
            case 'Odoo':
                uploadResult = await uploadOdoo(client,browser)
                ResultsLogger += `${client.name}=${uploadResult}\n`;
                break;
            default:
                console.log(`${client.name} = Invalid WebApp: ${client.Url}`)
                break;
        }
    }
    updateResultFile(ResultsLogger)
    browser.close();

}

async function uploadArcaDigital(client,browser){
    const page = await browser.newPage();
    
    //  Go to the Upload URL Address
    try{
        await page.goto(client.Url)
    
        //Login Required
        if (page.url().includes('login')){
            await loginArcaDgital(page,client)
        }
        let loginRetries = 0 ;
        while (loginRetries<4){
            try{
                if(!page.url().includes('items') ){
                    console.log(`URL Items Page Failed Retriying: ${loginRetries}`)
                    await loginArcaDgital(page,client);
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
        await beginUploadArcaDigital(page,client)
        console.error('SUCCESS OPERATION');
        
    }catch(error){
        console.error('FAILED OPERATION: Error At Data Uploading :', error.message);
    
    }finally{
        return "Success";
    }
}


async function beginUploadArcaDigital (page,client){
    // Click Export Button
    await page.getByText('Importar').click();

    //await page.click('button.btn.btn-custom.btn-sm.mt-2.mr-2.dropdown-toggle');

    // Products Dropdown Option 

    await page.locator(".dropdown-item.text-1").getByText('Productos').click();
    

    // Click on WareHouse Selector 
    await page.getByPlaceholder("Seleccionar").click();

    //Select Principal warehouse
    await page.locator(".el-select-dropdown__item").getByText('Almacén Oficina Principal').click();

    //Select the Upload File Element and uploads the webapp Uploads Format
    
    // Find the input element by CSS selector
    const inputElement = await page.$(".el-upload__input[name='file']");

    // Set the input files for the input element
    await inputElement.setInputFiles(Local_UploadFile_Path_ArcaDigital);
   
    // PROCEED BUTTON
    await page.locator(".el-button.el-button--primary.el-button--small").getByText('Procesar').click();
    // Wait network to end
    await page.waitForLoadState("networkidle");
}

async function loginArcaDgital(page,client){
    
    const urlRegex = /^(https?:\/\/)?((\d{1,3}\.){3}\d{1,3}|([a-zA-Z0-9-]+\.)+[a-zA-Z]{2,})(\/\S*)?$/;
    const match = client.Url.match(urlRegex);
    
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
    await page.goto(urlObject.protocol + urlObject.domain + "/login")
    await page.type('#email', client.User)
    await page.type('#password', client.Password)
    await page.click('.btn-signin')
    return;
}
