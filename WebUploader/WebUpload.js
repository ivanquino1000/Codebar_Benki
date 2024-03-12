//const { default: puppeteer, registerCustomQueryHandler } = require("puppeteer");
const { chromium, devices } = require('playwright'); // Install npx install chromium 
const os = require('os');
const fs = require('fs');
const path = require('path');
const { Console } = require("console");
const { errorMonitor } = require("events");
const { urlencoded } = require("express");
// Export Method 
// Allowed Extensions From:
//      - .arca.digital 
//      - Odoo - Local IP AdddesS
//      - Non Specified Won't be executed 
const URL_Address_Local_Path =  `${__dirname}/WebUploadUrls.json`;
const Local_UploadFile_Path_ArcaDigital =  path.resolve(__dirname, '..', 'Config', 'items.xlsx');//`${__dirname}/../Config/items.xlsx`;
const Local_UploadFile_Path_Odoo =  `${__dirname}/../Config/Odoo_Items.xls`;

var ClientsCredentials = {}


// //  LOCAL FILES TO UPLOAD - FORMATS BY PLATFORM  
// const Rantii_Arca ={
//     Local_Import_Items_Path: `${__dirname}/../config/items.xlsx`,
//     WebApp_User: '',
//     WebApp_Password: ''
// }
// const Jaiki_Arca ={
//     Local_Import_Items_Path: `${__dirname}/../config/items.xlsx`,
//     WebApp_User: '',
//     WebApp_Password: ''
// }
// const Soraza_Arca ={
//     Local_Import_Items_Path: `${__dirname}/../config/items.xlsx`,
//     WebApp_User: "admin.soraza@jbsistemas.com",
//     WebApp_Password: "@Misti456@"
// }

// const Larico_Arca ={
//     Local_Import_Items_Path: `${__dirname}/../config/items.xlsx`,
//     WebApp_User: "admin.larico@jbsistemas.com",
//     WebApp_Password: "@Misti456@"
// }

// const Gri_Arca ={
//     Local_Import_Items_Path: `${__dirname}/../config/items.xlsx`,
//     WebApp_User: '',
//     WebApp_Password: ''
// }

// // Note: Default "Import Plantilla " is MISSING A Concept = "Disponible en pdv" Boolean to TRUE For all new Products
// const Odoo ={
//     Local_Import_Items_Path: `${__dirname}/items.xlsx`,
//     WebApp_User: '',
//     WebApp_Password: ''
// }

// ************** ON DOMAIN CHANGE *********
// Change the URLs text file 
// Credentials in Login
// Buttons Remain the same




// Operation Result Set To Undefined
// file write Won't be reflected on debug mode dkw
fs.writeFile('WebUploadResult.txt', 'Undefined', (err) => {
    if (err) {
        console.error('Error writing file:', err);
    } else {
        console.log('Result File Reset successfully.');
    }
});


main();

async function main() {
    try {
        ClientsCredentials = await getClientsData();
        await UploadWebApp();
    } catch (error) {
        console.error('Error:', error);
    }
}

async function getClientsData() {
    return new Promise((resolve, reject) => {
        // Read the file asynchronously
        fs.readFile(URL_Address_Local_Path, 'utf8', (err, data) => {
            if (err) {
                reject(err); // Reject the promise if an error occurs
                return;
            }

            try {
                // Parse the JSON content and resolve the promise with the parsed data
                resolve(JSON.parse(data));
            } catch (parseError) {
                reject(parseError); // Reject the promise if parsing fails
            }
        });
    });
}



async function UploadWebApp(){
    //${__dirname} = C:\Users\ivan\Desktop\Codebar_Benki\WebExporter
    // Current Directory
    for (const client of ClientsCredentials.clients){
        switch(client.WebAppType){
            case 'ArcaDigital':
                await uploadArcaDigital(client)
                break;
            case 'Odoo':
                await uploadOdoo(client)
                break;
            default:
                break;
        }
    }

}
async function uploadArcaDigital(client){
    const browser = await chromium.launch({headless:false})
    const page = await browser.newPage();
    
    //Ensure Items Url Path to be loaded
    try{
        await page.goto(client.Url)
    
        //Log in Required
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
                fs.writeFileSync('WebUploadResult.txt', "Failed");
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
        await beginUploadArcaDigital(page,client)
        console.error('SUCCESS OPERATION');
        fs.writeFileSync('WebUploadResult.txt', "Success");
        
    }catch(error){
        console.error('FAILED OPERATION: Error At Data Uploading :', error.message);
        fs.writeFileSync('WebUploadResult.txt', "Failed");
    }finally{
        browser.close();
        
    }
}

// // Custom Export Method Depending on URL Extension
// async function ExportArcaDigital(url){
//     const browser = await chromium.launch({headless:false})
//     const page = await browser.newPage();
    
//     //Ensure Items Url Path to be loaded
//     try{
//         await page.goto(url.completeUrl)
    
//         //Log in Required
//         if (page.url().includes('login')){
//             await loginArcaDgital(page,client)
//         }
//         let loginRetries = 0 ;
//         while (loginRetries<4){
//             try{
//                 if(!page.url().includes('items') ){
//                     console.log(`URL Items Page Failed Retriying: ${loginRetries}`)
//                     await loginArcaDgital(page,client);
//                     await page.reload()
//                 }else{
//                     console.log(`Urls Items Page Load Succesfully`);
//                     break;
//                 }
//                 loginRetries++

//             }catch (error) {
//                 console.error('Error during Export:', error.message);
//                 fs.writeFileSync('WebUploadResult.txt', "Failed");
//                 await browser.close();
//             } 
//         }
//     }catch (error) {
//         console.error('Error during Upload:', error.message);
//         await browser.close();
//     } finally {
//         //await browser.close();
//     }

//     // Place the download from the Items page
//     try {
//         await upload_select_local (page,client)
//         console.error('SUCCESS OPERATION');
//         fs.writeFileSync('WebUploadtResult.txt', "Success");
        
//     }catch(error){
//         console.error('FAILED OPERATION: Error At Data Upload :', error.message);
//         fs.writeFileSync('WebUploadtResult.txt', "Failed");
//     }finally{
//         browser.close();
//     }
    
// }

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
    console.log("🚀 ~ beginUploadArcaDigital ~ Local_UploadFile_Path_ArcaDigital:", Local_UploadFile_Path_ArcaDigital)
    
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
