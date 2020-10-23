const { BrowserWindow, app, ipcMain } = require('electron');

const fs = require('fs');
const path = require('path');
const JSzip = require('jszip')
const xml = require('xml-js')
const handleSquirrelEvent = require('./update');

if (handleSquirrelEvent(app)) {
    return;
}

let mainWin = null


app.on('ready', function () {
    mainWin = new BrowserWindow({
        webPreferences: {
            nodeIntegration: true
        },
    })
    mainWin.loadURL(`file:///${path.join(__dirname, 'index.html')}`)
})


ipcMain.on('protect-sheet', async (evt, args) => {
    try {
        let passwords = await JSON.parse(fs.readFileSync(args.hashes)).sheets;
        let wb = await fs.readFileSync(args.wb)
        let zip = new JSzip();
        zip.loadAsync(wb).then(async function (contents) {

            let Promises = await Object.keys(contents.files).map(async function (filename) {
                return new Promise(async function (res, reject) {
                    if (filename.indexOf("xl/worksheets/") >= 0 && 
                        filename.substring(filename.lastIndexOf(".")) === ".xml") {
                        let content = await zip.file(filename).async('nodebuffer')
                        let data = JSON.parse(xml.xml2json(content, { compact: true }))
                        if (data.hasOwnProperty('worksheet')) {
                            data = xml.json2xml(data, {compact:true})
                            let protection = null;
                            try{
                                protection = {sheetProtection: passwords.filter(e => e.name === filename)[0].sheetProtection}
                            }catch{
                                protection = null;
                            }
                            
                             
                            if (protection.sheetProtection != undefined && 
                                protection.sheetProtection != null){
                                protection = xml.json2xml(JSON.stringify(protection), {compact:true});                            
                                data = data.replace("</sheetData>", `</sheetData>\n${protection}`)
                                data = data.replace("<sheetData/>", `<sheetData/>\n${protection}`)
                            }
                            
                        }
                        try {
                            await zip.file(filename, data, { binary: false });
                            res()
                        } catch (err) {
                            reject(err)
                        }
                    } else res()
                })
            })
            Promise.all(Promises).then(() => {
                zip
                    .generateNodeStream({ type: 'nodebuffer', streamFiles: true })
                    .pipe(fs.createWriteStream(args.wb))
                    .on('finish', function () {
                        evt.reply('succ');
                    });
            }).catch((err)=>{
                evt.reply('err', "Ocorreu um erro inesperado")
            })

        })

    } catch (err) {
        evt.reply("err", "Ocorreu um erro inesperado!")
    }
})

ipcMain.on('unProtect-sheet', async (evt, args) => {
    try {
        let passwords = { sheets: [] }
        let wb = await fs.readFileSync(args.wb)
        let zip = new JSzip();
        zip.loadAsync(wb).then(async function (contents) {
            let Promises = await Object.keys(contents.files).map(async function (filename) {
                return new Promise(async function (res, reject) {

                    if (filename.indexOf("xl/worksheets/") >= 0 && 
                        filename.substring(filename.lastIndexOf(".")) === ".xml") {
                        let content = await zip.file(filename).async('nodebuffer')
                        let data = JSON.parse(xml.xml2json(content, { compact: true }))
                        if (data.hasOwnProperty('worksheet')) {
                            let sheetProtection = data.worksheet.sheetProtection;
                            delete data.worksheet.sheetProtection;
                            passwords.sheets.push({ name: filename, sheetProtection: sheetProtection })
                        } else {
                            passwords.sheets.push({ name: filename, sheetProtection: {} })
                        }

                        let newContent = xml.json2xml(data, { compact: true })

                        try {
                            await zip.file(filename, newContent, { binary: false });
                            res()
                        } catch (err) {
                            reject(err)
                        }
                    } else res()
                })
            })
            Promise.all(Promises).then(() => {

                zip
                    .generateNodeStream({ type: 'nodebuffer', streamFiles: true })
                    .pipe(fs.createWriteStream(args.wb))
                    .on('finish', function () {
                        let ext = args.wb.substring(args.wb.lastIndexOf('.'));

                        fs.writeFile(args.wb.replace(ext, '.json'), JSON.stringify(passwords), function (err, succ) {
                            if (err) evt.reply('err', "Ocorreu um erro inesperado!")
                            else evt.reply('succ');
                        })
                    });


            }).catch(err => evt.reply('err', "Ocorreu um erro inesperado"))

        })

    } catch (err) {
        evt.reply("err", "Ocorreu um erro inesperado!")
    }
})



