const electron = require('electron');
const puppeteer = require('puppeteer');
const XLSX = require('xlsx');
var xl = require('excel4node');

const { app, BrowserWindow, ipcMain, Menu, dialog } = electron;

var concurentPupl, concurentLogin;
var delayInMilliseconds = 10000;
var defaultDelay = 10000;
var inputPhoneNumberArray = [];
let fileNametxt = "";
var wb;
var ws;
var defaultHeight = 15;
var username = "";
var password = "";
//danh sách số điện thoại
let tResult = [];

let mainWindow;
var mainBrowser = null;
var exPath = '';
//C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe
//C:\\Users\\Administrator\\AppData\\Local\\CocCoc\\Browser\\Application\\browser.exe
var startStartIndex = 0;
var rowSpacing = 2;
var directionToSource = "";
var lackPassword = "Mật khẩu";
var wrongLogin = "Tài khoản không hợp lệ, vui lòng thử lại";

var curentIndex = 0;
var headeTitle = "header", errorTitle = "error";

let sleepBetwwenClick = 1500;

const gotTheLock = app.requestSingleInstanceLock(); //singleton
var URL = {
    LOGIN: "https://10.156.0.19/Account/Subs_info_120days.aspx",
    HOME: "https://10.156.0.19/Account/Subs_info_120days.aspx",
    DISCOUNT: "https://10.156.0.19/Account/KMCB_HIST.aspx",
    SERVICE: "https://10.156.0.19/Account/Data_Packages_new.aspx",
};
var ERROR = "Server Error in '/' Application.";
var threshHoldeCount = 7;
const crawlCommand = {
    login: "crawl:login",
    openFile: "crawl:openFile",
    wrongPhoneNumber: "crawl:incorrect_number",
    hideBTN: "crawl:hideBTN",
    networkError: "crawl:network_error",
    result: "crawl:result",
    readError: "crawl:read_error",
    readErrorNull: "crawl:read_error_null",
    readSuccess: "crawl:read_sucess_new",
    readSuccessFirtTime: "crawl:read_sucess_first_time",
    inputfileNotexcel: "crawl:error_choose_not_xlsx",
    doCrawl: "crawl:do",
    runWithFile: "crawl:runwithfile",
    onRunning: "crawl:onrunning",
    currentCrawl: "crawl:currentCrawl",
    loginSuccess: "crawl:login_success",
    log: "crawl:log",
    notFoundNumber: "crawl:not_found_number",
    signalWrite: "crawl:signalWrite", // cho phép write hoặc không, mắc định là cho phép, chỉ khi có dialog , 
    //số không hợp lệ, -1 
    // không tìm thấy số -2
    // hoặc sesion timeout , -3
    //mất kết nối mạng -4 - trường hợp ít xảy ra, khồng xét
};

var canWrite = true, isFound = true;
var mCheckTrue = "Mở", mCheckFalse = "Đóng";
var xlStyleError;
var currentData = [

]

var page, pageLogin;
var breakPerSerrvice = 6;//có 6,5 cột dịch vụ
function createWindow() {
    mainWindow = new BrowserWindow({
        width: 800, height: 600, webPreferences: {
            nodeIntegration: true // dung được require trên html
        }
    });

    //dev tool
    //mainWindow.webContents.openDevTools();

    mainWindow.on('crashed', () => {
        win.destroy();
        createWindow();
    });

    mainWindow.loadURL(`file://${__dirname}/index.html`);

    // Build menu from template
    const mainMenu = Menu.buildFromTemplate(mainMenuTemplate);
    // Insert menu
    Menu.setApplicationMenu(mainMenu);

    mainWindow.on("uncaughtException", async function (e) {
        await mainWindow.webContents.send(crawlCommand.log, "error occurs " + e);
        return false;
    })

    mainWindow.on('closed', function () {
        mainWindow = null;
    })
}

//hàm nothing
function nothing() {

}

// Create menu template
const mainMenuTemplate = [
    {
        label: 'Chức năng',
        submenu: [
            {
                label: 'Chọn tệp chứa danh sách điện thoại',
                accelerator: process.platform == 'darwin' ? 'Command+F' : 'Ctrl+F',
                click() {
                    // if (crawling == false) {
                    chooseSource(readFile, nothing);
                    // }
                }
            },
            {
                label: 'Thoát',
                accelerator: process.platform == 'darwin' ? 'Command+Q' : 'Ctrl+Q',
                click() {
                    app.quit();
                }
            }
        ]
    }
];

if (!gotTheLock) {
    app.quit()
} else {
    app.on('second-instance', async (event, commandLine, workingDirectory) => {
        // Someone tried to run a second instance, we should focus our window.
        if (mainWindow) {
            await mainWindow.webContents.send(crawlCommand.log, "log second-instance");

            dialog.showMessageBox(mainWindow, {
                title: 'Không nên chạy nhiều cửa sổ',
                buttons: ['Đóng'],
                type: 'warning',
                message: 'Để tránh trang web từ chối truy cập nhiều, chỉ nên chạy 1 cửa sổ!',
            });

            if (mainWindow.isMinimized()) {
                myWindow.restore()
            }
            mainWindow.focus()
        }
    })

    // Create myWindow, load the rest of the app, etc...
    //app.whenReady().then(createWindow);
    app.on('ready', createWindow);
}

//app.on('ready', createWindow);

app.on('window-all-closed', function () {
    if (process.platform !== 'darwin') {
        app.quit();
    }
})

app.on('activate', function () {
    if (mainWindow === null) {
        createWindow();
    }
})

//bấm vào menu để mở file
//dùng tại 3 chỗ,
// 1 khi bấm vào menu -> đọc file , chỉ đọc file không làm gì cả
// 2 khi bấm vào nút chọn file khác , chỉ đọc file không làm gì cả
// 3 trường hợp chưa chọn file mà bấm vào nút lấy dữ liệu để crawl , mở đọc file , đọc xong rồi crawl
async function chooseSource(callback1, callback2) {
    dialog.showOpenDialog({
        title: "Chọn đường dẫn tới file text chứa danh sách số điện thoại",
        properties: ['openFile', 'multiSelections']
    }, function (files) {
        if (files !== undefined) {
            // handle files
        }
    }).then(async (result) => {
        if (!result.filePaths[0].endsWith(".xlsx")) {
            await mainWindow.webContents.send(crawlCommand.inputfileNotexcel, true);
        } else {
            directionToSource = result.filePaths[0];
            await mainWindow.webContents.send(crawlCommand.inputfileNotexcel, false);
            callback1(callback2);
        }
    }).catch(err => {
        ////console.log(err);
    });
};

//chuẩn bị file excel
async function prepareExxcel(callback) {

    //khởi tạo mảng
    inputPhoneNumberArray = [];
    tResult.forEach(element => {
        inputPhoneNumberArray.push(element);
    });

    await mainWindow.webContents.send(crawlCommand.log, 'ghi dữ liệu từ excel vào bộ nhớ tiến hành crawl  ' + inputPhoneNumberArray);

    cTotal = inputPhoneNumberArray.length;

    let cTimee = new Date();

    wb = new xl.Workbook();
    ws = wb.addWorksheet("Tra cuu");
    ws.column(1).setWidth(5);//STT

    //Thông tin thuê bao
    ws.column(2).setWidth(15);//Số thuê bao,
    ws.column(3).setWidth(15);//Lớp dịch vụ,
    ws.column(4).setWidth(15);//  Tài khoản chính,
    ws.column(5).setWidth(10);//  Tài khoản KM,
    ws.column(6).setWidth(10);//  Tài khoản KM1,
    ws.column(7).setWidth(10);//  Tài khoản KM2,
    ws.column(8).setWidth(10);// Tài khoản KM3,
    ws.column(9).setWidth(10);//  Tài khoản DK1,
    ws.column(10).setWidth(10);//  Tài khoản DK2,
    ws.column(11).setWidth(5);//  Trạng thái hiện tại,
    ws.column(12).setWidth(5);//   Trạng thái trước,
    ws.column(13).setWidth(20);//   Ngày hết hạn (yyyy-mm-dd),
    ws.column(14).setWidth(20);//   Ngày tạo thuê bao (yyyy-mm-dd),
    ws.column(15).setWidth(20);//   Ngày kích hoạt (yyyy-mm-dd),
    ws.column(16).setWidth(15);//   ACC ALO,
    ws.column(17).setWidth(15);//  Ala carte,
    ws.column(18).setWidth(35);//  Các gói KM được tham gia (09/2020),

    //Khuyến mãi
    ws.column(19).setWidth(15);//  MSISDN,
    ws.column(20).setWidth(25);//  Thời gian thực hiện giao dịch,
    ws.column(21).setWidth(10);//  Loại dịch vụ,
    ws.column(22).setWidth(10);// Bản tin đến,
    ws.column(23).setWidth(15);//  Bản tin phản hồi,
    ws.column(24).setWidth(15);//  Loại giao dịch,
    ws.column(25).setWidth(15);// A La Carte,
    ws.column(26).setWidth(20);// Ngày bắt đầu gói cước,
    ws.column(27).setWidth(20);// Ngày kết thsuc gói cước,
    ws.column(28).setWidth(15);// Trạng thái,
    ws.column(29).setWidth(20);// Lý do,
    ws.column(3).setWidth(25);// Chi tiết,

    //Dịch vụ
    ws.column(31).setWidth(15);// Số thuê bao,
    ws.column(32).setWidth(15);// Gói cước,
    ws.column(33).setWidth(20);// Ngày bắt đầu,
    ws.column(34).setWidth(20);// Ngày kết thúc,
    ws.column(35).setWidth(20);// Ngày thay đổi gần nhất,

    xlStyleSmall = wb.createStyle({
        alignment: {
            vertical: ['center'],
            horizontal: ['center'],
            wrapText: true,
        },
        font: {
            name: 'Arial',
            color: '#324b73',
            size: 12,
        }
    });

    xlStyleError = wb.createStyle({
        alignment: {
            vertical: ['center'],
            horizontal: ['center'],
            wrapText: true,
        },
        font: {
            name: 'Arial',
            color: 'red',
            size: 12,
        }
    });

    xlStyleBig = wb.createStyle({
        alignment: {
            vertical: ['center'],
            horizontal: ['center'],
            wrapText: true,
        },
        font: {
            name: 'Arial',
            color: '#4e3861',
            size: 12,
        }
    });

    xlStyleNone = wb.createStyle({
        alignment: {
            vertical: ['center'],
            horizontal: ['center'],
            wrapText: true,
        },
        font: {
            bold: true,
            name: 'Arial',
            color: '#4e3861',
            size: 12,
        },
    });

    fileNamexlxs = "(" + cTimee.getHours() + " Gio -" + cTimee.getMinutes() + " Phut Ngay " + cTimee.getDate() + " Thang " + (cTimee.getMonth() + 1) + " Nam " + cTimee.getFullYear() + ")   " + fileNametxt;

    let header = [
        "STT",
        //Thông tin thuê bao
        "Số thuê bao",
        "Lớp dịch vụ",
        "Tài khoản chính",
        "Tài khoản KM",
        "Tài khoản KM1",
        "Tài khoản KM2",
        "Tài khoản KM3",
        "Tài khoản DK1",
        "Tài khoản DK2",
        "Trạng thái hiện tại",
        "Trạng thái trước",
        "Ngày hết hạn (yyyy-mm-dd)",
        "Ngày tạo thuê bao (yyyy-mm-dd)",
        "Ngày kích hoạt (yyyy-mm-dd)",
        "ACC ALO",
        "Ala carte",
        "Các gói KM được tham gia (09/2020)",

        //Khuyến mãi
        "MSISDN",
        "Thời gian thực hiện giao dịch",
        "Loại dịch vụ",
        "Bản tin đến",
        "Bản tin phản hồi",
        "Loại giao dịch",
        "A La Carte",
        "Ngày bắt đầu gói cước",
        "Ngày kết thsuc gói cước",
        "Trạng thái",
        "Lý do",
        "Chi tiết",

        //dịch vụ
        "Số thuê bao",
        "Gói cước",
        "Ngày bắt đầu",
        "Ngày kết thúc",
        "Ngày thay đổi gần nhất",
    ];

    for (let i = 0; i < header.length; i++) {
        await mainWindow.webContents.send(crawlCommand.log, "vòng for ghi header  " + i + " title " + headeTitle + "-" + header[i]);
        await writeToXcell(1, Number.parseInt(i) + 1, headeTitle + "-" + header[i]);
    }

    await mainWindow.webContents.send(crawlCommand.log, "puppeteeer file ouput tên là  " + fileNamexlxs);
    ws.row(1).setHeight(defaultHeight);
    startStartIndex = 0;
    await mainWindow.webContents.send(crawlCommand.hideBTN, true);
    callback();
}

function specialForOnlyHitButton() {
    prepareExxcel(doCrawl);
}

async function readFile(callback) {

    let arraySourceFileName = directionToSource.split("\\");
    let isNew = false;
    if (fileNametxt != "") {
        isNew = true;
    }
    //tách tên file
    fileNametxt = arraySourceFileName[arraySourceFileName.length - 1];
    let fileNametxtRemoveExxtension = fileNametxt.replace('.xlsx', '');
    // if (err) {
    //     ////console.log("An error ocurred reading the file :" + err.message);
    //     await mainWindow.webContents.send(crawlCommand.readError, fileNametxt.replace('.xlsx', ''));
    //     return;
    // }

    let workbook = XLSX.readFile(directionToSource);//
    let sheet_name_list = workbook.SheetNames; // laasy cacs sheet
    data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]); //if you have multiple sheets

    if (data == '' || data == null) {
        await mainWindow.webContents.send(crawlCommand.readErrorNull, fileNametxtRemoveExxtension);
    } else {
        tResult = [];
        //assync 1 mảng
        await asyncReadFileExcel(data, function (item) {
            tResult.push(item);
        })
        console.log(tResult);
        await mainWindow.webContents.send(crawlCommand.log, 'dữ liệu trong tệp là  ' + tResult);

        if (isNew == true) {
            //await mainWindow.webContents.send(crawlCommand.log, 'đọc tệp lần đầu tiên thành công tên tệp là ' + fileNametxtRemoveExxtension);
            await mainWindow.webContents.send(crawlCommand.readSuccess, fileNametxtRemoveExxtension);
        }
        else {
            //await mainWindow.webContents.send(crawlCommand.log, 'đọc tệp lần nữa thành công tên tệp là ' + fileNametxtRemoveExxtension);
            await mainWindow.webContents.send(crawlCommand.readSuccessFirtTime, fileNametxtRemoveExxtension);
        }
        callback();
    }

    // fs.readFile(directionToSource, 'utf-8', async (err, data) => {

    //     // Change how to handle the file content
    //     if (data == '' || data == null) {
    //         await mainWindow.webContents.send(crawlCommand.readErrorNull, fileNametxt);
    //     } else {

    //     }
    // });
}

async function writeToFileXLSX() {
    await wb.write(fileNamexlxs);
}

async function writeToXcell(x, y, title) {
    //console.log("Ghi vao o ", x, y, "gia tri", title);
    //await mainWindow.webContents.send(crawlCommand.log, "Ghi vao o " + x + ":" + y + " gia tri " + title);
    title += "";

    if (title.startsWith(headeTitle)) {
        let tTitle = title.split("-")[1];
        title = JSON.stringify(title);
        //title.replace("\"/g","");
        ws.cell(x, y).string(tTitle).style(xlStyleNone);
    } else if (title.startsWith(errorTitle)) {
        let tTitle = title.split("-")[1];
        ws.cell(x, y).string('0' + tTitle).style(xlStyleError);
    } else {
        ws.cell(x, y).string(title).style(xlStyleSmall);
    }
    // }
}

async function writeNumberToCell(x, y, number) {
    await ws.cell(x, y).number(number).style(xlStyleSmall);
}

//sleep đi một vài giây
function timer(ms) {
    return new Promise(res => setTimeout(res, ms));
}

async function asyncReadFileExcel(array, callback) {
    for (let index = 0; index < array.length; index++) {
        await callback(array[index]["Số thuê bao"], index);
    }
}

async function asyncForEach(array, startIndex, callback) {
    let cIndex = 1;
    for (let index = startIndex; index < array.length; index++) {
        try {
            await mainWindow.webContents.send(crawlCommand.log, index + ' / ' + inputPhoneNumberArray.length);

            //đặt lại biến can write ặmc định là true
            canWrite = true;
            isFound = true;
            curentIndex = index;

            await mainWindow.webContents.send(crawlCommand.currentCrawl, (index + 1) + " " + inputPhoneNumberArray.length);
            await callback(array[index], index);
            //await mainWindow.webContents.send(crawlCommand.log, ' craw xong ' + index + ' / ' + inputPhoneNumberArray.length);
            await mainWindow.webContents.send(crawlCommand.onRunning, (index + 1) + " " + inputPhoneNumberArray.length);

            //sau ghi đủ số lượng threshHoldeCount sẽ ghi vào file excel
            if (index % threshHoldeCount === 0 && index > 0) {
                //await mainWindow.webContents.send(crawlCommand.log, 'đã đạt đủ thresshold  ' + threshHoldeCount);
                await writeToFileXLSX();
            }

            //crawl xong 1 số -> nghỉ await timer(delayInMilliseconds);
            if (index < array.length - 1) {
                await mainWindow.webContents.send(crawlCommand.log, 'delay   ' + delayInMilliseconds);
                await timer(delayInMilliseconds);
            }
        } catch (err) {
            await mainWindow.webContents.send(crawlCommand.log, 'err in crawl   ' + err);
        }
    }
}

//crawl

function doLogin(_username, _password) {
    concurentLogin = null;
    //đang login
    concurentLogin = puppeteer.launch({ headless: true, executablePath: exPath == "" ? "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe" : exPath }).then(async browser => {
        mainBrowser = browser;
        pageLogin = await browser.newPage();

        //pageLogin.setDefaultTimeout(0);

        await mainWindow.webContents.send(crawlCommand.loginSuccess, 2);
        await mainWindow.webContents.send(crawlCommand.log, 'doLogin');



        pageLogin.setViewport({ width: 2600, height: 3800 });

        //có dialog hiệnh lên
        //hầu hết các lỗi dialog, -> đóng trình duệt
        //dialog số không hợp lệ(sai định dang số, số quá ngắn, quá dài hoặc otp bị sai), không đóng google
        pageLogin.on('dialog', async dialog => {
            let mssg = dialog.message();
            await mainBrowser.close();
            concurentLogin = null;
            dialog.dismiss();
        });

        await pageLogin.goto(URL.LOGIN);//, { waitUntil: 'networkidle0' });
        await pageLogin.waitForNavigation({ waitUntil: 'networkidle0' });

        // await pageLogin.evaluate(({ _username, _password }) => {
        //     document.getElementById("txtUsername").value = _username;
        //     document.getElementById("txtPassword").value = _password;
        //     document.getElementById("btnLogin").click;
        // }, { _username, _password });

        await pageLogin.$eval('body #ctl01 .wrap-body .inner .tbl-login #txtUsername', (el, value) => el.value = value, _username);
        await pageLogin.$eval('body #ctl01 .wrap-body .inner .tbl-login #txtPassword', (el, value) => el.value = value, _password);


        //ngăn race condition
        await Promise.all([pageLogin.click('#ctl01 .wrap-login .inner .tbl-login #btnLogin'), pageLogin.waitForNavigation({ waitUntil: 'networkidle0' })]);

        //đăng nhập thành công
        await mainWindow.webContents.send(crawlCommand.loginSuccess, 1);

        //crawl data
        ipcMain.on(crawlCommand.doCrawl, async function (e, item) {
            ////console.log(e, item);
            delayInMilliseconds = item == null ? 10000 : item;
            //console.log("delayInMilliseconds", delayInMilliseconds,"directionToSource",directionToSource);
            await mainWindow.webContents.send(crawlCommand.log, 'bấm crawl đường dẫn đến thư mục ' + directionToSource);
            if (directionToSource == "" || directionToSource == null) {
                await chooseSource(readFile, specialForOnlyHitButton);
            } else {
                prepareExxcel(doCrawl);
            }

        })


    }).catch(async (err, browser) => {
        //các trường hợp do user đóng app, hoặc do mất mạng
        await mainWindow.webContents.send(crawlCommand.loginSuccess, -1);
        //await mainWindow.webContents.send(crawlCommand.otp, -1);
        await mainWindow.webContents.send(crawlCommand.result, false);

        await mainWindow.webContents.send(crawlCommand.log, 'uncaught exception ' + err);
        await mainBrowser.close();
        concurentLogin = null;
    });
}

async function doCrawl() {
    canWrite = true;
    //await page.goto(crawlUrl);
    //await mainWindow.webContents.send(crawlCommand.log, 'bắt đầu crawl ');
    const start = async () => {
        await asyncForEach(inputPhoneNumberArray, startStartIndex, async (element, index) => {

            await mainWindow.webContents.send(crawlCommand.log, 'crawl đến phần tử thứ  ' + index + " là số thuê bao " + inputPhoneNumberArray[index] + " = " + element);

            //lấy thông tin  thuê bao
            await pageLogin.goto(URL.LOGIN);

            //khuyến mại
            await pageLogin.goto(URL.DISCOUNT);

            //dịch vụ
            await pageLogin.goto(URL.SERVICE);

            //nhập vào số điện thoại
            await pageLogin.$eval('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtThueBao', (el, value) => el.value = value, element);

            // await pageLogin.evaluate(({ element }) => {
            //     document.getElementById("txtThueBao").value = element;
            //     document.getElementById("btnSearch").click();
            // }, { element });

            //ngăn race condition
            //await Promise.all([pageLogin.click('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #btnSearch'), pageLogin.waitForNavigation({ waitUntil: 'networkidle0' })]);

            //bấm nút tìm
            await pageLogin.click('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #btnSearch');


            //đợi page load
            //cần xử lý sleep vài giây

            //wwait for value change
            //await page.waitForFunction('document.getElementById("txtMSIN").value != "No Value"');
            //dialog không tìm thấy thuê bao hiện lên
            await timer(sleepBetwwenClick);

            let dialogNotFound = await pageLogin.$("body .panel .messager-body");


            let dialogNotFoundvalue = await (await dialogNotFound.getProperty('innerHTML')).jsonValue();
            //await mainWindow.webContents.send(crawlCommand.log, "dialog không tìm thấy số" + dialogNotFound + " value " + dialogNotFoundvalue);
            await mainWindow.webContents.send(crawlCommand.notFoundNumber, "không tìm thấy số" + inputPhoneNumberArray[index]);
            //await mainWindow.webContents.send(crawlCommand.log, "không tìm thấy số" + inputPhoneNumberArray[index]);
            let counterIndexNotFoundCannotFound = index + 1;
            await writeToXcell(index + rowSpacing, 1, errorTitle + "-" + counterIndexNotFoundCannotFound);//số thứ tự
            await writeToXcell(index + rowSpacing, 2, errorTitle + "-" + inputPhoneNumberArray[index] + " không tìm thấy");

            //bấm vào đóng
            await pageLogin.click('body .panel .dialog-button .l-btn span');

            await timer(sleepBetwwenClick);

            //làm mới mảng curent Data
            currentData.length = 0;

            //1
            currentData.push(index + 1);// số thứ tự

            //2
            currentData.push(inputPhoneNumberArray[index]);//"Số thuê bao",

            // await timer(5000);
            try {
                let domElement = "", value = "", valueAlt = "";

                //3
                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtMSIN");
                value = await (await domElement.getProperty('value')).jsonValue();//"MSIN",
                // await mainWindow.webContents.send(crawlCommand.log, 'MSIN  ' + value);
                await currentData.push(value);

                //4
                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtLoaiTB");
                value = await (await domElement.getProperty('value')).jsonValue();//"Loại thuê bao",
                // await mainWindow.webContents.send(crawlCommand.log, 'Loại thuê bao  ' + value);
                await currentData.push(value);

                //5
                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #chkGoiDi");
                value = await (await domElement.getProperty('checked')).jsonValue() === true ? mCheckTrue : mCheckFalse;//"Gọi đi",
                // await mainWindow.webContents.send(crawlCommand.log, 'Gọi đi  ' + value);
                await currentData.push(value);

                //6
                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #chkGoiDen");
                value = await (await domElement.getProperty('checked')).jsonValue() === true ? mCheckTrue : mCheckFalse;//"Gọi đến",
                // await mainWindow.webContents.send(crawlCommand.log, 'Gọi đến  ' + value);
                await currentData.push(value);

                //7
                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtSimType");
                value = await (await domElement.getProperty('value')).jsonValue();//"Loại SIM ",
                //await mainWindow.webContents.send(crawlCommand.log, 'Loại SIM  ' + value);
                await currentData.push(value);

                //8
                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtHangHoiVien");
                value = await (await domElement.getProperty('value')).jsonValue();//"Hạng hội viên",
                //await mainWindow.webContents.send(crawlCommand.log, ' ' + value);
                await currentData.push(value);

                //9
                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtTinh");
                value = await (await domElement.getProperty('value')).jsonValue();//"Hạng hội viên",
                // await mainWindow.webContents.send(crawlCommand.log, 'txtTinh  ' + value);
                await currentData.push(value);

                //10
                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtNgayKH");
                value = await (await domElement.getProperty('value')).jsonValue();//"Ngày KH",
                //await mainWindow.webContents.send(crawlCommand.log, 'Ngày KH  ' + value);
                await currentData.push(value);

                //11
                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtMaKH");
                value = await (await domElement.getProperty('value')).jsonValue();//"Mã KH",
                //await mainWindow.webContents.send(crawlCommand.log, 'Mã KH  ' + value);
                await currentData.push(value);

                //12
                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtMaCQ");
                value = await (await domElement.getProperty('value')).jsonValue();//""Mã CQ"",
                //await mainWindow.webContents.send(crawlCommand.log, 'Mã CQ ' + value);
                await currentData.push(value);

                //13
                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtTB");
                value = await (await domElement.getProperty('value')).jsonValue();//""Tên thuê bao"",
                //await mainWindow.webContents.send(crawlCommand.log, 'Tên thuê bao  ' + value);
                await currentData.push(value);


                //14
                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtNgaySinh");
                value = await (await domElement.getProperty('value')).jsonValue();//""Ngày sinh"",
                //await mainWindow.webContents.send(crawlCommand.log, 'Ngày sinh  ' + value);
                await currentData.push(value);

                //15
                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtSoGT");
                value = await (await domElement.getProperty('value')).jsonValue();//""Số GT"",
                //await mainWindow.webContents.send(crawlCommand.log, 'Số GT  ' + value);
                await currentData.push(value);

                //16
                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtNoiCap");
                value = await (await domElement.getProperty('value')).jsonValue();//""Ngày cấp"",
                //await mainWindow.webContents.send(crawlCommand.log, 'Ngày cấp  ' + value);
                await currentData.push(value);

                //17
                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtPIN");
                value = await (await domElement.getProperty('value')).jsonValue();//""PIN"",

                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtPUK");
                valueAlt = await (await domElement.getProperty('value')).jsonValue();//""PUK"",

                //await mainWindow.webContents.send(crawlCommand.log, 'Số PIN/PUK  ' + value + "/" + valueAlt);
                await currentData.push(value + "/" + valueAlt);//"Số PIN/PUK",

                //18
                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtPIN2");
                value = await (await domElement.getProperty('value')).jsonValue();//""PIN2"",

                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtPUK2");
                valueAlt = await (await domElement.getProperty('value')).jsonValue();//""PUK2"",

                // await mainWindow.webContents.send(crawlCommand.log, 'Số PIN2/PUK2  ' + value + "/" + valueAlt);
                await currentData.push(value + "/" + valueAlt);//"Số PIN2/PUK2",

                //19
                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtDoiTuong");
                value = await (await domElement.getProperty('value')).jsonValue();//""Đối tượng"",
                //await mainWindow.webContents.send(crawlCommand.log, 'Đối tượng  ' + value);
                await currentData.push(value);

                //20
                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtDiaChiChungTu");
                value = await (await domElement.getProperty('value')).jsonValue();//""Địa chỉ chứng từ"",
                //   await mainWindow.webContents.send(crawlCommand.log, 'Địa chỉ chứng từ ' + value);
                await currentData.push(value);

                //21
                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtDiaChiThanhToan");
                value = await (await domElement.getProperty('value')).jsonValue();//""Địa chỉ thanh toán",
                //   await mainWindow.webContents.send(crawlCommand.log, 'Địa chỉ thanh toán' + value);
                await currentData.push(value);

                //22
                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtDiaChiThuongTru");
                value = await (await domElement.getProperty('value')).jsonValue();//""Địa chỉ thường trú",
                //   await mainWindow.webContents.send(crawlCommand.log, 'Địa chỉ thường trú' + value);
                await currentData.push(value);

                //23
                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtTKC");
                value = await (await domElement.getProperty('value')).jsonValue();//""Tài khoản chính",
                //  await mainWindow.webContents.send(crawlCommand.log, 'Tài khoản chính' + value);
                await currentData.push(value);

                //24
                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtHSD");
                value = await (await domElement.getProperty('value')).jsonValue();//""Hạn sử dụng",
                //  await mainWindow.webContents.send(crawlCommand.log, 'Hạn sử dụng' + value);
                await currentData.push(value);

                //25
                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtKhuyenMai");
                value = await (await domElement.getProperty('value')).jsonValue();//""Thuê bao trả trước được tham gia khuyến mại",
                // await mainWindow.webContents.send(crawlCommand.log, 'Thuê bao trả trước được tham gia khuyến mại' + value);
                await currentData.push(value);

                //26
                domElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor #txtKhuyenNghi");
                value = await (await domElement.getProperty('value')).jsonValue();//""Gói cước trả trước ưu tiên mời KH đăng ký",
                //  await mainWindow.webContents.send(crawlCommand.log, 'Gói cước trả trước ưu tiên mời KH đăng ký' + value);
                await currentData.push(value);

                //bấm vào 3g tab
                //  await mainWindow.webContents.send(crawlCommand.log, 'click on 3G tab ');
                //3g tab tại tab thứ 2
                await pageLogin.click("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .box5 #tabtab :nth-child(1) span");
                await timer(sleepBetwwenClick);

                await pageLogin.click("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .box5 #tabtab :nth-child(2) span");
                await timer(sleepBetwwenClick);

                await pageLogin.click("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .box5 #tabtab :nth-child(3) span");
                await timer(sleepBetwwenClick);

                await pageLogin.click("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .box5 #tabtab :nth-child(2) span");
                await timer(sleepBetwwenClick);

                let dataFromTable3G = await pageLogin.$$eval("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .box5 #tabContent .midbox .myTbl tr td", tableData => tableData.map((td) => {
                    return td.innerHTML;
                }));

                // await mainWindow.webContents.send(crawlCommand.log, "dịch vụ 3g " + dataFromTable3G);

                //bấm vào lịch sử thuê bao
                // await mainWindow.webContents.send(crawlCommand.log, 'click lịch sử thuê bao ');
                //lịch sử thêu bao tab tại tab thứ 1
                await pageLogin.click("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .box5 #tabtab :nth-child(1) span");
                await timer(sleepBetwwenClick);

                await pageLogin.click("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .box5 #tabtab :nth-child(2) span");
                await timer(sleepBetwwenClick);

                await pageLogin.click("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .box5 #tabtab :nth-child(3) span");
                await timer(sleepBetwwenClick);

                await pageLogin.click("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .box5 #tabtab :nth-child(1) span");
                await timer(sleepBetwwenClick);

                let dataFromTableLSTB = await pageLogin.$$eval("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .box5 #tabContent .midbox .myTbl tr td", tableData => tableData.map((td) => {
                    return td.innerHTML;
                }));

                //  await mainWindow.webContents.send(crawlCommand.log, "lịch sử thuê bao " + dataFromTableLSTB);

                //bấm vào lịch sử nạp thẻ
                // await mainWindow.webContents.send(crawlCommand.log, 'click lịch sử nạp thẻ ');
                //lịch sử thêu bao tab tại tab thứ 3
                await pageLogin.click("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .box5 #tabtab :nth-child(1) span");
                await timer(sleepBetwwenClick);

                await pageLogin.click("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .box5 #tabtab :nth-child(2) span");
                await timer(sleepBetwwenClick);

                await pageLogin.click("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .box5 #tabtab :nth-child(3) span");
                await timer(sleepBetwwenClick);

                await pageLogin.click("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .box5 #tabtab :nth-child(3) span");
                await timer(sleepBetwwenClick);

                let dataFromTableLSNT = await pageLogin.$$eval("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .box5 #tabContent .midbox .myTbl tr td", tableData => tableData.map((td) => {
                    return td.innerHTML;
                }));

                // await mainWindow.webContents.send(crawlCommand.log, "lịch sử nạp thẻ " + dataFromTableLSNT);

                if (canWrite) {
                    //phần ghi ra file excel
                    //đến phẩn tử 26 là hết phần thông tin khách
                    // await mainWindow.webContents.send(crawlCommand.log, "ghi vào thông tin khách " + currentData);
                    let outerIndex = index;
                    for (let index = 0; index < 25; index++) {
                        await writeToXcell(outerIndex + rowSpacing, index + 1, currentData[index]);
                    }

                    let tempOnlyNeedDay = 0;

                    //3g
                    if (dataFromTable3G != undefined) {
                        let currentCollumn = 27;
                        breakPerSerrvice = 6;
                        let limitRange = dataFromTable3G.length > 18 ? 18 : dataFromTable3G.length; // do chỉ có 3 dịch vụ => 3 * 6 = 18
                        // await mainWindow.webContents.send(crawlCommand.log, "ghi vào thông tin dịch vụ ");
                        for (let index = 0; index < limitRange; index++) {
                            //dataFromTable3G
                            if (index % breakPerSerrvice == 0) {
                                tempOnlyNeedDay = 0;
                                continue;
                            } else {
                                tempOnlyNeedDay++;
                                if (tempOnlyNeedDay == 3 || tempOnlyNeedDay == 4) {
                                    let tDayInside = dataFromTable3G[index].split(" ")[0];
                                    //let tTimeInside = dataFromTable3G[index].split(" ")[1];
                                    await writeToXcell(outerIndex + rowSpacing, currentCollumn, tDayInside);
                                } else {
                                    await writeToXcell(outerIndex + rowSpacing, currentCollumn, dataFromTable3G[index]);
                                }

                                currentCollumn++;
                            }
                        }
                    }

                    //lịch sử thuê bao
                    if (dataFromTableLSTB != undefined) {
                        let currentCollumn = 42;
                        breakPerSerrvice = 5;
                        let startIndex = -1;
                        tempOnlyNeedDay = 0;
                        // await mainWindow.webContents.send(crawlCommand.log, "ghi vào thông tin lịch sử thuê bao gprs");
                        //tìm ra vị trí đầu tiên là dịch vụ gprs
                        dataFromTableLSTB.some((item, index) => {
                            if (item === "GPRS" && index % 9 == 0) {
                                startIndex = index;
                                return true;
                            }
                        });

                        if (startIndex != -1) {//tìm thấy
                            for (let index = startIndex; index < startIndex + 5; index++) {
                                //dataFromTableLSTB
                                // await mainWindow.webContents.send(crawlCommand.log, "ghi vào dịch vụ GPRS " + dataFromTableLSTB[index]);
                                tempOnlyNeedDay++;
                                if (tempOnlyNeedDay == 2) {
                                    let tDayInside = dataFromTableLSTB[index].split(" ")[0];
                                    //let tTimeInside = dataFromTable3G[index].split(" ")[1];
                                    await writeToXcell(outerIndex + rowSpacing, currentCollumn, tDayInside);
                                } else {
                                    await writeToXcell(outerIndex + rowSpacing, currentCollumn, dataFromTableLSTB[index]);
                                }
                                currentCollumn++;
                            }
                        }

                        //await mainWindow.webContents.send(crawlCommand.log, "ghi vào thông tin lịch sử thuê bao ic");

                        //tìm ra vị trí đầu tiên là dịch vụ ic
                        tempOnlyNeedDay = 0;
                        startIndex = -1;
                        dataFromTableLSTB.some((item, index) => {
                            if (item.includes("IC") && index % 9 == 0) {
                                startIndex = index;
                                return true;
                            }
                        });

                        //await mainWindow.webContents.send(crawlCommand.log, "tra ứu dịch vụ IC " + startIndex);
                        if (startIndex != -1) {//tìm thấy
                            for (let index = startIndex; index < startIndex + 5; index++) {
                                //dataFromTableLSTB
                                // await mainWindow.webContents.send(crawlCommand.log, "ghi vào dịch vụ IC " + dataFromTableLSTB[index]);
                                tempOnlyNeedDay++;
                                if (tempOnlyNeedDay == 2) {
                                    let tDayInside = dataFromTableLSTB[index].split(" ")[0];
                                    //let tTimeInside = dataFromTable3G[index].split(" ")[1];
                                    await writeToXcell(outerIndex + rowSpacing, currentCollumn, tDayInside);
                                } else {
                                    await writeToXcell(outerIndex + rowSpacing, currentCollumn, dataFromTableLSTB[index]);
                                }
                                currentCollumn++;
                            }
                        }

                        //await mainWindow.webContents.send(crawlCommand.log, "ghi vào thông tin lịch sử thuê bao oc");
                        //tìm ra vị trí đầu tiên là dịch vụ oc
                        startIndex = -1;
                        tempOnlyNeedDay = 0;
                        dataFromTableLSTB.some((item, index) => {
                            if (item.includes("OC") && index % 9 == 0) {
                                startIndex = index;
                                return true;
                            }
                        });

                        // await mainWindow.webContents.send(crawlCommand.log, "tra ứu dịch vụ OC " + startIndex);
                        if (startIndex != -1) {//tìm thấy
                            for (let index = startIndex; index < startIndex + 5; index++) {
                                //dataFromTableLSTB
                                // await mainWindow.webContents.send(crawlCommand.log, "ghi vào dịch vụ OC " + dataFromTableLSTB[index]);
                                tempOnlyNeedDay++;
                                if (tempOnlyNeedDay == 2) {
                                    let tDayInside = dataFromTableLSTB[index].split(" ")[0];
                                    //let tTimeInside = dataFromTable3G[index].split(" ")[1];
                                    await writeToXcell(outerIndex + rowSpacing, currentCollumn, tDayInside);
                                } else {
                                    await writeToXcell(outerIndex + rowSpacing, currentCollumn, dataFromTableLSTB[index]);
                                }
                                currentCollumn++;
                            }
                        }

                        // mainWindow.webContents.send(crawlCommand.log, "ghi vào thông tin lịch sử thuê bao can");
                        //tìm ra vị trí đầu tiên là dịch vụ oc
                        startIndex = -1;
                        tempOnlyNeedDay = 0;
                        dataFromTableLSTB.some((item, index) => {
                            if (item.includes("CAN") && index % 9 == 0) {
                                startIndex = index;
                                return true;
                            }
                        });

                        // await mainWindow.webContents.send(crawlCommand.log, "tra ứu dịch vụ CAN " + startIndex);
                        if (startIndex != -1) {//tìm thấy
                            for (let index = startIndex; index < startIndex + 5; index++) {
                                //dataFromTableLSTB
                                // await mainWindow.webContents.send(crawlCommand.log, "ghi vào dịch vụ OC " + dataFromTableLSTB[index]);
                                tempOnlyNeedDay++;
                                if (tempOnlyNeedDay == 2) {
                                    let tDayInside = dataFromTableLSTB[index].split(" ")[0];
                                    //let tTimeInside = dataFromTable3G[index].split(" ")[1];
                                    await writeToXcell(outerIndex + rowSpacing, currentCollumn, tDayInside);
                                } else {
                                    await writeToXcell(outerIndex + rowSpacing, currentCollumn, dataFromTableLSTB[index]);
                                }
                                currentCollumn++;
                            }
                        }


                    }

                    //lịch sử nạp thẻ
                    if (dataFromTableLSNT != undefined) {
                        let currentCollumn = 62;
                        breakPerSerrvice = 5;
                        let limitRange = dataFromTableLSNT.length > 15 ? 15 : dataFromTableLSNT.length; // do chỉ có 2 dịch vụ => 2 * 5 = 10
                        //await mainWindow.webContents.send(crawlCommand.log, "ghi vào thông tin nạp thẻ " + dataFromTableLSNT.length + " " + dataFromTableLSNT.leng);
                        tempOnlyNeedDay = 0;
                        if (dataFromTableLSNT.length == 0) {
                            //await mainWindow.webContents.send(crawlCommand.log, "thông tin nạp thẻ undefined");
                            //let noLSNTElement = await pageLogin.$("#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .box5 #tabContent .midbox div");
                            //let noLSNT = await (await noLSNTElement.getProperty('innerHTML')).jsonValue();
                            // await mainWindow.webContents.send(crawlCommand.log, "thông tin nạp thẻ undefined " + noLSNTElement);
                            let currentCollumn = 63;//ô 63 rộng hơn
                            let noLSNT = "Trong 30 ngày gần đây không có thông tin nạp thẻ";
                            await writeToXcell(outerIndex + rowSpacing, currentCollumn, noLSNT);
                        } else {

                            // do nạp thẻ header cũng kaf td, nên cần bắt đầu từ
                            for (let index = 5; index < limitRange; index++) {
                                //dataFromTable3G
                                if (index % 5 === 0) {
                                    tempOnlyNeedDay = 0;
                                }
                                tempOnlyNeedDay++;
                                if (tempOnlyNeedDay === 2) {
                                    let tDayInside = dataFromTableLSNT[index].split(" ")[0];
                                    //let tTimeInside = dataFromTable3G[index].split(" ")[1];
                                    await writeToXcell(outerIndex + rowSpacing, currentCollumn, tDayInside);
                                } else {
                                    await writeToXcell(outerIndex + rowSpacing, currentCollumn, dataFromTableLSNT[index]);
                                }
                                //await mainWindow.webContents.send(crawlCommand.log, "ghi vào thông tin nạp thẻ có nội dung " + dataFromTableLSNT[index]);
                                currentCollumn++;
                            }
                        }
                    }

                } else {
                    let counterIndexNotFound = index + 1;
                    await writeToXcell(index + rowSpacing, 1, errorTitle + "-" + counterIndexNotFound);//số thứ tự
                    await writeToXcell(index + rowSpacing, 2, errorTitle + "-" + inputPhoneNumberArray[index] + " bị lỗi, không tra cứu");
                }

            }
            catch (err) {
                await mainWindow.webContents.send(crawlCommand.log, 'lỗi  ' + err);
            }


        });


        await mainWindow.webContents.send(crawlCommand.log, 'end  ');
        await mainWindow.webContents.send(crawlCommand.log, 'write to excel  ');
        //lần chạy cuối cùng
        await writeToFileXLSX();

        //không được đóng
        // await mainBrowser.close();

        await mainWindow.webContents.send(crawlCommand.result, true);

        //concurentPup = null;
        //crawling = false;
    };

    start();
}

//liên lạc giữa index.js và index html
//open file
ipcMain.on(crawlCommand.openFile, async function (e, item) {
    chooseSource(readFile, nothing);
});

//login
ipcMain.on(crawlCommand.login, async function (e, item) {
    username = item.split(" ")[0];
    password = item.split(" ")[1];
    doLogin(username, password);
});