const electron = require('electron');
const puppeteer = require('puppeteer');
const XLSX = require('xlsx');
var xl = require('excel4node');

const { app, BrowserWindow, ipcMain, Menu, dialog } = electron;

var concurentPupl, concurentLogin;
var delayInMilliseconds = 10000;
var inputPhoneNumberArray = [];
let fileNametxt = "";
var wb;
var ws;
var defaultHeight = 15;
var username = "";
var password = "";
//danh sách số điện thoại
let tResult = [];
var optionPuppeteer = {
    args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--disable-accelerated-2d-canvas',
        '--no-first-run',
        '--no-zygote',
        '--single-process', // <- this one doesn't works in Windows
        '--disable-gpu',
        "--proxy-server='direct://'",
        '--proxy-bypass-list=*'
    ]
};
//args: ["--proxy-server='direct://'", '--proxy-bypass-list=*']
let mainWindow;
var mainBrowser = null;
var exPath = '';
//C:\Program Files\Google\Chrome\Application\\chrome.exe
//C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe
//C:\\Users\\Administrator\\AppData\\Local\\CocCoc\\Browser\\Application\\browser.exe
var startStartIndex = 0;
var rowSpacing = 2;
var directionToSource = "";
var lackPassword = "Mật khẩu";
var wrongLogin = "Tài khoản không hợp lệ, vui lòng thử lại";
var noneName = "Không có";
var curentIndex = 0;
var headeTitle = "header", errorTitle = "error";
var isRewrite = false;//quyeest ddinhj khi
let sleepBetwwenMain = 1000;
const gotTheLock = app.requestSingleInstanceLock(); //singleton
var URL = {
    LOGIN: "https://10.156.0.19/Account/Subs_info_120days.aspx",
    HOME: "https://10.156.0.19/Account/Subs_info_120days.aspx",
    SERVICE: "https://10.156.0.19/Account/Data_Packages_new.aspx",
    DISCOUNT: "https://10.156.0.19/Account/KMCB_2021.aspx",
    //"https://10.156.0.19/Account/KMCB_HIST.aspx",
    //https://10.156.0.19/Account/KMCB_2021.aspx
};
var ERROR = "Server Error in '/'";
var NOINFO = "Không có thông tin thuê bao ";
var WRONGINFO = " thuê bao bị sai số";
var threshHoldeCount = 15;
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
    thongtin: "crawl:thongtin",
    data: "crawl:data",
    khuyenmai: "crawl:khuyenmai",
    signalWrite: "crawl:signalWrite", // cho phép write hoặc không, mắc định là cho phép, chỉ khi có dialog , 
    //số không hợp lệ, -1 
    // không tìm thấy số -2
    // hoặc sesion timeout , -3
    //mất kết nối mạng -4 - trường hợp ít xảy ra, khồng xét
};

var thongtin = true,
    data = false,
    khuyenmai = false;

var canWrite = true;

var xlStyleError;
var currentData = [], currentDiscount = [], currentService = [];
var unitExcel = [], discountExcel1 = [], discountExcel2 = [], serviceExcel = [];//là mảng hai chiều chứa danh sách cá unit excel và thuộc tính
var discountHeader1 = [], discountHeader2 = [], infoHeader = [], serviceHeader = [], mainHeader = [];
var currentDiscount1 = [], currentDiscount2 = [];
var defaultHeader = [
    "STT",
    "Số thuê bao",
], defaultInfoHeader = [], defaultServiceHeader = [], defaultDiscountHeader1 = [], defaultDiscountHeader2 = [];
//giờ chỉ lấy 1 disocunt
var isDiscount2 = false;//true có 2 disocunt - false có 1 discount
var nameHeader = [
    //Thông tin thuê bao
    "Tài khoản chính",
    "Ngày hết hạn (yyyy-mm-dd)",
    "Các gói KM được tham gia (09/2020)", // cần thay đổi
]
var nameDiscount = [
    "MSISDN",
    "Thời gian bắt đầu đăng ký",
    "Gói cước",
    "Loại chu kỳ",
    "Chu kỳ hiện tại",
    "Số lần đã đăng ký trong ngày",
    "Thời gian thay đổi gần nhất",
    "Thời gian thay hết hạn chu kỳ",//-- mới thêm
    "Trạng thái",
    "Action",
];
// var nameDiscount1 = [
//     "MSISDN",//bỏ
//     "Thời gian bắt đầu đăng ký",
//     "Gói cước",
//     "Loại chu kỳ",
//     "Chu kỳ hiện tại",
//     "Số lần đã đăng ký trong ngày",
//     "Thời gian thay đổi gần nhất",
//     "Thời gian thay hết hạn chu kì", -- mới thêm
//     "Trạng thái",
//     "Action",
// ];
// var nameDiscount2 = [
//     "MSISDN",//bỏ
//     "Thời gian thực hiện giao dịch",
//     "Loại dịch vụ",
//     "Bản tin đến",
//     "Bản tin phản hồi",
//     "Loại giao dịch",
//     "A La Carte",
//     "Ngày bắt đầu gói cước",
//     "Ngày kết thúc gói cước",
//     "Trạng thái",
//     "Lý do",
//     "Chi tiết"
// ];
var nameService = [
    "STT",
    "Số thuê bao",
    "Gói cước",
    "Ngày bắt đầu",
    "Ngày kết thúc",
    "Ngày thay đổi gần nhất",
];
var curentProcerssingIndex = 0;
var page, pageLogin;
var currentDiscountCount = 1;
var currentServiceCount = 1;
//var breakPerSerrvice = 6;//có 6,5 cột dịch vụ
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

//singleton
// if (!gotTheLock) {
//     app.quit()
// } else {
//     app.on('second-instance', async (event, commandLine, workingDirectory) => {
//         // Someone tried to run a second instance, we should focus our window.
//         if (mainWindow) {
//             await mainWindow.webContents.send(crawlCommand.log, "log second-instance");

//             dialog.showMessageBox(mainWindow, {
//                 title: 'Không nên chạy nhiều cửa sổ',
//                 buttons: ['Đóng'],
//                 type: 'warning',
//                 message: 'Để tránh trang web từ chối truy cập nhiều, chỉ nên chạy 1 cửa sổ!',
//             });

//             if (mainWindow.isMinimized()) {
//                 myWindow.restore()
//             }
//             mainWindow.focus()
//         }
//     })

//     // Create myWindow, load the rest of the app, etc...
//     //app.whenReady().then(createWindow);
//     app.on('ready', createWindow);
// }

app.on('ready', createWindow);

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

async function createExcelMain() {
    wb = new xl.Workbook();
    ws = wb.addWorksheet("Tra cuu");
    ws.column(1).setWidth(5);//STT
    ws.column(2).setWidth(15);//Số thuê bao,
    // ws.column(3).setWidth(15);//Lớp dịch vụ,
    // ws.column(4).setWidth(15);//  Tài khoản chính,
    // ws.column(5).setWidth(25);//  Tài khoản KM,
    // ws.column(6).setWidth(25);//  Tài khoản KM1,
    // ws.column(7).setWidth(25);//  Tài khoản KM2,
    // ws.column(8).setWidth(25);// Tài khoản KM3,
    // ws.column(9).setWidth(25);//  Tài khoản DK1,
    // ws.column(10).setWidth(25);//  Tài khoản DK2,
    // ws.column(11).setWidth(25);//  Trạng thái hiện tại,
    // ws.column(12).setWidth(25);//   Trạng thái trước,
    // ws.column(13).setWidth(25);//   Ngày hết hạn (yyyy-mm-dd),
    // ws.column(14).setWidth(25);//   Ngày tạo thuê bao (yyyy-mm-dd),
    // ws.column(15).setWidth(25);//   Ngày kích hoạt (yyyy-mm-dd),
    // ws.column(16).setWidth(25);//   ACC ALO,
    // ws.column(17).setWidth(25);//  Ala carte,
    // ws.column(18).setWidth(35);//  Các gói KM được tham gia (09/2020),

    //Thông tin thuê bao
    ws.column(3).setWidth(15);//  Tài khoản chính,
    ws.column(4).setWidth(25);//   Ngày hết hạn (yyyy-mm-dd),
    ws.column(5).setWidth(35);//  Các gói KM được tham gia (09/2020),

    //Đăng ký gói cước Data
    //ws.column(6).setWidth(15);// Gói cước,
    //ws.column(8).setWidth(20);// Ngày bắt đầu,
    //ws.column(9).setWidth(20);// Ngày kết thúc,
    //ws.column(10).setWidth(20);// Ngày thay đổi gần nhất,

    //Khuyến mãi
    //ws.column(11).setWidth(20);//  Thời gian bắt đầu ĐK,
    //ws.column(12).setWidth(15);//  Gói cước,
    //ws.column(13).setWidth(10);//  Loại chu kỳ,
    //ws.column(14).setWidth(15);// Chu kỳ hiện tại,
    //ws.column(15).setWidth(15);//  Số lần đã ĐK trong ngày,
    //ws.column(16).setWidth(20);//  Thời gian thay đổi gần nhất,
    //ws.column(17).setWidth(20);// Thời gian hết hạn chu kỳ,
    //ws.column(18).setWidth(15);// Trạng thái,
    //ws.column(19).setWidth(10);// Action,
}

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

    createExcelMain();


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
    /*
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

        //dịch vụ
        //"Số thuê bao",
        //"Gói cước",
        //"Ngày bắt đầu",
       // "Ngày kết thúc",
       // "Ngày thay đổi gần nhất",

        //Khuyến mãi
       // "MSISDN",
       // "Thời gian thực hiện giao dịch",
       // "Loại dịch vụ",
       // "Bản tin đến",
       // "Bản tin phản hồi",
       // "Loại giao dịch",
       // "A La Carte",
       // "Ngày bắt đầu gói cước",
       // "Ngày kết thsuc gói cước",
       // "Trạng thái",
       // "Lý do",
       // "Chi tiết",
    ];
    
    for (let i = 0; i < header.length; i++) {
        await mainWindow.webContents.send(crawlCommand.log, "vòng for ghi header  " + i + " title " + headeTitle + "-" + header[i]);
        await writeToXcell(1, Number.parseInt(i) + 1, headeTitle + "-" + header[i]);
    }
    */
    mainHeader = [...defaultHeader];
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
    let indata = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]); //if you have multiple sheets

    if (indata == '' || indata == null) {
        await mainWindow.webContents.send(crawlCommand.readErrorNull, fileNametxtRemoveExxtension);
    } else {
        tResult = [];
        //assync 1 mảng
        await asyncReadFileExcel(indata, function (item) {
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
    //     if (indata == '' || indata == null) {
    //         await mainWindow.webContents.send(crawlCommand.readErrorNull, fileNametxt);
    //     } else {

    //     }
    // });
}

async function writeToFileXLSX() {
    await wb.write(fileNamexlxs);
}

async function changeSomeHTMLEntity(a) {
    a = a.replace(/&nbsp;/g, " ");
    a = a.replace(/&lt;/g, "<");
    a = a.replace(/&gt;/g, ">");
    a = a.replace(/&amp;/g, "&");
    a = a.replace(/&quot;/g, '"');
    a = a.replace(/&apos;/g, "'");
    a = a.replace(/&cent;/g, "cent");
    a = a.replace(/&pound;/g, "pound");
    a = a.replace(/&yen;/g, "yen");
    a = a.replace(/&euro;/g, "euro");
    a = a.replace(/&copy;/g, "copy");
    a = a.replace(/&cent;/g, "reg");
    return a;
}

async function writeToXcell(x, y, title) {
    //console.log("Ghi vao o ", x, y, "gia tri", title);
    //await mainWindow.webContents.send(crawlCommand.log, "Ghi vao o " + x + ":" + y + " gia tri " + title);
    title += "";
    try {
        if (title.startsWith(headeTitle)) {
            let tTitle = title.split("-")[1];
            title = JSON.stringify(title);
            //title.replace("\"/g","");
            ws.cell(x, y).string(tTitle).style(xlStyleNone);//xlStyleNone //xlStyleSmall
        } else if (title.startsWith(errorTitle)) {
            let tTitle = title.split("-")[1];
            ws.cell(x, y).string('0' + tTitle).style(xlStyleError);
        } else {
            title = await changeSomeHTMLEntity(title);
            ws.cell(x, y).string(title).style(xlStyleSmall);
        }
    } catch (e) {
        //await mainWindow.webContents.send(crawlCommand.log, 'error in write excel   ' + e);
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
    //C:\\Users\\Admin\\AppData\\Local\\CocCoc\\Browser\\Application\\browser.exe
    concurentLogin = puppeteer.launch({ args: ['--no-sandbox', "--proxy-server='direct://'", '--proxy-bypass-list=*'], headless: true, ignoreHTTPSErrors: true, executablePath: exPath == "" ? "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe" : exPath }).then(async browser => {
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
            await mainWindow.webContents.send(crawlCommand.log, 'Trang web hiện alert ' + mssg);
            await dialog.dismiss();
            //await mainWindow.webContents.send(crawlCommand.wrongPhoneNumber, inputPhoneNumberArray[curentIndex]);
            //await mainWindow.webContents.send(crawlCommand.signalWrite, -1);
            await mainWindow.webContents.send(crawlCommand.log, 'wronggg ' + mssg);
            await mainWindow.webContents.send(crawlCommand.notFoundNumber, "không tìm thấy số" + inputPhoneNumberArray[index]);
            canWrite = false;
        });

        await pageLogin.goto(URL.LOGIN);//, { waitUntil: 'networkidle0' });
        //await pageLogin.waitForNavigation({ waitUntil: 'networkidle0' });

        // await pageLogin.evaluate(({ _username, _password }) => {
        //     document.getElementById("txtUsername").value = _username;
        //     document.getElementById("txtPassword").value = _password;
        //     document.getElementById("btnLogin").click;
        // }, { _username, _password });
        //await mainWindow.webContents.send(crawlCommand.log, 'username ' + _username + "password" + _password);

        //let arrayName = await pageLogin.$$('body #ctl01 .page .main .accountInfo');     
        //let result = (await (await arrayName[0].getProperty('innerHTML')).jsonValue());
        //await mainWindow.webContents.send(crawlCommand.log, 'content ' + result);

        await pageLogin.$eval('body #ctl01 .page .main .accountInfo #MainContent_LoginUser_UserName', (el, value) => el.value = value, _username);
        await pageLogin.$eval('body #ctl01 .page .main .accountInfo #MainContent_LoginUser_Password', (el, value) => el.value = value, _password);


        //ngăn race condition
        await Promise.all([pageLogin.click('body #ctl01 .page .main .accountInfo #MainContent_LoginUser_LoginButton'), pageLogin.waitForNavigation({ waitUntil: 'networkidle0' })]);


        //đợi 1 vài giây
        await timer(sleepBetwwenMain);

        //trường hợp quên chưa nhập tên hoặc mật khẩu
        let dataFromLoginSummary = await pageLogin.$$eval("body #ctl01 .page .main #MainContent_LoginUser_LoginUserValidationSummary ul li", liData => liData.map((li) => {
            return li.innerHTML;
        }));

        //nếu text login hiện là faile thì gửi tín hiệu faile
        let dataFromLoginSummarySpan = await pageLogin.$$eval("body #ctl01 .page .main .failureNotification", spanData => spanData.map((span) => {
            return span.innerHTML;
        }));

        //await mainWindow.webContents.send(crawlCommand.log, 'dataFromLoginSummary ' + dataFromLoginSummary);
        //await mainWindow.webContents.send(crawlCommand.log, 'dataFromLoginSummary.length ' + dataFromLoginSummary.length);
        //await mainWindow.webContents.send(crawlCommand.log, 'dataFromLoginSummarySpan ' + dataFromLoginSummarySpan);
        //await mainWindow.webContents.send(crawlCommand.log, 'dataFromLoginSummarySpan.length ' + dataFromLoginSummarySpan.length);

        let isPass = true;
        if (isPass && dataFromLoginSummary != undefined) {
            if (dataFromLoginSummary.length > 0) {
                //sai tên đăng nhập hoặc mật khẩu
                isPass = false;
                //await mainWindow.webContents.send(crawlCommand.log, 'dataFromLoginSummary ' + dataFromLoginSummary);
                //await mainWindow.webContents.send(crawlCommand.log, 'wrongLogin');
                await mainWindow.webContents.send(crawlCommand.loginSuccess, 0);
            }
        }
        if (isPass && dataFromLoginSummarySpan != undefined) {
            if (dataFromLoginSummarySpan.length > 0) {
                //nhập sai ten đăng nhập hoặc mật khẩu
                isPass = false;
                //await mainWindow.webContents.send(crawlCommand.log, 'dataFromLoginSummarySpan ' + dataFromLoginSummarySpan);
                await mainWindow.webContents.send(crawlCommand.log, 'pasword or username wrong');
                await mainWindow.webContents.send(crawlCommand.loginSuccess, -3);
            }
        }
        if (isPass) {
            // await browser.close();
            // concurentLogin = null;

            //đăng nhập thành công
            await mainWindow.webContents.send(crawlCommand.loginSuccess, 1);
        }

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

    unitExcel = [...Array(0)];
    discountExcel1 = [...Array(0)];
    discountExcel2 = [...Array(0)];
    serviceExcel = [...Array(0)];

    defaultServiceHeader = [...Array(0)];
    defaultDiscountHeader1 = [...Array(0)];
    defaultDiscountHeader2 = [...Array(0)];

    let elementNoNumberContent;//Nội dung không có số này

    let elementWrongNumberContent;//Nội dung Số bị sai

    const start = async () => {
        await asyncForEach(inputPhoneNumberArray, startStartIndex, async (element, index) => {
            try {
                curentProcerssingIndex = index;
                await mainWindow.webContents.send(crawlCommand.log, '=====================================');
                await mainWindow.webContents.send(crawlCommand.log, '=====================================');
                await mainWindow.webContents.send(crawlCommand.log, '=====================================');

                await mainWindow.webContents.send(crawlCommand.log, 'crawl đến phần tử thứ  ' + index + " là số thuê bao " + inputPhoneNumberArray[index] + " = " + element);

                await mainWindow.webContents.send(crawlCommand.log, 'Số điện thoại ' + element + " type of " + typeof element);
                //reset currentHeader 
                //currentHeader = defaultHeader;

                //gặp alert chưa biết, có lẽ là lỗi,crawl sang cái tiếp theo

                isRewrite = false;

                defaultHeader = ["STT",
                    "Số thuê bao",];

                //làm mới mảng curent Data
                currentData = [...Array(0)];
                currentDiscount1 = [...Array(0)];
                currentDiscount2 = [...Array(0)];
                currentService = [...Array(0)];

                serviceHeader = [...Array(0)];
                discountHeader1 = [...Array(0)];
                discountHeader2 = [...Array(0)];

                //ép kiểu string
                element += "";
                //element 84913477588 0944854975
                if (element.startsWith("84") && element.length >= 11) {
                    element = element.substring(2, element.length);
                } else if (element.startsWith("0") && element.length >= 10) {
                    element = element.substring(1, element.length);
                }

                await mainWindow.webContents.send(crawlCommand.log, 'thongtin ' + thongtin);
                await mainWindow.webContents.send(crawlCommand.log, 'data ' + data);
                await mainWindow.webContents.send(crawlCommand.log, 'khuyenmai ' + khuyenmai);

                //1
                currentData.push(index + 1);// số thứ tự
                //await writeToXcell(index + rowSpacing, 1, index + 1);
                //2
                currentData.push(inputPhoneNumberArray[index]);//"Số thuê bao",

                //==========================================================================
                //lấy thông tin  thuê bao
                await pageLogin.goto(URL.HOME);
                //nhập vào số điện thoại
                await pageLogin.$eval('body #ctl01 .page .main #query .msisdn #MainContent_msisdn', (el, value) => el.value = value, element);
                await Promise.all([pageLogin.click('body #ctl01 .page .main #query #MainContent_submit_button'), pageLogin.waitForNavigation({ waitUntil: 'networkidle0' })]);

                //await page.waitForFunction("document.querySelector('.wrapper') && document.querySelector('.wrapper').clientHeight != 0");
                await timer(sleepBetwwenMain);

                //không có thông tin số
                elementNoNumberContent = await pageLogin.$$eval("body #ctl01 .page .main #query #MainContent_Grid2D", spanData => spanData.map((span) => {
                    return span.innerHTML;
                }));

                //số bị sai
                elementWrongNumberContent = await pageLogin.$$eval("body span h1", spanData => spanData.map((span) => {
                    return span.innerHTML;
                }));

                //đúng
                let dataFromTableHome = await pageLogin.$$eval("body #ctl01 .page .main #wrapper #MainContent_Grid2D tr td", tableData => tableData.map((td) => {
                    return td.innerHTML;
                }));

                if (thongtin) {
                    //thông tin thuê bao đều như nhau, không lo chuyện thêm header
                    let outerIndex = index;

                    //await mainWindow.webContents.send(crawlCommand.log, 'elementNoNumberContent  ' + elementNoNumberContent);
                    //await mainWindow.webContents.send(crawlCommand.log, 'elementWrongNumberContent length ' + elementWrongNumberContent.length + " elementWrongNumberContent " + elementWrongNumberContent + "check " + elementWrongNumberContent[0].includes(ERROR) + "type of");
                    //await mainWindow.webContents.send(crawlCommand.log, 'dataFromTableHome  ' + dataFromTableHome);

                    let tempOnlyNeedDay = 0;

                    let isPass = true;

                    let retryCount = 0;

                    while (isPass && elementNoNumberContent != undefined && retryCount == 3) {
                        if (isPass && elementNoNumberContent != undefined) {
                            if (elementNoNumberContent.length > 0) {
                                retryCount++;
                                //await mainWindow.webContents.send(crawlCommand.log, 'elementNoNumberContent ' + elementNoNumberContent);
                                await pageLogin.$eval('body #ctl01 .page .main #query .msisdn #MainContent_msisdn', (el, value) => el.value = value, element);
                                await Promise.all([pageLogin.click('body #ctl01 .page .main #query #MainContent_submit_button'), pageLogin.waitForNavigation({ waitUntil: 'networkidle0' })]);
                                await timer(sleepBetwwenMain);
                            }
                        }
                    }

                    if (retryCount == 3) {
                        currentData.push(errorTitle + "-" + NOINFO + " " + inputPhoneNumberArray[index]);
                        isPass = false;
                    }

                    retryCount = 0;
                    while (isPass && elementNoNumberContent != undefined && retryCount == 3) {
                        if (isPass && elementWrongNumberContent != undefined) {
                            if (elementWrongNumberContent.length > 0 && elementWrongNumberContent[0].includes(ERROR)) {
                                retryCount++;
                                await timer(sleepBetwwenMain);
                            }
                        }
                    }

                    if (retryCount == 3) {
                        currentData.push(errorTitle + "-" + inputPhoneNumberArray[index] + " " + WRONGINFO);
                        isPass = false;
                    }

                    if (isPass && dataFromTableHome != undefined) {
                        //let currentCollumn = 2;
                        //breakPerSerrvice = 6;
                        //let limitRange = dataFromTableHome.length > 18 ? 18 : dataFromTableHome.length; // do chỉ có 3 dịch vụ => 3 * 6 = 18
                        // for (let index = 0; index < dataFromTableHome.length; index++) {
                        //     //dataFromTableHome
                        //     if (index % 2 == 1) {//chỉ lẻ mới lấy
                        //         currentData.push(dataFromTableHome[index]);
                        //         //await writeToXcell(outerIndex + rowSpacing, currentCollumn, dataFromTableHome[index]);
                        //         //currentCollumn++;
                        //     }
                        // }
                        // for (let index = 0; index < dataFromTableHome.length; index++) {
                        //     if (index % 2 == 1) {//chỉ lẻ mới lấy
                        //         currentData.push(dataFromTableHome[index]);
                        //     }
                        // }
                        if (dataFromTableHome.length > 0) {
                            currentData.push(dataFromTableHome[5]);
                            await mainWindow.webContents.send(crawlCommand.log, 'tài khoản chính ' + dataFromTableHome[5]);
                            currentData.push(dataFromTableHome[23]);
                            await mainWindow.webContents.send(crawlCommand.log, 'ngày hết han ' + dataFromTableHome[23]);
                            currentData.push(dataFromTableHome[33]);
                            nameHeader[2] = dataFromTableHome[32];
                        } else {
                            //canWrite = false;
                            currentData.push(errorTitle + "-" + inputPhoneNumberArray[curentProcerssingIndex] + " bị lỗi, không tra cứu");
                            //currentData[1] = "";
                        }
                        //await mainWindow.webContents.send(crawlCommand.log, 'header ' + dataFromTableHome[32]);
                        //await mainWindow.webContents.send(crawlCommand.log, 'các gói khuyến mại được tham gia ' + dataFromTableHome[33]);
                    } else {
                        currentData.push(errorTitle + "-" + inputPhoneNumberArray[curentProcerssingIndex] + " bị lỗi, không tra cứu");
                    }
                }
                //await mainWindow.webContents.send(crawlCommand.log, 'currentData ' + currentData);

                if (data) {
                    //Dịch vụ 
                    await pageLogin.goto(URL.SERVICE);

                    //nhập vào số điện thoại
                    await pageLogin.$eval('body #ctl01 .page .main #query .msisdn #MainContent_msisdn', (el, value) => el.value = value, element);
                    await Promise.all([pageLogin.click('body #ctl01 .page .main #query #MainContent_submit_button'), pageLogin.waitForNavigation({ waitUntil: 'networkidle0' })]);

                    //await page.waitForFunction("document.querySelector('.wrapper') && document.querySelector('.wrapper').clientHeight != 0");
                    await timer(sleepBetwwenMain);


                    //không có thông tin số
                    elementNoNumberContent = await pageLogin.$$eval("body #ctl01 .page .main #query #MainContent_result_messages", spanData => spanData.map((span) => {
                        return span.innerHTML;
                    }));

                    let dataFromTableService = await pageLogin.$$eval("body #ctl01 .page .main #wrapper #MainContent_GridView1 tr td", tableData => tableData.map((td) => {
                        return td.innerHTML;
                    }));

                    // await mainWindow.webContents.send(crawlCommand.log, "ghi vào thông tin khách " + currentData);

                    //await mainWindow.webContents.send(crawlCommand.log, 'dataFromTableService  ' + dataFromTableService + " dataFromTableService != null " + dataFromTableService != null);

                    //await mainWindow.webContents.send(crawlCommand.log, 'elementNoNumberContent  ' + elementNoNumberContent);
                    //await mainWindow.webContents.send(crawlCommand.log, 'dataFromTableService  ' + dataFromTableService);

                    isPass = true;

                    let retryCount = 0;

                    while (isPass && elementNoNumberContent != undefined && retryCount == 3) {
                        if (isPass && elementNoNumberContent != undefined) {
                            //await mainWindow.webContents.send(crawlCommand.log, 'không có thông tin số length ' + elementNoNumberContent.length);
                            if (elementNoNumberContent.length > 0) {
                                retryCount++;
                                await pageLogin.$eval('body #ctl01 .page .main #query .msisdn #MainContent_msisdn', (el, value) => el.value = value, element);
                                await Promise.all([pageLogin.click('body #ctl01 .page .main #query #MainContent_submit_button'), pageLogin.waitForNavigation({ waitUntil: 'networkidle0' })]);
                                await timer(sleepBetwwenMain);
                            }
                        }
                    }

                    if (retryCount == 3) {
                        isPass = false;
                        //await mainWindow.webContents.send(crawlCommand.log, 'elementNoNumberContent ' + elementNoNumberContent);
                        //currentService.push(errorTitle + "-" + NOINFO + " " + inputPhoneNumberArray[index]);
                    }


                    //tối đa lấy 3 dòng
                    if (isPass && dataFromTableService != undefined) {
                        await mainWindow.webContents.send(crawlCommand.log, 'có data');
                        //listTempService = [];
                        let currentCollumn = 0;
                        //let limitRange = dataFromTableService.length > 18 ? 18 : dataFromTableService.length; // do chỉ có 3 dịch vụ => 3 * 6 = 18
                        currentServiceCount = 1;
                        for (let index = 0; index < dataFromTableService.length; index++) {
                            if (currentServiceCount > 3) {
                                break;
                            }
                            //dataFromTableService
                            if (currentCollumn > 1) {
                                currentService.push(dataFromTableService[index]);
                                serviceHeader.push(nameService[currentCollumn] + " " + currentServiceCount);
                            }
                            currentCollumn++;
                            if (currentCollumn == 6) {
                                currentServiceCount++;
                                currentCollumn = 0;
                            }
                        }
                        if(dataFromTableService.length == 0){
                            currentService.push(errorTitle + "-" + inputPhoneNumberArray[curentProcerssingIndex] + " bị lỗi, không tra cứu");
                        }
                    } else {
                        serviceHeader = [...nameDiscount];
                        currentService.push(errorTitle + "-" + inputPhoneNumberArray[curentProcerssingIndex] + " bị lỗi, không tra cứu");
                    }
                }
                await mainWindow.webContents.send(crawlCommand.log, 'currentService ' + currentService);
                await mainWindow.webContents.send(crawlCommand.log, 'serviceHeader ' + serviceHeader);

                //==========================================================================
                if (khuyenmai) {
                    //khuyến mại
                    await pageLogin.goto(URL.DISCOUNT);
                    //nhập vào số điện thoại
                    await pageLogin.$eval('body #ctl01 .page .main #query .msisdn #MainContent_msisdn', (el, value) => el.value = value, element);
                    await Promise.all([pageLogin.click('body #ctl01 .page .main #query #MainContent_submit_button'), pageLogin.waitForNavigation({ waitUntil: 'networkidle0' })]);

                    //await page.waitForFunction("document.querySelector('.wrapper') && document.querySelector('.wrapper').clientHeight != 0");
                    await timer(sleepBetwwenMain);

                    let dataFromTableDiscount1 = await pageLogin.$$eval("body #ctl01 .page .main #wrapper #MainContent_GridView1 tr td", tableData => tableData.map((td) => {
                        console.log("td", td, "td innerHTML", td.innerHTML);
                        return td.innerHTML;
                    }));

                    // let tt = await pageLogin.$$eval("body #ctl01 .page .main #wrapper #MainContent_GridView1", tableData => tableData.map((td) => {
                    //     return td.innerHTML;
                    // }));


                    // await mainWindow.webContents.send(crawlCommand.log, 'tt  ' + tt);

                    //bỏ qua bảng thứ 2
                    // let dataFromTableDiscount2 = await pageLogin.$$eval("body #ctl01 .page .main #wrapper #MainContent_GridView2 tr td", tableData => tableData.map((td) => {
                    //     return td.innerHTML;
                    // }));

                    //await mainWindow.webContents.send(crawlCommand.log, 'dataFromTableDiscount1 ' + dataFromTableDiscount1.length);
                    //await mainWindow.webContents.send(crawlCommand.log, 'dataFromTableDiscount2 ' + dataFromTableDiscount2.length);

                    //không có thông tin số
                    elementNoNumberContent = await pageLogin.$$eval("body #ctl01 .page .main #query #MainContent_result_check", spanData => spanData.map((span) => {
                        return span.innerHTML;
                    }));

                    //số bị sai
                    elementWrongNumberContent = await pageLogin.$$eval("body span h1", spanData => spanData.map((span) => {
                        return span.innerHTML;
                    }));

                    await mainWindow.webContents.send(crawlCommand.log, 'elementNoNumberContent  ' + elementNoNumberContent);
                    await mainWindow.webContents.send(crawlCommand.log, 'elementWrongNumberContent  ' + elementWrongNumberContent);

                    isPass = true;
                    let retryCount = 0;
                    while (isPass && elementNoNumberContent != undefined && retryCount == 3) {
                        if (isPass && elementNoNumberContent != undefined) {
                            if (elementNoNumberContent.length > 0) {
                                retryCount++;
                                await pageLogin.$eval('body #ctl01 .page .main #query .msisdn #MainContent_msisdn', (el, value) => el.value = value, element);
                                await Promise.all([pageLogin.click('body #ctl01 .page .main #query #MainContent_submit_button'), pageLogin.waitForNavigation({ waitUntil: 'networkidle0' })]);
                                await timer(sleepBetwwenMain);
                            }
                        }
                    }

                    if (retryCount == 3) {
                        //currentDiscount1.push(errorTitle + "-" + NOINFO + " " + inputPhoneNumberArray[index]);
                        isPass = false;
                    }

                    retryCount = 0;

                    while (isPass && elementWrongNumberContent != undefined && retryCount == 3) {
                        if (isPass && elementWrongNumberContent != undefined) {
                            if (elementWrongNumberContent.length > 0 && elementWrongNumberContent[0].includes(ERROR)) {
                                retryCount++;
                                await pageLogin.$eval('body #ctl01 .page .main #query .msisdn #MainContent_msisdn', (el, value) => el.value = value, element);
                                await Promise.all([pageLogin.click('body #ctl01 .page .main #query #MainContent_submit_button'), pageLogin.waitForNavigation({ waitUntil: 'networkidle0' })]);
                                await timer(sleepBetwwenMain);
                            }
                        }
                    }

                    if (retryCount == 3) {
                        //currentDiscount1.push(errorTitle + "-" + inputPhoneNumberArray[index] + " " + WRONGINFO);
                        isPass = false;
                    }

                    if (isPass) {
                        //1 không trống
                        if (dataFromTableDiscount1.length > 0) {
                            //await mainWindow.webContents.send(crawlCommand.log, 'dataFromTableDiscount1 không trống');
                            let currentCollumn = 0;
                            currentDiscountCount = 1;
                            for (let index = 0; index < dataFromTableDiscount1.length; index++) {
                                //dataFromTableDiscount1
                                if (currentCollumn > 0) {
                                    currentDiscount1.push(dataFromTableDiscount1[index]);
                                    discountHeader1.push(nameDiscount[currentCollumn] + " " + currentDiscountCount);
                                }
                                currentCollumn++;
                                if (currentCollumn == 10) {
                                    currentDiscountCount++;
                                    currentCollumn = 0;
                                }
                                if (currentDiscountCount > 5) {
                                    break;
                                }
                            }
                            if(dataFromTableDiscount1.length == 0){
                                currentDiscount1.push(errorTitle + "-" + inputPhoneNumberArray[curentProcerssingIndex] + " bị lỗi, không tra cứu");
                            }
                        } else {
                            discountHeader1 = [...nameDiscount];
                            currentDiscount1.push(errorTitle + "-" + inputPhoneNumberArray[curentProcerssingIndex] + " bị lỗi, không tra cứu");
                        }

                        //await mainWindow.webContents.send(crawlCommand.log, 'dataFromTableDiscount1  ' + dataFromTableDiscount1);
                        //bỏ qua bảng bên dưới
                        // if (dataFromTableDiscount2.length > 0) {
                        //     //listTempDiscount = [];
                        //     let currentCollumn = 0;
                        //     currentDiscountCount = 1;
                        //     for (let index = 0; index < dataFromTableDiscount2.length; index++) {
                        //         //dataFromTableDiscount1
                        //         if (currentCollumn > 0) {
                        //             currentDiscount2.push(dataFromTableDiscount2[index]);
                        //             discountHeader2.push(nameDiscount2[currentCollumn] + " " + currentDiscountCount);
                        //         }
                        //         currentCollumn++;
                        //         if (currentCollumn == 12) {
                        //             currentDiscountCount++;
                        //             currentCollumn = 0;
                        //         }
                        //     }
                        // }
                    } else {
                        currentDiscount1.push("...");
                    }
                }

                //await mainWindow.webContents.send(crawlCommand.log, 'currentDiscount ' + currentDiscount);
                //await mainWindow.webContents.send(crawlCommand.log, 'discountHeader ' + discountHeader);
                //==========================================================================
                if (thongtin) {
                    unitExcel.push([...currentData]);
                }
                if (data) {
                    serviceExcel.push([...currentService]);
                }
                if (khuyenmai) {
                    discountExcel1.push([...currentDiscount1]);
                }
                //discount1 không rỗng --- bản cũ, khi mà có số không có discount 1 có discount 2, có số lại có cả disocunt 1 và discount 2 
                // if (dataFromTableDiscount1 != null) {
                //     if (discountExcel1.length == 0) {
                //         //nghĩa là số điện thoại hiện tại có giảm giá 1,nhưng đây là số đầu tiên có => lần lượt push mảng rỗng vào các số trước
                //         for (var k = discountExcel1.length; k < index; k++) {
                //             discountExcel1.push([...Array(0)]);
                //         }
                //     }
                //     discountExcel1.push([...currentDiscount1]);
                // } else {
                //     //nếu trước đây mà đã có số điện thoại nào có giảm giá 1, thêm array rỗng
                //     if (discountExcel1.length > 0) {
                //         discountExcel1.push([...Array(0)]);
                //     }
                //     //nếu trước đây mà chưa có số điện thoại nào có giảm giá 1, không cần thêm
                // }

                //discountExcel2.push([...currentDiscount2]);

                let currentIndexHeader = 19;

                //await mainWindow.webContents.send(crawlCommand.log, 'unitExcel ' + JSON.stringify(unitExcel));
                //await mainWindow.webContents.send(crawlCommand.log, 'discountExcel ' + JSON.stringify(discountExcel));
                //await mainWindow.webContents.send(crawlCommand.log, 'serviceExcel ' + JSON.stringify(serviceExcel));

                //thêm lại header bị thiếu
                //await mainWindow.webContents.send(crawlCommand.log, 'currentServiceCount ' + currentServiceCount);
                //await mainWindow.webContents.send(crawlCommand.log, 'serviceHeader length ' + serviceHeader.length);
                //await mainWindow.webContents.send(crawlCommand.log, 'defaultServiceHeader length ' + defaultServiceHeader.length);
                if (thongtin) {
                    isRewrite = true;
                }
                if (data) {
                    if (serviceHeader.length > defaultServiceHeader.length) {
                        isRewrite = true;
                        defaultServiceHeader = [...Array(0)];
                        defaultServiceHeader = [...serviceHeader];

                        /*
                        serviceHeader.forEach((item, index) => {
                            ws.column(currentIndexHeader).setWidth(25);
                            currentIndexHeader++;
                        });
                        */
                        await mainWindow.webContents.send(crawlCommand.log, 'after serviceHeader length ' + serviceHeader.length);
                        await mainWindow.webContents.send(crawlCommand.log, 'after defaultServiceHeader length ' + defaultServiceHeader.length);
                    }
                }
                // serviceExcel = serviceExcel.map(async (item, index) => {
                //     //thêm service thiếu
                //     let tempService = [...Array(0)];
                //     if (item.length < defaultServiceHeader.length) {
                //         for (var j = item.length; j < defaultServiceHeader.length; j++) {
                //             tempService.push("");
                //         }
                //     }
                //     item = item.concat(tempService);
                //     await mainWindow.webContents.send(crawlCommand.log, 'service number ' + index + " add " + tempService.length + " item " + tempService);
                //     await mainWindow.webContents.send(crawlCommand.log, 'service number ' + index + " after added length " + item.length + " item " + item);
                // });

                //await mainWindow.webContents.send(crawlCommand.log, 'serviceExcel ' + JSON.stringify(serviceExcel));

                //await mainWindow.webContents.send(crawlCommand.log, 'currentDiscountCount ' + currentDiscountCount);
                //await mainWindow.webContents.send(crawlCommand.log, 'discountHeader1 length ' + discountHeader1.length);
                //await mainWindow.webContents.send(crawlCommand.log, 'discountHeader2 length ' + discountHeader2.length);
                //await mainWindow.webContents.send(crawlCommand.log, 'defaultDiscountHeader1 length ' + defaultDiscountHeader1.length);
                //await mainWindow.webContents.send(crawlCommand.log, 'defaultDiscountHeader2 length ' + defaultDiscountHeader2.length);
                if (khuyenmai) {
                    if (discountHeader1.length > defaultDiscountHeader1.length) {
                        isRewrite = true;
                        defaultDiscountHeader1 = [...Array(0)];
                        defaultDiscountHeader1 = [...discountHeader1];
                        /*
                        discountHeader.forEach((item, index) => {
                            ws.column(currentIndexHeader).setWidth(25);
                            currentIndexHeader++;
                        });
                        */
                        await mainWindow.webContents.send(crawlCommand.log, 'after discountHeader1 length ' + discountHeader1.length);
                        await mainWindow.webContents.send(crawlCommand.log, 'after defaultDiscountHeader1 length ' + defaultDiscountHeader1.length);
                    }
                }
                //bỏ qua bảng bên dưới

                //if (discountHeader2.length > defaultDiscountHeader2.length) {
                //     isRewrite = true;
                //    defaultDiscountHeader2 = [...Array(0)];
                //    defaultDiscountHeader2 = [...discountHeader2];
                /*
                discountHeader.forEach((item, index) => {
                    ws.column(currentIndexHeader).setWidth(25);
                    currentIndexHeader++;
                });
                */
                //await mainWindow.webContents.send(crawlCommand.log, 'after discountHeader2 length ' + discountHeader2.length);
                //await mainWindow.webContents.send(crawlCommand.log, 'after defaultDiscountHeader2 length ' + defaultDiscountHeader2.length);
                //}

                //đay là bản cũ
                if (isRewrite == true) {

                    if (data) {
                        serviceExcel = await Promise.all(serviceExcel.map(async (item, index) => {
                            //thêm service thiếu
                            let tempService = [...Array(0)];
                            if (item.length < defaultServiceHeader.length) {
                                for (var j = item.length; j < defaultServiceHeader.length; j++) {
                                    tempService.push("");
                                }
                            }
                            item = item.concat(tempService);
                            //await mainWindow.webContents.send(crawlCommand.log, 'service number ' + index + " add " + tempService.length + " item " + tempService);
                            //await mainWindow.webContents.send(crawlCommand.log, 'service number ' + index + " after added length " + item.length + " item " + item);
                            return item;
                        }));
                    }
                    if (khuyenmai) {
                        discountExcel1 = await Promise.all(discountExcel1.map(async (item, index) => {
                            //thêm discount thiếu
                            let tempDiscount = [...Array(0)];
                            if (item.length < defaultDiscountHeader1.length) {
                                for (var j = item.length; j < defaultDiscountHeader1.length; j++) {
                                    tempDiscount.push("");
                                }
                            }
                            item = item.concat(tempDiscount);
                            //await mainWindow.webContents.send(crawlCommand.log, 'discount header 1 number ' + index + " add " + tempDiscount);
                            return item;
                        }));
                    }
                    //await mainWindow.webContents.send(crawlCommand.log, 'discountExcel1 ' + JSON.stringify(discountExcel1));

                    //bỏ đi bảng dưới
                    // discountExcel2 = await Promise.all(discountExcel2.map(async (item, index) => {
                    //     //thêm discount thiếu
                    //     let tempDiscount = [...Array(0)];
                    //     if (item.length < defaultDiscountHeader2.length) {
                    //         for (var j = item.length; j < defaultDiscountHeader2.length; j++) {
                    //             tempDiscount.push("");
                    //         }
                    //     }
                    //     item = item.concat(tempDiscount);
                    //     //await mainWindow.webContents.send(crawlCommand.log, 'discount header 2 number ' + index + " add " + tempDiscount);
                    //     return item;
                    // }));

                    //await mainWindow.webContents.send(crawlCommand.log, 'discountExcel2 ' + JSON.stringify(discountExcel2));

                    //thêm lại header
                    createExcelMain();
                    currentIndexHeader = 0;
                    serviceHeader.forEach((item, index) => {
                        ws.column(currentIndexHeader).setWidth(25);
                        currentIndexHeader++;
                    });
                    discountHeader1.forEach((item, index) => {
                        ws.column(currentIndexHeader).setWidth(25);
                        currentIndexHeader++;
                    });
                    // discountHeader2.forEach((item, index) => {
                    //     ws.column(currentIndexHeader).setWidth(25);
                    //     currentIndexHeader++;
                    // });
                }
                //await mainWindow.webContents.send(crawlCommand.log, 'after adding ');
                //await mainWindow.webContents.send(crawlCommand.log, 'unitExcel ' + JSON.stringify(unitExcel));
                //await mainWindow.webContents.send(crawlCommand.log, 'discountExcel ' + JSON.stringify(discountExcel));
                //await mainWindow.webContents.send(crawlCommand.log, 'serviceExcel ' + JSON.stringify(serviceExcel));
                try {
                    if (canWrite) {
                        //write
                        //await mainWindow.webContents.send(crawlCommand.log, 'header length ' + defaultHeader.length);
                        //await mainWindow.webContents.send(crawlCommand.log, 'serviceHeader length ' + serviceHeader.length);
                        //await mainWindow.webContents.send(crawlCommand.log, 'discountHeader1 length ' + discountHeader1.length);
                        //await mainWindow.webContents.send(crawlCommand.log, 'discountHeader1  ' + discountHeader1);
                        //await mainWindow.webContents.send(crawlCommand.log, 'discountHeader2 ' + discountHeader2);
                        //lần chạy cuối cùng
                        //ghi header
                        try {
                            if (isRewrite == true) {
                                let indexCurrentHeader = 0;
                                //await mainWindow.webContents.send(crawlCommand.log, 'defaultHeader length ',defaultHeader);
                                if (thongtin) {
                                    //await mainWindow.webContents.send(crawlCommand.log, 'thong tin true ');
                                    if (defaultHeader.length == 2) {
                                        defaultHeader = defaultHeader.concat(nameHeader);
                                        //await mainWindow.webContents.send(crawlCommand.log, 'defaultHeader after concat ',defaultHeader);
                                    }
                                }
                                defaultHeader.forEach(async (item, index) => {
                                    indexCurrentHeader++;
                                    await writeToXcell(1, Number.parseInt(indexCurrentHeader), headeTitle + "-" + item);
                                    await mainWindow.webContents.send(crawlCommand.log, 'ghi vào header ' + (1) + " " + Number.parseInt(indexCurrentHeader) + " " + headeTitle + "-" + item);
                                });
                                if (data) {
                                    defaultServiceHeader.forEach(async (item, index) => {
                                        indexCurrentHeader++;
                                        //if (index > 1) {//bỏ qua so thứ tự và số điện thoại
                                        await writeToXcell(1, Number.parseInt(indexCurrentHeader), headeTitle + "-" + item);
                                        await mainWindow.webContents.send(crawlCommand.log, 'ghi vào header service ' + (1) + " " + Number.parseInt(indexCurrentHeader) + " " + headeTitle + "-" + item);
                                        //}
                                    });
                                }
                                if (khuyenmai) {
                                    defaultDiscountHeader1.forEach(async (item, index) => {
                                        indexCurrentHeader++;
                                        //if (index > 0) {//bỏ qua so dien thoai
                                        await writeToXcell(1, Number.parseInt(indexCurrentHeader), headeTitle + "-" + item);
                                        await mainWindow.webContents.send(crawlCommand.log, 'ghi vào header Discount' + (1) + " " + Number.parseInt(indexCurrentHeader) + " " + headeTitle + "-" + item);
                                        //}
                                    });
                                }
                                // defaultDiscountHeader2.forEach(async (item, index) => {
                                //     indexCurrentHeader++;
                                //     //if (index > 0) {//bỏ qua so dien thoai
                                //     await writeToXcell(1, Number.parseInt(indexCurrentHeader), headeTitle + "-" + item);
                                //     //await mainWindow.webContents.send(crawlCommand.log, 'ghi vào header ' + (1) + " " + Number.parseInt(indexCurrentHeader) + " " + headeTitle + "-" + item);
                                //     //}
                                // });
                            }
                        } catch (e) {
                            await mainWindow.webContents.send(crawlCommand.log, 'error when writing header ' + e);
                        }
                        //ghi content
                        let indexCurrent = 0;
                        //await mainWindow.webContents.send(crawlCommand.log, 'content');
                        for (let j = 0; j < unitExcel.length; j++) {
                            indexCurrent = 0;
                            //await mainWindow.webContents.send(crawlCommand.log, 'index ' + indexCurrent);
                            //await mainWindow.webContents.send(crawlCommand.log, 'unitExcel ' + unitExcel[j].length);
                            //await mainWindow.webContents.send(crawlCommand.log, 'serviceExcel ' + serviceExcel[j].length);
                            //await mainWindow.webContents.send(crawlCommand.log, 'discountExcel ' + discountExcel[j].length);
                            //await mainWindow.webContents.send(crawlCommand.log, 'write unit');
                            if (thongtin) {
                                unitExcel[j].forEach(async (item, index) => {
                                    indexCurrent++;//await writeToXcell(index + rowSpacing, 1, index + 1);
                                    //await mainWindow.webContents.send(crawlCommand.log, 'index inside unitExcel ' + indexCurrent);
                                    if (item == "&nbsp;") {
                                        await writeToXcell((j + rowSpacing), Number.parseInt(indexCurrent), noneName);
                                    } else {
                                        await writeToXcell((j + rowSpacing), Number.parseInt(indexCurrent), item);
                                    }
                                    //await mainWindow.webContents.send(crawlCommand.log, 'ghi vào excel unit ' + (j + rowSpacing) + " " + Number.parseInt(indexCurrent) + " " + item);
                                });
                            }
                            //await mainWindow.webContents.send(crawlCommand.log, 'write service '+serviceExcel[j].length);
                            if (data) {
                                serviceExcel[j].forEach(async (item, index) => {
                                    indexCurrent++;
                                    //if (index <= 1) {
                                    //await writeToXcell(j + rowSpacing, Number.parseInt(indexCurrent), "Dich vu");
                                    //} else {
                                    //await mainWindow.webContents.send(crawlCommand.log, 'index inside serviceExcel ' + indexCurrent);
                                    if (item == "&nbsp;") {
                                        await writeToXcell((j + rowSpacing), Number.parseInt(indexCurrent), noneName);
                                    } else {
                                        await writeToXcell((j + rowSpacing), Number.parseInt(indexCurrent), item);
                                    }
                                    //}
                                    //await mainWindow.webContents.send(crawlCommand.log, 'ghi vào excel service ' + (j + rowSpacing) + " " + Number.parseInt(indexCurrent) + " " + item);

                                });
                            }
                            //await mainWindow.webContents.send(crawlCommand.log, 'write discount 1 ');
                            if (khuyenmai) {
                                if (discountExcel1[j] != undefined) {
                                    discountExcel1[j].forEach(async (item, index) => {
                                        indexCurrent++;
                                        //if (index == 0) {
                                        // await writeToXcell(j + rowSpacing, Number.parseInt(indexCurrent), "Khuyen mai");
                                        //} else {
                                        //await mainWindow.webContents.send(crawlCommand.log, 'index inside discountExcel ' + indexCurrent);
                                        if (item == "&nbsp;") {
                                            await writeToXcell((j + rowSpacing), Number.parseInt(indexCurrent), noneName);
                                        } else {
                                            await writeToXcell((j + rowSpacing), Number.parseInt(indexCurrent), item);
                                        }
                                        //}
                                        //await mainWindow.webContents.send(crawlCommand.log, 'ghi vào excel discount ' + (j + rowSpacing) + " " + Number.parseInt(indexCurrent) + " " + item);
                                    });
                                }
                            }
                            //await mainWindow.webContents.send(crawlCommand.log, 'write discount 2 ');
                            // if (discountExcel2[j] != undefined) {
                            //     discountExcel2[j].forEach(async (item, index) => {
                            //         indexCurrent++;
                            //         //if (index == 0) {
                            //         // await writeToXcell(j + rowSpacing, Number.parseInt(indexCurrent), "Khuyen mai");
                            //         //} else {
                            //         //await mainWindow.webContents.send(crawlCommand.log, 'index inside discountExcel ' + indexCurrent);
                            //         if (item == "&nbsp;") {
                            //             await writeToXcell((j + rowSpacing), Number.parseInt(indexCurrent), noneName);
                            //         } else {
                            //             await writeToXcell((j + rowSpacing), Number.parseInt(indexCurrent), item);
                            //         }
                            //         //}
                            //         //await mainWindow.webContents.send(crawlCommand.log, 'ghi vào excel discount ' + (j + rowSpacing) + " " + Number.parseInt(indexCurrent) + " " + item);
                            //     });
                            // }
                        }

                    }
                    // else {
                    //     await mainWindow.webContents.send(crawlCommand.log, 'số  ' + inputPhoneNumberArray[index] + ' bị lỗi ');
                    //     await writeToXcell(index + rowSpacing, 1, errorTitle + "-" + inputPhoneNumberArray[index] + " bị lỗi, không tra cứu");

                    // }
                }
                catch (err) {
                    await mainWindow.webContents.send(crawlCommand.log, 'lỗi  ' + err);
                    //await writeToXcell(index + rowSpacing, 1, errorTitle + "-" + counterIndexNotFound);//số thứ tự
                    //currentData[0] = errorTitle + "-" + inputPhoneNumberArray[curentProcerssingIndex] + " bị lỗi, không tra cứu";
                    //currentData.push(errorTitle + "-" + inputPhoneNumberArray[curentProcerssingIndex] + " bị lỗi, không tra cứu");
                    let errorInffo = errorTitle + "-" + inputPhoneNumberArray[curentProcerssingIndex] + " bị lỗi không tra cứu" + err;
                    if (thongtin) {
                        unitExcel[curentProcerssingIndex] = [errorInffo];
                    }
                    if (data) {
                        serviceExcel[curentProcerssingIndex] = [errorInffo];
                    }

                    if (khuyenmai) {
                        discountExcel1[curentProcerssingIndex] = [errorInffo];
                    }
                    await writeToXcell(curentProcerssingIndex + rowSpacing, 1, errorInffo);
                }
            } catch (error) {
                await mainWindow.webContents.send(crawlCommand.log, 'lỗi  ' + error);
                //await writeToXcell(index + rowSpacing, 1, errorTitle + "-" + counterIndexNotFound);//số thứ tự
                //currentData[0] = errorTitle + "-" + inputPhoneNumberArray[curentProcerssingIndex] + " bị lỗi, không tra cứu";
                let errorInffo = errorTitle + "-" + inputPhoneNumberArray[curentProcerssingIndex] + " bị lỗi không tra cứu" + error;
                await writeToXcell(curentProcerssingIndex + rowSpacing, 1, errorInffo);
                if (thongtin) {
                    unitExcel[curentProcerssingIndex] = [errorInffo];
                }
                if (data) {
                    serviceExcel[curentProcerssingIndex] = [errorInffo];
                }

                if (khuyenmai) {
                    discountExcel1[curentProcerssingIndex] = [errorInffo];
                }
            }

        });

        await mainWindow.webContents.send(crawlCommand.log, 'end  ');
        await mainWindow.webContents.send(crawlCommand.log, 'write to excel  ');

        //await timer(2500);
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

//thongtin
ipcMain.on(crawlCommand.thongtin, async function (e, item) {
    thongtin = item;
});

//data 
ipcMain.on(crawlCommand.data, async function (e, item) {
    data = item;
});

//khuyenmai 
ipcMain.on(crawlCommand.khuyenmai, async function (e, item) {
    khuyenmai = item;
});