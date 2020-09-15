// 导入大漠插件版本
const path = require("path");

const dm = require('@loverto/dm.dll')
const sleep = require('./sleep');
const fs = require('fs')
const readline = require('readline');

const dmExt = require('./dm.dll.ext')
const coreldraw = require('./coreldraw')
const common = require('./common')
const _ = require('lodash')

const keycode = require('keycode')

const log4js = require('log4js');
log4js.configure({
    appenders: {
        out: { type: 'stdout' },
        stb13: { type: 'file', filename: process.env.storePath+path.sep+'stb-13.log' } },
    categories: { default: { appenders: ['out','stb13'], level: 'debug' } }
});

const logger = log4js.getLogger('stb');

const {DB} = require('../universal/database')

let keepTable = new Map()
let value = [1, 2]
let value1 = [1, 3]
let value2 = [4]

keepTable.set("1", value)
keepTable.set("2", value1)
keepTable.set("3", value2)

// 保存尺寸与位置的映射关系
let keepPicTable = new Map();

keepPicTable.set("15",1);
keepPicTable.set("14",2);
keepPicTable.set("13",3);
keepPicTable.set("12",4);

// 获取大漠插件的版本
logger.debug(dm.dll.ver())

let db = null;
// 蒙泰彩色电子出版系统 V6.0(普及版) - [40模板(1).tpf (0.7%)]
const windowTitle = "蒙泰彩色电子出版系统 V6.0(普及版)";
let devWidth = 1440
let devHeight = 900
let screenSize = dm.getScreenSize();
logger.debug(screenSize)
let winW = screenSize.width/devWidth;
let winH = screenSize.height/devHeight;

// dpi 界面坐标
let dpiUIPosition = [1021,629]

let dpiSwitchCoordinate = [871,529]

let dpiSwitchFlagCoordinate =  [869,548]
let dpiSwitchFlagColor =  "000000"
let dpiSwitchFlagSim =  0.8

// 需要修改的dpi的值
let dpiValue = 96

let fileSuff = ".cdr";

let modelFilePath = ""
let textFilePath = ""
let imageFilePath = ""
let exportModelFilePath = ""
let pch = ""

// 手提包 15面的 9个
let modelPosition = ["432,298","555,298","676,296","432,387","554,386","675,386","432,480","554,481","676,481"]
// 鼠标垫开始位置
let mouseInitPosition = "693,115";
let mouseInitSecondPosition = "693,335";
let mouseInitThirdPosition = "693,536";
// 鼠标垫位置数组
let mousePosition = [];
// 鼠标垫的列数
let mouse = 3;
// 鼠标垫的行数
let mouseRow = 65;

let x = 12
let y = 8;
// mousePosition.push(mouseInitPosition);

let positionArr;
let positionX;
let positionY;
let currentPosition = mouseInitPosition;

/**
 * 初始化坐标
 * @param mouseRow
 */
function initPosition(mouseRow) {
    for (let j = 0; j < mouseRow; j++) {
        positionArr = currentPosition.split(",");
        positionX = Number(positionArr[0]);
        positionY = Number(positionArr[1]) + (j * y);
        for (let i = 0; i < mouse; i++) {
            let temp = positionX + (i * x) + "," + (positionY);
            mousePosition.push(temp);
        }
    }
}

initPosition(mouseRow);
// currentPosition = mouseInitSecondPosition;
// initPosition(22);
// currentPosition = mouseInitThirdPosition;
// initPosition(22);

logger.info(JSON.stringify(mousePosition))


// 导出需要选择的坐标点
let exportCoordinate = ["408,272","583,435"]

// 手提包 13面的正常的坐标 4个
let modelNormal = ["530,314","530,340","530,369","530,395"]

// 旋转坐标点
let spinCoordinate = [366,96]

// 箭头坐标
let arrowCoordinate = [12,136]


// 导入图片坐标
let importImagePositionCoordinate = ["149,481","369,527","520,523","669,518","818,522","210,532"]


// 左上角点击坐标
let leftClickCoordinate = importImagePositionCoordinate[0].split(",")

// 空白位置坐标
let clickWhite = [207,328]


// 模板坐标坐标
let modelCoordinate = importImagePositionCoordinate


// 宽和高位置
let widthAndHeightPosition = ["180,88","180,108"]
// 宽高参数
let widthHeightParam = ["389,247","357,255","347,245","325,229.5"]
// 参考行坐标
let refRowCoordinate =["396,495","526,498","667,503","800,498"]
// 批量复制左上右下坐标
let batchCopyCoordinate = ["120,453", "407,519"]

// 替换坐标
// 替换文本查找坐标, 替换文本替换坐标, 全部替换坐标,替换完成，替换关闭
let replaceCoordinate = common.ratioConversion(["614,396","610,426","907,455","748,484","966,363"],winW,winH)
// 替换需要查找的文本
let findText = "编号位置";
// 批次号增量标记
let pchIncreateFlag = -1


let refYValueOne = 103 * (winH)
let refYValueTwo = 76 * (winH)
let refYValueThree = 133 * (winH)
let refYValueFour = 64 * (winH)

let diffOne = 0;
let diffTwo = 0;

/*
* 按行读取文件内容
* 返回：字符串数组
* 参数：fReadName:文件名路径
*      callback:回调函数
* */
function readFileToArr(fReadName,callback){
    const fRead = fs.createReadStream(fReadName);
    const objReadline = readline.createInterface({
        input: fRead
    });
    const arr = new Array();
    objReadline.on('line',function (line) {
        arr.push(line);
        //console.log('line:'+ line);
    });
    objReadline.on('close',function () {
        // console.log(arr);
        callback(arr);
    });
}


/**
 * 主方法
 */
function main(configObject) {
    logger.debug("sbd 40 is starting up ")
    if (!configObject){
        let storePath = process.env.storePath;
        db = new DB(storePath);
        logger.debug("from db config")
        configObject  = db.get("configObject");
        logger.debug("config value " + JSON.stringify(configObject))
    }


    initConfig(configObject);
    // 如果没有找到窗口，则退出
    if (!coreldraw.findCorelDrawAndFullScreen(windowTitle)){
        logger.debug("corel draw window is not find")
        return;
    };

    //activeInput(windowTitle,"US")
    // return;
    logger.debug("corel draw eas")
    coreldraw.eas();
    logger.debug("corel draw start open model")
    coreldraw.openUModel(modelFilePath)
    sleep.msleep(1000)
    // logger.debug("corel draw mouse is move arrow")
    // // 设置为可移动
    // coreldraw.moveAndClick(arrowCoordinate)
    if (fs.existsSync(textFilePath)){
        logger.debug("file is exists")
        //let readFileSync = fs.readFileSync(textFilePath);
        // 按行读取数据
        readFileToArr(textFilePath,function (dataa) {
            logger.debug("data length"+dataa.length)

            let i =0;
            let j = 0
            // 先屏蔽该逻辑
            if (db.has('mt') && false){
                let mt = db.get('mt');
                logger.debug("mt from db config"+JSON.stringify(mt));
                // 当前执行的批次数
                i = mt.currentBatch;
                // 当前执行的条数
                j = mt.currentRow;
                // 获取缓存的批次号
                pch = mt.pch;
                // 获取缓存的批次号量
                pchIncreateFlag = mt.pchIncreateFlag;
            }
            let dataArr = _.chunk(dataa,63*3);
            logger.debug("dataArr ：" + dataArr.length);
            // 遍历按行读取的数据
            for (i; i<dataArr.length; i++){
                //db.set('mt',{currentBatch: i,pch:pch,pchIncreateFlag:pchIncreateFlag})
                let data = dataArr[i]
                for (let j = 0;j<data.length;j++){

                    let picfilename = data[j];
                    let picPath = common.getFilePathByFileName(imageFilePath,picfilename);
                    // 图片路径和模板路径都存在
                    if (fs.existsSync(picPath) && fs.existsSync(modelFilePath)){
                        handler(picPath,null,false,modelCoordinate,picfilename,j)
                    }
                    // logger.debug("开始点击空白位置")
                    // sleep.msleep(500)
                    // coreldraw.moveAndClick(clickWhite)
                    // sleep.msleep(500)

                    // 执行完数组中的值，就保存够13张则保存图片
                    // if (j==data.length-1){
                    //     dm.keyPress(keycode('shift'));
                    //     logger.debug("张数够了，开始保存")
                    //     logger.debug("开始获取序列号"+pch+pchIncreateFlag)
                    //     // 获取序列号
                    //     let result = common.getSequenceNumber(pch,pchIncreateFlag);
                    //     logger.debug("获取序列号后的结果"+JSON.stringify(result))
                    //     pch = result.pch;
                    //     pchIncreateFlag = result.pchIncreateFlag
                    //     let mt = {currentBatch:i,currentRow: j,pch:pch,pchIncreateFlag:pchIncreateFlag};
                    //     logger.debug("开始存储序列号到数据库中"+JSON.stringify(mt))
                    //     db.set('mt',mt)
                    //     logger.debug("开始替换编号")
                    //     // 替换编号
                    //     coreldraw.findAndReplaceText(replaceCoordinate,findText,pch);
                    //
                    //     sleep.msleep(200);
                    //     logger.debug("开始保存文件")
                    //     let exportPathAbsout = exportModelFilePath + path.sep + pch + fileSuff;
                    //     logger.debug("开始另存为"+exportPathAbsout)
                    //     coreldraw.saveAsPath(exportPathAbsout);
                    //     sleep.msleep(3000)
                    //     logger.debug("保存完毕，开始关闭当前标签页")
                    //     coreldraw.closeModel();
                    //     coreldraw.closeModel();
                    //     coreldraw.closeModel();
                    //     sleep.msleep(500)
                    //     logger.debug("关闭完毕")
                    //     coreldraw.eas();
                    //     // 保存当前的序列号
                    //     if(db.has("configObject")){
                    //         configObject.pch = pch;
                    //         db.set("configObject",configObject)
                    //         logger.debug("把pch保存到数据库中"+JSON.stringify(configObject))
                    //     }
                    //
                    // }

                }


                // // 执行完重置该行数据
                // j = 0;
                // logger.debug("保存之后，判断是否需要打开新的模板")
                // if (i<=dataArr.length-1){
                //     logger.debug("执行完毕，开始保存，共执行"+i+"版");
                //     sleep.msleep(500)
                //     coreldraw.openUModel(modelFilePath);
                //     sleep.msleep(200)
                //     // 可移动坐标
                //     coreldraw.moveAndClick(arrowCoordinate)
                // }

            }

            // 最后执行完当前所有的图片后，编号自动更新一位，避免下次重命名
            // 获取序列号
            let result = common.getSequenceNumber(pch,pchIncreateFlag);
            pch = result.pch;
            pchIncreateFlag = result.pchIncreateFlag

            // 保存当前的序列号
            if(db.has("configObject")){
                configObject.pch = pch;
                db.set("configObject",configObject)
            }

            process.send({totalSize:dataa.length})

        })
    }

}

/**
 * 激活输入法
 * 该发放暂时不可用
 * @param windowTitle
 * @param input
 */
function activeInput(windowTitle,input) {
    const hwnd = dm.findWindow("", windowTitle);
    if (dmExt.checkInputMethod(hwnd, input) == 0) {
        dmExt.activeInputMethod(hwnd, input)
    }
}

/**
 * 导入模型并解锁
 * @param coreldrawHandlerFilePath
 */
function importModelAndUnLock (coreldrawHandlerFilePath) {
    coreldraw.importUModel(coreldrawHandlerFilePath)
    // sleep.msleep(500)
    // logger.debug('start position' + JSON.stringify(leftClickCoordinate))
    // dm.moveTo(leftClickCoordinate[0], leftClickCoordinate[1])
    // sleep.msleep(200)
    // dm.leftClick()
    // sleep.msleep(2000)
    // logger.debug('开始解除组合')
    // // 解锁
    // coreldraw.ctrlAndU()
    //
    // logger.debug('点击空白坐标')
    // // 点击空白坐标
    // sleep.msleep(500)
    // coreldraw.moveAndClick(clickWhite)
    // sleep.msleep(500)
    //
    // logger.debug('开始删除无关的图')
}

/**
 * 核心处理业务方法
 * @param coreldrawHandlerFilePath 文件路径
 * @param model 模型
 * @param flag 标志位
 * @param coordinateArray 坐标点数组
 * @param filename 文件名称
 * @param number 当前张数
 */
function handler(coreldrawHandlerFilePath,model,flag,coordinateArray,filename,number) {
    logger.debug("corelDrawHandlerFilePath:"+coreldrawHandlerFilePath+"model:"+model
    + "flag:" + flag+"coordinateArray:"+coordinateArray+"filename:"+filename+
        "number:"+number
    )
    // 点击设定的坐标点
    let mousePositionElement = mousePosition[number];
    coreldraw.moveAndClick(mousePositionElement.split(","));
    sleep.msleep(200)
    logger.debug("start import model")
    importModelAndUnLock(coreldrawHandlerFilePath)
    coreldraw.enter();
    // // 删除不相关的图
    // let moveCoordinate = [];
    // // 15 寸
    // let keepPic = 1;
    // moveCoordinate = coreldraw.deleteOtherObject(coordinateArray, keepPic);
    //
    // let endCoordinate = modelPosition[number].split(",");
    //
    // logger.debug("开始移动图片")
    // common.selectAreaByPointArray(moveCoordinate,endCoordinate);

}

/**
 * 初始化配置文件
 * @param configObject 配置对象
 */
function initConfig(configObject) {
    modelFilePath = configObject.modelFilePath;
    imageFilePath = configObject.imageFilePath;
    exportModelFilePath = configObject.exportModelFilePath;
    textFilePath = configObject.textFilePath;
    pch = configObject.pch;
}

exports.main = main


main();
