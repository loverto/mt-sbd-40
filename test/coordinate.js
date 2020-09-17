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
let y = 9;
let y1 = 8;
let y2 = 9;
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
    let tempy = 0;
    let randomW = Math.floor(Math.random() * (2)) + 2;
    console.warn("随机数是:"+randomW);
    for (let j =0,w = 0; j < mouseRow; j++) {
        positionArr = currentPosition.split(",");
        if(w<10&& randomW>=0){
            y = y1;
            randomW--;
        }else {
            y = y2;
        }
        if (w==9){
            w=0;
            randomW = Math.floor(Math.random() * (2)) + 2;
            console.warn("随机数是:"+randomW);
        }
        positionX = Number(positionArr[0]);
        positionY = Number(positionArr[1]) + (tempy);
        tempy+=y;
        w++;
        for (let i = 0; i < mouse; i++) {
            let temp = positionX + (i * x) + "," + (positionY);
            mousePosition.push(temp);
        }
    }
}


initPosition(mouseRow);
console.log(mousePosition)
