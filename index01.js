//引入npm包
const xlsx = require('node-xlsx');
const fs = require('fs');

//读取文件内容
const obj = xlsx.parse('./data/客里表.xls');

// 获取Excel表中第一个sheet数据
const excelObj = obj[0].data;

let resultData = [];

for (let i = 1; i < excelObj.length; i++) {
    let obj = {
        name:null,
        startStation:[],
        KM:[]
    };
    if(resultData.length > 0){
        for(let j = 0; j < resultData.length; j++){
            if(excelObj[i][4] == resultData[j].name){
                resultData[j].startStation.push(excelObj[i][8]);
                resultData[j].KM.push(excelObj[i][9]);
                continue
            }else if(excelObj[i][4] != resultData[j].name && j == resultData.length - 1){
                obj.name = excelObj[i][4];
                obj.startStation.push(excelObj[i][8]);
                obj.KM.push(excelObj[i][9]);
                resultData.push(obj);
            }
        }
    }else{
        // 首次进入循环，默认添加第一个数据
        obj.name = excelObj[i][4];
        obj.startStation.push(excelObj[i][8]);
        obj.KM.push(excelObj[i][9]);
        resultData.push(obj);
    }
}

// 此时resultData已经处理完毕，后面对数据进行重新处理
let resultTXT = '';
for(let i = 0 ; i < resultData.length; i++){
    let content = '线路名称：' + resultData[i].name + ";";
    content += ' 起点站：' + resultData[i].startStation[0] + ";";
    content += ' 终点站：' + resultData[i].startStation[resultData[i].startStation.length - 1] + ";";
    content += ' 起点里程：' + resultData[i].KM[0] + ";";
    content += ' 终点里程：' + resultData[i].KM[resultData[i].KM.length - 1] + ";" + '\n';
    console.log(content);
    resultTXT += content;
    // fs.writeFile('货运里程表.doc','测试',function(error){
    //     if(error){
    //         console.log(error);
    //         return false;
    //     }
    //     console.log('写入成功' + i);
    // })
}
fs.writeFile('客里表.txt',resultTXT,function(error){
        if(error){
            console.log(error);
            return false;
        }
        console.log('写入成功');
    })


