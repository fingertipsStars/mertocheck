
const express = require('express');
const bodyParser = require('body-parser');
const app = express();
var multer = require('multer');
var upload = multer();
var fs = require('fs');
const ExcelJS = require('exceljs');
const nodeXlxs = require("node-xlsx");
let setMonth = new Date().getMonth() + 1;
let currentMonth = new Date().getMonth() + 1;
let currentYear = new Date().getFullYear();
const daysInMonth = new Date(currentYear, setMonth, 0).getDate();//当前月份天数
console.log(daysInMonth);
let limitEnd = daysInMonth+5;
const port = 3000;
// 绩效的excel表头
let newExcelHeader = ['序号','工号','姓名','车站/队/组','日期','事件内容','考核条款','考核分值','检查人','总分','员工签字'];
// 绩效excel 标题
let newExcelTitle = '站务车间杨湾中心站'+ currentMonth +'月份月度绩效考核汇总';
// 考勤excel标题名称
let newExcelHoursTitle = `洛阳轨道交通集团有限责任公司运营分公司${currentYear}年${currentMonth}月考勤表`;
// 考勤的excel部分表头
let newExcelHoursBefore = ['序号','工号','日期    姓名','班组','岗位'];
// 考勤的excel部分表头
let newExcelHoursBack = [
  '实际出勤(h)','标准工时(h)',	'超缺工时(h)',	'法定节假日加班小时',	'倒班人员必填  夜班(个)',	'员工签名',	'休息',	
  '调休假',	'迟到',	'早退 平时',	'旷工'	,'培训'	,'出差'	,'丧假'	,'病假'	,'婚假',	'年休假',	'事假'	,'产假',	'孕检假'	,'工伤假',	'陪护假',
  	'节育假',	'哺乳时间',	'备注'];
// 夜班上班情况统计汇总表
let nightTitle = `${currentYear}年${currentMonth}月站务车间夜班上班情况统计汇总表`;
let nightExcelHeader = [ "序号",	"部门",	"工号",	"姓名",	"科室/车间",	"岗位",	`${currentMonth}月夜班次数`	,`${currentMonth}月夜班上班时间`];
let maxMonth = [1,3,5,7,8,10,12];
let formalMonth = [4,6,9,11];

// 设置静态文件目录
app.use(express.static('public'));

// 使用body-parser中间件解析请求体
app.use(bodyParser.urlencoded({ extended: false }));
app.use(bodyParser.json());
// 报错处理
app.use((err, req, res, next) => {
  // 错误处理逻辑
  res.status(500).send('Something went wrong.');
});

// 为app添加中间件处理跨域请求
app.use(function(req,res,next){
    res.header('Access-Control-Allow-Origin','*');
    res.header('Access-Control-Allow-Methods','PUT, GET, POST, DELETE, OPTIONS');
    res.header('Access-Control-Allow-Headers','X-Requested-With');
    res.header('Access-Control-Allow-Headers','Content-Type');
    next();
})

  // 定义路由
app.get('/', (req, res) => {
  // 渲染主页模板并发送给客户端
  res.render('index');
});
  
// POST接口(绩效)
app.post('/api/upload/check', upload.any(),(req, res) => {
  // console.log(req.files)
  let sheets = [];
  let sheetData = req.files;
  if(sheetData.length && sheetData.length <=1){
    sheets = nodeXlxs.parse(sheetData[0].buffer);
  }else{
    for(let i = 0;i < sheetData.length;i++){
      sheets.push(nodeXlxs.parse(sheetData[i].buffer))
    }
  }
  
  // 处理接收到的数据
  // console.log(sheets);
  let todoData = [];
  let totalData = [];
  let qmArr = [];
  let twArr = [];
  let sjwArr = [];
  let ywArr = [];
  sheets.forEach(function(sheet,index){
    let singleSheet = sheet[0];
      if(singleSheet['name'] =='检查通报' || singleSheet['name'] == 'Sheet1'){
        for(var rowId in singleSheet['data']){
          // 每一行数据
          var row = singleSheet['data'][rowId];
          if(row.length == 12 && row[1] != "工号" && row[7] != "/" && row[6] != "/"){
            // console.log(row)
            let newRow = row.splice(0,row.length-3);
            todoData.push(newRow);
          }
        }
      }
  })

  sortStation(todoData, 3);
  sortName(qmArr, 1);
  sortName(twArr, 1);
  sortName(sjwArr, 1);
  sortName(ywArr, 1);
  totalData = qmArr.concat(twArr,sjwArr,ywArr);
  // console.log(totalData)
  // 排序 先按车站排序，再按姓名排序
  // 排序 按姓名排序
  function sortName(arr, col) {
    for (var i = 0; i < arr.length - 1; i++) {
      for (var j = 0; j < arr.length - 1 - i; j++){
        if (arr[j][col] > arr[j + 1][col]) {
          var temp = arr[j];
          arr[j] = arr[j + 1];
          arr[j + 1] = temp;
        }
      }
    }
    return arr;
  }
 // 排序 先按车站排序
  function sortStation(arr, col) {
    for (let i = 0; i < arr.length; i++) {
      if (arr[i][col] == "启明南路站") {
        qmArr.push(arr[i])
      }else if(arr[i][col] == "塔湾站"){
        twArr.push(arr[i])
      }else if(arr[i][col] == "史家湾站"){
        sjwArr.push(arr[i])
      }else{
        ywArr.push(arr[i])
      }
    }
  }
// 创建一个新的工作簿
const workbook = new ExcelJS.Workbook();
// 添加一个工作表  绩效表
const worksheet = workbook.addWorksheet('Sheet1');
// 设置表头
worksheet.addRow([newExcelTitle]);
worksheet.addRow(newExcelHeader);
// 添加数据到工作表
totalData.forEach(function(rowData) {
  const excelSerialDateToJSDate = (serialDate) => {
    const utcDays = Math.floor(serialDate - 25569);
    const utcValue = utcDays * 86400;
    const dateInfo = new Date(utcValue * 1000);
    const year = dateInfo.getFullYear();
    const month = dateInfo.getMonth() + 1;
    const day = dateInfo.getDate();
    return `${month}月${day}日`;
  };
  let str =  "" + rowData[4];
  str.indexOf("月") != -1 ? rowData[4]= rowData[4] : rowData[4] = excelSerialDateToJSDate(rowData[4]) ;
  worksheet.addRow(rowData);
});

// 设置合并单元格
worksheet.mergeCells('A1:K1');//表头合并
 
  // 循环遍历数组处理（合并单元格）
  let mergeStartRow = 3;
  let mergeEndRow = 3;
  let score = 100;//初始分数
  let index = 1;//初始序号
  changeInfo(totalData);
  function changeInfo(arr){
    var mark = 0;
    let s3 = 0;
    for (let i = 0; i < arr.length - 1; i++) {
        if (arr[i][1] == arr[i + 1][1]) {
          mergeEndRow++;
        }else{
          index++;
          if (mergeStartRow !== mergeEndRow) {
            worksheet.mergeCells(`A${mergeStartRow}:A${mergeEndRow}`); // 合并单元格
            worksheet.mergeCells(`B${mergeStartRow}:B${mergeEndRow}`); // 合并单元格
            worksheet.mergeCells(`C${mergeStartRow}:C${mergeEndRow}`); // 合并单元格
            worksheet.mergeCells(`D${mergeStartRow}:D${mergeEndRow}`); // 合并单元格
            worksheet.mergeCells(`J${mergeStartRow}:J${mergeEndRow}`); // 合并单元格
            // let s1 = worksheet.getCell(`H${mergeStartRow}`);
            // let s2 = worksheet.getCell(`H${mergeEndRow}`);
            // mark = Number(s2)+ Number(s1);
          }else{
            // mark = Number(arr[i][7]);
          }
          
          mergeStartRow = mergeEndRow + 1;
          mergeEndRow++;
        }
          // 获取合并单元格的主单元格
          const mainCellOrder = worksheet.getCell(`A${mergeStartRow}`);
          // 修改主单元格的内容
          mainCellOrder.value = index;
          if (mergeStartRow !== mergeEndRow) {
            let s1 = worksheet.getCell(`H${mergeStartRow}`).value;
            let s2 = worksheet.getCell(`H${mergeEndRow}`).value;
            s3 = 0;
            if(mergeEndRow - mergeStartRow > 1){
              for(let j = mergeStartRow +1;j < mergeEndRow;j++){
                s3 += Number(worksheet.getCell(`H${j}`).value);
              }
            }
           
            mark = Number(s2) + Number(s1) + Number(s3);
          }else{
            let s1 = worksheet.getCell(`H${mergeStartRow}`);
            mark = Number(s1);
          }
         
          let mainCellScore = worksheet.getCell(`J${mergeStartRow}`);
            mainCellScore.value = score + mark;
    }
    if (arr[arr.length-1][1] == arr[arr.length - 2][1]) {
      worksheet.mergeCells(`A${mergeStartRow}:A${mergeEndRow}`); // 合并单元格
      worksheet.mergeCells(`B${mergeStartRow}:B${mergeEndRow}`); // 合并单元格
      worksheet.mergeCells(`C${mergeStartRow}:C${mergeEndRow}`); // 合并单元格
      worksheet.mergeCells(`D${mergeStartRow}:D${mergeEndRow}`); // 合并单元格
      worksheet.mergeCells(`J${mergeStartRow}:J${mergeEndRow}`); // 合并单元格
    }
    
  }
// 设置单元格样式
const mergedCell = worksheet.getCell('A1');
mergedCell.fill = {
  type: 'pattern',
  pattern: 'solid',
  // fgColor: { argb: 'FFFF0000' } // 设置背景颜色为红色
};
 // 设置所有单元格的对齐方式为居中并自动换行
 worksheet.eachRow(function(row) {
  row.eachCell(function(cell) {
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  });
});
 // 设置所有行的高度
 worksheet.eachRow(function(row) {
  row.height = null; // 设置行的高度为自动
});

// 设置列宽度
worksheet.getColumn(1).width = 10; // 设置第1列的宽度为15
worksheet.getColumn(2).width = 10; 
worksheet.getColumn(3).width = 10; 
worksheet.getColumn(4).width = 16; 
worksheet.getColumn(5).width = 12; 
worksheet.getColumn(6).width = 55; 
worksheet.getColumn(7).width = 65; 
worksheet.getColumn(8).width = 15; 
worksheet.getColumn(9).width = 10; 
worksheet.getColumn(10).width = 10; 
worksheet.getColumn(11).width = 25; 

worksheet.getCell('A1').font = { name: 'Arial',bold: true, size: 20 }; // 设置加粗和字体大小为12
 // 设置某一行中多个单元格的字体样式
 const rowIndex = 2; // 设置要修改的行的索引
 const font = {
   name: 'Arial',
   size: 11,
   bold: true,
  //  color: { argb: 'FF0000FF' }
 };
 const row = worksheet.getRow(rowIndex);
 row.eachCell(function(cell) {
   cell.font = font;
 });
  // 设置所有数据行的边框为黑色
  worksheet.eachRow(function(row, rowNumber) {
    if (rowNumber > 1) { // 跳过标题行
      row.border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } }
      };
    }
  });
// 保存工作簿为Excel文件
workbook.xlsx.writeFile(`杨湾中心站${currentMonth}月绩效.xlsx`)
  .then(() => {
    console.log('绩效Excel文件已生成');
    res.status(200).send({ message: 'upload successful' });
  })
  .catch((error) => {
    console.log('生成绩效Excel文件时出错：', error);
    res.status(404).send({ message: 'upload check error' });
  });
  
});
// 考勤接口
app.post('/api/upload/hours', upload.any(),(req, res) => {
  // console.log(req.files)
  let sheets = [];
  let sheetData = req.files;
  if(sheetData.length && sheetData.length <=1){
    sheets = nodeXlxs.parse(sheetData[0].buffer);
  }else{
    for(let i = 0;i < sheetData.length;i++){
      sheets.push(nodeXlxs.parse(sheetData[i].buffer))
    }
  }
  
  // 处理接收到的数据
  // console.log(sheets);
  let todoData = [];
  let newtodoData = [];
  // sheets.forEach(function(sheet,index){
    let singleSheet = sheets[0];
      if( singleSheet['name'] == 'Sheet1'){
        for(let rowId in singleSheet['data']){
          // 每一行数据
          let newRow = singleSheet['data'][rowId];
          if(newRow.length > 7){
            // console.log(newRow)
            todoData.push(newRow);
          }
        }
      }
  // })
// 创建一个新的工作簿
const workbook = new ExcelJS.Workbook();
// 添加一个工作表   考勤表
const worksheet = workbook.addWorksheet('Sheet1');
 // 添加一个工作表   夜班统计表
 const nightworksheet = workbook.addWorksheet('Sheet2');

// 夜班统计表设置表头
nightworksheet.addRow([nightTitle]);
nightworksheet.addRow(nightExcelHeader);

// 考勤表设置表头
let newExcelHoursMid=[];

// 添加数据到工作表
todoData.shift()
todoData.forEach(function(rowData,index) {
  if(index == 0){
    newExcelHoursMid = rowData.splice(5,rowData.length-1);
  }else{
    // 换位置
    let temp = rowData[1];
    rowData[1] = rowData[3];
    rowData[3] = temp;
    let temp1 = rowData[2];
    rowData[2] = rowData[4];
    rowData[4] = temp1;
  }
});
for(let q = 0;q < newExcelHoursMid.length;q++){
  newExcelHoursMid[q] = newExcelHoursMid[q]+"";
  newExcelHoursMid[q] = newExcelHoursMid[q].replace("号", ""); //将ele中的所有号去掉
}

let tempHeader = newExcelHoursBefore.concat(newExcelHoursMid).concat(newExcelHoursBack)
worksheet.addRow([newExcelHoursTitle]);
worksheet.addRow(tempHeader);
// 先构造数组，在改变数组对应的白夜年为数字
for (let k = 1; k < todoData.length; k++) {
  if(k == 0){
    newtodoData.push(todoData[k].splice(5,todoData[k].length-1));
  }else{
    newtodoData.push(todoData[k]);
    newtodoData.push(todoData[k]);
  }
}
let nightNum = 0;//夜班个数
let yearNum = 0;//年假个数
let marriageNum = 0;//婚假个数
let illNum = 0;//病假个数
let nightArr = [];
for (let n = 0; n < newtodoData.length; n++) {
  nightNum = 0;
  yearNum = 0;
  marriageNum = 0;
  illNum = 0;
  if(newtodoData[n].length){
    let nightObj = {};
    nightObj.yearRange = [];
    nightObj.nightRange = [];
    if (n % 2 == 0) {
      for(let m = 0; m < newtodoData[n].length; m++){
        if((newtodoData[n][m]+"").indexOf("A2") != -1 || (newtodoData[n][m]+"").indexOf("B2") != -1 || (newtodoData[n][m]+"").indexOf("C2") != -1 || (newtodoData[n][m]+"").indexOf("D2") != -1 || (newtodoData[n][m]+"").indexOf("E2") != -1 ){
          nightNum++;
          nightObj.nightRange.push(m - 4);
        }
        if((newtodoData[n][m]+"").indexOf("年") != -1 ){
          yearNum++;
          nightObj.yearRange.push(m - 4);
        }

        if((newtodoData[n][m]+"").indexOf("婚") != -1 ){
          marriageNum++;
        }
        if((newtodoData[n][m]+"").indexOf("病") != -1 ){
          illNum++;
        }
        newtodoData[n][m] = (newtodoData[n][m]+"").indexOf("A1") != -1 || (newtodoData[n][m]+"").indexOf("B1") != -1 || (newtodoData[n][m]+"").indexOf("C1") != -1 || (newtodoData[n][m]+"").indexOf("D1") != -1 || (newtodoData[n][m]+"").indexOf("E1") != -1 || (newtodoData[n][m]+"").indexOf("F") != -1 ? "白" : (newtodoData[n][m]+"").indexOf("A2") != -1 || (newtodoData[n][m]+"").indexOf("B2") != -1 || (newtodoData[n][m]+"").indexOf("C2") != -1 || (newtodoData[n][m]+"").indexOf("D2") != -1 || (newtodoData[n][m]+"").indexOf("E2") != -1 ? "夜" : newtodoData[n][m] == "Z" ? "日" : newtodoData[n][m] == "调" ? "":newtodoData[n][m];
      }
      nightObj.id = newtodoData[n][1];
      nightObj.night = nightNum;
      nightObj.year = yearNum;
      nightObj.name = newtodoData[n][2];
      nightObj.job = newtodoData[n][4];
      nightObj.marry = marriageNum;
      nightObj.ill = illNum;
      nightArr.push(nightObj);
    }
    
    if(n % 2 != 0){
      for(let m = 0; m < newtodoData[n].length; m++){
        newtodoData[n][m] = newtodoData[n][m] == "白" || newtodoData[n][m] == "夜"  ? 12 : newtodoData[n][m] == "日" ? 8 : (newtodoData[n][m]+"").indexOf("年") != -1 ? 8 : (newtodoData[n][m]+"").indexOf("产") != -1 || (newtodoData[n][m]+"").indexOf("婚") != -1 || (newtodoData[n][m]+"").indexOf("育") != -1 || (newtodoData[n][m]+"").indexOf("孕") != -1 ? 8 : (newtodoData[n][m]+"").indexOf("病") != -1 || (newtodoData[n][m]+"").indexOf("事") != -1 ? "" : newtodoData[n][m];

      }
    }
    worksheet.addRow(newtodoData[n]);
  }
}
const isMaxMonth = maxMonth.indexOf(setMonth);
// 设置合并单元格
worksheet.mergeCells('A1:BH1');//表头合并
  // 循环遍历数组处理（合并单元格）
  let mergeStartRow = 3;
  let mergeEndRow = 3;
  let index = 1;//初始序号
  changeInfo(newtodoData);
  function changeInfo(arr){
    for (let i = 0; i < arr.length - 1; i++) {
        if (arr[i][1] == arr[i + 1][1]) {
          mergeEndRow++;
        }else{
          index++;
          if (mergeStartRow !== mergeEndRow) {
            myMerge(mergeStartRow,mergeEndRow);
          }
          mergeStartRow = mergeEndRow + 1;
          mergeEndRow++;
        }
         // 获取合并单元格的主单元格
         const mainCellOrder = worksheet.getCell(`A${mergeStartRow}`);
         // 修改主单元格的内容
         mainCellOrder.value = index;
    }
    if (arr[arr.length - 2][1] == arr[arr.length - 1][1]) {
      myMerge(mergeStartRow,mergeEndRow);
    }
    
  }

  // merge单元格
  function myMerge(start,end){
    worksheet.mergeCells(`A${start}:A${end}`); // 合并单元格
    worksheet.mergeCells(`B${start}:B${end}`); // 合并单元格
    worksheet.mergeCells(`C${start}:C${end}`); // 合并单元格
    worksheet.mergeCells(`D${start}:D${end}`); 
    worksheet.mergeCells(`E${start}:E${end}`); 
     // 先判断是否2月份，再判断大月和小月
     if(currentMonth == 2 ){
      if(getFebruaryDays(currentYear) == 29){
        // 闰年  29天
        worksheet.mergeCells(`AI${start}:AI${end}`); 
        worksheet.mergeCells(`AJ${start}:AJ${end}`); 
        worksheet.mergeCells(`AK${start}:AK${end}`);
        worksheet.mergeCells(`AL${start}:AL${end}`); 
        worksheet.mergeCells(`AM${start}:AM${end}`); 
        worksheet.mergeCells(`AN${start}:AN${end}`); 
        worksheet.mergeCells(`AO${start}:AO${end}`); 
        worksheet.mergeCells(`AP${start}:AP${end}`); 
        worksheet.mergeCells(`AQ${start}:AQ${end}`); 
        worksheet.mergeCells(`AR${start}:AR${end}`); 
        worksheet.mergeCells(`AS${start}:AS${end}`); 
        worksheet.mergeCells(`AT${start}:AT${end}`); 
        worksheet.mergeCells(`AU${start}:AU${end}`); 
        worksheet.mergeCells(`AV${start}:AV${end}`); 
        worksheet.mergeCells(`AW${start}:AW${end}`); 
        worksheet.mergeCells(`AX${start}:AX${end}`); 
        worksheet.mergeCells(`AY${start}:AY${end}`); 
        worksheet.mergeCells(`AZ${start}:AZ${end}`); 
        worksheet.mergeCells(`BA${start}:BA${end}`); 
        worksheet.mergeCells(`BB${start}:BB${end}`); 
        worksheet.mergeCells(`BC${start}:BC${end}`); 
        worksheet.mergeCells(`BD${start}:BD${end}`); 
        worksheet.mergeCells(`BE${start}:BE${end}`); 
        worksheet.mergeCells(`BF${start}:BF${end}`); 
        worksheet.mergeCells(`BG${start}:BG${end}`); 
      }else{
        // 平年 28天
        worksheet.mergeCells(`AH${start}:AH${end}`);
        worksheet.mergeCells(`AI${start}:AI${end}`); 
        worksheet.mergeCells(`AJ${start}:AJ${end}`); 
        worksheet.mergeCells(`AK${start}:AK${end}`);
        worksheet.mergeCells(`AL${start}:AL${end}`); 
        worksheet.mergeCells(`AM${start}:AM${end}`); 
        worksheet.mergeCells(`AN${start}:AN${end}`); 
        worksheet.mergeCells(`AO${start}:AO${end}`); 
        worksheet.mergeCells(`AP${start}:AP${end}`); 
        worksheet.mergeCells(`AQ${start}:AQ${end}`); 
        worksheet.mergeCells(`AR${start}:AR${end}`); 
        worksheet.mergeCells(`AS${start}:AS${end}`); 
        worksheet.mergeCells(`AT${start}:AT${end}`); 
        worksheet.mergeCells(`AU${start}:AU${end}`); 
        worksheet.mergeCells(`AV${start}:AV${end}`); 
        worksheet.mergeCells(`AW${start}:AW${end}`); 
        worksheet.mergeCells(`AX${start}:AX${end}`); 
        worksheet.mergeCells(`AY${start}:AY${end}`); 
        worksheet.mergeCells(`AZ${start}:AZ${end}`); 
        worksheet.mergeCells(`BA${start}:BA${end}`); 
        worksheet.mergeCells(`BB${start}:BB${end}`); 
        worksheet.mergeCells(`BC${start}:BC${end}`); 
        worksheet.mergeCells(`BD${start}:BD${end}`); 
        worksheet.mergeCells(`BE${start}:BE${end}`); 
        worksheet.mergeCells(`BF${start}:BF${end}`); 
      }
    }else{
      if(isMaxMonth != -1){
        worksheet.mergeCells(`AK${start}:AK${end}`); 
        worksheet.mergeCells(`AL${start}:AL${end}`); 
        worksheet.mergeCells(`AM${start}:AM${end}`); 
        worksheet.mergeCells(`AN${start}:AN${end}`); 
        worksheet.mergeCells(`AO${start}:AO${end}`); 
        worksheet.mergeCells(`AP${start}:AP${end}`); 
        worksheet.mergeCells(`AQ${start}:AQ${end}`); 
        worksheet.mergeCells(`AR${start}:AR${end}`); 
        worksheet.mergeCells(`AS${start}:AS${end}`); 
        worksheet.mergeCells(`AT${start}:AT${end}`); 
        worksheet.mergeCells(`AU${start}:AU${end}`); 
        worksheet.mergeCells(`AV${start}:AV${end}`); 
        worksheet.mergeCells(`AW${start}:AW${end}`); 
        worksheet.mergeCells(`AX${start}:AX${end}`); 
        worksheet.mergeCells(`AY${start}:AY${end}`); 
        worksheet.mergeCells(`AZ${start}:AZ${end}`); 
        worksheet.mergeCells(`BA${start}:BA${end}`); 
        worksheet.mergeCells(`BB${start}:BB${end}`); 
        worksheet.mergeCells(`BC${start}:BC${end}`); 
        worksheet.mergeCells(`BD${start}:BD${end}`); 
        worksheet.mergeCells(`BE${start}:BE${end}`); 
        worksheet.mergeCells(`BF${start}:BF${end}`); 
        worksheet.mergeCells(`BG${start}:BG${end}`); 
        worksheet.mergeCells(`BH${start}:BH${end}`); 
        worksheet.mergeCells(`BI${start}:BI${end}`); 
    }else{
        worksheet.mergeCells(`AJ${start}:AJ${end}`); 
        worksheet.mergeCells(`AK${start}:AK${end}`); 
        worksheet.mergeCells(`AL${start}:AL${end}`); 
        worksheet.mergeCells(`AM${start}:AM${end}`); 
        worksheet.mergeCells(`AN${start}:AN${end}`); 
        worksheet.mergeCells(`AO${start}:AO${end}`); 
        worksheet.mergeCells(`AP${start}:AP${end}`); 
        worksheet.mergeCells(`AQ${start}:AQ${end}`); 
        worksheet.mergeCells(`AR${start}:AR${end}`); 
        worksheet.mergeCells(`AS${start}:AS${end}`); 
        worksheet.mergeCells(`AT${start}:AT${end}`); 
        worksheet.mergeCells(`AU${start}:AU${end}`); 
        worksheet.mergeCells(`AV${start}:AV${end}`); 
        worksheet.mergeCells(`AW${start}:AW${end}`); 
        worksheet.mergeCells(`AX${start}:AX${end}`); 
        worksheet.mergeCells(`AY${start}:AY${end}`); 
        worksheet.mergeCells(`AZ${start}:AZ${end}`); 
        worksheet.mergeCells(`BA${start}:BA${end}`); 
        worksheet.mergeCells(`BB${start}:BB${end}`); 
        worksheet.mergeCells(`BC${start}:BC${end}`); 
        worksheet.mergeCells(`BD${start}:BD${end}`); 
        worksheet.mergeCells(`BE${start}:BE${end}`); 
        worksheet.mergeCells(`BF${start}:BF${end}`); 
        worksheet.mergeCells(`BG${start}:BG${end}`); 
        worksheet.mergeCells(`BH${start}:BH${end}`); 
        worksheet.mergeCells(`BI${start}:BI${end}`); 
     }
   }
  }

  

// 根据单元格内容设置单元格样式
let reality = 0;//实际出勤
let realNight;
let realYear;
let realMarry;
let realIll;
let realCellOrder;// 获取实际出勤的单元格
let standardCellOrder;//  获取标准工时单元格
let oversizeCellOrder;//  获取超缺工时单元格
let nightCellOrder;//  获取夜班单元格
let yearCellOrder;//  获取年休假单元格
let marryCellOrder;//  获取婚假单元格
let illCellOrder;//  获取病假单元格
let paramOrder;
for (let n = 0; n < newtodoData.length; n++) {
  paramOrder = n + 3;
  if(currentMonth == 2){
    if(getFebruaryDays(currentYear) == 29){
      // 闰年  29天
      nightCellOrder = worksheet.getCell(`AM${paramOrder}`);
      yearCellOrder = worksheet.getCell(`AY${paramOrder}`);
      marryCellOrder =worksheet.getCell(`AX${paramOrder}`);
      illCellOrder = worksheet.getCell(`AW${paramOrder}`);
    }else{
      // 平年  28天
      nightCellOrder = worksheet.getCell(`AL${paramOrder}`);
      yearCellOrder = worksheet.getCell(`AX${paramOrder}`);
      marryCellOrder =worksheet.getCell(`AW${paramOrder}`);
      illCellOrder = worksheet.getCell(`AV${paramOrder}`);
    }
  }else{
    if(isMaxMonth != -1){
      nightCellOrder = worksheet.getCell(`AO${paramOrder}`);
      yearCellOrder = worksheet.getCell(`BA${paramOrder}`);
      marryCellOrder = worksheet.getCell(`AZ${paramOrder}`);
      illCellOrder = worksheet.getCell(`AY${paramOrder}`);
   }else{
      nightCellOrder = worksheet.getCell(`AN${paramOrder}`);
      yearCellOrder = worksheet.getCell(`AZ${paramOrder}`);
      marryCellOrder =worksheet.getCell(`AY${paramOrder}`);
      illCellOrder = worksheet.getCell(`AX${paramOrder}`);
   }
  }
  for(let i = 0;i < nightArr.length;i++){
    if(newtodoData[n][1] == nightArr[i].id){
      // 循环体代码
      realNight = nightArr[i].night;
      realYear = nightArr[i].year;
      realMarry = nightArr[i].marry;
      realIll = nightArr[i].ill;
    }
     // 修改夜班
     nightCellOrder.value = realNight;
     // 修改年休假
     yearCellOrder.value = realYear == 0 ? "" : realYear;
     // 修改婚假
     marryCellOrder.value = realMarry == 0 ? "" : realMarry;
     // 修改病假
     illCellOrder.value = realIll == 0 ? "" : realIll;
  }
}

worksheet.eachRow(function(row, rowNumber) {
  row.height = null; // 设置行的高度为自动
  if(currentMonth == 2){
    if(getFebruaryDays(currentYear) == 29){
      // 闰年  29天
      realCellOrder = worksheet.getCell(`AI${rowNumber}`);
      standardCellOrder =  worksheet.getCell(`AJ${rowNumber}`);
      oversizeCellOrder = worksheet.getCell(`AK${rowNumber}`);
      nightCellOrder = worksheet.getCell(`AM${rowNumber}`);
      yearCellOrder = worksheet.getCell(`AY${rowNumber}`);
      marryCellOrder =worksheet.getCell(`AX${rowNumber}`);
      illCellOrder = worksheet.getCell(`AW${paramOrder}`);
    }else{
      // 平年  28天
      realCellOrder = worksheet.getCell(`AH${rowNumber}`);
      standardCellOrder =  worksheet.getCell(`AI${rowNumber}`);
      oversizeCellOrder = worksheet.getCell(`AJ${rowNumber}`);
      nightCellOrder = worksheet.getCell(`AL${rowNumber}`);
      yearCellOrder = worksheet.getCell(`AX${rowNumber}`);
      marryCellOrder =worksheet.getCell(`AW${rowNumber}`);
      illCellOrder = worksheet.getCell(`AV${paramOrder}`);
    }
  }else{
    if(isMaxMonth != -1){
      realCellOrder = worksheet.getCell(`AK${rowNumber}`);
      standardCellOrder = worksheet.getCell(`AL${rowNumber}`);
      oversizeCellOrder = worksheet.getCell(`AM${rowNumber}`);
      nightCellOrder = worksheet.getCell(`AO${rowNumber}`);
      yearCellOrder = worksheet.getCell(`BA${rowNumber}`);
      marryCellOrder = worksheet.getCell(`AZ${rowNumber}`);
      illCellOrder = worksheet.getCell(`AY${paramOrder}`);
   }else{
      realCellOrder = worksheet.getCell(`AJ${rowNumber}`);
      standardCellOrder =  worksheet.getCell(`AK${rowNumber}`);
      oversizeCellOrder = worksheet.getCell(`AL${rowNumber}`);
      nightCellOrder = worksheet.getCell(`AN${rowNumber}`);
      yearCellOrder = worksheet.getCell(`AZ${rowNumber}`);
      marryCellOrder =worksheet.getCell(`AY${rowNumber}`);
      illCellOrder = worksheet.getCell(`AX${paramOrder}`);
   }
  }
  reality = 0;
  row.eachCell(function(cell, colNumber) {
    // 设置所有单元格的对齐方式为居中并自动换行
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    // 判断单元格内容是否符合条件
    if (((cell.value+"").indexOf('年') != -1 || (cell.value+"").indexOf('病') != -1 || (cell.value+"").indexOf('事') != -1) && rowNumber > 2 && colNumber > 5) {
      // 修改单元格背景色
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {argb: 'FF00B04E'} // 使用绿色作为背景色
      };
    }
    if (((cell.value+"").indexOf('产') != -1 || (cell.value+"").indexOf('婚') != -1 || (cell.value+"").indexOf('育') != -1 || (cell.value+"").indexOf('孕') != -1) && rowNumber > 2 && colNumber > 5) {
      // 修改单元格背景色
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {argb: 'FF00B04E'} // 使用绿色作为背景色
      };
    }
    // 计算实际出勤
    if(colNumber > 5 && rowNumber > 2 && colNumber <= limitEnd && (cell.value === 8 || cell.value === 12 || cell.value === 4)){
      reality+= Number(cell.value);
    }
  });
  // console.log(reality)

    // 从第三行开始赋值，前两行为标题和表头
    if(rowNumber > 3){
      // 修改实际出勤单元格的内容
      realCellOrder.value = reality;
      
     if(maxMonth.indexOf(currentMonth) != -1){
       standardCellOrder.value = 186;
      } else if(formalMonth.indexOf(currentMonth) != -1){
       standardCellOrder.value = 180;
      }else{
       // 2月份的标准工时
       standardCellOrder.value = getFebruaryDays(currentYear) * 6;
      }
      //  超缺工时
    oversizeCellOrder.value = reality - standardCellOrder.value;
    }
    
});


// 设置单元格样式
const mergedCell = worksheet.getCell('A1');
mergedCell.fill = {
  type: 'pattern',
  pattern: 'solid'
};
 

// 设置列宽度
worksheet.getColumn(1).width = 5; // 设置第1列的宽度为5
worksheet.getColumn(2).width = 10; 
worksheet.getColumn(3).width = 10; 
worksheet.getColumn(4).width = 16; 

worksheet.getCell('A1').font = { name: 'Arial',bold: true, size: 20 }; // 设置加粗和字体大小为12
 // 设置某一行中多个单元格的字体样式
 const rowIndex = 2; // 设置要修改的行的索引
 const font = {
   name: 'Arial',
   size: 11,
   bold: true,
  //  color: { argb: 'FF0000FF' }
 };
 const row = worksheet.getRow(rowIndex);
 row.eachCell(function(cell) {
   cell.font = font;
 });
  // 设置所有数据行的边框为黑色
  worksheet.eachRow(function(row, rowNumber) {
    if (rowNumber > 1) { // 跳过标题行
      row.border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } }
      };
    }
  });

// 添加数据到工作表
let orderNum = 1;
for (let n = 0; n < nightArr.length; n++) {
      let temrarr = [];
      temrarr.push(orderNum);
      temrarr.push("客运部");
      temrarr.push(nightArr[n].id);
      temrarr.push(nightArr[n].name);
      temrarr.push("站务车间1号线");
      temrarr.push(nightArr[n].job);
      temrarr.push(nightArr[n].night);
      if(nightArr[n].nightRange.length){
        let tempele = nightArr[n].nightRange;
        let tempLength = tempele.length;
        for(let m = 0;m < tempLength; m++){
          //  考勤周期为上月26号至本月25号
          // if(maxMonth.indexOf(setMonth) != -1){
          //   if(tempele[m] < 7){
          //     tempele[m] = setMonth + "月" + (daysInMonth - 6 + tempele[m]) + "日";
          //   }else{
          //     tempele[m] = currentMonth + "月" + (tempele[m] - 6) + "日";
          //   }
          //  } else if(formalMonth.indexOf(setMonth) != -1){
          //   if(tempele[m] < 6){
          //     tempele[m] = setMonth + "月" + (daysInMonth - 5 + tempele[m]) + "日";
          //   }else{
          //     tempele[m] = currentMonth + "月" + (tempele[m] - 5) + "日";
          //   }
          //  }else{
          //   // 2月份28天
          //   if(tempele[m] < 4){
          //     tempele[m] = setMonth + "月" + (daysInMonth - 3 + tempele[m]) + "日";
          //   }else{
          //     tempele[m] = currentMonth + "月" + (tempele[m] - 3) + "日";
          //   }
          //  }
          //  考勤周期为1号至月底
          tempele[m] = currentMonth + "月" + tempele[m] + "日";
        }
        temrarr.push(nightArr[n].nightRange);
      }
    nightworksheet.addRow(temrarr);
    orderNum++;
}
nightworksheet.mergeCells("A1:H1"); // 合并单元格
// 设置单元格样式
const mergedCellNight = nightworksheet.getCell('A1');
mergedCellNight.fill = {
  type: 'pattern',
  pattern: 'solid'
};
 

// 设置列宽度
nightworksheet.getColumn(1).width = 10; // 设置第1列的宽度为5
nightworksheet.getColumn(2).width = 10; 
nightworksheet.getColumn(3).width = 10; 
nightworksheet.getColumn(4).width = 12; 
nightworksheet.getColumn(5).width = 26; 
nightworksheet.getColumn(7).width = 20; 
nightworksheet.getColumn(8).width = 40; 
 // 设置所有单元格的对齐方式为居中并自动换行,所有行的高度
 nightworksheet.eachRow(function(row) {
  row.height = 40; // 设置行的高度25
  row.eachCell(function(cell) {
    // 去除引号和方括号
     const oldValue = cell.value;
     const newValue = (oldValue+"").replace(/[\[\]"]/g, ''); // 去除方括号 [] 和引号
     cell.value = newValue;
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  });
});


nightworksheet.getCell('A1').font = { name: 'Arial',bold: true, size: 20 }; // 设置加粗和字体大小为12
 // 设置某一行中多个单元格的字体样式
 const nightRowIndex = 2; // 设置要修改的行的索引
 const nightFont = {
   name: 'Arial',
   size: 14,
  //  bold: true,
  //  color: { argb: 'FF0000FF' }
 };
 const nightRow = nightworksheet.getRow(nightRowIndex);
 nightRow.eachCell(function(cell) {
   cell.font = nightFont;
 });
  // 设置所有数据行的边框为黑色
  nightworksheet.eachRow(function(row, rowNumber) {
    // 设置边框
    if (rowNumber > 1) { // 跳过标题行
      row.border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } }
      };
    }
  });

// 保存工作簿为考勤Excel文件
workbook.xlsx.writeFile(`杨湾中心站${currentMonth}月考勤统计表.xlsx`)
  .then(() => {
    console.log('考勤Excel文件和夜班统计Excel已生成');
    res.status(200).send({ message: 'upload successful' });
  })
  .catch((error) => {
    console.log('生成考勤Excel文件和夜班统计Excel时出错：', error);
    res.status(404).send({ message: 'upload hours error' });
  });



}); 
// 判断是否为闰年，计算2月份天数
function getFebruaryDays(year) {
  if (year % 4 === 0 && year % 100 !== 0 || year % 400 === 0) {
    return 29; // 闰年的2月份有29天
  } else {
    return 28; // 平年的2月份有28天
  }
}

// 启动服务器
app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
  console.log(`Server is running at http://localhost:${port}`);
});


