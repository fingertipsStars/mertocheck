
const express = require('express');
const bodyParser = require('body-parser');
const app = express();
var multer = require('multer');
var upload = multer();
var fs = require('fs');
const ExcelJS = require('exceljs');
const nodeXlxs = require("node-xlsx");
let setMonth = new Date().getMonth() + 1;
const port = 3000;
let newExcelHeader = ['序号','工号','姓名','车站/队/组','日期','事件内容','考核条款','考核分值','检查人','总分','员工签字'];
let newExcelTitle = '站务车间杨湾中心站'+ setMonth +'月份月度绩效考核汇总';

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
  
// POST接口示例
app.post('/api/upload', upload.any(),(req, res) => {
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
  
  // 在这里处理接收到的数据
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
// 添加一个工作表
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
workbook.xlsx.writeFile('output.xlsx')
  .then(() => {
    console.log('Excel文件已生成');
    res.status(200).send({ message: 'upload successful' });
  })
  .catch((error) => {
    console.log('生成Excel文件时出错：', error);
    res.status(404).send({ message: 'upload error' });
  });
  
});
// 启动服务器
app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
  console.log(`Server is running at http://localhost:${port}`);
});


