// triger ghi dữ liệu vào SQL
var insert_trigger = false;			// Trigger
var old_insert_trigger = false;		// Trigger old
// triger ghi dữ liệu vào SQL1
var insert_trigger1 = false;			// Trigger
var old_insert_trigger1 = false;		// Trigger old
// Mảng xuất dữ liệu report Excel
var SQL_Excel = [];  // Dữ liệu nhập kho
// /////////////////////////////////////////////////++THIẾT LẬP KẾT NỐI WEB++////////////////////////////////////////////////////////
var express = require("express");
var app = express();
app.use(express.static("public"));
app.set("view engine", "ejs");
app.set("views", "./views");
var server = require("http").Server(app);
var io = require("socket.io")(server);
server.listen(3000);
// Gọi Home khi truy cập webserver 
app.get("/", function(req, res){
    res.render("home")
});
//
// KHỞI TẠO KẾT NỐI PLC
var nodes7 = require('nodes7');  
var conn_plc = new nodes7; //PLC1
// Tạo kết nối plc (slot = 2 nếu là 300/400, slot = 1 nếu là 1200/1500)
conn_plc.initiateConnection({port: 102, host: '192.168.1.10', rack: 0, slot: 1}, PLC_connected); 

///////////////////////////////////// Bảng tag trong Visual studio code////////////////////////////////////////////
const plc_tags_list = require('./my_modules/plc_tags_list');
const tags_list = plc_tags_list.tags_list();

// ////////////////////////////////////////////GỬI DỮ LIỆu TAG CHO PLC//////////////////////////////////////////////
const plc_tags_array = require('./my_modules/plc_tags_array');
const tags_array = plc_tags_array.tags_array();
function PLC_connected(err) {
    if (typeof(err) !== "undefined") {
        console.log(err); // Hiển thị lỗi nếu không kết nối đƯỢc với PLC
    }
    conn_plc.setTranslationCB(function(tag) {return tags_list[tag];});  // Đưa giá trị đọc lên từ PLC và mảng
    conn_plc.addItems(tags_array);
}
////////////////////////////////////////// Đọc dữ liệu từ PLC và đưa vào array tags////////////////////////////////
var arr_tag_value = []; // Tạo một mảng lưu giá trị tag đọc về
function valuesReady(anythingBad, values) {
    if (anythingBad) { console.log("Lỗi khi đọc dữ liệu tag"); } // Cảnh báo lỗi
    var lodash = require('lodash'); // Chuyển variable sang array
    arr_tag_value = lodash.map(values, (item) => item);
    console.log(values); // Hiển thị giá trị để kiểm tra
}
////////////////////////////////////////////// Hàm chức năng scan giá trị////////////////////////////////////////
function fn_read_data_scan(){
    conn_plc.readAllItems(valuesReady);
    fn_sql_insert();
    fn_sql_insert1();
}
// Time cập nhật mỗi 1s
setInterval(
    () => fn_read_data_scan(),
    1000 // 1s = 1000ms
);
// //////////////////////////////////////LẬP BẢNG TAG ĐỂ GỬI QUA CLIENT (TRÌNH DUYỆT)///////////////////////////////////
const plc_fn_tag = require('./my_modules/plc_fn_tag');
function fn_tag(socket){
    plc_fn_tag.fn_tag(socket, arr_tag_value);
}
// /////////////////////////////////////////////// GỬI DỮ LIỆU BẢNG TAG ĐẾN CLIENT (TRÌNH DUYỆT) ///////////////////////
io.on("connection", function(socket){
    socket.on("Client-send-data", function(data){
        fn_tag(socket);
    });
    fn_SQLSearch();         // Hàm tìm kiếm SQL
    fn_SQLSearch_ByTime();  // Hàm tìm kiếm SQL theo thời gian
    fn_Require_ExcelExport(); // Nhận yêu cầu xuất Excel
    fn_SQLSearch2();         // Hàm tìm kiếm SQL
    fn_SQLSearch_ByTime2();  // Hàm tìm kiếm SQL theo thời gian
    fn_Require_ExcelExport2(); // Nhận yêu cầu xuất Excel
});
//////////////////////////////////////////////////////// CƠ SỞ DỮ LIỆU SQL ///////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Khởi tạo SQL
var mysql = require('mysql');
var sqlcon = mysql.createConnection({
  host: "localhost",
  user: "root",
  password: "123456",
  database: "sql_plc",
  dateStrings:true // Hiển thị không có T và Z
});
var mysql1 = require('mysql');
var sqlcon1 = mysql1.createConnection({
    host: "localhost",
    user: "root",
    password: "123456",
    database: "sql_plc",
    dateStrings:true // Hiển thị không có T và Z
  });
/////////////////////////////////////////////////////////////////// ****///////////////////////////////////////////////////////
function fn_sql_insert(){
    insert_trigger = arr_tag_value[0];		// Read trigger from PLC
    var sqltable_Name = "plc_data";
    // Lấy thời gian hiện tại
	var tzoffset = (new Date()).getTimezoneOffset() * 60000; //Vùng Việt Nam (GMT7+)
	var temp_datenow = new Date();
	var timeNow = (new Date(temp_datenow - tzoffset)).toISOString().slice(0, -1).replace("T"," ");
	var timeNow_toSQL = "'" + timeNow + "',";
 
    // Dữ liệu đọc lên từ các tag
    var Current_Avg = "'" + arr_tag_value[4] + "',";
    var Voltage_LL_Avg = "'" + arr_tag_value[8] + "',";
    var Active_Power_Total = "'" + arr_tag_value[9] + "',";
    var Frequency = "'" + arr_tag_value[12] + "',";
    var Power_Factor_Total = "'" + arr_tag_value[13] + "'";
    // Ghi dữ liệu vào SQL
    if (insert_trigger && !old_insert_trigger)
    {
        var sql_write_str11 = "INSERT INTO " + sqltable_Name + " (date_time, Current_Avg, Voltage_LL_Avg, Active_Power_Total, Frequency, Power_Factor_Total) VALUES (";
        var sql_write_str12 = timeNow_toSQL 
                            + Current_Avg //////////Current_Avg Voltage_LL_Avg Active_Power_Total Frequency cosphi Times
                            + Voltage_LL_Avg
                            + Active_Power_Total
                            + Frequency
                            + Power_Factor_Total
                            ;
        var sql_write_str1 = sql_write_str11 + sql_write_str12 + ");";
        // Thực hiện ghi dữ liệu vào SQL
		sqlcon.query(sql_write_str1, function (err, result) {
            if (err) {
                console.log(err);
             } else {
                console.log("SQL - Ghi dữ liệu thành công");
              } 
			});
    }
    old_insert_trigger = insert_trigger;
}
// Đọc dữ liệu từ SQL
function fn_SQLSearch(){
    io.on("connection", function(socket){
        socket.on("msg_SQL_Show", function(data)
        {
            var sqltable_Name = "plc_data";
            var queryy1 = "SELECT * FROM " + sqltable_Name + ";" 
            sqlcon.query(queryy1, function(err, results, fields) {
                if (err) {
                    console.log(err);
                } else {
                    const objectifyRawPacket = row => ({...row});
                    const convertedResponse = results.map(objectifyRawPacket);
                    socket.emit('SQL_Show', convertedResponse);
                } 
            });
        });
    });
    }
// Đọc dữ liệu SQL theo thời gian
function fn_SQLSearch_ByTime(){
    io.on("connection", function(socket){
        socket.on("msg_SQL_ByTime", function(data)
        {
            var tzoffset = (new Date()).getTimezoneOffset() * 60000; //offset time Việt Nam (GMT7+)
            // Lấy thời gian tìm kiếm từ date time piker
            var timeS = new Date(data[0]); // Thời gian bắt đầu
            var timeE = new Date(data[1]); // Thời gian kết thúc
            // Quy đổi thời gian ra định dạng cua MySQL
            var timeS1 = "'" + (new Date(timeS - tzoffset)).toISOString().slice(0, -1).replace("T"," ")	+ "'";
            var timeE1 = "'" + (new Date(timeE - tzoffset)).toISOString().slice(0, -1).replace("T"," ") + "'";
            var timeR = timeS1 + "AND" + timeE1; // Khoảng thời gian tìm kiếm (Time Range)
            var sqltable_Name = "plc_data"; // Tên bảng
            var dt_col_Name = "date_time";  // Tên cột thời gian
            var Query1 = "SELECT * FROM " + sqltable_Name + " WHERE "+ dt_col_Name + " BETWEEN ";
            var Query = Query1 + timeR + ";";
            
            sqlcon.query(Query, function(err, results, fields) {
                if (err) {
                    console.log(err);
                } else {
                    const objectifyRawPacket = row => ({...row});
                    const convertedResponse = results.map(objectifyRawPacket);
                    SQL_Excel = convertedResponse; // Xuất báo cáo Excel
                    socket.emit('SQL_ByTime', convertedResponse);
                } 
            });
        });
    });
}
//////////////////////////////////////////////////////////////*************//////////////////////////////////////////////////////////////////
function fn_sql_insert1(){
    insert_trigger1 = arr_tag_value[23];		// Read trigger from PLC
    var sqltable_Name = "plcc_data";
    // Lấy thời gian hiện tại
	var tzoffset = (new Date()).getTimezoneOffset() * 60000; //Vùng Việt Nam (GMT7+)
	var temp_datenow = new Date();
	var timeNow = (new Date(temp_datenow - tzoffset)).toISOString().slice(0, -1).replace("T"," ");
	var timeNow_toSQL = "'" + timeNow + "',";
 
    // Dữ liệu đọc lên từ các tag
    var Active_Load_Timer_Minute = "'" + arr_tag_value[20] + "',";
    var Stop_Load_Timer = "'" + arr_tag_value[21] + "',";
    var Hieu_suat_lam_viec = "'" + arr_tag_value[22] + "'";
    // Ghi dữ liệu vào SQL
    if (insert_trigger1 && !old_insert_trigger1)
    {
        var sql_write_str11 = "INSERT INTO " + sqltable_Name + " (date_time, Active_Load_Timer_Minute, Stop_Load_Timer, Hieu_suat_lam_viec) VALUES (";
        var sql_write_str12 = timeNow_toSQL 
                            + Active_Load_Timer_Minute 
                            + Stop_Load_Timer
                            + Hieu_suat_lam_viec
                            ;
        var sql_write_str1 = sql_write_str11 + sql_write_str12 + ");";
        // Thực hiện ghi dữ liệu vào SQL
		sqlcon1.query(sql_write_str1, function (err, result) {
            if (err) {
                console.log(err);
             } else {
                console.log("SQL - Ghi dữ liệu thành công");
              } 
			});
    }
    old_insert_trigger1 = insert_trigger1;
}
// Đọc dữ liệu từ SQL
function fn_SQLSearch2(){
  io.on("connection", function(socket){
      socket.on("msg_SQL_Show2", function(data)
      {
          var sqltable_Name = "plcc_data";
          var queryy1 = "SELECT * FROM " + sqltable_Name + ";" 
          sqlcon.query(queryy1, function(err, results, fields) {
              if (err) {
                  console.log(err);
              } else {
                  const objectifyRawPacket = row => ({...row});
                  const convertedResponse = results.map(objectifyRawPacket);
                  socket.emit('SQL_Show2', convertedResponse);
              } 
          });
      });
  });
  }
  // Đọc dữ liệu SQL theo thời gian
function fn_SQLSearch_ByTime2(){
  io.on("connection", function(socket){
      socket.on("msg_SQL_ByTime2", function(data)
      {
          var tzoffset = (new Date()).getTimezoneOffset() * 60000; //offset time Việt Nam (GMT7+)
          // Lấy thời gian tìm kiếm từ date time piker
          var timeS = new Date(data[0]); // Thời gian bắt đầu
          var timeE = new Date(data[1]); // Thời gian kết thúc
          // Quy đổi thời gian ra định dạng cua MySQL
          var timeS1 = "'" + (new Date(timeS - tzoffset)).toISOString().slice(0, -1).replace("T"," ")	+ "'";
          var timeE1 = "'" + (new Date(timeE - tzoffset)).toISOString().slice(0, -1).replace("T"," ") + "'";
          var timeR = timeS1 + "AND" + timeE1; // Khoảng thời gian tìm kiếm (Time Range)

          var sqltable_Name = "plcc_data"; // Tên bảng
          var dt_col_Name = "date_time";  // Tên cột thời gian

          var Query1 = "SELECT * FROM " + sqltable_Name + " WHERE "+ dt_col_Name + " BETWEEN ";
          var Query = Query1 + timeR + ";";
          
          sqlcon.query(Query, function(err, results, fields) {
              if (err) {
                  console.log(err);
              } else {
                  const objectifyRawPacket = row => ({...row});
                  const convertedResponse = results.map(objectifyRawPacket);
                  SQL_Excel = convertedResponse; // Xuất báo cáo Excel
                  socket.emit('SQL_ByTime2', convertedResponse);
              } 
          });
      });
  });
}
// ////////////////////////////////////////////////// BÁO CÁO EXCEL ///////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
const Excel = require('exceljs');
const { CONNREFUSED } = require('dns');
function fn_excelExport(){
    // =====================CÁC THUỘC TÍNH CHUNG=====================
        // Lấy ngày tháng hiện tại
        let date_ob = new Date();
        let date = ("0" + date_ob.getDate()).slice(-2);
        let month = ("0" + (date_ob.getMonth() + 1)).slice(-2);
        let year = date_ob.getFullYear();
        let hours = date_ob.getHours();
        let minutes = date_ob.getMinutes();
        let seconds = date_ob.getSeconds();
        let day = date_ob.getDay();
        var dayName = '';
        if(day == 0){dayName = 'Chủ nhật,'}
        else if(day == 1){dayName = 'Thứ hai,'}
        else if(day == 2){dayName = 'Thứ ba,'}
        else if(day == 3){dayName = 'Thứ tư,'}
        else if(day == 4){dayName = 'Thứ năm,'}
        else if(day == 5){dayName = 'Thứ sáu,'}
        else if(day == 6){dayName = 'Thứ bảy,'}
        else{};
    // Tạo và khai báo Excel
    let workbook = new Excel.Workbook()
    let worksheet =  workbook.addWorksheet('Máy 1-Nhà máy sợi', {
      pageSetup:{paperSize: 9, orientation:'landscape'},
      properties:{tabColor:{argb:'FFC0000'}},
    });
    // Page setup (cài đặt trang)
    worksheet.properties.defaultRowHeight = 20;
    worksheet.pageSetup.margins = {
      left: 0.3, right: 0.25,
      top: 0.75, bottom: 0.75,
      header: 0.3, footer: 0.3
    };
    // =====================THẾT KẾ HEADER=====================
    // Logo công ty
    const imageId1 = workbook.addImage({
        filename: 'public/images/Logo.png',
        extension: 'png',
      });
    worksheet.addImage(imageId1, 'A1:B3');
    // Thông tin công ty
    worksheet.getCell('C1').value = 'CÔNG TY CỔ PHẦN DỆT MAY HUẾ';
    worksheet.getCell('C1').style = { font:{bold: true,size: 14},alignment: {vertical: 'middle'}} ;
    worksheet.getCell('C2').value = 'Địa chỉ: 122 Dương Thiệu Tước - P. Thủy Dương - TX Hương Thủy - TT. Huế';
    worksheet.getCell('C3').value = 'ĐT: 0234.3.864.337 - Fax: 0234.3.864.338';
    // Tên báo cáo
    worksheet.getCell('A5').value = 'BÁO CÁO CÁC CHỈ SỐ ĐIỆN MÁY 1';
    worksheet.mergeCells('A5:H5');
    worksheet.getCell('A5').style = { font:{name: 'Times New Roman', bold: true,size: 14},alignment: {horizontal:'center',vertical: 'middle'}} ;
    // Ngày in biểu
    worksheet.getCell('H6').value = "Ngày in báo cáo: " + dayName +" " + date + "/" + month + "/" + year + " " + hours + ":" + minutes + ":" + seconds;
    worksheet.getCell('H6').style = { font:{bold: false, italic: true},alignment: {horizontal:'right',vertical: 'bottom',wrapText: false}} ;
     
    // Tên nhãn các cột
    var rowpos = 7; 
    var collumName = ["STT","Thời gian", "Dòng điện (A)", "Điện áp (V)", "Công suất (kW)", "Tần số (Hz)", "Hệ số cosphi", "Ghi chú"]
    worksheet.spliceRows(rowpos, 1, collumName);
     
    // =====================XUẤT DỮ LIỆU EXCEL SQL=====================
    // Dump all the data into Excel
    var rowIndex = 0;
    SQL_Excel.forEach((e, index) => {
    // row 1 is the header.
    rowIndex =  index + rowpos;
    // worksheet1 collum
    worksheet.columns = [
          {key: 'STT'},
          {key: 'date_time'}, ///////date_time, Current_Avg, Voltage_LL_Avg, Active_Power_Total, Frequency, cosphi, Times
          {key: 'Current_Avg'},
          {key: 'Voltage_LL_Avg'},
          {key: 'Active_Power_Total'},
          {key: 'Frequency'},
          {key: 'Power_Factor_Total'}
        ]
    worksheet.addRow({
          STT: {
            formula: index + 1
          },
          ...e
        })
    })
    // Lấy tổng số hàng
    const totalNumberOfRows = worksheet.rowCount; 
    // Tính tổng
    worksheet.addRow([
        'Trung bình:',
        '',
      {formula: `=round(average(C${rowpos + 1}:C${totalNumberOfRows}),2)`},
      {formula: `=round(average(D${rowpos + 1}:D${totalNumberOfRows}),2)`},
      {formula: `=round(average(E${rowpos + 1}:E${totalNumberOfRows}),2)`},
      {formula: `=round(average(F${rowpos + 1}:F${totalNumberOfRows}),2)`},
      {formula: `=round(average(G${rowpos + 1}:G${totalNumberOfRows}),2)`},
    ])
    // Style cho hàng total (Tổng cộng)
    worksheet.getCell(`A${totalNumberOfRows+1}`).style = { font:{bold: true,size: 12},alignment: {horizontal:'center',}} ;
    // Tô màu cho hàng total (Tổng cộng)
    const total_row = ['A','B', 'C', 'D', 'E','F','G','H']
    total_row.forEach((v) => {
        worksheet.getCell(`${v}${totalNumberOfRows+1}`).fill = {type: 'pattern',pattern:'solid',fgColor:{ argb:'f2ff00' }}
    })
    // =====================STYLE CHO CÁC CỘT/HÀNG=====================
    // Style các cột nhãn
    const HeaderStyle = ['A','B', 'C', 'D', 'E','F','G','H']
    HeaderStyle.forEach((v) => {
        worksheet.getCell(`${v}${rowpos}`).style = { font:{bold: true},alignment: {horizontal:'center',vertical: 'middle',wrapText: true}} ;
        worksheet.getCell(`${v}${rowpos}`).border = {
          top: {style:'thin'},
          left: {style:'thin'},
          bottom: {style:'thin'},
          right: {style:'thin'}
        }
    })
    // Cài đặt độ rộng cột
    worksheet.columns.forEach((column, index) => {
        column.width = 15;
    })
    // Set width header
    worksheet.getColumn(1).width = 12;
    worksheet.getColumn(2).width = 20;
    worksheet.getColumn(8).width = 30;
     
    // ++++++++++++Style cho các hàng dữ liệu++++++++++++
    worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      var datastartrow = rowpos;
      var rowindex = rowNumber + datastartrow;
      const rowlength = datastartrow + SQL_Excel.length
      if(rowindex >= rowlength+1){rowindex = rowlength+1}
      const insideColumns = ['A','B', 'C', 'D', 'E','F','G','H']
    // Tạo border
      insideColumns.forEach((v) => {
          // Border
        worksheet.getCell(`${v}${rowindex}`).border = {
          top: {style: 'thin'},
          bottom: {style: 'thin'},
          left: {style: 'thin'},
          right: {style: 'thin'}
        },
        // Alignment
        worksheet.getCell(`${v}${rowindex}`).alignment = {horizontal:'center',vertical: 'middle',wrapText: true}
      })
    })
    // =====================THẾT KẾ FOOTER=====================
    worksheet.getCell(`H${totalNumberOfRows+2}`).value = 'Ngày …………tháng ……………năm 20………';
    worksheet.getCell(`H${totalNumberOfRows+2}`).style = { font:{bold: true, italic: false},alignment: {horizontal:'right',vertical: 'middle',wrapText: false}} ;
     
    worksheet.getCell(`B${totalNumberOfRows+3}`).value = 'Giám đốc';
    worksheet.getCell(`B${totalNumberOfRows+4}`).value = '(Ký, ghi rõ họ tên)';
    worksheet.getCell(`B${totalNumberOfRows+3}`).style = { font:{bold: true, italic: false},alignment: {horizontal:'center',vertical: 'bottom',wrapText: false}} ;
    worksheet.getCell(`B${totalNumberOfRows+4}`).style = { font:{bold: false, italic: true},alignment: {horizontal:'center',vertical: 'top',wrapText: false}} ;
     
    worksheet.getCell(`E${totalNumberOfRows+3}`).value = 'Trưởng ca';
    worksheet.getCell(`E${totalNumberOfRows+4}`).value = '(Ký, ghi rõ họ tên)';
    worksheet.getCell(`E${totalNumberOfRows+3}`).style = { font:{bold: true, italic: false},alignment: {horizontal:'center',vertical: 'bottom',wrapText: false}} ;
    worksheet.getCell(`E${totalNumberOfRows+4}`).style = { font:{bold: false, italic: true},alignment: {horizontal:'center',vertical: 'top',wrapText: false}} ;
     
    worksheet.getCell(`H${totalNumberOfRows+3}`).value = 'Người in biểu';
    worksheet.getCell(`H${totalNumberOfRows+4}`).value = '(Ký, ghi rõ họ tên)';
    worksheet.getCell(`H${totalNumberOfRows+3}`).style = { font:{bold: true, italic: false},alignment: {horizontal:'center',vertical: 'bottom',wrapText: false}} ;
    worksheet.getCell(`H${totalNumberOfRows+4}`).style = { font:{bold: false, italic: true},alignment: {horizontal:'center',vertical: 'top',wrapText: false}} ;
    // =====================THỰC HIỆN XUẤT DỮ LIỆU EXCEL=====================
    // Export Link
    var currentTime = year + "_" + month + "_" + date + "_" + hours + "h" + minutes + "m" + seconds + "s";
    var saveasDirect = "Report/Report_" + currentTime + ".xlsx";
    SaveAslink = saveasDirect; // Send to client
    var booknameLink = "public/" + saveasDirect;
    var Bookname = "Bao_cao_dien_" + currentTime + ".xlsx";
    // Write book name
    workbook.xlsx.writeFile(booknameLink)  
    // Return
    return [SaveAslink, Bookname]  
    } // Đóng fn_excelExport 1
// =====================TRUYỀN NHẬN DỮ LIỆU VỚI TRÌNH DUYỆT=====================
// Hàm chức năng truyền nhận dữ liệu với trình duyệt
function fn_Require_ExcelExport(){
    io.on("connection", function(socket){
        socket.on("msg_Excel_Report", function(data)
        {
            const [SaveAslink1, Bookname] = fn_excelExport();
            var data = [SaveAslink1, Bookname];
            socket.emit('send_Excel_Report', data);
        });
    });
}
//////////////////////////////
// /////////////////////////////// BÁO CÁO EXCEL ///////////////////////////////
function fn_excelExport2(){
    // =====================CÁC THUỘC TÍNH CHUNG=====================
        // Lấy ngày tháng hiện tại
        let date_ob = new Date();
        let date = ("0" + date_ob.getDate()).slice(-2);
        let month = ("0" + (date_ob.getMonth() + 1)).slice(-2);
        let year = date_ob.getFullYear();
        let hours = date_ob.getHours();
        let minutes = date_ob.getMinutes();
        let seconds = date_ob.getSeconds();
        let day = date_ob.getDay();
        var dayName = '';
        if(day == 0){dayName = 'Chủ nhật,'}
        else if(day == 1){dayName = 'Thứ hai,'}
        else if(day == 2){dayName = 'Thứ ba,'}
        else if(day == 3){dayName = 'Thứ tư,'}
        else if(day == 4){dayName = 'Thứ năm,'}
        else if(day == 5){dayName = 'Thứ sáu,'}
        else if(day == 6){dayName = 'Thứ bảy,'}
        else{};
    // Tạo và khai báo Excel
    let workbook = new Excel.Workbook()
    let worksheet =  workbook.addWorksheet('Máy 1-Nhà máy sợi', {
      pageSetup:{paperSize: 9, orientation:'landscape'},
      properties:{tabColor:{argb:'FFC0000'}},
    });
    // Page setup (cài đặt trang)
    worksheet.properties.defaultRowHeight = 20;
    worksheet.pageSetup.margins = {
      left: 0.3, right: 0.25,
      top: 0.75, bottom: 0.75,
      header: 0.3, footer: 0.3
    };
    // =====================THẾT KẾ HEADER=====================
    // Logo công ty
    const imageId1 = workbook.addImage({
        filename: 'public/images/Logo.png',
        extension: 'png',
      });
      worksheet.addImage(imageId1, 'A1:B3');
      // Thông tin công ty
      worksheet.getCell('C1').value = 'CÔNG TY CỔ PHẦN DỆT MAY HUẾ';
      worksheet.getCell('C1').style = { font:{name: 'Times New Roman', bold: true,size: 14},alignment: {vertical: 'middle'}} ;
      worksheet.getCell('C2').value = 'Địa chỉ: 122 Dương Thiệu Tước - P. Thủy Dương - TX Hương Thủy - TT. Huế';
      worksheet.getCell('C3').value = 'ĐT: 0234.3.864.337 - Fax: 0234.3.864.338';
    // Tên báo cáo
    worksheet.getCell('A5').value = 'BÁO CÁO VẬN HÀNH MÁY 1';
    worksheet.mergeCells('A5:F5');
    worksheet.getCell('A5').style = { font:{name: 'Times New Roman', bold: true,size: 14},alignment: {horizontal:'center',vertical: 'middle'}} ;
    // Ngày in biểu
    worksheet.getCell('F6').value = "Ngày in báo cáo: " + dayName +" " + date + "/" + month + "/" + year + " " + hours + ":" + minutes + ":" + seconds;
    worksheet.getCell('F6').style = { font:{bold: false, italic: true},alignment: {horizontal:'right',vertical: 'bottom',wrapText: false}} ;
     
    // Tên nhãn các cột
    var rowpos = 7; 
    var collumName = ["STT","Thời gian ca làm", "Thời gian tải chạy (Phút)", "Thời gian tải dừng (Phút)", "Hiệu suất (%)", "Ghi chú"]
    worksheet.spliceRows(rowpos, 1, collumName);
     
    // =====================XUẤT DỮ LIỆU EXCEL SQL=====================
    // Dump all the data into Excel
    var rowIndex = 0;
    SQL_Excel.forEach((e, index) => {
    // row 1 is the header.
    rowIndex =  index + rowpos;
    // worksheet1 collum
    worksheet.columns = [
          {key: 'STT'},
          {key: 'date_time'}, ///////date_time, Current_Avg, Voltage_LL_Avg, Active_Power_Total, Frequency, cosphi, Times
          {key: 'Active_Load_Timer_Minute'},
          {key: 'Stop_Load_Timer'},
          {key: 'Hieu_suat_lam_viec'}
        ]
    worksheet.addRow({
          STT: {
            formula: index + 1
          },
          ...e
        })
    })
    // Lấy tổng số hàng
    const totalNumberOfRows = worksheet.rowCount; 
    // Tính tổng
    worksheet.addRow([
        'Trung bình:',
        '',
      {formula: `=round(average(C${rowpos + 1}:C${totalNumberOfRows}),2)`},
      {formula: `=round(average(D${rowpos + 1}:D${totalNumberOfRows}),2)`},
      {formula: `=round(average(E${rowpos + 1}:E${totalNumberOfRows}),2)`},
    ])
    // Style cho hàng total (Tổng cộng)
    worksheet.getCell(`A${totalNumberOfRows+1}`).style = { font:{bold: true,size: 12},alignment: {horizontal:'center',}} ;
    // Tô màu cho hàng total (Tổng cộng)
    const total_row = ['A','B', 'C', 'D', 'E','F']
    total_row.forEach((v) => {
        worksheet.getCell(`${v}${totalNumberOfRows+1}`).fill = {type: 'pattern',pattern:'solid',fgColor:{ argb:'f2ff00' }}
    })
    // =====================STYLE CHO CÁC CỘT/HÀNG=====================
    // Style các cột nhãn
    const HeaderStyle = ['A','B', 'C', 'D', 'E','F']
    HeaderStyle.forEach((v) => {
        worksheet.getCell(`${v}${rowpos}`).style = { font:{bold: true},alignment: {horizontal:'center',vertical: 'middle',wrapText: true}} ;
        worksheet.getCell(`${v}${rowpos}`).border = {
          top: {style:'thin'},
          left: {style:'thin'},
          bottom: {style:'thin'},
          right: {style:'thin'}
        }
    })
    // Cài đặt độ rộng cột
    worksheet.columns.forEach((column, index) => {
        column.width = 15;
    })
    // Set width header
    worksheet.getColumn(1).width = 12;
    worksheet.getColumn(2).width = 20;
    worksheet.getColumn(3).width = 30;
    worksheet.getColumn(4).width = 30;
    worksheet.getColumn(6).width = 30;
    // ++++++++++++Style cho các hàng dữ liệu++++++++++++
    worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      var datastartrow = rowpos;
      var rowindex = rowNumber + datastartrow;
      const rowlength = datastartrow + SQL_Excel.length
      if(rowindex >= rowlength+1){rowindex = rowlength+1}
      const insideColumns = ['A','B', 'C', 'D', 'E','F']
    // Tạo border
      insideColumns.forEach((v) => {
          // Border
        worksheet.getCell(`${v}${rowindex}`).border = {
          top: {style: 'thin'},
          bottom: {style: 'thin'},
          left: {style: 'thin'},
          right: {style: 'thin'}
        },
        // Alignment
        worksheet.getCell(`${v}${rowindex}`).alignment = {horizontal:'center',vertical: 'middle',wrapText: true}
      })
    })
    // =====================THẾT KẾ FOOTER=====================
    worksheet.getCell(`F${totalNumberOfRows+2}`).value = 'Ngày …………tháng ……………năm 20………';
    worksheet.getCell(`F${totalNumberOfRows+2}`).style = { font:{bold: true, italic: false},alignment: {horizontal:'right',vertical: 'middle',wrapText: false}} ;
     
    worksheet.getCell(`B${totalNumberOfRows+3}`).value = 'Giám đốc';
    worksheet.getCell(`B${totalNumberOfRows+4}`).value = '(Ký, ghi rõ họ tên)';
    worksheet.getCell(`B${totalNumberOfRows+3}`).style = { font:{bold: true, italic: false},alignment: {horizontal:'center',vertical: 'bottom',wrapText: false}} ;
    worksheet.getCell(`B${totalNumberOfRows+4}`).style = { font:{bold: false, italic: true},alignment: {horizontal:'center',vertical: 'top',wrapText: false}} ;
     
    worksheet.getCell(`D${totalNumberOfRows+3}`).value = 'Trưởng ca';
    worksheet.getCell(`D${totalNumberOfRows+4}`).value = '(Ký, ghi rõ họ tên)';
    worksheet.getCell(`D${totalNumberOfRows+3}`).style = { font:{bold: true, italic: false},alignment: {horizontal:'center',vertical: 'bottom',wrapText: false}} ;
    worksheet.getCell(`D${totalNumberOfRows+4}`).style = { font:{bold: false, italic: true},alignment: {horizontal:'center',vertical: 'top',wrapText: false}} ;
     
    worksheet.getCell(`F${totalNumberOfRows+3}`).value = 'Người in biểu';
    worksheet.getCell(`F${totalNumberOfRows+4}`).value = '(Ký, ghi rõ họ tên)';
    worksheet.getCell(`F${totalNumberOfRows+3}`).style = { font:{bold: true, italic: false},alignment: {horizontal:'center',vertical: 'bottom',wrapText: false}} ;
    worksheet.getCell(`F${totalNumberOfRows+4}`).style = { font:{bold: false, italic: true},alignment: {horizontal:'center',vertical: 'top',wrapText: false}} ;
     
    // =====================THỰC HIỆN XUẤT DỮ LIỆU EXCEL=====================
    // Export Link
    var currentTime = year + "_" + month + "_" + date + "_" + hours + "h" + minutes + "m" + seconds + "s";
    var saveasDirect = "Report/Report_" + currentTime + ".xlsx";
    SaveAslink = saveasDirect; // Send to client
    var booknameLink = "public/" + saveasDirect;
    var Bookname = "Bao_cao_van_hanh_" + currentTime + ".xlsx";
    // Write book name
    workbook.xlsx.writeFile(booknameLink)
    // Return
    return [SaveAslink, Bookname]
    } // Đóng fn_excelExport
// =====================TRUYỀN NHẬN DỮ LIỆU VỚI TRÌNH DUYỆT=====================
// Hàm chức năng truyền nhận dữ liệu với trình duyệt
function fn_Require_ExcelExport2(){
    io.on("connection", function(socket){
        socket.on("msg_Excel_Report2", function(data)
        {
            const [SaveAslink1, Bookname] = fn_excelExport2();
            var data = [SaveAslink1, Bookname];
            socket.emit('send_Excel_Report2', data);
        });
    });
}