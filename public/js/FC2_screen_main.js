// Chương trình con đọc dữ liệu SQL
function fn_Table01_SQL_Show(){
    socket.emit("msg_SQL_Show", "true");
    socket.on('SQL_Show',function(data){
        fn_table_01(data);
    }); 
}
// Chương trình con hiển thị SQL ra bảng
function fn_table_01(data){
    if(data){
        $("#table_01 tbody").empty(); 
        var len = data.length;
        var txt = "<tbody>";
        if(len > 0){
            for(var i=0;i<len;i++){
                    txt += "<tr><td>"+data[i].date_time     ///////Current_Avg Voltage_LL_Avg Active_Power_Total Frequency cosphi Times
                        +"</td><td>"+data[i].Current_Avg
                        +"</td><td>"+data[i].Voltage_LL_Avg
                        +"</td><td>"+data[i].Active_Power_Total
                        +"</td><td>"+data[i].Frequency
                        +"</td><td>"+data[i].Power_Factor_Total
                        +"</td></tr>";
                    }
            if(txt != ""){
            txt +="</tbody>"; 
            $("#table_01").append(txt);
            }
        }
    }   
}
// Tìm kiếm SQL theo khoảng thời gian
function fn_SQL_By_Time()
{
    var val = [document.getElementById('dtpk_Search_Start').value,
               document.getElementById('dtpk_Search_End').value];
    socket.emit('msg_SQL_ByTime', val);
    socket.on('SQL_ByTime', function(data){
        fn_table_01(data); // Show sdata
    });
}

////////////// CODE TẠO BẢNG TRONG SQL //////////////
// CREATE TABLE `web_plc`.`plc_data` (
//     `date_time` DATETIME NOT NULL,
//     `Current_Avg` FLOAT NULL,
//     `Voltage_LL_Avg` FLOAT NULL,
//     `Active_Power_Total` FLOAT NULL,
//     `Frequency` FLOAT NULL,
//     `cosphi` FLOAT NULL,
//     `Times` FLOAT NULL)
//   ENGINE = InnoDB
//   DEFAULT CHARACTER SET = utf8
//   COLLATE = utf8_bin;

// Chương trình con đọc dữ liệu SQL
function fn_Table02_SQL_Show(){
    socket.emit("msg_SQL_Show2", "true");
    socket.on('SQL_Show2',function(data){
        fn_table_02(data);
    }); 
}
// Chương trình con hiển thị SQL ra bảng
function fn_table_02(data){
    if(data){
        $("#table_02 tbody").empty(); 
        var len = data.length;
        var txt = "<tbody>";
        if(len > 0){
            for(var i=0;i<len;i++){
                    txt += "<tr><td>"+data[i].date_time     ///////Current_Avg Voltage_LL_Avg Active_Power_Total Frequency cosphi Times
                        +"</td><td>"+data[i].Active_Load_Timer_Minute
                        +"</td><td>"+data[i].Stop_Load_Timer
                        +"</td><td>"+data[i].Hieu_suat_lam_viec
                        +"</td></tr>";
                    }
            if(txt != ""){
            txt +="</tbody>"; 
            $("#table_02").append(txt);
            }
        }
    }   
}
// Tìm kiếm SQL theo khoảng thời gian
function fn_SQL_By_Time2()
{
    var val = [document.getElementById('dtpk_Search_Start2').value,
               document.getElementById('dtpk_Search_End2').value];
    socket.emit('msg_SQL_ByTime2', val);
    socket.on('SQL_ByTime2', function(data){
        fn_table_02(data); // Show sdata
    });
}