 ////////////// CÁC KHỐI CHƯƠNG TRÌNH CON //////////////-->
   
        // Chương trình con đọc dữ liệu lên IO Field
        function fn_IOFieldDataShow(tag, IOField, tofix){
            socket.on(tag,function(data){
                if(tofix == 0){
                    document.getElementById(IOField).value = data;
                } else{
                document.getElementById(IOField).value = data.toFixed(tofix);
                }
            });
        }
////////////// YÊU CẦU DỮ LIỆU TỪ SERVER- REQUEST DATA //////////////-->

        var myVar = setInterval(myTimer, 100);
        function myTimer() {
            socket.emit("Client-send-data", "Request data client");
        }
// Chương trình con chuyển trang
function fn_ScreenChange(scr_1, scr_2, scr_3)
{
    document.getElementById(scr_1).style.visibility = 'visible';   // Hiển thị trang được chọn
    document.getElementById(scr_2).style.visibility = 'hidden';    // Ẩn trang 1
    document.getElementById(scr_3).style.visibility = 'hidden';    // Ẩn trang 2
}
