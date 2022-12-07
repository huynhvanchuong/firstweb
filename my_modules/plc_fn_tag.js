exports.fn_tag = fn_tag;

function fn_tag(socket, arr_tag_value){
        socket.emit("sql_insert_Trigger", arr_tag_value[0]); 
        socket.emit("Current_A", arr_tag_value[1]);  
        socket.emit("Current_B", arr_tag_value[2]);
        socket.emit("Current_C", arr_tag_value[3]);
        socket.emit("Current_Avg", arr_tag_value[4]);
        socket.emit("Voltage_AB", arr_tag_value[5]);
        socket.emit("Voltage_BC", arr_tag_value[6]);  
        socket.emit("Voltage_CA", arr_tag_value[7]);
        socket.emit("Voltage_LL_Avg", arr_tag_value[8]);
        socket.emit("Active_Power_Total", arr_tag_value[9]);
        socket.emit("Reactive_Power_Total", arr_tag_value[10]);
        socket.emit("Apparent_Power_Total", arr_tag_value[11]);  
        socket.emit("Frequency", arr_tag_value[12]);
        socket.emit("Power_Factor_Total", arr_tag_value[13]);
        socket.emit("THD_Current_A", arr_tag_value[14]);
        socket.emit("THD_Current_B", arr_tag_value[15]);
        socket.emit("THD_Current_C", arr_tag_value[16]);  
        socket.emit("THD_Voltage_AB", arr_tag_value[17]);
        socket.emit("THD_Voltage_BC", arr_tag_value[18]);
        socket.emit("THD_Voltage_CA", arr_tag_value[19]);
        socket.emit("Active_Load_Timer_Minute", arr_tag_value[20]);
        socket.emit("Stop_Load_Timer", arr_tag_value[21]);
        socket.emit("Hieu_suat_lam_viec", arr_tag_value[22]);
        socket.emit("sql_insert_Trigger_1D", arr_tag_value[23]); 

};