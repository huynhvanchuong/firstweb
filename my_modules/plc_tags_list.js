exports.tags_list = function(){
    return tags_list;
};
var tags_list = { 
    sql_insert_Trigger: 'DB20,X0.0', // Trigger ghi dữ liệu xuống SQL 
    Current_A: 'DB20,REAL2', 
    Current_B: 'DB20,REAL6',
    Current_C: 'DB20,REAL10',
    Current_Avg: 'DB20,REAL22',
    Voltage_AB: 'DB20,REAL42', 
    Voltage_BC: 'DB20,REAL46',
    Voltage_CA: 'DB20,REAL50',
    Voltage_LL_Avg: 'DB20,REAL54',
    Active_Power_Total: 'DB20,REAL122', 
    Reactive_Power_Total: 'DB20,REAL138',
    Apparent_Power_Total: 'DB20,REAL154',
    Frequency: 'DB20,REAL222', 
    Power_Factor_Total:'DB20,REAL226',
    THD_Current_A: 'DB20,REAL230',
    THD_Current_B: 'DB20,REAL234',
    THD_Current_C: 'DB20,REAL238',
    THD_Voltage_AB: 'DB20,REAL274',
    THD_Voltage_BC: 'DB20,REAL278',          
    THD_Voltage_CA: 'DB20,REAL282',                
    Active_Load_Timer_Minute: 'DB20,DINT366',
    Stop_Load_Timer: 'DB20,DINT370',
    Hieu_suat_lam_viec: 'DB20,REAL374',
    sql_insert_Trigger_1D: 'DB20,X378.0'
};
