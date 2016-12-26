<?php 
//Error_Reporting(E_ALL & ~E_NOTICE);

if (!extension_loaded('dbase')) 
	{
		dl('php_dbase.dll');
	}

set_time_limit (600);

// переменные

$exp_file = 'ZMIN_05.dbf';


$nom = $nr;
$dat;
$srk = 3;
$kvk = 350;
$udk = 4;
$kpk;


$regions = array (
	"1" => array("KPOL" => "06901", "VDK" => "51", "sheet" => "ВІЛЬН"),
    "2" => array("KPOL" => "06902", "VDK" => "52", "sheet" => "ДНІПРДЗ"),
    "3" => array("KPOL" => "06903", "VDK" => "54", "sheet" => "ЖОВТІВ"),
    "4" => array("KPOL" => "06904", "VDK" => "55", "sheet" => "Кр.Р"),
    "5" => array("KPOL" => "06905", "VDK" => "56", "sheet" => "Марг."),
    "6" => array("KPOL" => "06906", "VDK" => "57", "sheet" => "Нік."),
    "7" => array("KPOL" => "06907", "VDK" => "58", "sheet" => "Новом."),
    "8" => array("KPOL" => "06908", "VDK" => "59", "sheet" => "Орд."),
    "9" => array("KPOL" => "06909", "VDK" => "60", "sheet" => "Павл."),
    "10" => array("KPOL" => "06910", "VDK" => "61", "sheet" => "Перш."),
    "11" => array("KPOL" => "06911", "VDK" => "62", "sheet" => "Синельн."),
    "12" => array("KPOL" => "06912", "VDK" => "63", "sheet" => "Терн."),
    "13" => array("KPOL" => "06913", "VDK" => "01", "sheet" => "Апост."),
    "14" => array("KPOL" => "06914", "VDK" => "02", "sheet" => "Васильк."),
    "15" => array("KPOL" => "06915", "VDK" => "03", "sheet" => "Верх."),
    "16" => array("KPOL" => "06916", "VDK" => "04", "sheet" => "Дніпр."),
    "17" => array("KPOL" => "06917", "VDK" => "05", "sheet" => "Крив.р-н"),
    "18" => array("KPOL" => "06918", "VDK" => "06", "sheet" => "Крин."),
    "19" => array("KPOL" => "06919", "VDK" => "07", "sheet" => "Магд."),
    "20" => array("KPOL" => "06920", "VDK" => "08", "sheet" => "Меж."),
    "21" => array("KPOL" => "06921", "VDK" => "09", "sheet" => "Нікоп.р-н"),
    "22" => array("KPOL" => "06922", "VDK" => "10", "sheet" => "Новом. р-н"),
    "23" => array("KPOL" => "06923", "VDK" => "11", "sheet" => "Павл.р-н"),
    "24" => array("KPOL" => "06924", "VDK" => "12", "sheet" => "Петрик."),
    "25" => array("KPOL" => "06925", "VDK" => "13", "sheet" => "Петроп."),
    "26" => array("KPOL" => "06926", "VDK" => "14", "sheet" => "Покров."),
    "27" => array("KPOL" => "06927", "VDK" => "15", "sheet" => "Пятих."),
    "28" => array("KPOL" => "06928", "VDK" => "16", "sheet" => "Синельнр-н"),
    "29" => array("KPOL" => "06929", "VDK" => "17", "sheet" => "Солон."),
    "30" => array("KPOL" => "06930", "VDK" => "18", "sheet" => "Соф."),
    "31" => array("KPOL" => "06931", "VDK" => "19", "sheet" => "Томак."),
    "32" => array("KPOL" => "06932", "VDK" => "20", "sheet" => "Царич."),
    "33" => array("KPOL" => "06933", "VDK" => "21", "sheet" => "Широк."),
    "34" => array("KPOL" => "06934", "VDK" => "22", "sheet" => "Юр'їв."),
    "35" => array("KPOL" => "06935", "VDK" => "81", "sheet" => "АНД"),
    "36" => array("KPOL" => "06936", "VDK" => "82", "sheet" => "Бабуш."),
    "37" => array("KPOL" => "06937", "VDK" => "83", "sheet" => "Жовт.ДП"),
    "38" => array("KPOL" => "06938", "VDK" => "84", "sheet" => "Індуст."),
    "39" => array("KPOL" => "06939", "VDK" => "85", "sheet" => "Кіров."),
    "40" => array("KPOL" => "06940", "VDK" => "86", "sheet" => "Красн."),
    "41" => array("KPOL" => "06941", "VDK" => "87", "sheet" => "Лен."),
    "43" => array("KPOL" => "06943", "VDK" => "90", "sheet" => "Жовт.К.Р."),
    "44" => array("KPOL" => "06944", "VDK" => "91", "sheet" => "Інгул."),
    "45" => array("KPOL" => "06945", "VDK" => "92", "sheet" => "Тернівс."),
    "46" => array("KPOL" => "06946", "VDK" => "93", "sheet" => "Центр-гор."),
    "69" => array("KPOL" => "00069", "VDK" => "00", "sheet" => "УДК")
    );


// функции

function print_name ($Worksheet, $sheet)
{
    global $nom, $nr, $dat, $srk, $kvk, $udk, $kpk;

    $j = 14;

    $coord_name = "A" . $j;

    $excel_cell_name = $Worksheet->Range($coord_name);
    $excel_cell_name->activate;
    $excel_result_name = $excel_cell_name->value;

    print "$excel_result_name\n";
} 

function print_data ($j, $nom, $nr, $dat, $srk, $kvk, $udk, $kpk, $kpol, $vdk, $Worksheet)
{
    $coord_kekv = "A" . $j;

    $SUTVSG = 0;
    $SUTVS = 0;
    $SUTVV = 0;
    $SUTVD = 0;
    $SUTVB = 0;
    $SUTVI = 0;

    $coord_SUTVR = "O" . $j;

    $coord_SUTV1 = "C" . $j;
    $coord_SUTV2 = "D" . $j;
    $coord_SUTV3 = "E" . $j;
    $coord_SUTV4 = "F" . $j;
    $coord_SUTV5 = "G" . $j;
    $coord_SUTV6 = "H" . $j;
    $coord_SUTV7 = "I" . $j;
    $coord_SUTV8 = "J" . $j;
    $coord_SUTV9 = "K" . $j;
    $coord_SUTV10 = "L" . $j;
    $coord_SUTV11 = "M" . $j;
    $coord_SUTV12 = "N" . $j;

    $excel_cell_kekv = $Worksheet->Range($coord_kekv);
    $excel_cell_kekv->activate;
    $kekv = $excel_cell_kekv->value;

    $excel_cell_SUTVR = $Worksheet->Range($coord_SUTVR);
    $excel_cell_SUTVR->activate;
    $SUTVR = $excel_cell_SUTVR->value;

    if ($kekv == 1120) {
        $SUTVG = $SUTVR;
    } else {
        $SUTVG = 0;
    } 

    $excel_cell_SUTV1 = $Worksheet->Range($coord_SUTV1);
    $excel_cell_SUTV1->activate;
    $SUTV1 = $excel_cell_SUTV1->value;

    $excel_cell_SUTV2 = $Worksheet->Range($coord_SUTV2);
    $excel_cell_SUTV2->activate;
    $SUTV2 = $excel_cell_SUTV2->value;

    $excel_cell_SUTV3 = $Worksheet->Range($coord_SUTV3);
    $excel_cell_SUTV3->activate;
    $SUTV3 = $excel_cell_SUTV3->value;

    $excel_cell_SUTV4 = $Worksheet->Range($coord_SUTV4);
    $excel_cell_SUTV4->activate;
    $SUTV4 = $excel_cell_SUTV4->value;

    $excel_cell_SUTV5 = $Worksheet->Range($coord_SUTV5);
    $excel_cell_SUTV5->activate;
    $SUTV5 = $excel_cell_SUTV5->value;

    $excel_cell_SUTV6 = $Worksheet->Range($coord_SUTV6);
    $excel_cell_SUTV6->activate;
    $SUTV6 = $excel_cell_SUTV6->value;

    $excel_cell_SUTV7 = $Worksheet->Range($coord_SUTV7);
    $excel_cell_SUTV7->activate;
    $SUTV7 = $excel_cell_SUTV7->value;

    $excel_cell_SUTV8 = $Worksheet->Range($coord_SUTV8);
    $excel_cell_SUTV8->activate;
    $SUTV8 = $excel_cell_SUTV8->value;

    $excel_cell_SUTV9 = $Worksheet->Range($coord_SUTV9);
    $excel_cell_SUTV9->activate;
    $SUTV9 = $excel_cell_SUTV9->value;

    $excel_cell_SUTV10 = $Worksheet->Range($coord_SUTV10);
    $excel_cell_SUTV10->activate;
    $SUTV10 = $excel_cell_SUTV10->value;

    $excel_cell_SUTV11 = $Worksheet->Range($coord_SUTV11);
    $excel_cell_SUTV11->activate;
    $SUTV11 = $excel_cell_SUTV11->value;

    $excel_cell_SUTV12 = $Worksheet->Range($coord_SUTV12);
    $excel_cell_SUTV12->activate;
    $SUTV12 = $excel_cell_SUTV12->value;

    $insert_array = array ($srk, $nr, $nom, $dat, $kvk, $kpol, $udk, $vdk, $kpk, $kekv, $SUTVG, $SUTVSG, $SUTVS, $SUTVV, $SUTVD, $SUTVB, $SUTVI, $SUTVR, $SUTV1 , $SUTV2, $SUTV3, $SUTV4, $SUTV5, $SUTV6, $SUTV7, $SUTV8, $SUTV9, $SUTV10, $SUTV11, $SUTV12);
    return $insert_array;
} 

function open_export_file()
{
    global $exp_file;

    $id = dbase_open ($exp_file, 2) or die ("Could not open dbf file <i>$exp_file</i>.");

    for ($ndel = 1; $ndel <= dbase_numrecords ($id); $ndel++) {
        dbase_delete_record ($id, $ndel);
    } 
    dbase_pack ($id);
    return $id;
} 

function insert_into_export_file($id, $ins_data_array)
{
    dbase_add_record ($id, $ins_data_array);
} 

function close_export_file($id)
{
    global $exp_file, $dat;
    dbase_close($id);
    if (!copy($exp_file, 'export/' . substr($exp_file, 0, -4) . '_' . $dat . '.dbf')) {
        echo "failed to copy $exp_$file...\n";
    } 
} 

function insert_data ($trs)
{
    global $regions, $Workbook, $id, $nom, $nr, $dat, $srk, $kvk, $udk, $kpk;

    $kpol = $regions[$trs]['KPOL'];
    $vdk = $regions[$trs]['VDK'];
    $sheet = $regions[$trs]['sheet'];

    $Worksheet = $Workbook->Worksheets($sheet);
    $Worksheet->activate;

    print_name($Worksheet, $sheet);

    for ($j = 22; $j < 26; $j++) {
        $ins_data_array = print_data ($j, $nom, $nr, $dat, $srk, $kvk, $udk, $kpk, $kpol, $vdk, $Worksheet);

        insert_into_export_file($id, $ins_data_array);
    } ;
} 

// открытие XLS файла

$filename = "D:/zmin/dod7.xls";

$excel_app = new COM("Excel.application") or Die ("Did not connect");

$excel_app->Visible = 1;

$Workbook = $excel_app->Workbooks->Open("$filename") or Die("Did not open $filename $Workbook");



$id = open_export_file();

// чтение XLS файла

if ($t69) {
	insert_data ($t69);
} 
if ($t1) {
	insert_data ($t1);
}
if ($t2) {
    insert_data ($t2);
} 
if ($t3) {
    insert_data ($t3);
} 
if ($t4) {
    insert_data ($t4);
} 
if ($t5) {
    insert_data ($t5);
} 
if ($t6) {
    insert_data ($t6);
} 
if ($t7) {
    insert_data ($t7);
} 
if ($t8) {
    insert_data ($t8);
} 
if ($t9) {
    insert_data ($t9);
} 
if ($t10) {
    insert_data ($t10);
} 
if ($t11) {
    insert_data ($t11);
} 
if ($t12) {
    insert_data ($t12);
} 
if ($t13) {
    insert_data ($t13);
} 
if ($t14) {
    insert_data ($t14);
} 
if ($t15) {
    insert_data ($t15);
} 
if ($t16) {
    insert_data ($t16);
} 
if ($t17) {
    insert_data ($t17);
} 
if ($t18) {
    insert_data ($t18);
} 
if ($t19) {
    insert_data ($t19);
} 
if ($t20) {
    insert_data ($t20);
} 
if ($t21) {
    insert_data ($t21);
} 
if ($t22) {
    insert_data ($t22);
} 
if ($t23) {
    insert_data ($t23);
} 
if ($t24) {
    insert_data ($t24);
} 
if ($t25) {
    insert_data ($t25);
} 
if ($t26) {
    insert_data ($t26);
} 
if ($t27) {
    insert_data ($t27);
} 
if ($t28) {
    insert_data ($t28);
} 
if ($t29) {
    insert_data ($t29);
} 
if ($t30) {
    insert_data ($t30);
} 
if ($t31) {
    insert_data ($t31);
} 
if ($t32) {
    insert_data ($t32);
} 
if ($t33) {
    insert_data ($t33);
} 
if ($t34) {
    insert_data ($t34);
} 
if ($t35) {
    insert_data ($t35);
} 
if ($t36) {
    insert_data ($t36);
} 
if ($t37) {
    insert_data ($t37);
} 
if ($t38) {
    insert_data ($t38);
} 
if ($t39) {
    insert_data ($t39);
} 
if ($t40) {
    insert_data ($t40);
} 
if ($t41) {
    insert_data ($t41);
} 
if ($t43) {
    insert_data ($t43);
} 
if ($t44) {
    insert_data ($t44);
} 
if ($t45) {
    insert_data ($t45);
} 
if ($t46) {
    insert_data ($t46);
} 


close_export_file($id);



//closing excel

$excel_app->Quit();

//free the object

//$excel_app->Release();
$excel_app = null;

?>