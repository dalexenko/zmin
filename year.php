<?php 
 Error_Reporting(E_ALL & ~E_NOTICE);

if (!extension_loaded('dbase')) 
	{
		dl('php_dbase.dll');
	}


set_time_limit (600);

// переменные

echo $db_file = "zmin_05.dbf";


$kkv = array ("1111" => "4", "1120" => "5", "1131" => "6", "1135" => "7", "1136" => "8", "1137" => "9", "1138" => "10", "1139" => "11", "1140" => "12", "1161" => "13", "1162" => "14", "1163" => "15", "1164" => "16", "1165" => "17", "2110" => "18", "2120" => "19", "2122" => "20", "2123" => "21", "2130" => "22", "2132" => "23", "2133" => "24", "2300" => "25");

// функции


function empty_dbf_file($id)
{
	for ($ndel = 1; $ndel <= dbase_numrecords ($id); $ndel++) 
		{
        dbase_delete_record ($id, $ndel);
		}
		dbase_pack ($id);
}


function insert_data ($sdvig)
{
    global $nr, $dat, $kpk, $kkv, $id, $Workbook, $sheet1, $sheet2;

    $nom = $nr;
	$srk = 3;
    $kvk = 350;
    $udk = 4;

    $SUTVSG = 0;
    $SUTVS = 0;
    $SUTVV = 0;
    $SUTVD = 0;
    $SUTVB = 0;
    $SUTVI = 0;
    $SUTVR = 0;
    $SUTV1 = 0;
    $SUTV2 = 0;
    $SUTV3 = 0;
    $SUTV4 = 0;
    $SUTV5 = 0;
    $SUTV6 = 0;
    $SUTV7 = 0;
    $SUTV8 = 0;
    $SUTV9 = 0;
    $SUTV10 = 0;
    $SUTV11 = 0;
    $SUTV12 = 0;

    for ($i = 46; $i < 1333; 28) {
        $coord_cod_terr = "A" . $i;
        $coord_name = "B" . $i;
        $coord_kpol = "E" . $i;

        $kekv = $i + $sdvig;

        $coord_kekv = "C" . $kekv;
        $coord_SUTVG = "D" . $kekv;

        $Worksheet = $Workbook->Worksheets($sheet1);
        $Worksheet->activate;

        $excel_cell_cod_terr = $Worksheet->Range($coord_cod_terr);
        $excel_cell_cod_terr->activate;
        $excel_result_cod_terr = substr ($excel_cell_cod_terr->value, 2, 4);

        $excel_cell_name = $Worksheet->Range($coord_name);
        $excel_cell_name->activate;
        $excel_result_name = $excel_cell_name->value;

        $excel_cell_kpol = $Worksheet->Range($coord_kpol);
        $excel_cell_kpol->activate;
        $excel_result_kpol = $excel_cell_kpol->value;

        $excel_cell_kekv = $Worksheet->Range($coord_kekv);
        $excel_cell_kekv->activate;
        $excel_result_kekv = $excel_cell_kekv->value;

        $excel_cell_SUTVG = $Worksheet->Range($coord_SUTVG);
        $excel_cell_SUTVG->activate;
        $excel_result_SUTVG = $excel_cell_SUTVG->value;

        print $excel_result_cod_terr . ": " . $excel_result_name . "\n";

        $vdk = substr($excel_result_cod_terr, 0, 2);

        print "$srk: $nr: $nom: $dat: $kvk: $excel_result_kpol: $udk: $kpk: $excel_result_kekv: $excel_result_SUTVG:\n";
      // запись в DBF файл
        dbase_add_record ($id, array ($srk, $nr, $nom, $dat, $kvk, $excel_result_kpol, $udk, $vdk, $kpk, $excel_result_kekv, $excel_result_SUTVG, $SUTVSG, $SUTVS, $SUTVV, $SUTVD, $SUTVB, $SUTVI, $SUTVR, $SUTV1 , $SUTV2, $SUTV3, $SUTV4, $SUTV5, $SUTV6, $SUTV7, $SUTV8, $SUTV9, $SUTV10, $SUTV11, $SUTV12)) or die ("Could not add a record to dbf file <i>$db_file</i>.");

        $i = $i + 28;
    } 
} 

// открытие DBF файла для записи


$id = dbase_open ($db_file, 2) or die ("Could not open dbf file <i>$db_file</i>."); 


// открытие и чтение XLS файла
$filename = "D:/zmin/dod6.xls";
$sheet1 = "дод6";
$sheet2 = "лист2";

$excel_app = new COM("Excel.application") or Die ("Did not connect");

$excel_app->Visible = 1;

$Workbook = $excel_app->Workbooks->Open("$filename") or Die("Did not open $filename $Workbook");


if ($action=='cleardbf') {
   empty_dbf_file($id);
} 


// чтение

if ($k1111) {
    insert_data ($kkv['1111']);
} 
if ($k1120) {
	insert_data ($kkv['1120']);
} 
if ($k1131) {
	insert_data ($kkv['1131']);
} 
if ($k1135) {
    insert_data ($kkv['1135']);
} 
if ($k1136) {
    insert_data ($kkv['1136']);
} 
if ($k1137) {
	insert_data ($kkv['1137']);
} 
if ($k1138) {
    insert_data ($kkv['1138']);
} 
if ($k1139) {
    insert_data ($kkv['1139']);
} 
if ($k1140) {
    insert_data ($kkv['1140']);
} 
if ($k1161) {
    insert_data ($kkv['1161']);
} 
if ($k1162) {
    insert_data ($kkv['1162']);
} 
if ($k1163) {
    insert_data ($kkv['1163']);
} 
if ($k1164) {
    insert_data ($kkv['1164']);
} 
if ($k1165) {
    insert_data ($kkv['1165']);
} 
if ($k2110) {
    insert_data ($kkv['2110']);
} 
if ($k2120) {
    insert_data ($kkv['2120']);
} 
if ($k2122) {
    insert_data ($kkv['2122']);
} 
if ($k2123) {
    insert_data ($kkv['2123']);
} 
if ($k2130) {
    insert_data ($kkv['2130']);
} 
if ($k2132) {
    insert_data ($kkv['2132']);
} 
if ($k2133) {
    insert_data ($kkv['2133']);
} 
if ($k2300) {
    insert_data ($kkv['2300']);
} 

// закрыть DBF файл для записи
$id_close = dbase_close ($id) or die ("Could not close dbf file <i>$db_file</i>."); 

if (!copy($db_file, 'export/' . substr($db_file, 0, -4) . '_' . $dat . '.dbf')) {
        echo "failed to copy $db_$file...<br />\n";
    } 

// closing excel
$excel_app->Quit();

// free the object
//$excel_app->Release();
$excel_app = null;

?>