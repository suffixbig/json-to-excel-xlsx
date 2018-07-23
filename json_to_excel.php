<?php
header ( "Content-type: text/html; charset=utf-8" );
//date_default_timezone_set ( 'PRC' ); //設置中國時區
date_default_timezone_set("Asia/Shanghai");//設置台灣時區
$dir=dirname(__FILE__);//目前路徑


//目的
/*
$filename = "/Users/jiangminghui/Documents/test/hao.xlsx";

$objPHPExcel = PHPExcel_IOFactory::load($filename);//加載文件

$sheet = $objPHPExcel->getSheet(0);
$highestRow = $sheet->getHighestRow(); // 取得總行數
$highestColumn = $sheet->getHighestColumn(); // 取得總列數
$city = array();
for($i=1;$i<=$highestRow;$i++) {
	$key = $objPHPExcel->getActiveSheet()->getCell("A".$i)->getValue();
	$val = $objPHPExcel->getActiveSheet()->getCell("B".$i)->getValue();
	$city[$val] = $key;
}
print_r($city);
*/

include 'PHPExcel.php';
$br="<br>\n";
//來源
$filename = $dir."/test01.json";
$str = file_get_contents($filename);//開檔
/*
注意有檔頭會解不成功，有換行也會解不成功
*/
$str = str_replace("\r\n", "", $str);
$str = str_replace("\n", "", $str);
$str = preg_replace('/\s/', '', $str);

//print_r($str);
$array = json_decode($str,true);
echo "讀取".$filename.$br;
echo "開始解開JSON".$br;
//print_r($array);exit;

if (!is_array($array)){
	die('無法解開JSON');
}

/*
echo "<pre>";
print_r($array);
echo "</pre>";
*/
$savefile=$dir."/".pathinfo($filename, PATHINFO_EXTENSION).date("Y-m-d").".xlsx";
echo "存成檔案".$savefile.$br;
$flightsArray = $array;
// echo "<pre>";
// print_r($flightsArray);
// echo "</pre>";


$sortFlightsArray = array();
//第1行
$tempArray = array();
foreach($flightsArray[0] as $key =>$v){
$tempArray[$key]=$key;
}

$sortFlightsArray[0] = $tempArray;//第1行

for ($i=0; $i < count($flightsArray); $i++) {
	$f=$flightsArray[$i];

	$tempArray = array();
	foreach($f as $key =>$v){
	$tempArray[$key]=(string)$v;
	}
	$sortFlightsArray[$i+1] = $tempArray;
}


// echo "<pre>";
// print_r($sortFlightsArray);
// echo "</pre>";

//********************************************************************************

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

// Fill worksheet from values in array
$objPHPExcel->getActiveSheet()->fromArray($sortFlightsArray, null, 'A1');//取陣列

// Rename worksheet
$objPHPExcel->getActiveSheet()->setTitle('Members');

// Set AutoSize for name and email fields
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);

//設定背景顏色單色
$objPHPExcel->getActiveSheet(0)->getStyle('A1:I1')->applyFromArray(
	array('fill'     => array(
	'type' => PHPExcel_Style_Fill::FILL_SOLID,
	'color' => array('argb' => 'D1EEEE')
	),
	)
	);
/*
$objPHPExcel->getActiveSheet()->setCellValueExplicit('H1', (string)'0919636153',PHPExcel_Cell_DataType::TYPE_STRING); 
*/

// Save Excel 2007 file
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save($savefile);//存檔名稱

?>