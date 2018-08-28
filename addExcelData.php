<?php

require 'Classes/PHPExcel.php';
require 'Classes/PHPExcel/IOFactory.php';

$filename = 'excel.xlsx';

$serv = 'localhost';
$name = 'root';
$password = '';

$arr = [
    [
        '0' => 'Дата',
        '1' => 'Имя',
        '2' => 'Возраст',
        '3' => 'Логин',
        '4' => 'Баланс'
    ]
];

try{
    $conn = new PDO("mysql:host=$serv;dbname=excel", $name, $password);
    $conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    $stmt = $conn->prepare("SELECT * FROM excel_data");
    $stmt->execute();
    $result = $stmt->setFetchMode(PDO::FETCH_ASSOC);
    foreach(new RecursiveArrayIterator($stmt->fetchAll()) as $v) {
        $arr_for = [
            [
                '0' => $v['data'],
                '1' => $v['name'],
                '2' => $v['age'],
                '3' => $v['login'],
                '4' => $v['balans'],
            ]
        ];
        $arr = array_merge($arr, $arr_for);

    }
}catch (PDOException $e){
    echo 'Connection failed' .$e->getMessage();
}


$php_excel = new PHPExcel();

$i = 0;
$cell = 1;
foreach ($arr as $row){
//    if($i == 0){
//        $php_excel -> getActiveSheet()->setCellValueExplicit("A$cell", $row[0], PHPExcel_Cell_DataType::TYPE_STRING);
//    }else{
//        $date = new DateTime($row[0]);
//        $date = $date->format('d:m:Y');
//        $php_excel -> getActiveSheet()->setCellValueExplicit("A$cell", $date, PHPExcel_Cell_DataType::TYPE_STRING);
//    }
    $php_excel -> getActiveSheet()->setCellValueExplicit("A$cell", $row[0], PHPExcel_Cell_DataType::TYPE_STRING);
    $php_excel->getActiveSheet()->setCellValueExplicit("B$cell", $row[1], PHPExcel_Cell_DataType::TYPE_STRING);
    $php_excel->getActiveSheet()->setCellValue("C$cell", $row[2]);
    $php_excel->getActiveSheet()->setCellValueExplicit("D$cell", $row[3], PHPExcel_Cell_DataType::TYPE_STRING);
    $php_excel->getActiveSheet()->setCellValue("E$cell", $row[4]);
    $i++; $cell++;
}

$php_excel->getActiveSheet()->getColumnDimension("A")->setWidth(16);
$php_excel->getActiveSheet()->getColumnDimension("B")->setWidth(16);
$php_excel->getActiveSheet()->getColumnDimension("C")->setWidth(16);
$php_excel->getActiveSheet()->getColumnDimension("D")->setWidth(16);
$php_excel->getActiveSheet()->getColumnDimension("E")->setWidth(16);
$page = $php_excel->setActiveSheetIndex();
$page->setTitle('first_page');
$obj_write = PHPExcel_IOFactory::createWriter($php_excel, 'Excel2007');
$obj_write->save($filename);

header ("Location: ".$_SERVER['HTTP_REFERER']);

?>