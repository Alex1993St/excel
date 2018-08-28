<?php

require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/IOFactory.php';

$filename = 'excel.xlsx';

if(isset($_POST['changeExcel'])){
    $arr_for = [];
    $arr_for_prepear = [];
    $arr =[
        [
            '0' => $_POST['title_date'],
            '1' => $_POST['title_name'],
            '2' => $_POST['title_age'],
            '3' => $_POST['title_login'],
            '4' => $_POST['title_balans'],
        ]
    ];
    for($i = 0; $i < count($_POST['date']); $i++){

        $arr_for_prepear = [
            [
                '0' =>  $_POST['date'][$i],
                '1' => $_POST['name'][$i],
                '2' => $_POST['age'][$i],
                '3' => $_POST['login'][$i],
                '4' => $_POST['balans'][$i]
            ]
        ];

        //array_push($arr_for, $_POST['date'][$i], $_POST['name'][$i], $_POST['age'][$i], $_POST['login'][$i], $_POST['balans'][$i]);
        $arr_for = array_merge($arr_for, $arr_for_prepear);




    }

    $all_array = array_merge($arr, $arr_for);

    $php_excel = new PHPExcel();

    $cell = 1;
    foreach ($all_array as $row){
        $php_excel->getActiveSheet()->setCellValueExplicit("A$cell", $row[0], PHPExcel_Cell_DataType::TYPE_STRING);
        $php_excel->getActiveSheet()->setCellValueExplicit("B$cell", $row[1], PHPExcel_Cell_DataType::TYPE_STRING);
        $php_excel->getActiveSheet()->setCellValue("C$cell", $row[2]);
        $php_excel->getActiveSheet()->setCellValueExplicit("D$cell", $row[3], PHPExcel_Cell_DataType::TYPE_STRING);
        $php_excel->getActiveSheet()->setCellValue("E$cell", $row[4]);
        $cell++;
    }

    $php_excel->getActiveSheet()->getColumnDimension("A")->setWidth(16);
    $php_excel->getActiveSheet()->getColumnDimension("B")->setWidth(16);
    $php_excel->getActiveSheet()->getColumnDimension("C")->setWidth(16);
    $php_excel->getActiveSheet()->getColumnDimension("D")->setWidth(16);
    $php_excel->getActiveSheet()->getColumnDimension("E")->setWidth(16);
    $page = $php_excel->setActiveSheetIndex();
    $page->setTitle('first_page');
    $obj_write = PHPExcel_IOFactory::createWriter($php_excel, 'Excel2007');
    if (file_exists($filename)) {
        unlink($filename);
    }
    $obj_write->save($filename);

    header ("Location: /changeExcelData.php");

//        $pExcel = PHPExcel_IOFactory::createReader('Excel2007');
//        $pExcel = $pExcel->load('excel.xlsx');
//        $pExcel->setActiveSheetIndex(0);
//        $aSheets = $pExcel->getActiveSheet();
//        $aSheets->setCellValue('B8', '123');
//        $objWriter = PHPExcel_IOFactory::createWriter($pExcel, 'Excel2007');
//        $file = "excel.xlsx";
//        $objWriter->save($file);
}

?>