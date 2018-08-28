<?php

require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/IOFactory.php';

$filename = 'excel.xlsx';
if(file_exists($filename)){
    $file_type = PHPExcel_IOFactory::identify( $filename );
    $objReader = PHPExcel_IOFactory::createReader( $file_type );
    $objPHPExcel = $objReader->load( $filename );
    $allPhpExcelData = $objPHPExcel->getActiveSheet()->toArray();

    $worksheetData = $objReader->listWorksheetInfo($filename);
    $totalRows     = $worksheetData[0]['totalRows'];
    $totalColumns  = $worksheetData[0]['totalColumns'];
   // $i = $totalRows + 1;
}else{
    $title = array(
        array(
            'name' => 'Дата',
            'cell' => 'A'
        ),
        array(
            'name' => 'Имя',
            'cell' => 'B'
        ),
        array(
            'name' => 'Возраст',
            'cell' => 'C'
        ),
        array(
            'name' => 'Логин',
            'cell' => 'D'
        ),
        array(
            'name' => 'Баланс',
            'cell' => 'E'
        ),
    );

    $i = 2;
}


if(isset($_POST['send'])){
   $form_data = array(
        array(
            '0' => $_POST['date'],
            '1' => $_POST['name'],
            '2' => $_POST['age'],
            '3' => $_POST['login'],
            '4' => $_POST['balans']
        ),
   );
}

$phpexcel = new PHPExcel();

if(!file_exists($filename)) {
    for ($j = 0; $j < count($title); $j++) {
        $string = $title[$j]['name'];
        $cellLetter = $title[$j]['cell'] . 1;
        $phpexcel->getActiveSheet()->setCellValueExplicit($cellLetter, $string, PHPExcel_Cell_DataType::TYPE_STRING);
    }

    foreach ($form_data as $row) {
        $date = new DateTime($row['0']);
        $date = $date->format('d.m.Y');
        $phpexcel->getActiveSheet()->setCellValueExplicit("A$i", $date, PHPExcel_Cell_DataType::TYPE_STRING);
        $string = $row['1'];
        //$string = mb_convert_encoding($string, 'UTF-8', 'Windows-1251');
        $phpexcel->getActiveSheet()->setCellValueExplicit("B$i", $string, PHPExcel_Cell_DataType::TYPE_STRING);
        $phpexcel->getActiveSheet()->setCellValue("C$i", $row['2']); // setCellValue для чисел
        $string = $row['3'];
        //$string = mb_convert_encoding($string, 'UTF-8', 'Windows-1251');
        $phpexcel->getActiveSheet()->setCellValueExplicit("D$i", $string, PHPExcel_Cell_DataType::TYPE_STRING);
        $phpexcel->getActiveSheet()->setCellValue("E$i", $row['4']);
        $i++;
    }
}else{
    $arr  = array_merge($allPhpExcelData, $form_data);

    $i = 0;
    $cell = 1;
    foreach($arr as $row_data){
        //var_dump($row_data);
//        if($i == 0 ){
//            $phpexcel->getActiveSheet()->setCellValueExplicit("A$cell", $row_data[0], PHPExcel_Cell_DataType::TYPE_STRING);
//        }else{
//            $date = new DateTime($row_data[0]);
//            $date = $date->format('d:m:Y');

            $phpexcel->getActiveSheet()->setCellValueExplicit("A$cell", $row_data[0], PHPExcel_Cell_DataType::TYPE_STRING);
//        }
               $string = $row_data[1];
               $phpexcel->getActiveSheet()->setCellValueExplicit("B$cell", $string, PHPExcel_Cell_DataType::TYPE_STRING);
               $phpexcel->getActiveSheet()->setCellValue("C$cell", $row_data[2]);
               $string = $row_data[3];
               $phpexcel->getActiveSheet()->setCellValueExplicit("D$cell", $string, PHPExcel_Cell_DataType::TYPE_STRING);
               $phpexcel->getActiveSheet()->setCellValue("E$cell", $row_data[4]);
        $i++;
        $cell++;
    }


//    $i = 2;
//    for($index = 0; $index < count($arr); $index++){
//       if($index == 0){
//           $phpexcel->getActiveSheet()->setCellValueExplicit("A1", $arr[0][0], PHPExcel_Cell_DataType::TYPE_STRING);
//           $phpexcel->getActiveSheet()->setCellValueExplicit("B1", $arr[0][1], PHPExcel_Cell_DataType::TYPE_STRING);
//           $phpexcel->getActiveSheet()->setCellValueExplicit("C1", $arr[0][2], PHPExcel_Cell_DataType::TYPE_STRING);
//           $phpexcel->getActiveSheet()->setCellValueExplicit("D1", $arr[0][3], PHPExcel_Cell_DataType::TYPE_STRING);
//           $phpexcel->getActiveSheet()->setCellValueExplicit("E1", $arr[0][4], PHPExcel_Cell_DataType::TYPE_STRING);
//       }else{
//               $date = new DateTime($arr[$index][0]);
//               $date = $date->format('d.m.Y');
//               $phpexcel->getActiveSheet()->setCellValueExplicit("A$i", $date, PHPExcel_Cell_DataType::TYPE_STRING);
//               $string = $arr[$index][1];
//               $phpexcel->getActiveSheet()->setCellValueExplicit("B$i", $string, PHPExcel_Cell_DataType::TYPE_STRING);
//               $phpexcel->getActiveSheet()->setCellValue("C$i", $arr[$index][2]);
//               $string = $arr[$index][3];
//               $phpexcel->getActiveSheet()->setCellValueExplicit("D$i", $string, PHPExcel_Cell_DataType::TYPE_STRING);
//               $phpexcel->getActiveSheet()->setCellValue("E$i", $arr[$index][4]);
//           $i++;
//       }
//    }

}
    $phpexcel->getActiveSheet()->getColumnDimension('A')->setWidth(16);
    $phpexcel->getActiveSheet()->getColumnDimension('B')->setWidth(16);
    $phpexcel->getActiveSheet()->getColumnDimension('C')->setWidth(16);
    $phpexcel->getActiveSheet()->getColumnDimension('D')->setWidth(16);
    $phpexcel->getActiveSheet()->getColumnDimension('E')->setWidth(16);
    $page = $phpexcel->setActiveSheetIndex();
    $page->setTitle('first-list');
    $objWriter = PHPExcel_IOFactory::createWriter($phpexcel, 'Excel2007');
    if (file_exists($filename)) {
        unlink($filename);
    }
    $objWriter->save($filename);


header ("Location: ".$_SERVER['HTTP_REFERER']);

//    $phpexcel->getActiveSheet()->setCellValueExplicit($cellLetter, $string, PHPExcel_Cell_DataType::TYPE_STRING); // 1) ячейка 2) дани 3) строка // getActiveSheet -> текущий лист // setCellValueExplicit - вставить данные в ячейку
?>