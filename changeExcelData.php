<?php

require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/IOFactory.php';

function parse_excel_file( $filename ){
    // получаем тип файла (xls, xlsx), чтобы правильно его обработать
    $file_type = PHPExcel_IOFactory::identify( $filename );
    // создаем объект для чтения
    $objReader = PHPExcel_IOFactory::createReader( $file_type );
    $objPHPExcel = $objReader->load( $filename ); // загружаем данные файла в объект
    $result = $objPHPExcel->getActiveSheet()->toArray(); // выгружаем данные из объекта в массив

    return $result;
}

$filename = 'excel.xlsx';
if(file_exists($filename)){
    $res = parse_excel_file( $filename );
}


?>
<!--    $xls->setActiveSheetIndex(0)->setCellValue('A1', 'Hello world!')-->
<form method="post" action="/change_excle.php">

    <table>
        <tr>
            <th><input type="text" name="title_date" value="<?php echo $res[0][0] ?>"></th>
            <th><input type="text" name="title_name" value="<?php echo $res[0][1] ?>"></th>
            <th><input type="text" name="title_age" value="<?php echo $res[0][2] ?>"></th>
            <th><input type="text" name="title_login" value="<?php echo $res[0][3] ?>"></th>
            <th><input type="text" name="title_balans" value="<?php echo $res[0][4] ?>"></th>
        </tr>
        <?php for($i = 1; $i <count($res); $i++ ): ?>
            <tr>
                <th><input type="date" name="date[]" value="<?php echo $res[$i][0] ?>"></th>
                <th><input type="text" name="name[]" value="<?php echo $res[$i][1] ?>"></th>
                <th><input type="number" name="age[]" value="<?php echo $res[$i][2] ?>"></th>
                <th><input type="text" name="login[]" value="<?php echo $res[$i][3] ?>"></th>
                <th><input type="number" name="balans[]" value="<?php echo $res[$i][4] ?>"></th>
            </tr>
        <?php endfor ?>
    </table>
        <input type="submit" name="changeExcel" value="changeExcel">
    </form>
    <a href="/">Назад</a>

