<!DOCTYPE html>
<html>
<head>
    <title></title>
</head>

<body>
    <form method="post" action="addExcelData.php">
        <input type="submit" name="Excel" value="Excel">
    </form>




<?php
require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/IOFactory.php';

function parse_excel_file( $filename ){
    $file_type = PHPExcel_IOFactory::identify( $filename );
    $objReader = PHPExcel_IOFactory::createReader( $file_type );
    $objPHPExcel = $objReader->load( $filename );
    $result = $objPHPExcel->getActiveSheet()->toArray();

    return $result;
}

$filename = 'excel.xlsx';
if(file_exists($filename)){
    $res = parse_excel_file( $filename );
}



?>
<form action="/phpToExcel.php" method="post">
    <input type="date" name="date">
    <input type="text" name="name" placeholder="Name">
    <input type="number" name="age" placeholder="Age">
    <input type="text" name="login" placeholder="Login">
    <input type="number" name="balans" placeholder="Balans">
    <input type="submit" name="send" value="Send">
</form>



<table>
    <tr>
        <th><?php echo $res[0][0] ?></th>
        <th><?php echo $res[0][1] ?></th>
        <th><?php echo $res[0][2] ?></th>
        <th><?php echo $res[0][3] ?></th>
        <th><?php echo $res[0][4] ?></th>
    </tr>
    <?php for($i = 1; $i <count($res); $i++ ): ?>
    <tr>
        <th><?php echo $res[$i][0] ?></th>
        <th><?php echo $res[$i][1] ?></th>
        <th><?php echo $res[$i][2] ?></th>
        <th><?php echo $res[$i][3] ?></th>
        <th><?php echo $res[$i][4] ?></th>
    </tr>
    <?php endfor ?>
</table>

<a href="/changeExcelData.php">Изменить</a>


</body>
</html>