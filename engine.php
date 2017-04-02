<?php

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
include 'apis/PHPExcel/Classes/PHPExcel.php';

readExcel();

function readExcel() {

    $objReader = PHPExcel_IOFactory::createReader('Excel2007');
    $objReader->setReadDataOnly(true);

    $objPHPExcel = $objReader->load("PreciosAbril3.xlsx");
    $objWorksheet = $objPHPExcel->getActiveSheet();

    $highestRow = $objWorksheet->getHighestRow();
    $highestColumn = $objWorksheet->getHighestColumn();

    $highestColumnIndex = PHPExcel_Cell::columnIndexFromString($highestColumn);

    echo '<table border="1">' . "\n";
    for ($row = 1; $row <= $highestRow; ++$row) {

//  for ($col = 0; $col <= $highestColumnIndex; ++$col) {
        $model = $objWorksheet->getCellByColumnAndRow(4, $row)->getValue();
        $newPrice = $objWorksheet->getCellByColumnAndRow(8, $row)->getValue();
//        if($col==7){
        if(!strlen($model) > 0)
            return 0;
//        echo "<br><br>Buscando: ".$model;
                searh($model, $newPrice);
//        }
//        else{
//            echo " $col: $model ";
//        }
  }
        sleep(.3);
//    }

    echo '</table>' . "\n";
}

function searh($model, $newPrice) {
    $servername = "127.0.0.1";
    $username = "root";
    $password = "root";
    $db = "miele_020417";

// Create connection
    $conn = new mysqli($servername, $username, $password, $db);

// Check connection
    if ($conn->connect_error) {
        die("Connection failed: " . $conn->connect_error);
    }

    $sql = "SELECT * FROM productos where modelo='$model';";
    
    $result = $conn->query($sql);

    if ($result->num_rows > 0) {
        // output data of each row
        while ($row = $result->fetch_assoc()) {
            echo "<br>Encontrado: id: " . $row["id"] . " - Model: " . $row["modelo"]." oldPrice:". $row["precio"] ."    newPrice: $newPrice<br>";
            $id = $row["id"];
            $update = "UPDATE productos SET precio = $newPrice where id = $id";
            
            if(!$conn->query($update))
                echo "Error al actualizar producto";
            else
                echo "Producto actualizado";
        }
    } else {
    }
    $conn->close();
}
