<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

function leerExcel($archivo)
{
    // Cargar el archivo Excel
    $spreadsheet = IOFactory::load($archivo);
    $hoja = $spreadsheet->getActiveSheet();

    // Obtener el número máximo de filas y columnas
    $maxFila = $hoja->getHighestRow();
    $maxColumna = $hoja->getHighestColumn();

    // Recorrer las filas
    echo "<table border='1'>";
    for ($fila = 1; $fila <= $maxFila; $fila++) {
        echo "<tr>";
        for ($columna = 'A'; $columna <= $maxColumna; $columna++) {
            $valor = $hoja->getCell($columna . $fila)->getValue();
            echo "<td>$valor</td>";
        }
        echo "</tr>";
    }
    echo "</table>";
}

// Llamar a la función con el archivo de prueba
$archivoExcel = 'document\data17-02-25.xlsx'; // Asegúrate de que este archivo esté en la misma carpeta
leerExcel($archivoExcel);
