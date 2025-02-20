<?php

require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

// Configuración de la base de datos
$serverName = "10.0.3.16, 1433";
$connectionOptions = [
    "Database" => "DBACINHOUSE_TEST",
    "Uid" => "sa",
    "PWD" => "4cf4rm4",
    "CharacterSet" => "UTF-8"
];


$conn = sqlsrv_connect($serverName, $connectionOptions);
if ($conn === false) {
    die(print_r(sqlsrv_errors(), true));
}

// Cargar el archivo Excel
$inputFileName = "document/rol-osac-final.xlsx"; // Ruta del archivo de entrada
$spreadsheet = IOFactory::load($inputFileName);
$worksheet = $spreadsheet->getActiveSheet();
$rows = $worksheet->toArray();

$createdRecords = [];
$errorRecords = [];

foreach ($rows as $index => $row) {
    if ($index === 0) continue; // Saltar encabezados
    
    list($usuario, $rol, $sucursal) = $row;
    
    $query = "SELECT COUNT(*) AS count FROM TBSEGTBLUSUROL WHERE USUINIDUSUARIO = ? AND ROLINIDROL = ? AND SUCINIDSUCURSAL = ?";
    $params = [$usuario, $rol, $sucursal];
    $stmt = sqlsrv_query($conn, $query, $params);
    
    if ($stmt === false) {
        $errorRecords[] = array_merge($row, ["Error en la consulta"]);
        continue;
    }
    
    $rowCount = sqlsrv_fetch_array($stmt, SQLSRV_FETCH_ASSOC)['count'];
    
    if ($rowCount == 0) {
        // Insertar el nuevo registro
        $insertQuery = "INSERT INTO TBSEGTBLUSUROL (USUINIDUSUARIO, ROLINIDROL, SUCINIDSUCURSAL) VALUES (?, ?, ?)";
        $insertStmt = sqlsrv_query($conn, $insertQuery, $params);
        
        if ($insertStmt) {
            $createdRecords[] = $row;
        } else {
            $errorRecords[] = array_merge($row, ["Error al insertar"]);
        }
    }
}

// Guardar archivo con registros creados
if (!empty($createdRecords)) {
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->fromArray([['USUINIDUSUARIO', 'ROLINIDROL', 'SUCINIDSUCURSAL']], NULL, 'A1');
    $sheet->fromArray($createdRecords, NULL, 'A2');
    $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
    $writer->save("document/CREATE-USUROL-" . date('Y-m-d') . ".xlsx");
}

// Guardar archivo de errores
if (!empty($errorRecords)) {
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->fromArray([['USUINIDUSUARIO', 'ROLINIDROL', 'SUCINIDSUCURSAL', 'Error']], NULL, 'A1');
    $sheet->fromArray($errorRecords, NULL, 'A2');
    $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
    $writer->save("document/ERROR-USUROL-" . date('Y-m-d') . ".xlsx");
}

sqlsrv_close($conn);

echo "Proceso completado.";
