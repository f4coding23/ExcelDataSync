<?php

require 'vendor/autoload.php'; // Asegúrate de instalar phpspreadsheet
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Shared\Date;


$serverName = "10.0.3.16, 1433";
$connectionOptions = [
    "Database" => "DBACOSAC_TEST",
    "Uid" => "sa",
    "PWD" => "4cf4rm4",
    "CharacterSet" => "UTF-8"
];


$conn = sqlsrv_connect($serverName, $connectionOptions);
if ($conn === false) {
    die(print_r(sqlsrv_errors(), true));
}

$file = 'document\data-osac-final.xlsx';
$spreadsheet = IOFactory::load($file);
$sheet = $spreadsheet->getActiveSheet();
$data = $sheet->toArray(null, true, true, true);

function safe_strlen($value)
{
    return !is_null($value) ? strlen($value) : 0;
}

foreach ($data as $key => $row) {
    if ($key == 1) continue; // Omitir encabezados

    $codigo = $row['A'];
    $nombre = $row['B'];
    $idCondicion = $row['C'];
    $idUsuario = $row['D'];
    $idArea = $row['E'];
    $idLocal = $row['F'];
    $piso = $row['G'];
    $proveedor = $row['H'];
    $marca = $row['I'];
    $modelo = $row['J'];
    $serie = $row['K'];
    $ip = $row['L'];
    $session = $row['M'];
    $codigoInventario = $row['N'];
    $otros = $row['O'];
    $idTipo = $row['P'];
    $estado = $row['Q'];
    $password = $row['R'];
    $mac = $row['S'];

    $convertirFecha = fn($valor, $esFechaExcel) =>
    empty($valor) ? null : (is_numeric($valor) && $esFechaExcel ? Date::excelToDateTimeObject((float)$valor)->format('Y-m-d') : (preg_match('/^\d{4}-\d{2}-\d{2}$/', $valor) ? $valor : (DateTime::createFromFormat('d/m/Y', trim($valor))?->format('Y-m-d') ?? null)));

    $garantiaInicio = $convertirFecha($row['T'], Date::isDateTime($sheet->getCell('T' . $key)));
    $garantiaFin = $convertirFecha($row['U'], Date::isDateTime($sheet->getCell('U' . $key)));

    $propiedad = $row['V'];
    $numero = $row['W'];
    $duracion = $row['X'];


    $sql = "INSERT INTO TBOSA_EQUIPOS (codigo, nombre, idCondicion, idUsuario, idArea, idLocal, piso, proveedor, marca, modelo, serie, ip, session, codigoInventario, otros, idTipo, estado, password, mac, garantiaInicio, garantiaFin, propiedad, numero, duracion) 
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

    $params = [$codigo, $nombre, $idCondicion, $idUsuario, $idArea, $idLocal, $piso, $proveedor, $marca, $modelo, $serie, $ip, $session, $codigoInventario, $otros, $idTipo, $estado, $password, $mac, $garantiaInicio, $garantiaFin, $propiedad, $numero, $duracion];

    $stmt = sqlsrv_query($conn, $sql, $params);
    if ($stmt === false) {
        echo "Fila $key -> Código: " . safe_strlen($row['A']) . ", Nombre: " . safe_strlen($row['B']) . ", Proveedor: " . safe_strlen($proveedor) . ", Marca: " . safe_strlen($marca) . ", Modelo: " . safe_strlen($modelo) . ", IP: " . safe_strlen($ip) . ", Código Inventario: " . safe_strlen($codigoInventario) . ", Password: " . safe_strlen($password) . ", MAC: " . safe_strlen($mac) . "\n";
        echo "Error en la fila $key: " . print_r(sqlsrv_errors(), true) . "\n";
    }
}

sqlsrv_close($conn);

echo "Carga completada.";
