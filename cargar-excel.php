<?php

require 'vendor/autoload.php'; // Asegúrate de instalar phpspreadsheet
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

$servername = "10.0.3.16, 1433"; 
$username = "sa";
$password = "4cf4rm4";
$dbname = "DBACOSAC_TEST";

$conn = new mysqli($servername, $username, $password, $dbname);
if ($conn->connect_error) {
    die("Conexión fallida: " . $conn->connect_error);
}

$file = 'document\data-osac-final.xlsx';
$spreadsheet = IOFactory::load($file);
$sheet = $spreadsheet->getActiveSheet();
$data = $sheet->toArray(null, true, true, true);

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
    $garantiaInicio = !empty($row['T']) ? DateTime::createFromFormat('d/m/Y', $row['T'])->format('Y-m-d') : null;
    $garantiaFin = !empty($row['U']) ? DateTime::createFromFormat('d/m/Y', $row['U'])->format('Y-m-d') : null;
    $propiedad = $row['V'];
    $numero = $row['W'];
    $duracion = $row['X'];

    $sql = "INSERT INTO TBOSA_EQUIPOS (codigo, nombre, idCondicion, idUsuario, idArea, idLocal, piso, proveedor, marca, modelo, serie, ip, session, codigoInventario, otros, idTipo, estado, password, mac, garantiaInicio, garantiaFin, propiedad, numero, duracion) 
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

    $stmt = $conn->prepare($sql);
    $stmt->bind_param("ssiiiiissssssssiiisssssssi", $codigo, $nombre, $idCondicion, $idUsuario, $idArea, $idLocal, $piso, $proveedor, $marca, $modelo, $serie, $ip, $session, $codigoInventario, $otros, $idTipo, $estado, $password, $mac, $garantiaInicio, $garantiaFin, $propiedad, $numero, $duracion);
    
    if (!$stmt->execute()) {
        echo "Error en la fila $key: " . $stmt->error . "\n";
    }
}

$stmt->close();
$conn->close();

echo "Carga completada.";
?>
