<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

// Configuración de conexión a SQL Server
$serverName = "10.0.3.16, 1433"; 
$connectionOptions = [
    "Database" => "DBACOSAC_TEST",
    "Uid" => "sa",
    "PWD" => "4cf4rm4",
    "CharacterSet" => "UTF-8"
];

$conn = sqlsrv_connect($serverName, $connectionOptions);
if (!$conn) {
    die("❌ Error de conexión a SQL Server (DBACOSAC_TEST): " . print_r(sqlsrv_errors(), true));
}

// Configuración de conexión a DBACINHOUSE_TEST
$connectionOptionsInhouse = [
    "Database" => "DBACINHOUSE_TEST",
    "Uid" => "sa",
    "PWD" => "4cf4rm4",
    "CharacterSet" => "UTF-8"
];

$connInhouse = sqlsrv_connect($serverName, $connectionOptionsInhouse);
if (!$connInhouse) {
    die("❌ Error de conexión a SQL Server (DBACINHOUSE_TEST): " . print_r(sqlsrv_errors(), true));
}

// Cargar el archivo Excel
$archivoExcel = 'document/data17-02-25.xlsx';
$spreadsheet = IOFactory::load($archivoExcel);
$hoja = $spreadsheet->getActiveSheet();

// Obtener la última fila con datos
$maxFila = $hoja->getHighestRow();
$encabezados = $hoja->rangeToArray("A1:" . $hoja->getHighestColumn() . "1")[0];

// Insertar la cabecera "USUINIDUSUARIO" en la columna Q
array_splice($encabezados, 16, 0, "USUINIDUSUARIO");  // Coloca "USUINIDUSUARIO" en la columna Q

$filasErrores = [];
$filasErroresB = [];
$filasProcesadas = [];

// **Procesar cada fila**
for ($fila = 2; $fila <= $maxFila; $fila++) {
    $filaDatos = $hoja->rangeToArray("A$fila:" . $hoja->getHighestColumn() . $fila)[0];

    $valorB = trim($filaDatos[1] ?? '');
    $valorE = trim($filaDatos[4] ?? '');
    $valorF = trim($filaDatos[5] ?? '');
    $valorG = strtolower(trim($filaDatos[6] ?? ''));
    $valorL = trim($filaDatos[11] ?? '');
    $valorM = trim($filaDatos[12] ?? '');
    $valorO = trim($filaDatos[14] ?? '');

    // **Convertir Estado en columna G**
    if ($valorG === "activo") {
        $filaDatos[6] = 1;
    } elseif ($valorG === "inactivo") {
        $filaDatos[6] = 2;
    } elseif ($valorG === "baja contable") {
        $filaDatos[6] = 3;
    }

    // **Validar y reemplazar la columna B**
    $stmtB = sqlsrv_query($conn, "SELECT idtipo FROM TBOSA_TIPOS WHERE descripcion = ?", [$valorB]);
    if ($stmtB && $row = sqlsrv_fetch_array($stmtB, SQLSRV_FETCH_ASSOC)) {
        $filaDatos[1] = $row['idtipo'];
    } else {
        $filaConErrorB = $filaDatos;
        $filaConErrorB[] = "Fila $fila";
        $filasErroresB[] = $filaConErrorB;
    }

    // **Validar y reemplazar la columna E**
    $stmtE = sqlsrv_query($conn, "SELECT idCondicion FROM TBOSA_CONDICION WHERE condicion = ?", [$valorE]);
    if ($stmtE && $row = sqlsrv_fetch_array($stmtE, SQLSRV_FETCH_ASSOC)) {
        $filaDatos[4] = $row['idCondicion'];
    } else {
        $filasErrores[] = array_merge(["E$fila"], $filaDatos);
    }

    // **Validar y reemplazar la columna F**
    $stmtF = sqlsrv_query($connInhouse, "SELECT EMPINIDEMPRESA FROM TBSEGMAEEMPRESA WHERE EMPVCNOEMPRESA = ? OR EMPVCNOCORTO = ?", [$valorF, $valorF]);
    if ($stmtF && $row = sqlsrv_fetch_array($stmtF, SQLSRV_FETCH_ASSOC)) {
        $filaDatos[5] = $row['EMPINIDEMPRESA'];
    } else {
        $filasErrores[] = array_merge(["F$fila"], $filaDatos);
    }

    // **Validar y reemplazar la columna L**
    $stmtL = sqlsrv_query($connInhouse, "SELECT EMPINIDEMPRESA FROM TBSEGMAEEMPRESA WHERE EMPVCNOEMPRESA = ? OR EMPVCNOCORTO = ?", [$valorL, $valorL]);
    if ($stmtL && $row = sqlsrv_fetch_array($stmtL, SQLSRV_FETCH_ASSOC)) {
        $filaDatos[11] = $row['EMPINIDEMPRESA'];
    } else {
        $filasErrores[] = array_merge(["L$fila"], $filaDatos);
    }

    // **Validar y reemplazar la columna M**
    $stmtM = sqlsrv_query($connInhouse, "SELECT SUCINIDSUCURSAL FROM TBSEGMAESUCURSAL WHERE SUCINIDSUCURSAL != 23 AND EMPINIDEMPRESA = ?", [$filaDatos[11]]);
    if ($stmtM && $row = sqlsrv_fetch_array($stmtM, SQLSRV_FETCH_ASSOC)) {
        $filaDatos[12] = $row['SUCINIDSUCURSAL'];
    } else {
        $filasErrores[] = array_merge(["M$fila"], $filaDatos);
    }

    // **Obtener USUINIDUSUARIO en nueva columna P**
    $nuevoValorP = "";
    if ($valorO !== "SIN USUARIO") {
        $consultaSQL = "
            DECLARE @NombreCompleto NVARCHAR(255) = ?;
            DECLARE @SQLQuery NVARCHAR(MAX) = N'';

            SET @SQLQuery = N'SELECT USUINIDUSUARIO FROM TBSEGMAEUSUARIO WHERE ';

            DECLARE @Index INT = 1;
            DECLARE @Word NVARCHAR(50);
            DECLARE @TempNombre NVARCHAR(255) = @NombreCompleto + ' ';

            WHILE CHARINDEX(' ', @TempNombre, @Index) > 0
            BEGIN
                SET @Word = SUBSTRING(@TempNombre, @Index, CHARINDEX(' ', @TempNombre, @Index) - @Index);
                SET @Index = CHARINDEX(' ', @TempNombre, @Index) + 1;

                IF LEN(@Word) > 0
                    SET @SQLQuery = @SQLQuery + N'CHARINDEX(''' + @Word + ''', USUVCNOUSUARIO) > 0 AND ';
            END

            SET @SQLQuery = LEFT(@SQLQuery, LEN(@SQLQuery) - 4);

            EXEC sp_executesql @SQLQuery;
        ";

        $stmtP = sqlsrv_query($connInhouse, $consultaSQL, [$valorO]);
        if ($stmtP && $row = sqlsrv_fetch_array($stmtP, SQLSRV_FETCH_ASSOC)) {
            $nuevoValorP = $row['USUINIDUSUARIO'];
        } else {
            $filasErrores[] = array_merge(["P$fila"], $filaDatos);
        }
    }
    array_splice($filaDatos, 16, 0, $nuevoValorP);

    $filasProcesadas[] = $filaDatos;
}

// **Función para guardar archivos Excel**
function guardarExcel($nombre, $encabezado, $data) {
    if (!empty($data)) {
        $spreadsheet = new Spreadsheet();
        $hoja = $spreadsheet->getActiveSheet();
        $hoja->fromArray([$encabezado], null, 'A1');
        $hoja->fromArray($data, null, 'A2');
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save("document/$nombre-" . date("Y-m-d") . ".xlsx");
    }
}

// **Guardar archivos corregidos**
guardarExcel("EXCELDATASYNC", $encabezados, $filasProcesadas);
guardarExcel("TIPOS-ERRORES", array_merge($encabezados, ["Posición en EXCELDATASYNC"]), $filasErroresB);
guardarExcel("DATA-FINAL-ERRORES", array_merge(["Celda"], $encabezados), $filasErrores);

echo "✅ Proceso completado. Archivos generados en 'document/'";
?>
