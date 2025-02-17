<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

// Configuración de conexión a SQL Server para DBACOSAC_TEST (Tipos, Condiciones y Estado)
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

// Configuración de conexión a SQL Server para DBACINHOUSE_TEST (Empresas)
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
$valoresTipos = [];
$valoresCondicion = [];
$valoresEmpresa = [];
$valoresEstado = [];
$filasProcesadas = [];

// Leer las columnas B, E, F y G desde la fila 2
for ($fila = 2; $fila <= $maxFila; $fila++) {
    $valorB = strtolower(trim($hoja->getCell("B" . $fila)->getValue()));
    $valorE = strtolower(trim($hoja->getCell("E" . $fila)->getValue()));
    $valorF = strtolower(trim($hoja->getCell("F" . $fila)->getValue()));
    $valorG = strtolower(trim($hoja->getCell("G" . $fila)->getValue()));

    if (!empty($valorB)) $valoresTipos[$fila] = $valorB;
    if (!empty($valorE)) $valoresCondicion[$fila] = $valorE;
    if (!empty($valorF)) $valoresEmpresa[$fila] = $valorF;
    if (!empty($valorG)) $valoresEstado[$fila] = $valorG;
}

// **Consulta SQL para validar la columna B (Tipos de equipo)**
$sqlB = "SELECT LOWER(descripcion) AS descripcion_normalizada, idtipo FROM [dbo].[TBOSA_TIPOS] WHERE LOWER(descripcion) IN (" . implode(",", array_fill(0, count($valoresTipos), "?")) . ")";
$stmtB = sqlsrv_query($conn, $sqlB, array_values($valoresTipos));

$valoresBDTipos = [];
while ($row = sqlsrv_fetch_array($stmtB, SQLSRV_FETCH_ASSOC)) {
    $valoresBDTipos[$row['descripcion_normalizada']] = $row['idtipo'];
}

// **Consulta SQL para validar la columna E (Condiciones)**
$sqlE = "SELECT LOWER(condicion) AS condicion_normalizada, idCondicion FROM [dbo].[TBOSA_CONDICION] WHERE LOWER(condicion) IN (" . implode(",", array_fill(0, count($valoresCondicion), "?")) . ")";
$stmtE = sqlsrv_query($conn, $sqlE, array_values($valoresCondicion));

$valoresBDCondicion = [];
while ($row = sqlsrv_fetch_array($stmtE, SQLSRV_FETCH_ASSOC)) {
    $valoresBDCondicion[$row['condicion_normalizada']] = $row['idCondicion'];
}

// **Consulta SQL para validar la columna F (Empresa) en DBACINHOUSE_TEST**
$sqlF = "SELECT E.EMPINIDEMPRESA, LOWER(E.EMPVCNOEMPRESA) AS empresa_normalizada, LOWER(E.EMPVCNOCORTO) AS empresa_corto 
         FROM [DBACINHOUSE_TEST].[dbo].[TBSEGMAESISTEMA] SIS
         INNER JOIN [DBACINHOUSE_TEST].[dbo].[TBLSEGTBLSISSUC] SS ON SIS.SISINIDSISTEMA = SS.SISINIDSISTEMA
         INNER JOIN [DBACINHOUSE_TEST].[dbo].[TBSEGMAESUCURSAL] SUC ON SS.SUCINIDSUCURSAL = SUC.SUCINIDSUCURSAL
         INNER JOIN [DBACINHOUSE_TEST].[dbo].[TBSEGMAEEMPRESA] E ON SUC.EMPINIDEMPRESA = E.EMPINIDEMPRESA
         WHERE SIS.SISCHCDSISTEMA = 'OSA'
         AND (LOWER(E.EMPVCNOEMPRESA) IN (" . implode(",", array_fill(0, count($valoresEmpresa), "?")) . ")
         OR LOWER(E.EMPVCNOCORTO) IN (" . implode(",", array_fill(0, count($valoresEmpresa), "?")) . "))";
$stmtF = sqlsrv_query($connInhouse, $sqlF, array_merge(array_values($valoresEmpresa), array_values($valoresEmpresa)));

$valoresBDEmpresa = [];
while ($row = sqlsrv_fetch_array($stmtF, SQLSRV_FETCH_ASSOC)) {
    $valoresBDEmpresa[$row['empresa_normalizada']] = $row['EMPINIDEMPRESA'];
    $valoresBDEmpresa[$row['empresa_corto']] = $row['EMPINIDEMPRESA'];
}

// **Procesar y modificar EXCELDATASYNC con las correcciones**
foreach ($valoresTipos as $filaOriginal => $valorB) {
    $filaProcesada = $hoja->rangeToArray("A$filaOriginal:" . $hoja->getHighestColumn() . $filaOriginal)[0];

    // Reemplazar columna B con idtipo
    if (isset($valoresBDTipos[$valorB])) {
        $filaProcesada[1] = $valoresBDTipos[$valorB];
    }

    // Reemplazar columna E con idCondicion
    if (isset($valoresBDCondicion[$valoresCondicion[$filaOriginal]])) {
        $filaProcesada[4] = $valoresBDCondicion[$valoresCondicion[$filaOriginal]];
    }

    // Reemplazar columna F con EMPINIDEMPRESA
    if (isset($valoresBDEmpresa[$valoresEmpresa[$filaOriginal]])) {
        $filaProcesada[5] = $valoresBDEmpresa[$valoresEmpresa[$filaOriginal]];
    }

    // Reemplazar columna G con estado convertido (1 o 2)
    $estadoConvertido = ($valoresEstado[$filaOriginal] === "activo") ? 1 : 2;
    $filaProcesada[6] = $estadoConvertido;

    $filasProcesadas[] = $filaProcesada;
}

// **Guardar archivo EXCELDATASYNC**
$fechaActual = date("Y-m-d");
$excelProcesado = new Spreadsheet();
$hojaProcesado = $excelProcesado->getActiveSheet();
$hojaProcesado->fromArray([$hoja->rangeToArray("A1:" . $hoja->getHighestColumn() . "1")[0]], null, 'A1');
$hojaProcesado->fromArray($filasProcesadas, null, 'A2');
$writerProcesado = IOFactory::createWriter($excelProcesado, 'Xlsx');
$archivoProcesado = "document/EXCELDATASYNC-$fechaActual.xlsx";
$writerProcesado->save($archivoProcesado);

echo "✅ Proceso completado. Archivos generados en 'document/'";
?>
