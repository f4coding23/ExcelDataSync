<?php
header('Content-Type: application/json');

// Definir la ruta completa de tech-info.json
$techInfoFile = __DIR__ . '/tech-info.json';

// Si el archivo no existe, devuelve un error en JSON
if (!file_exists($techInfoFile)) {
    echo json_encode(["error" => "Archivo tech-info.json no encontrado"]);
    exit;
}

// Leer y devolver el contenido del JSON
$jsonContent = file_get_contents($techInfoFile);
$jsonData = json_decode($jsonContent, true);

// Si hay un error al decodificar JSON, devuelve un error
if ($jsonData === null) {
    echo json_encode(["error" => "Error al decodificar tech-info.json"]);
    exit;
}

echo json_encode($jsonData, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES);
exit;
