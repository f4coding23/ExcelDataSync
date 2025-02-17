<?php
$dataFile = 'tech-info.json';

// Si el archivo ya existe, cargarlo para mantener `created_at`
if (file_exists($dataFile)) {
    $existingData = json_decode(file_get_contents($dataFile), true);
    $createdAt = $existingData['created_at'] ?? date('Y-m-d H:i:s');
} else {
    $createdAt = date('Y-m-d H:i:s');
}

$data = [];

// Fecha de creación y última modificación
$data['created_at'] = $createdAt;
$data['updated_at'] = date('Y-m-d H:i:s');

// Obtener el nombre del proyecto desde la carpeta actual
$data['project_name'] = basename(getcwd());

// Función segura para ejecutar shell_exec() y evitar errores
function safeShellExec($command)
{
    $output = shell_exec($command);
    return $output !== null ? trim($output) : "No disponible";
}

// Comprobar si el directorio es un repositorio Git antes de ejecutar comandos Git
$isGitRepo = safeShellExec('git rev-parse --is-inside-work-tree 2>NUL');

if ($isGitRepo === "true") {
    $data['git_user'] = safeShellExec('git config user.name');
    $data['git_email'] = safeShellExec('git config user.email');
    $data['git_last_commit'] = safeShellExec('git log -1 --pretty=format:"%h - %s"');
} else {
    $data['git_user'] = "No disponible";
    $data['git_email'] = "No disponible";
    $data['git_last_commit'] = "No disponible";
}

// Obtener versión de PHP
$data['php_version'] = phpversion();

// Obtener versión de Composer sin caracteres ANSI
$data['composer_version'] = safeShellExec('composer --version --no-ansi');

// Obtener paquetes y versiones de Composer
$composerPackagesJson = safeShellExec('composer show --format=json --no-ansi');
$data['composer_packages'] = [];

if ($composerPackagesJson !== "No disponible") {
    $composerPackages = json_decode($composerPackagesJson, true);
    if (isset($composerPackages['installed'])) {
        foreach ($composerPackages['installed'] as $package) {
            $data['composer_packages'][$package['name']] = $package['version'];
        }
    }
}

// Obtener sistema operativo
$data['os'] = php_uname();

// Guardar en un archivo JSON con formato legible
file_put_contents($dataFile, json_encode($data, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES));

echo "Archivo tech-info.json actualizado correctamente.\n";
