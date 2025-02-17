<?php
// Leer el archivo tech-info.json
$techInfoFile = 'tech-info.json';
$techInfo = file_exists($techInfoFile) ? json_decode(file_get_contents($techInfoFile), true) : [];
?>

<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Información del Proyecto - ExcelDataSync</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://unpkg.com/alpinejs" defer></script>
</head>
<body class="bg-gray-100 p-5">

    <div class="max-w-4xl mx-auto bg-white p-6 rounded-lg shadow-md" x-data="techInfo">
        <h1 class="text-2xl font-bold text-gray-700">Información del Proyecto - ExcelDataSync</h1>
        
        <div class="mt-4 border-t border-gray-200 pt-4">
            <p><strong>Nombre del Proyecto:</strong> <span x-text="info.project_name"></span></p>
            <p><strong>Creado el:</strong> <span x-text="info.created_at"></span></p>
            <p><strong>Última Actualización:</strong> <span x-text="info.updated_at"></span></p>
            <p><strong>Sistema Operativo:</strong> <span x-text="info.os"></span></p>
        </div>

        <div class="mt-4 border-t border-gray-200 pt-4">
            <h2 class="text-xl font-semibold text-gray-600">PHP y Composer</h2>
            <p><strong>Versión PHP:</strong> <span x-text="info.php_version"></span></p>
            <p><strong>Versión Composer:</strong> <span x-text="info.composer_version"></span></p>
        </div>

        <div class="mt-4 border-t border-gray-200 pt-4">
            <h2 class="text-xl font-semibold text-gray-600">Git</h2>
            <p><strong>Usuario Git:</strong> <span x-text="info.git_user"></span></p>
            <p><strong>Email Git:</strong> <span x-text="info.git_email"></span></p>
            <p><strong>Último Commit:</strong> <span x-text="info.git_last_commit"></span></p>
        </div>

        <div class="mt-4 border-t border-gray-200 pt-4">
            <h2 class="text-xl font-semibold text-gray-600">Dependencias Composer</h2>
            <ul class="list-disc list-inside">
                <template x-for="(version, package) in info.composer_packages">
                    <li><span x-text="package"></span>: <span x-text="version"></span></li>
                </template>
            </ul>
        </div>

        <button class="mt-5 bg-blue-500 text-white px-4 py-2 rounded" @click="loadInfo()">
            🔄 Actualizar Información
        </button>
    </div>

    <script>
    document.addEventListener("alpine:init", () => {
        Alpine.data("techInfo", () => ({
            info: {},
            async loadInfo() {
                try {
                    let response = await fetch('http://localhost:81/ExcelDataSync/get-tech-info.php');
                    let text = await response.text();
                    console.log("📄 Respuesta del servidor:", text);

                    // Si la respuesta incluye HTML, hay un problema
                    if (text.trim().startsWith("<")) {
                        throw new Error("El servidor devolvió HTML en lugar de JSON.");
                    }

                    let json = JSON.parse(text);
                    this.info = json;
                } catch (error) {
                    console.error("❌ Error cargando el JSON:", error);
                }
            },
            init() {
                this.loadInfo();
            }
        }));
    });
</script>




</body>
</html>
