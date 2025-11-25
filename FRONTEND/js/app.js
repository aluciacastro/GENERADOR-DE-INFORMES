// ===================================
// CONFIGURACIÓN
// ===================================

const API_URL = '/api';


// Estado de la aplicación
let selectedFiles = [];
let generatedFiles = [];

// Elementos del DOM
const uploadForm = document.getElementById('uploadForm');
const excelFileInput = document.getElementById('excelFile');
const fileListDiv = document.getElementById('fileList');
const downloadLinksDiv = document.getElementById('downloadLinks');
const loadedCount = document.getElementById('loadedCount');
const generatedCount = document.getElementById('generatedCount');
const generateBtn = document.getElementById('generateBtn');
const clearBtn = document.getElementById('clearBtn');

// ===================================
// FILE INPUTS
// ===================================

/**
 * Inicializa los inputs de archivo
 */
function initializeFileInputs() {
    // Excel files (multiple)
    excelFileInput.addEventListener('change', (e) => {
        selectedFiles = Array.from(e.target.files);
        updateFileList();
    });
}

/**
 * Actualiza la lista visual de archivos seleccionados en la columna izquierda
 */
function updateFileList() {
    // Actualizar contador
    loadedCount.textContent = `${selectedFiles.length} archivo${selectedFiles.length !== 1 ? 's' : ''}`;
    
    if (selectedFiles.length === 0) {
        fileListDiv.classList.add('empty-state');
        fileListDiv.innerHTML = `
            <div class="empty-icon">📭</div>
            <p class="empty-text">No hay archivos cargados</p>
            <p class="empty-hint">Arrastra o selecciona archivos Excel</p>
        `;
        generateBtn.disabled = true;
        return;
    }

    fileListDiv.classList.remove('empty-state');
    fileListDiv.innerHTML = '';
    generateBtn.disabled = false;

    selectedFiles.forEach((file, index) => {
        const fileItem = document.createElement('div');
        fileItem.className = 'file-item-col';
        fileItem.dataset.index = index;
        fileItem.innerHTML = `
            <span class="file-item-col-icon">📊</span>
            <div class="file-item-col-info">
                <div class="file-item-col-name" title="${file.name}">${file.name}</div>
                <div class="file-item-col-meta">${formatFileSize(file.size)}</div>
            </div>
            <button type="button" class="file-item-col-action" data-index="${index}">✕</button>
        `;
        fileListDiv.appendChild(fileItem);
    });

    // Agregar event listeners a botones de eliminar
    document.querySelectorAll('.file-item-col-action').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const index = parseInt(e.target.dataset.index);
            removeFile(index);
        });
    });
}

/**
 * Elimina un archivo de la lista
 */
function removeFile(index) {
    selectedFiles.splice(index, 1);
    
    // Actualizar el input
    const dataTransfer = new DataTransfer();
    selectedFiles.forEach(file => dataTransfer.items.add(file));
    excelFileInput.files = dataTransfer.files;
    
    updateFileList();
}

/**
 * Actualiza la columna de documentos generados
 */
function updateDownloadLinks() {
    // Actualizar contador
    generatedCount.textContent = `${generatedFiles.length} documento${generatedFiles.length !== 1 ? 's' : ''}`;
    
    if (generatedFiles.length === 0) {
        downloadLinksDiv.classList.add('empty-state');
        downloadLinksDiv.innerHTML = `
            <div class="empty-icon">📝</div>
            <p class="empty-text">Sin documentos generados</p>
            <p class="empty-hint">Genera informes para verlos aquí</p>
        `;
        return;
    }

    downloadLinksDiv.classList.remove('empty-state');
    downloadLinksDiv.innerHTML = '';

    generatedFiles.forEach((file, index) => {
        const downloadItem = document.createElement('div');
        downloadItem.className = 'download-item-col';
        downloadItem.innerHTML = `
            <span class="download-item-col-icon">📄</span>
            <div class="download-item-col-info">
                <div class="download-item-col-name" title="${file.filename}">${file.filename}</div>
                <div class="download-item-col-meta">${formatFileSize(file.size)}</div>
            </div>
            <button class="download-item-col-btn" data-index="${index}">
                <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                    <polyline points="7 10 12 15 17 10"/>
                    <line x1="12" y1="15" x2="12" y2="3"/>
                </svg>
                Descargar
            </button>
        `;
        downloadLinksDiv.appendChild(downloadItem);
    });

    // Agregar event listeners a botones de descarga
    document.querySelectorAll('.download-item-col-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const index = parseInt(e.target.closest('button').dataset.index);
            downloadFile(index);
        });
    });
}

// ===================================
// FORM SUBMISSION
// ===================================

/**
 * Maneja el envío del formulario
 */
async function handleFormSubmit(e) {
    e.preventDefault();

    // Validar que se hayan seleccionado archivos
    if (selectedFiles.length === 0) {
        alert('Por favor selecciona al menos un archivo Excel');
        return;
    }

    // Verificar que el backend esté disponible
    const isHealthy = await checkAPIHealth();
    if (!isHealthy) {
        alert('El servidor backend no está disponible. Intenta de nuevo.');
        return;
    }

    // Deshabilitar botón y mostrar estado de carga
    generateBtn.disabled = true;
    const originalBtnHTML = generateBtn.innerHTML;
    generateBtn.innerHTML = `
        <svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor" style="animation: spin 1s linear infinite;">
            <circle cx="12" cy="12" r="10" stroke="currentColor" stroke-width="2" fill="none" opacity="0.25"/>
            <path d="M12 2a10 10 0 0 1 10 10" stroke="currentColor" stroke-width="2" fill="none"/>
        </svg>
        <span>Procesando...</span>
    `;

    generatedFiles = [];
    let errorCount = 0;

    try {
        // Procesar cada archivo
        for (let i = 0; i < selectedFiles.length; i++) {
            const file = selectedFiles[i];
            const fileElement = document.querySelector(`.file-item-col[data-index="${i}"]`);
            
            // Marcar como procesando
            if (fileElement) {
                fileElement.classList.add('processing');
                fileElement.querySelector('.file-item-col-meta').innerHTML = `
                    <span class="file-item-col-status processing">⏳ Procesando...</span>
                `;
            }
            
            try {
                // Preparar FormData
                const formData = new FormData();
                formData.append('excel_file', file);

                // Enviar petición
                const response = await fetch(`${API_URL}/generate`, {
                    method: 'POST',
                    body: formData
                });

                if (!response.ok) {
                    const errorData = await response.json().catch(() => ({ error: 'Error desconocido' }));
                    throw new Error(errorData.error || `Error HTTP ${response.status}`);
                }

                // Obtener el archivo generado
                const blob = await response.blob();
                
                // Extraer nombre del archivo del header Content-Disposition
                const contentDisposition = response.headers.get('Content-Disposition');
                let filename = null;
                
                if (contentDisposition) {
                    const filenameMatch = contentDisposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
                    if (filenameMatch && filenameMatch[1]) {
                        filename = filenameMatch[1].replace(/['"]/g, '');
                    }
                }
                
                // Si no se pudo extraer del header, generar nombre basado en el archivo original
                if (!filename) {
                    const nombreBase = file.name.replace(/\.(xlsx?|xls)$/i, '');
                    const nombreFormateado = nombreBase.replace(/_/g, ' ')
                        .split(' ')
                        .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
                        .join(' ');
                    filename = `informe ${nombreFormateado}.docx`;
                }

                // Guardar archivo generado
                generatedFiles.push({
                    blob: blob,
                    filename: filename,
                    originalName: file.name,
                    size: blob.size
                });

                // Marcar como completado
                if (fileElement) {
                    fileElement.classList.remove('processing');
                    fileElement.querySelector('.file-item-col-meta').innerHTML = `
                        <span class="file-item-col-status completed">✓ Completado</span>
                    `;
                }

                // Actualizar columna de documentos generados
                updateDownloadLinks();

                // Pequeño delay entre peticiones (500ms)
                if (i < selectedFiles.length - 1) {
                    await new Promise(resolve => setTimeout(resolve, 500));
                }

            } catch (fileError) {
                console.error(`Error procesando ${file.name}:`, fileError);
                errorCount++;
                
                // Marcar como error
                if (fileElement) {
                    fileElement.classList.remove('processing');
                    fileElement.querySelector('.file-item-col-meta').innerHTML = `
                        <span class="file-item-col-status error">✕ Error</span>
                    `;
                }
            }
        }

        // Restaurar botón
        generateBtn.innerHTML = originalBtnHTML;
        generateBtn.disabled = false;

        // Mostrar notificación
        if (generatedFiles.length > 0) {
            if (errorCount > 0) {
                alert(`Se generaron ${generatedFiles.length} de ${selectedFiles.length} informes. ${errorCount} archivo(s) tuvieron errores.`);
            } else {
                alert(`¡${generatedFiles.length} informe(s) generado(s) exitosamente!`);
            }
        } else {
            alert('No se pudo generar ningún informe. Verifica que el backend esté corriendo y los archivos sean válidos.');
        }

    } catch (error) {
        console.error('Error:', error);
        alert(`Error: ${error.message}`);
        generateBtn.innerHTML = originalBtnHTML;
        generateBtn.disabled = false;
    }
}

// ===================================
// DESCARGA
// ===================================

/**
 * Descarga un archivo específico
 */
function downloadFile(index) {
    const file = generatedFiles[index];
    if (!file) {
        alert('No hay archivo para descargar');
        return;
    }

    // Crear URL temporal
    const url = window.URL.createObjectURL(file.blob);
    
    // Crear enlace temporal
    const a = document.createElement('a');
    a.href = url;
    a.download = file.filename;
    document.body.appendChild(a);
    a.click();
    
    // Limpiar
    window.URL.revokeObjectURL(url);
    document.body.removeChild(a);
}

// ===================================
// RESET Y LIMPIEZA
// ===================================

/**
 * Limpia todos los archivos y documentos
 */
function clearAll() {
    // Limpiar archivos seleccionados
    selectedFiles = [];
    generatedFiles = [];
    
    // Resetear input
    excelFileInput.value = '';
    
    // Actualizar vistas
    updateFileList();
    updateDownloadLinks();
    
    console.log('✨ Todo limpiado');
}

// ===================================
// UTILIDADES
// ===================================

/**
 * Formatea el tamaño de archivo
 */
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

/**
 * Verifica la salud de la API
 */
async function checkAPIHealth() {
    try {
        const response = await fetch(`${API_URL}/health`, {
            method: 'GET',
            headers: {
                'Accept': 'application/json'
            }
        });
        
        if (!response.ok) {
            console.error('❌ API respondió con error:', response.status);
            return false;
        }
        
        const data = await response.json();
        console.log('✅ API Status:', data);
        return data.status === 'healthy';
    } catch (error) {
        console.error('❌ API no disponible:', error);
        return false;
    }
}

// ===================================
// EVENT LISTENERS
// ===================================

/**
 * Inicializa los event listeners
 */
function initializeEventListeners() {
    uploadForm.addEventListener('submit', handleFormSubmit);
    clearBtn.addEventListener('click', clearAll);
}

// ===================================
// INICIALIZACIÓN
// ===================================

document.addEventListener('DOMContentLoaded', () => {
    initializeFileInputs();
    initializeEventListeners();
    
    // Verificar API al cargar
    checkAPIHealth().then(isHealthy => {
        if (!isHealthy) {
            console.warn('⚠️ Advertencia: Backend no disponible. Inicia el servidor backend antes de usar la aplicación.');
        }
    });
});