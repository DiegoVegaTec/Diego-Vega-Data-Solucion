// Configurar PDF.js Worker globalmente
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.worker.min.js';

// --- EFECTO DE FONDO DIGITAL MATRIX ---
const canvas = document.getElementById('matrix-canvas');
const ctx = canvas.getContext('2d');

function resizeCanvas() {
    canvas.width = window.innerWidth;
    canvas.height = window.innerHeight;
}
resizeCanvas();
window.addEventListener('resize', resizeCanvas);

const katakana = 'ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝ1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ';
const alphabet = katakana.split('');

const fontSize = 16;
let columns = canvas.width / fontSize;
let rainDrops = Array(Math.floor(columns)).fill(1);

function drawMatrix() {
    ctx.fillStyle = 'rgba(10, 15, 12, 0.05)'; // Crea el rastro de desvanecimiento
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    ctx.fillStyle = '#00ff41'; // Color verde Matrix clásico
    ctx.font = fontSize + 'px monospace';

    for (let i = 0; i < rainDrops.length; i++) {
        const text = alphabet[Math.floor(Math.random() * alphabet.length)];
        ctx.fillText(text, i * fontSize, rainDrops[i] * fontSize);

        if (rainDrops[i] * fontSize > canvas.height && Math.random() > 0.975) {
            rainDrops[i] = 0;
        }
        rainDrops[i]++;
    }
}
setInterval(drawMatrix, 30);


// --- GESTIÓN DE PESTAÑAS ---
function switchTab(tabId) {
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
    
    document.getElementById(tabId).classList.add('active');
    event.currentTarget.classList.add('active');
}

// Variables Globales de almacenamiento temporal de archivos
let cacheFiles = {
    unir: [],
    limpiar: null,
    duplicados: null,
    pdf: null,
    json: null
};

// Inicialización de Drag and Drop y selectores nativos
setupModule('unir', true);
setupModule('limpiar', false);
setupModule('duplicados', false);
setupModule('pdf', false);
setupModule('json', false);

function setupModule(id, isMultiple) {
    const zone = document.getElementById(`drop-${id}`);
    const input = document.getElementById(`file-${id}`);
    const listDisplay = document.getElementById(`file-${id === 'unir' ? 'list' : 'info'}-${id}`);
    const btn = document.getElementById(`btn-${id}`);

    if(!zone) return;

    zone.addEventListener('click', () => input.click());
    zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('dragover'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
    zone.addEventListener('drop', (e) => {
        e.preventDefault();
        zone.classList.remove('dragover');
        handleFiles(e.dataTransfer.files, id, isMultiple, listDisplay, btn);
    });

    input.addEventListener('change', (e) => {
        handleFiles(e.target.files, id, isMultiple, listDisplay, btn);
    });
}

function handleFiles(files, id, isMultiple, listDisplay, btn) {
    if(files.length === 0) return;
    
    listDisplay.style.display = 'block';
    if(isMultiple) {
        cacheFiles[id] = Array.from(files);
        listDisplay.innerHTML = `<strong>Archivos cargados:</strong><br>` + cacheFiles[id].map(f => `- ${f.name}`).join('<br>');
        btn.disabled = cacheFiles[id].length < 2;
    } else {
        cacheFiles[id] = files[0];
        listDisplay.innerHTML = `<strong>Archivo seleccionado:</strong> ${files[0].name} (${(files[0].size/1024).toFixed(1)} KB)`;
        btn.disabled = false;
    }
}

// --- LOGICA DE TRATAMIENTO DE DATOS ---

// 1. UNIR EXCELS
document.getElementById('btn-unir').addEventListener('click', async () => {
    const files = cacheFiles.unir;
    const masterWorkbook = XLSX.utils.book_new();
    let sheetsData = {};

    for (let file of files) {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array' });
        
        workbook.SheetNames.forEach(sheetName => {
            const jsonSheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
            if (!sheetsData[sheetName]) {
                sheetsData[sheetName] = [];
            }
            if (sheetsData[sheetName].length === 0) {
                sheetsData[sheetName] = sheetsData[sheetName].concat(jsonSheet);
            } else {
                sheetsData[sheetName] = sheetsData[sheetName].concat(jsonSheet.slice(1));
            }
        });
    }

    Object.keys(sheetsData).forEach(sheetName => {
        const ws = XLSX.utils.aoa_to_sheet(sheetsData[sheetName]);
        XLSX.utils.book_append_sheet(masterWorkbook, ws, sheetName);
    });

    XLSX.writeFile(masterWorkbook, 'Excel_Maestro_Consolidado.xlsx');
});

// 2. ELIMINAR VALORES NEGATIVOS Y EN 0
document.getElementById('btn-limpiar').addEventListener('click', async () => {
    const file = cacheFiles.limpiar;
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, { type: 'array' });
    const masterWorkbook = XLSX.utils.book_new();

    wb.SheetNames.forEach(name => {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1 });
        if(rows.length === 0) return;
        
        const headers = rows[0];
        const filteredRows = rows.slice(1).filter(row => {
            return !row.some(cell => typeof cell === 'number' && cell <= 0);
        });

        filteredRows.unshift(headers);
        const ws = XLSX.utils.aoa_to_sheet(filteredRows);
        XLSX.utils.book_append_sheet(masterWorkbook, ws, name);
    });

    XLSX.writeFile(masterWorkbook, 'Excel_Depurado_Sin_Negativos.xlsx');
});

// 3. ELIMINAR FILAS DUPLICADAS EXACTAS
document.getElementById('btn-duplicados').addEventListener('click', async () => {
    const file = cacheFiles.duplicados;
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, { type: 'array' });
    const masterWorkbook = XLSX.utils.book_new();

    wb.SheetNames.forEach(name => {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1 });
        if(rows.length === 0) return;

        const uniqueRowsMap = new Map();
        rows.forEach(row => {
            const stringifiedRow = JSON.stringify(row);
            if(!uniqueRowsMap.has(stringifiedRow)) {
                uniqueRowsMap.set(stringifiedRow, row);
            }
        });

        const uniqueRows = Array.from(uniqueRowsMap.values());
        const ws = XLSX.utils.aoa_to_sheet(uniqueRows);
        XLSX.utils.book_append_sheet(masterWorkbook, ws, name);
    });

    XLSX.writeFile(masterWorkbook, 'Excel_Sin_Duplicados.xlsx');
});

// 4. CONVERTIR PDF A EXCEL
document.getElementById('btn-pdf').addEventListener('click', async () => {
    const file = cacheFiles.pdf;
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    let excelRows = [];

    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        
        let lastY = null;
        let line = [];
        
        textContent.items.forEach(item => {
            let currentY = item.transform[5];
            if (lastY !== null && Math.abs(currentY - lastY) > 5) {
                excelRows.push(line);
                line = [];
            }
            line.push(item.str.trim());
            lastY = currentY;
        });
        if(line.length > 0) excelRows.push(line);
    }

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(excelRows);
    XLSX.utils.book_append_sheet(wb, ws, "Datos Extraídos PDF");
    XLSX.writeFile(wb, 'PDF_Convertido_a_Excel.xlsx');
});

// 5. CONVERTIR JSON A EXCEL
document.getElementById('btn-json').addEventListener('click', async () => {
    const file = cacheFiles.json;
    const text = await file.text();
    try {
        const jsonParsed = JSON.parse(text);
        const dataArray = Array.isArray(jsonParsed) ? jsonParsed : [jsonParsed];
        
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(dataArray);
        XLSX.utils.book_append_sheet(wb, ws, "JSON Data");
        XLSX.writeFile(wb, 'JSON_Convertido_a_Excel.xlsx');
    } catch (err) {
        alert("Error de sintaxis: El archivo JSON no tiene una estructura válida.");
    }
});