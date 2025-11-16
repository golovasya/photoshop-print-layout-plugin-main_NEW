// =====================================================
// Print Layout Manager - UXP Plugin for Photoshop
// =====================================================

const { app } = require('photoshop');

// Глобальные переменные
let tableData = [];
let currentFile = null;
let selectedPrintIndex = null;
let layerToPrintMap = new Map();
let printToLayerMap = new Map();

// Элементы UI
let loadXlsxBtn, runScriptBtn, clearFileBtn;
let fileInfo, fileName, printsList, printDetails;
let searchInput, statusText, printCount;
let detailArticle, detailSize, detailColor, mockupImage;
let physicalWidth, physicalHeight, applySizeBtn;

// =====================================================
// Инициализация
// =====================================================

function init() {
    loadXlsxBtn = document.getElementById('loadXlsxBtn');
    runScriptBtn = document.getElementById('runScriptBtn');
    clearFileBtn = document.getElementById('clearFileBtn');
    fileInfo = document.getElementById('fileInfo');
    fileName = document.getElementById('fileName');
    printsList = document.getElementById('printsList');
    printDetails = document.getElementById('printDetails');
    searchInput = document.getElementById('searchInput');
    statusText = document.getElementById('statusText');
    printCount = document.getElementById('printCount');
    
    detailArticle = document.getElementById('detailArticle');
    detailSize = document.getElementById('detailSize');
    detailColor = document.getElementById('detailColor');
    mockupImage = document.getElementById('mockupImage');
    physicalWidth = document.getElementById('physicalWidth');
    physicalHeight = document.getElementById('physicalHeight');
    applySizeBtn = document.getElementById('applySizeBtn');

    loadXlsxBtn.addEventListener('click', loadXlsxFile);
    runScriptBtn.addEventListener('click', runLayoutScript);
    clearFileBtn.addEventListener('click', clearFile);
    searchInput.addEventListener('input', filterPrints);
    applySizeBtn.addEventListener('click', applyPhysicalSize);

    // Проверяем, что XLSX библиотека загружена
    if (typeof XLSX === 'undefined') {
        updateStatus('ОШИБКА: Библиотека XLSX не загружена!');
        console.error('XLSX library not found! Make sure lib/xlsx.full.min.js exists and is loaded in index.html');
    } else {
        console.log('XLSX library loaded successfully');
        updateStatus('Плагин готов к работе');
    }

    checkDocument();
    refreshPrintsList();
}

// =====================================================
// Загрузка XLSX файла
// =====================================================

async function loadXlsxFile() {
    try {
        // Проверяем наличие библиотеки
        if (typeof XLSX === 'undefined') {
            updateStatus('ОШИБКА: Библиотека XLSX не загружена');
            return;
        }

        updateStatus('Выбор файла...');
        
        const fs = require('uxp').storage.localFileSystem;
        
        const file = await fs.getFileForOpening({
            types: ['xlsx', 'xls']
        });

        if (!file) {
            updateStatus('Выбор файла отменён');
            return;
        }

        updateStatus('Чтение файла...');
        
        const arrayBuffer = await file.read();
        
        console.log('File size:', arrayBuffer.byteLength);
        
        if (!arrayBuffer || arrayBuffer.byteLength === 0) {
            throw new Error('Файл пустой');
        }
        
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        
        console.log('Sheets:', workbook.SheetNames);
        
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
        
        console.log('Rows:', jsonData.length);
        
        parseTableData(jsonData);
        
        currentFile = file;
        fileName.textContent = file.name;
        fileInfo.classList.remove('hidden');
        runScriptBtn.disabled = false;
        
        updateStatus(`Загружено ${tableData.length} записей из ${file.name}`);
        refreshPrintsList();
        
    } catch (error) {
        console.error('ОШИБКА:', error);
        updateStatus('Ошибка: ' + error.message);
    }
}

// =====================================================
// Парсинг данных таблицы
// =====================================================

function parseTableData(jsonData) {
    tableData = [];
    
    for (let i = 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        
        if (!row || row.length === 0) continue;
        
        const printData = {
            rowIndex: i,
            photo: row[0] || null,
            size: row[1] || 'Unknown',
            orderId: row[2] || '',
            name: row[3] || '',
            color: row[4] || '',
            article: row[5] || 'Unknown',
            physicalWidth: null,
            physicalHeight: null,
            layerId: null
        };
        
        tableData.push(printData);
    }
    
    console.log('Parsed records:', tableData.length);
}

// =====================================================
// Очистка файла
// =====================================================

function clearFile() {
    currentFile = null;
    tableData = [];
    fileName.textContent = 'Файл не загружен';
    fileInfo.classList.add('hidden');
    runScriptBtn.disabled = true;
    refreshPrintsList();
    updateStatus('Файл очищен');
}

// =====================================================
// Запуск скрипта раскладки
// =====================================================

async function runLayoutScript() {
    try {
        updateStatus('Запуск скрипта раскладки...');
        
        if (tableData.length === 0) {
            updateStatus('Сначала загрузите таблицу XLSX');
            return;
        }
        
        const fs = require('uxp').storage.localFileSystem;
        
        const scriptFile = await fs.getFileForOpening({
            types: ['jsx']
        });
        
        if (!scriptFile) {
            updateStatus('Выбор скрипта отменён');
            return;
        }
        
        updateStatus('Выполнение скрипта...');
        
        const scriptContent = await scriptFile.read({ format: require('uxp').storage.formats.utf8 });
        
        const { executeAsModal } = require('photoshop').core;
        
        await executeAsModal(async () => {
            const batchPlay = require('photoshop').action.batchPlay;
            
            await batchPlay([{
                _obj: "AdobeScriptAutomation Scripts",
                javaScriptMessage: scriptContent,
                _options: { dialogOptions: "dontDisplay" }
            }], {});
        });
        
        await new Promise(resolve => setTimeout(resolve, 1000));
        await refreshPrintsList();
        
        updateStatus('Скрипт выполнен успешно');
        
    } catch (error) {
        console.error('Ошибка выполнения скрипта:', error);
        updateStatus('Ошибка: ' + error.message);
    }
}

// =====================================================
// Обновление списка принтов
// =====================================================

async function refreshPrintsList() {
    printsList.innerHTML = '';
    
    if (!app.activeDocument) {
        printsList.innerHTML = '<div class="hint" style="padding: 20px; text-align: center;">Нет открытого документа</div>';
        printCount.textContent = '0';
        return;
    }
    
    try {
        const doc = app.activeDocument;
        const layers = doc.layers;
        
        layerToPrintMap.clear();
        printToLayerMap.clear();
        
        let matchCount = 0;
        
        for (let i = 0; i < layers.length; i++) {
            const layer = layers[i];
            
            if (layer.isBackgroundLayer) continue;
            
            const layerName = layer.name;
            
            for (let j = 0; j < tableData.length; j++) {
                const printData = tableData[j];
                
                if (layerName.includes(printData.article)) {
                    printData.layerId = layer.id;
                    
                    try {
                        const bounds = layer.bounds;
                        printData.physicalWidth = Math.round((bounds.right - bounds.left) * 0.352778 * 10) / 10;
                        printData.physicalHeight = Math.round((bounds.bottom - bounds.top) * 0.352778 * 10) / 10;
                    } catch (err) {
                        console.error('Error getting layer bounds:', err);
                    }
                    
                    layerToPrintMap.set(layer.id, printData);
                    printToLayerMap.set(j, layer.id);
                    matchCount++;
                    break;
                }
            }
        }
        
        printCount.textContent = matchCount.toString();
        
        const matchedPrints = tableData.filter(p => p.layerId !== null);
        
        if (matchedPrints.length === 0) {
            printsList.innerHTML = '<div class="hint" style="padding: 20px; text-align: center;">Нет сопоставленных слоёв.<br>Слои должны содержать артикулы в названии.</div>';
            return;
        }
        
        matchedPrints.forEach((printData, index) => {
            const item = createPrintItem(printData, index);
            printsList.appendChild(item);
        });
        
        updateStatus(`Найдено ${matchCount} принтов на холсте`);
        
    } catch (error) {
        console.error('Error refreshing prints list:', error);
        printsList.innerHTML = '<div class="hint" style="padding: 20px; text-align: center; color: red;">Ошибка: ' + error.message + '</div>';
    }
}

// =====================================================
// Создание элемента принта
// =====================================================

function createPrintItem(printData, index) {
    const item = document.createElement('div');
    item.className = 'print-item';
    item.dataset.index = index;
    item.dataset.layerId = printData.layerId;
    
    const thumbnail = document.createElement('div');
    thumbnail.className = 'print-thumbnail';
    thumbnail.innerHTML = '<span style="font-size: 20px;">🖼️</span>';
    
    const info = document.createElement('div');
    info.className = 'print-info';
    
    const article = document.createElement('div');
    article.className = 'print-article';
    article.textContent = printData.article;
    
    const meta = document.createElement('div');
    meta.className = 'print-meta';
    
    const sizeBadge = document.createElement('span');
    sizeBadge.className = 'print-size-badge';
    sizeBadge.textContent = printData.size;
    
    const dimensions = document.createElement('span');
    if (printData.physicalWidth && printData.physicalHeight) {
        dimensions.textContent = `${printData.physicalWidth}×${printData.physicalHeight} мм`;
    } else {
        dimensions.textContent = 'Размер не определён';
    }
    
    meta.appendChild(sizeBadge);
    meta.appendChild(dimensions);
    
    info.appendChild(article);
    info.appendChild(meta);
    
    item.appendChild(thumbnail);
    item.appendChild(info);
    
    item.addEventListener('click', () => selectPrint(index, printData));
    
    return item;
}

// =====================================================
// Выбор принта
// =====================================================

async function selectPrint(index, printData) {
    selectedPrintIndex = index;
    
    document.querySelectorAll('.print-item').forEach(item => {
        item.classList.remove('selected');
    });
    
    const selectedItem = document.querySelector(`[data-index="${index}"]`);
    if (selectedItem) {
        selectedItem.classList.add('selected');
    }
    
    showPrintDetails(printData);
    
    try {
        if (printData.layerId && app.activeDocument) {
            const layer = app.activeDocument.layers.find(l => l.id === printData.layerId);
            if (layer) {
                app.activeDocument.activeLayers = [layer];
                updateStatus(`Выбран: ${printData.article}`);
            }
        }
    } catch (error) {
        console.error('Error selecting layer:', error);
    }
}

// =====================================================
// Показ деталей принта
// =====================================================

function showPrintDetails(printData) {
    printDetails.classList.remove('hidden');
    
    detailArticle.textContent = printData.article;
    detailSize.textContent = printData.size;
    detailColor.textContent = printData.color || 'Не указан';
    
    physicalWidth.value = printData.physicalWidth || '';
    physicalHeight.value = printData.physicalHeight || '';
    
    mockupImage.src = '';
    mockupImage.alt = 'Мокап недоступен';
}

// =====================================================
// Применение физического размера
// =====================================================

async function applyPhysicalSize() {
    if (selectedPrintIndex === null) {
        updateStatus('Сначала выберите принт из списка');
        return;
    }
    
    const width = parseFloat(physicalWidth.value);
    const height = parseFloat(physicalHeight.value);
    
    if (isNaN(width) || isNaN(height) || width <= 0 || height <= 0) {
        updateStatus('Введите корректные размеры (мм)');
        return;
    }
    
    try {
        const printData = tableData.filter(p => p.layerId !== null)[selectedPrintIndex];
        
        if (!printData || !printData.layerId) {
            updateStatus('Слой не найден');
            return;
        }
        
        const doc = app.activeDocument;
        const layer = doc.layers.find(l => l.id === printData.layerId);
        
        if (!layer) {
            updateStatus('Слой не найден в документе');
            return;
        }
        
        const widthPx = width / 0.352778;
        const heightPx = height / 0.352778;
        
        const bounds = layer.bounds;
        const currentWidth = bounds.right - bounds.left;
        const currentHeight = bounds.bottom - bounds.top;
        
        const scaleX = (widthPx / currentWidth) * 100;
        const scaleY = (heightPx / currentHeight) * 100;
        
        await layer.scale(scaleX, scaleY);
        
        printData.physicalWidth = width;
        printData.physicalHeight = height;
        
        updateStatus(`Размер изменён: ${width}×${height} мм`);
        
        refreshPrintsList();
        
    } catch (error) {
        console.error('Error applying size:', error);
        updateStatus('Не удалось применить размер: ' + error.message);
    }
}

// =====================================================
// Фильтрация принтов
// =====================================================

function filterPrints() {
    const query = searchInput.value.toLowerCase();
    
    document.querySelectorAll('.print-item').forEach(item => {
        const article = item.querySelector('.print-article').textContent.toLowerCase();
        
        if (article.includes(query)) {
            item.style.display = 'flex';
        } else {
            item.style.display = 'none';
        }
    });
}

// =====================================================
// Проверка документа
// =====================================================

function checkDocument() {
    if (!app.activeDocument) {
        updateStatus('Нет открытого документа');
    }
}

// =====================================================
// Утилиты
// =====================================================

function updateStatus(message) {
    statusText.textContent = message;
    console.log('Status:', message);
}

// =====================================================
// Запуск при загрузке
// =====================================================

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
} else {
    init();
}
