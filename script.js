const backgrounds = {
    1: 'radial-gradient(circle at 15% 20%, rgba(14, 165, 233, 0.28) 0%, transparent 42%), radial-gradient(circle at 85% 80%, rgba(99, 102, 241, 0.22) 0%, transparent 46%), linear-gradient(135deg, #020617 0%, #0b1530 42%, #0a1f3f 100%)',
    2: 'linear-gradient(160deg, #001b2e 0%, #003049 34%, #005f73 72%, #0a9396 100%)',
    3: 'radial-gradient(circle at 20% 10%, rgba(217, 70, 239, 0.2) 0%, transparent 48%), linear-gradient(135deg, #1e1b4b 0%, #312e81 45%, #4c1d95 100%)',
    4: 'linear-gradient(135deg, #f0fdf4 0%, #d1fae5 45%, #bae6fd 100%)',
    5: 'radial-gradient(circle at 80% 15%, rgba(253, 186, 116, 0.45) 0%, transparent 38%), linear-gradient(145deg, #7c2d12 0%, #c2410c 45%, #f59e0b 100%)',
    6: 'linear-gradient(155deg, #f8fafc 0%, #e2e8f0 52%, #cbd5e1 100%)',
    7: 'linear-gradient(145deg, #fff7ed 0%, #ffedd5 40%, #fde68a 100%)',
    8: 'radial-gradient(circle at 18% 20%, rgba(74, 222, 128, 0.3) 0%, transparent 45%), linear-gradient(160deg, #052e16 0%, #14532d 46%, #166534 100%)',
    9: 'repeating-linear-gradient(90deg, rgba(56, 189, 248, 0.08) 0px, rgba(56, 189, 248, 0.08) 1px, transparent 1px, transparent 16px), repeating-linear-gradient(0deg, rgba(99, 102, 241, 0.08) 0px, rgba(99, 102, 241, 0.08) 1px, transparent 1px, transparent 16px), linear-gradient(145deg, #020617 0%, #111827 100%)',
    10: 'linear-gradient(135deg, #111827 0%, #1f2937 35%, #b45309 100%)'
};

let currentBg = 1;
let renderTimeout = null;
let animationInterval = null;
let adjustSizesTimeout = null;
let isAnimating = false;
let isExporting = false;
let currentColumnKeysSignature = '';
let currentColumnTabKey = '';
let tableScrollY = 0;
const customBackgroundImage = {
    dataUrl: '',
    image: null,
    opacity: 0.45
};
const columnConfigs = {};
const rowAnimationClasses = [
    'anim-type-expand',
    'anim-type-slide-left',
    'anim-type-slide-right',
    'anim-type-rise-up',
    'anim-type-zoom-in'
];

function escapeHTML(value) {
    return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

let cachedTableDataText = null;
let cachedTableData = null;

function parseTableData() {
    const dataInput = document.getElementById('dataInput');
    const text = dataInput?.value || '';
    if (text && text === cachedTableDataText && cachedTableData) {
        return cachedTableData;
    }

    try {
        const data = JSON.parse(text);
        if (!Array.isArray(data) || data.length === 0) {
            cachedTableDataText = null;
            cachedTableData = null;
            return null;
        }
        cachedTableDataText = text;
        cachedTableData = data;
        return data;
    } catch (e) {
        // JSON 未完成输入时也会触发 change/input；避免反复报错和无意义的重渲染。
        return null;
    }
}

function getDisplayKeys(data) {
    if (!Array.isArray(data) || data.length === 0) return [];
    return Object.keys(data[0]).filter(key => key !== '排名变动');
}

function clamp(value, min, max) {
    return Math.max(min, Math.min(max, value));
}

function createDefaultColumnConfig(columnCount) {
    const width = columnCount > 0 ? 100 / columnCount : 20;
    return {
        width,
        cellFontSize: 0,
        cellColor: '#f1f5f9',
        cellBold: false,
        headerFontSize: 0,
        headerColor: '#fbbf24',
        visible: true
    };
}

function syncColumnConfigs(keys) {
    const keySet = new Set(keys);
    Object.keys(columnConfigs).forEach(key => {
        if (!keySet.has(key)) {
            delete columnConfigs[key];
        }
    });
    keys.forEach(key => {
        if (!columnConfigs[key]) {
            columnConfigs[key] = createDefaultColumnConfig(keys.length);
        }
    });
}

function getNormalizedColumnWidths(keys) {
    if (keys.length === 0) return {};
    const rawWidths = keys.map(key => {
        const configured = Number(columnConfigs[key]?.width);
        return Number.isFinite(configured) ? clamp(configured, 1, 1000) : 1;
    });
    const total = rawWidths.reduce((sum, value) => sum + value, 0) || 1;
    const normalized = {};
    keys.forEach((key, index) => {
        normalized[key] = (rawWidths[index] / total) * 100;
    });
    return normalized;
}

function buildColumnStyles(config) {
    const headerStyles = [];
    const cellStyles = [];
    if (Number.isFinite(config.headerFontSize) && config.headerFontSize > 0) {
        headerStyles.push(`font-size:${config.headerFontSize}px`);
    }
    if (config.headerColor) {
        headerStyles.push(`color:${config.headerColor}`);
    }
    if (Number.isFinite(config.cellFontSize) && config.cellFontSize > 0) {
        cellStyles.push(`font-size:${config.cellFontSize}px`);
    }
    if (config.cellColor) {
        cellStyles.push(`color:${config.cellColor}`);
    }
    if (config.cellBold) {
        cellStyles.push('font-weight:700');
    }
    return {
        headerStyle: headerStyles.join(';'),
        cellStyle: cellStyles.join(';')
    };
}

function renderColumnConfigPanel(allKeys) {
    const list = document.getElementById('columnConfigList');
    if (!list) return;
    list.innerHTML = '';
    const visibleKeys = allKeys.filter(key => columnConfigs[key]?.visible !== false);

    if (!allKeys.length) {
        const empty = document.createElement('div');
        empty.className = 'column-config-empty';
        empty.textContent = '请先输入有效JSON数据，系统会自动生成列配置项';
        list.appendChild(empty);
        currentColumnTabKey = '';
        return;
    }

    if (!currentColumnTabKey || !allKeys.includes(currentColumnTabKey)) {
        currentColumnTabKey = allKeys[0];
    }

    const header = document.createElement('div');
    header.className = 'column-list-header';
    header.textContent = '点击列名进入设置，勾选控制显示/隐藏';
    list.appendChild(header);

    const columnList = document.createElement('div');
    columnList.className = 'column-list';
    allKeys.forEach(key => {
        const isVisible = visibleKeys.includes(key);
        const item = document.createElement('div');
        item.className = `column-list-item${key === currentColumnTabKey ? ' active' : ''}${!isVisible ? ' hidden' : ''}`;

        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.checked = isVisible;
        checkbox.className = 'column-visibility-checkbox';
        checkbox.addEventListener('change', (e) => {
            e.stopPropagation();
            columnConfigs[key].visible = e.target.checked;
            renderTable();
        });

        const label = document.createElement('span');
        label.className = 'column-list-label';
        label.textContent = key;
        label.addEventListener('click', () => {
            if (currentColumnTabKey === key) return;
            currentColumnTabKey = key;
            renderColumnConfigPanel(allKeys);
        });

        item.appendChild(checkbox);
        item.appendChild(label);
        columnList.appendChild(item);
    });
    list.appendChild(columnList);

    const config = columnConfigs[currentColumnTabKey];
    const card = document.createElement('div');
    card.className = 'column-config-card';

    const title = document.createElement('div');
    title.className = 'column-config-title';
    title.textContent = `${currentColumnTabKey} 列设置`;
    card.appendChild(title);

    const grid = document.createElement('div');
    grid.className = 'column-config-grid';
    card.appendChild(grid);

    const fields = [
        { label: '列宽权重', prop: 'width', type: 'number', min: 1, max: 1000, step: 1, value: Math.round(config.width) },
        { label: '列字体大小', prop: 'cellFontSize', type: 'number', min: 0, max: 40, step: 1, value: config.cellFontSize || 0 },
        { label: '列字体颜色', prop: 'cellColor', type: 'color', value: config.cellColor || '#f1f5f9' },
        { label: '列字体加粗', prop: 'cellBold', type: 'checkbox', checked: !!config.cellBold },
        { label: '表头字体大小', prop: 'headerFontSize', type: 'number', min: 0, max: 40, step: 1, value: config.headerFontSize || 0 },
        { label: '表头字体颜色', prop: 'headerColor', type: 'color', value: config.headerColor || '#fbbf24' }
    ];

    fields.forEach(field => {
        const wrapper = document.createElement('div');
        wrapper.className = 'column-config-field';
        const label = document.createElement('label');
        label.textContent = field.label;
        const input = document.createElement('input');
        input.type = field.type;
        if (field.type === 'checkbox') {
            input.checked = !!field.checked;
            wrapper.classList.add('checkbox-field');
        } else {
            input.value = String(field.value);
        }
        if (field.type === 'number') {
            input.min = String(field.min);
            input.max = String(field.max);
            input.step = String(field.step);
        }
        const handler = event => {
            const nextValue = field.type === 'checkbox' ? event.target.checked : event.target.value;
            handleColumnConfigInputChange(currentColumnTabKey, field.prop, nextValue);
        };
        input.addEventListener(field.type === 'checkbox' ? 'change' : 'input', handler);
        wrapper.appendChild(label);
        wrapper.appendChild(input);
        grid.appendChild(wrapper);
    });

    list.appendChild(card);
}

function handleColumnConfigInputChange(key, prop, rawValue) {
    if (!columnConfigs[key]) return;
    if (prop === 'cellColor' || prop === 'headerColor') {
        columnConfigs[key][prop] = rawValue;
    } else if (prop === 'cellBold' || prop === 'visible') {
        columnConfigs[key][prop] = !!rawValue;
    } else if (prop === 'width') {
        const value = Number(rawValue);
        if (!Number.isFinite(value)) return;
        columnConfigs[key][prop] = clamp(value, 1, 1000);
    } else if (prop === 'cellFontSize' || prop === 'headerFontSize') {
        const value = Number(rawValue);
        if (!Number.isFinite(value)) return;
        columnConfigs[key][prop] = clamp(value, 0, 40);
    } else {
        return;
    }
    renderTable();
    updatePreview();
}

function getBarrageLines() {
    const rawValue = document.getElementById('barrageInput')?.value || '';
    return rawValue
        .split(/\r?\n/)
        .map(item => item.trim())
        .filter(Boolean);
}

function renderBarrage() {
    const barrageLayer = document.getElementById('barrageLayer');
    if (!barrageLayer) return;
    barrageLayer.innerHTML = '';

    // 如果不是在动画或导出状态，就不生成弹幕DOM，保持画面干净
    if (!isAnimating && !isExporting) return;
    
    // 如果弹幕开关未开启，也不生成弹幕
    const barrageEnabled = document.getElementById('barrageEnabled')?.checked;
    if (!barrageEnabled) return;

    const lines = getBarrageLines();
    if (!lines.length) return;

    const sizeInput = Number(document.getElementById('barrageFontSize')?.value || 16);
    const fontSize = Number.isFinite(sizeInput) ? clamp(sizeInput, 10, 48) : 16;
    const color = document.getElementById('barrageColor')?.value || '#ffffff';
    const layerHeight = Math.max(120, barrageLayer.clientHeight || 560);
    const capsuleHeight = fontSize + 14;
    const minTop = 78;
    const maxTop = Math.max(minTop, layerHeight - capsuleHeight - 12);
    const usedTops = [];

    lines.forEach((line, index) => {
        const item = document.createElement('div');
        item.className = 'barrage-item';
        item.textContent = line;
        item.style.fontSize = `${fontSize}px`;
        item.style.color = color;
        let top = minTop;
        if (maxTop > minTop) {
            for (let attempt = 0; attempt < 8; attempt++) {
                const candidate = Math.round(minTop + Math.random() * (maxTop - minTop));
                const overlaps = usedTops.some(existing => Math.abs(existing - candidate) < capsuleHeight * 0.85);
                if (!overlaps || attempt === 7) {
                    top = candidate;
                    break;
                }
            }
        }
        usedTops.push(top);
        item.style.top = `${top}px`;
        item.style.animationDuration = `${8 + (index % 4) * 1.2 + (index % 3) * 0.8}s`;
        item.style.animationDelay = `${-(index * 1.35)}s`;
        barrageLayer.appendChild(item);
    });
}

function scheduleAdjustTableSizes(delay = 120) {
    if (adjustSizesTimeout) {
        clearTimeout(adjustSizesTimeout);
    }
    adjustSizesTimeout = setTimeout(() => {
        if (isAnimating || isExporting) return;
        adjustTableSizes();
    }, delay);
}

function init() {
    setupEventListeners();
    renderTable();
    updateBackground();
    updatePreview();
    updateSortColumnOptions();
    scheduleAdjustTableSizes(320);
}

function setupEventListeners() {
    document.getElementById('titleContent').addEventListener('input', updatePreview);
    document.getElementById('titleSize').addEventListener('input', updatePreview);
    document.getElementById('titleColor').addEventListener('input', updatePreview);
    document.getElementById('tableHeaderSize').addEventListener('input', () => scheduleAdjustTableSizes(50));
    
    const dataInput = document.getElementById('dataInput');
    dataInput.addEventListener('input', handleDataChange);
    dataInput.addEventListener('change', handleDataChange);

    document.getElementById('barrageInput')?.addEventListener('input', renderBarrage);
    document.getElementById('barrageFontSize')?.addEventListener('input', renderBarrage);
    document.getElementById('barrageColor')?.addEventListener('input', renderBarrage);
    document.getElementById('barrageEnabled')?.addEventListener('change', renderBarrage);

    document.querySelectorAll('.background-option').forEach(option => {
        option.addEventListener('click', function() {
            document.querySelectorAll('.background-option').forEach(o => o.classList.remove('selected'));
            this.classList.add('selected');
            currentBg = parseInt(this.dataset.bg);
            updateBackground();
        });
    });

    document.querySelector('.background-option[data-bg="1"]').classList.add('selected');

    const backgroundImageInput = document.getElementById('backgroundImageInput');
    const backgroundImageOpacity = document.getElementById('backgroundImageOpacity');
    const clearBackgroundImageBtn = document.getElementById('clearBackgroundImageBtn');
    if (backgroundImageInput) {
        backgroundImageInput.addEventListener('change', handleBackgroundImageUpload);
    }
    if (backgroundImageOpacity) {
        backgroundImageOpacity.addEventListener('input', handleBackgroundImageOpacityChange);
    }
    if (clearBackgroundImageBtn) {
        clearBackgroundImageBtn.addEventListener('click', clearBackgroundImage);
    }

    document.getElementById('playAnimationBtn').addEventListener('click', playAnimation);
    document.getElementById('exportAnimationBtn').addEventListener('click', exportAnimationVideo);

    document.getElementById('sortColumn')?.addEventListener('change', handleSortChange);
    document.getElementById('sortOrder')?.addEventListener('change', handleSortChange);
}

function handleDataChange() {
    if (renderTimeout) {
        clearTimeout(renderTimeout);
    }
    renderTimeout = setTimeout(() => {
        renderTable();
        updateSortColumnOptions();
    }, 100);
}

function updateSortColumnOptions() {
    const select = document.getElementById('sortColumn');
    if (!select) return;

    const data = parseTableData();
    const keys = data && data.length > 0 ? Object.keys(data[0]) : [];
    const numericKeys = [];
    keys.forEach(key => {
        if (key === '排名变动') return;
        let isNumeric = false;
        for (let i = 0; i < data.length; i++) {
            const v = data[i]?.[key];
            if (typeof v === 'number') {
                isNumeric = true;
                break;
            }
            if (!isNaN(Number(v))) {
                isNumeric = true;
                break;
            }
        }
        if (isNumeric) numericKeys.push(key);
    });

    const currentValue = select.value;
    select.innerHTML = '<option value="">-- 不排序 --</option>';
    numericKeys.forEach(key => {
        const option = document.createElement('option');
        option.value = key;
        option.textContent = key;
        select.appendChild(option);
    });

    if (numericKeys.includes(currentValue)) {
        select.value = currentValue;
    } else {
        select.value = '';
    }
}

function handleSortChange() {
    renderTable();
}

function getSortedData(data) {
    const sortColumn = document.getElementById('sortColumn')?.value;
    const sortOrder = document.getElementById('sortOrder')?.value || 'asc';

    if (!sortColumn || !data || data.length === 0) {
        return data;
    }

    const sorted = [...data];
    sorted.sort((a, b) => {
        const aVal = a[sortColumn];
        const bVal = b[sortColumn];
        const aNum = Number(aVal);
        const bNum = Number(bVal);

        if (isNaN(aNum) || isNaN(bNum)) {
            return String(aVal).localeCompare(String(bVal));
        }

        return sortOrder === 'asc' ? aNum - bNum : bNum - aNum;
    });

    return sorted;
}

function updatePreview() {
    const title = document.getElementById('titleContent').value;
    const size = document.getElementById('titleSize').value;
    const color = document.getElementById('titleColor').value;

    const previewTitle = document.getElementById('previewTitle');
    previewTitle.textContent = title;
    previewTitle.style.fontSize = size + 'px';
    previewTitle.style.color = color;
    renderBarrage();
    
    scheduleAdjustTableSizes(200);
}

function updateBackground() {
    const previewPanel = document.getElementById('previewPanel');
    previewPanel.style.background = backgrounds[currentBg];
    const backgroundImageLayer = document.getElementById('backgroundImageLayer');
    if (!backgroundImageLayer) return;
    if (!customBackgroundImage.dataUrl) {
        backgroundImageLayer.style.backgroundImage = 'none';
        backgroundImageLayer.style.opacity = '0';
        return;
    }
    backgroundImageLayer.style.backgroundImage = `url("${customBackgroundImage.dataUrl}")`;
    backgroundImageLayer.style.opacity = `${customBackgroundImage.opacity}`;
}

function handleBackgroundImageUpload(event) {
    const file = event.target.files && event.target.files[0];
    if (!file) return;
    if (!file.type.startsWith('image/')) {
        alert('请选择图片文件');
        return;
    }
    const reader = new FileReader();
    reader.onload = () => {
        const dataUrl = typeof reader.result === 'string' ? reader.result : '';
        if (!dataUrl) return;
        const image = new Image();
        image.onload = () => {
            customBackgroundImage.dataUrl = dataUrl;
            customBackgroundImage.image = image;
            updateBackground();
        };
        image.onerror = () => {
            alert('背景图加载失败，请重试');
        };
        image.src = dataUrl;
    };
    reader.readAsDataURL(file);
}

function handleBackgroundImageOpacityChange(event) {
    const nextOpacity = Number(event.target.value);
    if (!Number.isFinite(nextOpacity)) return;
    customBackgroundImage.opacity = Math.max(0, Math.min(1, nextOpacity / 100));
    updateBackground();
}

function clearBackgroundImage() {
    customBackgroundImage.dataUrl = '';
    customBackgroundImage.image = null;
    const backgroundImageInput = document.getElementById('backgroundImageInput');
    if (backgroundImageInput) {
        backgroundImageInput.value = '';
    }
    updateBackground();
}

function getSelectedAnimationClass() {
    const animationType = document.getElementById('animationType')?.value || 'expand';
    const className = `anim-type-${animationType}`;
    if (!rowAnimationClasses.includes(className)) {
        return 'anim-type-expand';
    }
    return className;
}

function applyDynamicTableStyle(tableContainer, metrics) {
    const { bodyRowHeight, lastRowHeight, paddingVertical, thFontSize, fontSize } = metrics;
    const style = document.createElement('style');
    style.id = 'dynamic-table-styles';
    style.textContent = `
        #tableContainer {
            --row-expand-max: ${bodyRowHeight}px;
            --row-expand-max-last: ${lastRowHeight}px;
            --cell-padding-vertical: ${paddingVertical}px;
        }
        table {
            width: 100%;
            table-layout: fixed;
            border-collapse: collapse;
            transition: transform 0.28s cubic-bezier(0.25, 0.1, 0.25, 1);
        }
        tbody tr.static-row,
        tbody tr.expanded-row,
        tbody tr.collapsed-row {
            height: ${bodyRowHeight}px;
            box-sizing: border-box;
        }
        tbody tr.static-row td,
        tbody tr.expanded-row td,
        tbody tr.collapsed-row td {
            height: ${bodyRowHeight}px;
            box-sizing: border-box;
            overflow: hidden;
        }
        tbody tr.static-row:last-child,
        tbody tr.expanded-row:last-child,
        tbody tr.collapsed-row:last-child {
            height: ${lastRowHeight}px;
        }
        tbody tr.static-row:last-child td,
        tbody tr.expanded-row:last-child td,
        tbody tr.collapsed-row:last-child td {
            height: ${lastRowHeight}px;
        }
        th {
            font-size: ${thFontSize}px;
            padding: 6px 4px;
            line-height: 1.2;
            box-sizing: border-box;
            vertical-align: middle;
        }
        td {
            font-size: ${fontSize}px;
            padding: ${paddingVertical}px 4px;
            line-height: 1.2;
            box-sizing: border-box;
            vertical-align: middle;
        }
        .rank-cell {
            font-size: ${fontSize + 2}px;
        }
        .wealth-cell {
            font-size: ${fontSize + 1}px;
        }
        .change-up, .change-down {
            font-size: ${Math.max(8, fontSize)}px;
        }
    `;

    const oldStyle = document.getElementById('dynamic-table-styles');
    if (oldStyle) {
        oldStyle.remove();
    }
    document.head.appendChild(style);
    document.body.style.setProperty('--fixed-table-height', `${tableContainer.clientHeight}px`);
}

function adjustTableSizes() {
    try {
        const data = parseTableData();
        if (!data) return;

        const rowCount = data.length;
        const tableScrollWrapper = document.getElementById('tableScrollWrapper');
        const table = tableScrollWrapper.querySelector('table');
        
        if (!tableScrollWrapper || !table) return;
        
        const containerHeight = tableScrollWrapper.clientHeight;
        
        if (containerHeight <= 0) {
            setTimeout(adjustTableSizes, 50);
            return;
        }

        const thead = table.querySelector('thead');
        const headerHeight = thead ? Math.ceil(thead.getBoundingClientRect().height) : 0;
        const rowDividerTotal = Math.max(0, rowCount - 1);
        const safeBottomInset = 8;
        const availableBodyHeight = Math.max(0, Math.floor(containerHeight - headerHeight - rowDividerTotal - safeBottomInset));

        let fontSize, paddingVertical, thFontSize, bodyRowHeight, lastRowHeight;
        
        if (rowCount <= 8) {
            fontSize = 13;
            paddingVertical = 6;
            thFontSize = 14;
            bodyRowHeight = Math.max(1, Math.floor(availableBodyHeight / rowCount));
        } else if (rowCount <= 10) {
            fontSize = 12;
            paddingVertical = 5;
            thFontSize = 13;
            bodyRowHeight = Math.max(1, Math.floor(availableBodyHeight / rowCount));
        } else if (rowCount <= 15) {
            fontSize = 12;
            paddingVertical = 5;
            thFontSize = 13;
            bodyRowHeight = 36;
        } else if (rowCount <= 20) {
            fontSize = 11;
            paddingVertical = 4;
            thFontSize = 12;
            bodyRowHeight = 32;
        } else {
            fontSize = 10;
            paddingVertical = 3;
            thFontSize = 11;
            bodyRowHeight = 28;
        }
        lastRowHeight = bodyRowHeight;

        const customHeaderSize = parseInt(document.getElementById('tableHeaderSize')?.value || '0', 10);
        if (Number.isFinite(customHeaderSize) && customHeaderSize > 0) {
            thFontSize = customHeaderSize;
        }

        applyDynamicTableStyle(tableScrollWrapper, {
            bodyRowHeight,
            lastRowHeight,
            paddingVertical,
            thFontSize,
            fontSize
        });
    } catch (e) {
        console.error('Error adjusting table sizes:', e);
    }
}

function renderTable() {
    try {
        const data = parseTableData();
        if (!data) return;

        const sortedData = getSortedData(data);
        const allKeys = getDisplayKeys(sortedData);
        syncColumnConfigs(allKeys);
        const visibleKeys = allKeys.filter(key => columnConfigs[key]?.visible !== false);
        const keySignature = allKeys.join('||');
        if (keySignature !== currentColumnKeysSignature) {
            renderColumnConfigPanel(allKeys, visibleKeys);
            currentColumnKeysSignature = keySignature;
        }
        const normalizedWidths = getNormalizedColumnWidths(visibleKeys);
        const tableScrollWrapper = document.getElementById('tableScrollWrapper');
        
        let html = '<table><colgroup>';
        visibleKeys.forEach(key => {
            const width = normalizedWidths[key] || 0;
            html += `<col style="width:${width.toFixed(4)}%;">`;
        });
        html += '</colgroup><thead><tr>';
        
        visibleKeys.forEach(key => {
            const styles = buildColumnStyles(columnConfigs[key]);
            html += `<th style="${styles.headerStyle}">${escapeHTML(key)}</th>`;
        });
        
        html += '</tr></thead><tbody>';
        
        sortedData.forEach(row => {
            html += '<tr class="static-row">';
            
            visibleKeys.forEach(key => {
                let cellClass = '';
                let cellContent = escapeHTML(row[key] ?? '');
                const styles = buildColumnStyles(columnConfigs[key]);
                
                if (key === '排名') {
                    cellClass = 'rank-cell';
                    if (row['排名变动']) {
                        if (row['排名变动'] === '上升') {
                            cellContent += '<span class="change-up">↑</span>';
                        } else if (row['排名变动'] === '下降') {
                            cellContent += '<span class="change-down">↓</span>';
                        }
                    }
                } else if (key === '财富(亿元人民币)') {
                    cellClass = 'wealth-cell';
                } else if (key === '主要公司') {
                    cellClass = 'company-cell';
                }
                
                html += `<td class="${cellClass}" style="${styles.cellStyle}">${cellContent}</td>`;
            });
            
            html += '</tr>';
        });
        
        html += '</tbody></table>';
        tableScrollWrapper.innerHTML = html;
        
        scheduleAdjustTableSizes(200);
    } catch (e) {
        console.error('Error rendering table:', e);
    }
}

function wait(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function getSupportedRecorderProfile() {
    const profiles = [
        { mimeType: 'video/mp4;codecs="avc1.42E01E,mp4a.40.2"', fileExtension: 'mp4' },
        { mimeType: 'video/mp4;codecs="avc1.42E01E"', fileExtension: 'mp4' },
        { mimeType: 'video/mp4', fileExtension: 'mp4' },
        { mimeType: 'video/webm;codecs=vp9', fileExtension: 'webm' },
        { mimeType: 'video/webm;codecs=vp8', fileExtension: 'webm' },
        { mimeType: 'video/webm', fileExtension: 'webm' }
    ];

    for (const profile of profiles) {
        if (MediaRecorder.isTypeSupported(profile.mimeType)) {
            return profile;
        }
    }

    return { mimeType: '', fileExtension: 'webm' };
}

function downloadVideo(blob, fileExtension = 'webm') {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    link.href = url;
    link.download = `dynamic-table-${timestamp}.${fileExtension}`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function drawRoundedRect(context, x, y, width, height, radius, fillColor, strokeColor, strokeWidth = 1) {
    const r = Math.max(0, Math.min(radius, Math.min(width, height) / 2));
    context.beginPath();
    context.moveTo(x + r, y);
    context.arcTo(x + width, y, x + width, y + height, r);
    context.arcTo(x + width, y + height, x, y + height, r);
    context.arcTo(x, y + height, x, y, r);
    context.arcTo(x, y, x + width, y, r);
    context.closePath();

    if (fillColor) {
        context.fillStyle = fillColor;
        context.fill();
    }
    if (strokeColor && strokeWidth > 0) {
        context.strokeStyle = strokeColor;
        context.lineWidth = strokeWidth;
        context.stroke();
    }
}

function clipRoundedRect(context, x, y, width, height, radius) {
    const r = Math.max(0, Math.min(radius, Math.min(width, height) / 2));
    context.beginPath();
    context.moveTo(x + r, y);
    context.arcTo(x + width, y, x + width, y + height, r);
    context.arcTo(x + width, y + height, x, y + height, r);
    context.arcTo(x, y + height, x, y, r);
    context.arcTo(x, y, x + width, y, r);
    context.closePath();
    context.clip();
}

function drawPanelBackground(context, width, height) {
    const makeLinear = (x0, y0, x1, y1, stops) => {
        const gradient = context.createLinearGradient(x0, y0, x1, y1);
        stops.forEach(stop => gradient.addColorStop(stop[0], stop[1]));
        return gradient;
    };
    const makeRadial = (x, y, radius, stops) => {
        const gradient = context.createRadialGradient(x, y, 0, x, y, radius);
        stops.forEach(stop => gradient.addColorStop(stop[0], stop[1]));
        return gradient;
    };

    if (currentBg === 1) {
        context.fillStyle = makeLinear(0, 0, width, height, [[0, '#020617'], [0.45, '#0b1530'], [1, '#0a1f3f']]);
        context.fillRect(0, 0, width, height);
        context.fillStyle = makeRadial(width * 0.15, height * 0.2, width * 0.55, [[0, 'rgba(14, 165, 233, 0.28)'], [1, 'rgba(14, 165, 233, 0)']]);
        context.fillRect(0, 0, width, height);
        context.fillStyle = makeRadial(width * 0.85, height * 0.8, width * 0.55, [[0, 'rgba(99, 102, 241, 0.22)'], [1, 'rgba(99, 102, 241, 0)']]);
    } else if (currentBg === 2) {
        context.fillStyle = makeLinear(0, 0, width, height, [[0, '#001b2e'], [0.34, '#003049'], [0.72, '#005f73'], [1, '#0a9396']]);
    } else if (currentBg === 3) {
        context.fillStyle = makeLinear(0, 0, width, height, [[0, '#1e1b4b'], [0.45, '#312e81'], [1, '#4c1d95']]);
        context.fillRect(0, 0, width, height);
        context.fillStyle = makeRadial(width * 0.2, height * 0.1, width * 0.5, [[0, 'rgba(217, 70, 239, 0.2)'], [1, 'rgba(217, 70, 239, 0)']]);
    } else if (currentBg === 4) {
        context.fillStyle = makeLinear(0, 0, width, height, [[0, '#f0fdf4'], [0.45, '#d1fae5'], [1, '#bae6fd']]);
    } else if (currentBg === 5) {
        context.fillStyle = makeLinear(0, 0, width, height, [[0, '#7c2d12'], [0.45, '#c2410c'], [1, '#f59e0b']]);
        context.fillRect(0, 0, width, height);
        context.fillStyle = makeRadial(width * 0.8, height * 0.15, width * 0.42, [[0, 'rgba(253, 186, 116, 0.45)'], [1, 'rgba(253, 186, 116, 0)']]);
    } else if (currentBg === 6) {
        context.fillStyle = makeLinear(0, 0, width, height, [[0, '#f8fafc'], [0.52, '#e2e8f0'], [1, '#cbd5e1']]);
    } else if (currentBg === 7) {
        context.fillStyle = makeLinear(0, 0, width, height, [[0, '#fff7ed'], [0.4, '#ffedd5'], [1, '#fde68a']]);
    } else if (currentBg === 8) {
        context.fillStyle = makeLinear(0, 0, width, height, [[0, '#052e16'], [0.46, '#14532d'], [1, '#166534']]);
        context.fillRect(0, 0, width, height);
        context.fillStyle = makeRadial(width * 0.18, height * 0.2, width * 0.52, [[0, 'rgba(74, 222, 128, 0.3)'], [1, 'rgba(74, 222, 128, 0)']]);
    } else if (currentBg === 9) {
        context.fillStyle = makeLinear(0, 0, width, height, [[0, '#020617'], [1, '#111827']]);
        context.fillRect(0, 0, width, height);
        context.strokeStyle = 'rgba(56, 189, 248, 0.08)';
        context.lineWidth = 1;
        for (let x = 0; x < width; x += 16) {
            context.beginPath();
            context.moveTo(x, 0);
            context.lineTo(x, height);
            context.stroke();
        }
        context.strokeStyle = 'rgba(99, 102, 241, 0.08)';
        for (let y = 0; y < height; y += 16) {
            context.beginPath();
            context.moveTo(0, y);
            context.lineTo(width, y);
            context.stroke();
        }
        return;
    } else {
        context.fillStyle = makeLinear(0, 0, width, height, [[0, '#111827'], [0.35, '#1f2937'], [1, '#b45309']]);
    }
    context.fillRect(0, 0, width, height);
}

function drawCustomBackgroundImage(context, width, height) {
    const image = customBackgroundImage.image;
    if (!image || !image.complete || !image.naturalWidth || !image.naturalHeight) return;
    const imageRatio = image.naturalWidth / image.naturalHeight;
    const targetRatio = width / height;
    let sx = 0;
    let sy = 0;
    let sw = image.naturalWidth;
    let sh = image.naturalHeight;
    if (imageRatio > targetRatio) {
        sw = image.naturalHeight * targetRatio;
        sx = (image.naturalWidth - sw) / 2;
    } else {
        sh = image.naturalWidth / targetRatio;
        sy = (image.naturalHeight - sh) / 2;
    }
    context.save();
    context.globalAlpha = customBackgroundImage.opacity;
    context.drawImage(image, sx, sy, sw, sh, 0, 0, width, height);
    context.restore();
}

function hasVisibleColor(color) {
    if (!color) return false;
    if (color === 'transparent') return false;
    if (color === 'rgba(0, 0, 0, 0)') return false;
    return true;
}

function drawTextInRect(context, text, rect, style, align = 'left', renderState = null, cacheSeed = '', stableHeight = null) {
    const paddingLeft = parseFloat(style.paddingLeft) || 0;
    const paddingRight = parseFloat(style.paddingRight) || 0;
    const paddingTop = parseFloat(style.paddingTop) || 0;
    const paddingBottom = parseFloat(style.paddingBottom) || 0;
    const fontSize = parseFloat(style.fontSize) || 12;
    const lineHeight = Number.isFinite(parseFloat(style.lineHeight)) ? parseFloat(style.lineHeight) : fontSize * 1.2;
    const shadowOffsetX = parseFloat(style.textShadow === 'none' ? '0' : (style.textShadow.match(/(-?\d+(\.\d+)?)px/g)?.[0] || '0')) || 0;
    const shadowOffsetY = parseFloat(style.textShadow === 'none' ? '0' : (style.textShadow.match(/(-?\d+(\.\d+)?)px/g)?.[1] || '0')) || 0;
    const shadowBlur = parseFloat(style.textShadow === 'none' ? '0' : (style.textShadow.match(/(-?\d+(\.\d+)?)px/g)?.[2] || '0')) || 0;
    const shadowColor = style.textShadow === 'none'
        ? 'rgba(0, 0, 0, 0)'
        : (style.textShadow.match(/rgba?\([^)]+\)|#[0-9a-fA-F]+/)?.[0] || 'rgba(0, 0, 0, 0)');
    context.font = `${style.fontWeight} ${fontSize}px ${style.fontFamily}`;
    context.fillStyle = style.color;
    context.shadowOffsetX = shadowOffsetX;
    context.shadowOffsetY = shadowOffsetY;
    context.shadowBlur = shadowBlur;
    context.shadowColor = shadowColor;
    context.textBaseline = 'middle';
    context.textAlign = align;

    if (align === 'center') {
        context.fillText(text, rect.x + rect.width / 2, rect.y + rect.height / 2);
        return;
    }

    const x = rect.x + Math.max(2, paddingLeft);
    const maxWidth = Math.max(10, Math.floor(rect.width - paddingLeft - paddingRight - 4));
    const availableHeightCurrent = Math.max(lineHeight, rect.height - paddingTop - paddingBottom);
    const availableHeightForLayout = Math.max(
        lineHeight,
        (stableHeight !== null ? stableHeight : rect.height) - paddingTop - paddingBottom
    );
    const stableMaxLines = Math.max(1, Math.floor(availableHeightForLayout / lineHeight));
    const normalizedText = (text || '').toString();
    const cacheKey = `${cacheSeed}|${normalizedText}|${style.fontWeight}|${fontSize}|${style.fontFamily}|${maxWidth}|${stableMaxLines}`;
    let lines = renderState && renderState.textLayoutCache ? renderState.textLayoutCache.get(cacheKey) : null;
    if (!lines) {
        const chars = Array.from(normalizedText);
        lines = [];
        let currentLine = '';

        for (const ch of chars) {
            const testLine = currentLine + ch;
            if (context.measureText(testLine).width <= maxWidth || currentLine.length === 0) {
                currentLine = testLine;
            } else {
                lines.push(currentLine);
                currentLine = ch;
            }
        }
        if (currentLine) {
            lines.push(currentLine);
        }
        if (renderState && renderState.textLayoutCache) {
            renderState.textLayoutCache.set(cacheKey, lines);
        }
    }

    context.save();
    context.beginPath();
    context.rect(rect.x, rect.y, rect.width, rect.height);
    context.clip();
    context.textBaseline = 'top';
    context.textAlign = 'left';
    const renderLineCount = Math.min(stableMaxLines, lines.length);
    const textBlockHeight = renderLineCount * lineHeight;
    const startY = rect.y + paddingTop + Math.max(0, (availableHeightForLayout - textBlockHeight) / 2);
    const maxVisibleY = rect.y + availableHeightCurrent + paddingTop;
    for (let i = 0; i < renderLineCount; i++) {
        const lineY = startY + i * lineHeight;
        if (lineY > maxVisibleY) break;
        context.fillText(lines[i], x, lineY, maxWidth);
    }
    context.restore();
}

function buildExportBarrageRenderCache(previewPanel) {
    const barrageLayer = previewPanel.querySelector('#barrageLayer');
    if (!barrageLayer) return null;
    const barrageItems = Array.from(barrageLayer.querySelectorAll('.barrage-item'));
    return {
        items: barrageItems.map(item => {
            const itemStyle = getComputedStyle(item);
            return {
                text: item.textContent || '',
                borderWidth: parseFloat(itemStyle.borderWidth) || 1,
                styleSnapshot: {
                    backgroundColor: itemStyle.backgroundColor,
                    borderColor: itemStyle.borderColor,
                    color: itemStyle.color,
                    fontSize: itemStyle.fontSize,
                    lineHeight: itemStyle.lineHeight,
                    fontWeight: itemStyle.fontWeight,
                    fontFamily: itemStyle.fontFamily,
                    paddingLeft: itemStyle.paddingLeft,
                    paddingRight: itemStyle.paddingRight,
                    paddingTop: itemStyle.paddingTop,
                    paddingBottom: itemStyle.paddingBottom,
                    textShadow: itemStyle.textShadow
                }
            };
        })
    };
}

function buildExportTableRenderCache(previewPanel) {
    const tableContainer = previewPanel.querySelector('#tableContainer');
    if (!tableContainer) return null;

    const containerStyle = getComputedStyle(tableContainer);
    const rowHeightDefault = parseFloat(containerStyle.getPropertyValue('--row-expand-max')) || null;
    const rowHeightLast = parseFloat(containerStyle.getPropertyValue('--row-expand-max-last')) || rowHeightDefault;

    const scrollWrapper = tableContainer.querySelector('#tableScrollWrapper');
    const table = scrollWrapper?.querySelector('table');
    const thead = table?.querySelector('thead');
    const tbody = table?.querySelector('tbody');
    if (!scrollWrapper || !table || !thead || !tbody) return null;

    const wrapperRectAbs = scrollWrapper.getBoundingClientRect();
    const headerRectAbs = thead.getBoundingClientRect();

    const clipRectAbs = {
        left: wrapperRectAbs.left,
        top: headerRectAbs.bottom,
        width: wrapperRectAbs.width,
        height: wrapperRectAbs.bottom - headerRectAbs.bottom
    };

    const headerCells = [];
    const headerRows = Array.from(thead.querySelectorAll('tr'));
    for (const rowEl of headerRows) {
        const cells = Array.from(rowEl.children);
        for (const cellEl of cells) {
            const cellRectAbs = cellEl.getBoundingClientRect();
            if (cellRectAbs.width <= 0 || cellRectAbs.height <= 0) continue;
            const cellStyle = getComputedStyle(cellEl);
            headerCells.push({
                cellIndex: cellEl.cellIndex || 0,
                text: cellEl.textContent || '',
                rectAbs: { left: cellRectAbs.left, top: cellRectAbs.top, width: cellRectAbs.width, height: cellRectAbs.height },
                styleSnapshot: {
                    backgroundColor: cellStyle.backgroundColor,
                    borderBottomWidth: cellStyle.borderBottomWidth,
                    borderBottomColor: cellStyle.borderBottomColor,
                    color: cellStyle.color,
                    fontSize: cellStyle.fontSize,
                    lineHeight: cellStyle.lineHeight,
                    fontWeight: cellStyle.fontWeight,
                    fontFamily: cellStyle.fontFamily,
                    paddingLeft: cellStyle.paddingLeft,
                    paddingRight: cellStyle.paddingRight,
                    paddingTop: cellStyle.paddingTop,
                    paddingBottom: cellStyle.paddingBottom,
                    textShadow: cellStyle.textShadow
                }
            });
        }
    }

    const bodyRows = [];
    const bodyRowEls = Array.from(tbody.querySelectorAll('tr'));
    bodyRowEls.forEach((rowEl, rowIndex) => {
        const rowRectAbs = rowEl.getBoundingClientRect();
        const rowStyle = getComputedStyle(rowEl);
        const cells = Array.from(rowEl.children);

        const baseRowWidth = rowRectAbs.width || 1;
        const baseRowHeight = rowRectAbs.height || 1;

        bodyRows.push({
            rowEl,
            rowIndex,
            baseRowRectAbs: { left: rowRectAbs.left, top: rowRectAbs.top, width: baseRowWidth, height: baseRowHeight },
            backgroundColor: rowStyle.backgroundColor,
            cells: cells.map((cellEl, cellIndex) => {
                const cellRectAbs = cellEl.getBoundingClientRect();
                return {
                    cellIndex,
                    text: cellEl.textContent || '',
                    // 保存相对行的偏移：用于在动画缩放/位移时快速恢复像素位置
                    offsetX: cellRectAbs.left - rowRectAbs.left,
                    offsetY: cellRectAbs.top - rowRectAbs.top,
                    width: cellRectAbs.width,
                    height: cellRectAbs.height,
                    styleSnapshot: (() => {
                        const cs = getComputedStyle(cellEl);
                        return {
                            backgroundColor: cs.backgroundColor,
                            borderBottomWidth: cs.borderBottomWidth,
                            borderBottomColor: cs.borderBottomColor,
                            color: cs.color,
                            fontSize: cs.fontSize,
                            lineHeight: cs.lineHeight,
                            fontWeight: cs.fontWeight,
                            fontFamily: cs.fontFamily,
                            paddingLeft: cs.paddingLeft,
                            paddingRight: cs.paddingRight,
                            paddingTop: cs.paddingTop,
                            paddingBottom: cs.paddingBottom,
                            textShadow: cs.textShadow
                        };
                    })()
                };
            })
        });
    });

    return {
        clipRectAbs,
        headerCells,
        bodyRows,
        rowHeightDefault,
        rowHeightLast,
        lastRowIndex: bodyRowEls.length - 1
    };
}

function drawBarrageLayerToCanvas(previewPanel, context, toLocalRect, renderState = null) {
    const barrageLayer = previewPanel.querySelector('#barrageLayer');
    if (!barrageLayer) return;
    const barrageCache = renderState && renderState.exportBarrageCache ? renderState.exportBarrageCache : null;
    const barrageItems = Array.from(barrageLayer.querySelectorAll('.barrage-item'));
    barrageItems.forEach((item, index) => {
        const itemRect = item.getBoundingClientRect();
        if (itemRect.width <= 0 || itemRect.height <= 0) return;
        const localRect = toLocalRect(itemRect);
        const cachedItem = barrageCache?.items?.[index];
        const itemStyle = cachedItem?.styleSnapshot || getComputedStyle(item);
        const borderWidth = cachedItem?.borderWidth ?? (parseFloat(itemStyle.borderWidth) || 1);
        drawRoundedRect(
            context,
            localRect.x,
            localRect.y,
            localRect.width,
            localRect.height,
            localRect.height / 2,
            itemStyle.backgroundColor,
            itemStyle.borderColor,
            borderWidth
        );
        const text = cachedItem?.text ?? (item.textContent || '');
        drawTextInRect(context, text, localRect, itemStyle, 'center', renderState, `barrage-${index}`);
    });
}

function drawPreviewPanelToCanvas(previewPanel, context, width, height, renderState = null, drawBackground = true) {
    const panelRect = previewPanel.getBoundingClientRect();
    const toLocalRect = rect => ({
        x: rect.left - panelRect.left,
        y: rect.top - panelRect.top,
        width: rect.width,
        height: rect.height
    });

    context.clearRect(0, 0, width, height);
    if (drawBackground) {
        const panelStyle = getComputedStyle(previewPanel);
        const panelRadius = parseFloat(panelStyle.borderRadius) || 12;
        context.save();
        clipRoundedRect(context, 0, 0, width, height, panelRadius);
        drawPanelBackground(context, width, height);
        drawCustomBackgroundImage(context, width, height);
        context.restore();
    }

    const titleWrapper = previewPanel.querySelector('.title-wrapper');
    const title = previewPanel.querySelector('.title');
    if (titleWrapper && title) {
        const titleRect = toLocalRect(titleWrapper.getBoundingClientRect());
        const titleWrapperStyle = getComputedStyle(titleWrapper);
        drawRoundedRect(
            context,
            titleRect.x,
            titleRect.y,
            titleRect.width,
            titleRect.height,
            parseFloat(titleWrapperStyle.borderRadius) || 8,
            titleWrapperStyle.backgroundColor,
            titleWrapperStyle.borderColor,
            parseFloat(titleWrapperStyle.borderWidth) || 0
        );

        const titleStyle = getComputedStyle(title);
        const titleTextRect = toLocalRect(title.getBoundingClientRect());
        drawTextInRect(context, title.textContent || '', titleTextRect, titleStyle, 'center', renderState, 'title');
    }

    const tableContainer = previewPanel.querySelector('#tableContainer');
    if (!tableContainer) return;

    const containerRect = toLocalRect(tableContainer.getBoundingClientRect());
    const containerStyle = getComputedStyle(tableContainer);
    const rowHeightDefault = parseFloat(containerStyle.getPropertyValue('--row-expand-max')) || null;
    const rowHeightLast = parseFloat(containerStyle.getPropertyValue('--row-expand-max-last')) || rowHeightDefault;
    const scrollWrapper = tableContainer.querySelector('#tableScrollWrapper');
    const table = scrollWrapper?.querySelector('table');
    const thead = table?.querySelector('thead');
    const tbody = table?.querySelector('tbody');
    
    drawRoundedRect(
        context,
        containerRect.x,
        containerRect.y,
        containerRect.width,
        containerRect.height,
        parseFloat(containerStyle.borderRadius) || 12,
        containerStyle.backgroundColor,
        containerStyle.borderColor,
        parseFloat(containerStyle.borderWidth) || 1
    );

    context.save();
    context.beginPath();
    context.rect(containerRect.x, containerRect.y, containerRect.width, containerRect.height);
    context.clip();

    // 绘制表头（固定在顶部）
    if (thead) {
        const headerRows = Array.from(thead.querySelectorAll('tr'));
        for (const row of headerRows) {
            const rowStyle = getComputedStyle(row);
            const rowOpacity = Number.parseFloat(rowStyle.opacity || '1');
            if (rowOpacity <= 0.02) continue;

            const cells = Array.from(row.children);
            for (const cell of cells) {
                const cellRect = cell.getBoundingClientRect();
                if (cellRect.width <= 0 || cellRect.height <= 0) continue;

                const localRect = toLocalRect(cellRect);
                const cellStyle = getComputedStyle(cell);

                context.save();
                context.globalAlpha = rowOpacity;

                if (hasVisibleColor(rowStyle.backgroundColor)) {
                    context.fillStyle = rowStyle.backgroundColor;
                    context.fillRect(localRect.x, localRect.y, localRect.width, localRect.height);
                }

                if (hasVisibleColor(cellStyle.backgroundColor)) {
                    context.fillStyle = cellStyle.backgroundColor;
                    context.fillRect(localRect.x, localRect.y, localRect.width, localRect.height);
                }

                if (cellStyle.borderBottomWidth !== '0px' && cellStyle.borderBottomColor) {
                    context.strokeStyle = cellStyle.borderBottomColor;
                    context.lineWidth = parseFloat(cellStyle.borderBottomWidth) || 1;
                    context.beginPath();
                    context.moveTo(localRect.x, localRect.y + localRect.height);
                    context.lineTo(localRect.x + localRect.width, localRect.y + localRect.height);
                    context.stroke();
                }

                const cellSeed = `header-${cell.cellIndex || 0}`;
                drawTextInRect(context, cell.textContent || '', localRect, cellStyle, 'left', renderState, cellSeed, cellRect.height);
                context.restore();
            }
        }
    }

    // 绘制表体（受滚动影响）
    if (tbody && scrollWrapper && thead) {
        const exportTableCache = renderState && renderState.exportTableCache ? renderState.exportTableCache : null;

        // 获取表头的实际位置，表体应该从表头下边缘开始裁剪
        const headerRect = thead.getBoundingClientRect();
        const wrapperRect = scrollWrapper.getBoundingClientRect();

        const fallbackClipLeft = wrapperRect.left - panelRect.left;
        const fallbackClipTop = headerRect.bottom - panelRect.top; // 从表头下边缘开始
        const fallbackClipWidth = wrapperRect.width;
        const fallbackClipHeight = wrapperRect.bottom - headerRect.bottom; // 表头下边缘到 wrapper 底部

        const useCacheClip = exportTableCache && exportTableCache.clipRectAbs && exportTableCache.clipRectAbs.height > 0;
        const clipRect = useCacheClip
            ? {
                x: exportTableCache.clipRectAbs.left - panelRect.left,
                y: exportTableCache.clipRectAbs.top - panelRect.top,
                w: exportTableCache.clipRectAbs.width,
                h: exportTableCache.clipRectAbs.height
            }
            : {
                x: fallbackClipLeft,
                y: fallbackClipTop,
                w: fallbackClipWidth,
                h: fallbackClipHeight
            };

        if (clipRect.h > 0) {
            context.save();
            context.beginPath();
            context.rect(clipRect.x, clipRect.y, clipRect.w, clipRect.h);
            context.clip();

            if (exportTableCache && exportTableCache.bodyRows && exportTableCache.bodyRows.length) {
                for (const cachedRow of exportTableCache.bodyRows) {
                    const rowEl = cachedRow.rowEl;
                    const rowRect = rowEl.getBoundingClientRect();
                    if (rowRect.width <= 0 || rowRect.height <= 0) continue;

                    const rowOpacity = Number.parseFloat(getComputedStyle(rowEl).opacity || '1');
                    if (rowOpacity <= 0.02) continue;

                    const baseW = cachedRow.baseRowRectAbs.width || 1;
                    const baseH = cachedRow.baseRowRectAbs.height || 1;
                    const scaleX = rowRect.width / baseW;
                    const scaleY = rowRect.height / baseH;

                    const rowAbsLeft = rowRect.left;
                    const rowAbsTop = rowRect.top;

                    const rowBgVisible = hasVisibleColor(cachedRow.backgroundColor);
                    const stableHeight = cachedRow.rowIndex === exportTableCache.lastRowIndex ? rowHeightLast : rowHeightDefault;

                    context.save();
                    context.globalAlpha = rowOpacity;

                    for (const cachedCell of cachedRow.cells) {
                        const cellAbsLeft = rowAbsLeft + cachedCell.offsetX * scaleX;
                        const cellAbsTop = rowAbsTop + cachedCell.offsetY * scaleY;

                        const localRect = {
                            x: cellAbsLeft - panelRect.left,
                            y: cellAbsTop - panelRect.top,
                            width: cachedCell.width * scaleX,
                            height: cachedCell.height * scaleY
                        };

                        if (rowBgVisible) {
                            context.fillStyle = cachedRow.backgroundColor;
                            context.fillRect(localRect.x, localRect.y, localRect.width, localRect.height);
                        }

                        const cellStyle = cachedCell.styleSnapshot;
                        if (hasVisibleColor(cellStyle.backgroundColor)) {
                            context.fillStyle = cellStyle.backgroundColor;
                            context.fillRect(localRect.x, localRect.y, localRect.width, localRect.height);
                        }

                        if (cellStyle.borderBottomWidth !== '0px' && cellStyle.borderBottomColor) {
                            context.strokeStyle = cellStyle.borderBottomColor;
                            context.lineWidth = parseFloat(cellStyle.borderBottomWidth) || 1;
                            context.beginPath();
                            context.moveTo(localRect.x, localRect.y + localRect.height);
                            context.lineTo(localRect.x + localRect.width, localRect.y + localRect.height);
                            context.stroke();
                        }

                        const cellSeed = `${cachedRow.rowIndex}-${cachedCell.cellIndex}`;
                        drawTextInRect(
                            context,
                            cachedCell.text,
                            localRect,
                            cellStyle,
                            'left',
                            renderState,
                            cellSeed,
                            stableHeight
                        );
                    }

                    context.restore();
                }

            } else {
                // 回退：保持原有 DOM + 逐单元格绘制逻辑
                const bodyRows = Array.from(tbody.querySelectorAll('tr'));
                for (const row of bodyRows) {
                    const rowStyle = getComputedStyle(row);
                    const rowOpacity = Number.parseFloat(rowStyle.opacity || '1');
                    if (rowOpacity <= 0.02) continue;

                    const cells = Array.from(row.children);
                    for (const cell of cells) {
                        const cellRect = cell.getBoundingClientRect();
                        if (cellRect.width <= 0 || cellRect.height <= 0) continue;

                        const localRect = toLocalRect(cellRect);
                        const cellStyle = getComputedStyle(cell);

                        context.save();
                        context.globalAlpha = rowOpacity;

                        if (hasVisibleColor(rowStyle.backgroundColor)) {
                            context.fillStyle = rowStyle.backgroundColor;
                            context.fillRect(localRect.x, localRect.y, localRect.width, localRect.height);
                        }

                        if (hasVisibleColor(cellStyle.backgroundColor)) {
                            context.fillStyle = cellStyle.backgroundColor;
                            context.fillRect(localRect.x, localRect.y, localRect.width, localRect.height);
                        }

                        if (cellStyle.borderBottomWidth !== '0px' && cellStyle.borderBottomColor) {
                            context.strokeStyle = cellStyle.borderBottomColor;
                            context.lineWidth = parseFloat(cellStyle.borderBottomWidth) || 1;
                            context.beginPath();
                            context.moveTo(localRect.x, localRect.y + localRect.height);
                            context.lineTo(localRect.x + localRect.width, localRect.y + localRect.height);
                            context.stroke();
                        }

                        const isLastBodyRow = tbody.lastElementChild === row;
                        const stableHeight = isLastBodyRow ? rowHeightLast : rowHeightDefault;
                        const rowClassSeed = `${row.rowIndex || 0}`;
                        const cellSeed = `${rowClassSeed}-${cell.cellIndex || 0}`;
                        drawTextInRect(context, cell.textContent || '', localRect, cellStyle, 'left', renderState, cellSeed, stableHeight);
                        context.restore();
                    }
                }
            }

            context.restore();
        }
    }
    context.restore();

    drawBarrageLayerToCanvas(previewPanel, context, toLocalRect, renderState);
}

async function startAutomaticRecorder(previewPanel, shouldRecord = true) {
    const rect = previewPanel.getBoundingClientRect();
    const width = Math.max(1, Math.floor(rect.width));
    const height = Math.max(1, Math.floor(rect.height));
    if (width <= 0 || height <= 0) {
        throw new Error('预览区域未准备完成');
    }

    const recordCanvas = document.createElement('canvas');

    const renderScale = Math.max(2, Math.min(3, window.devicePixelRatio || 1));
    recordCanvas.width = Math.round(width * renderScale);
    recordCanvas.height = Math.round(height * renderScale);
    const context = recordCanvas.getContext('2d', { alpha: false });
    if (!context) {
        throw new Error('无法初始化录制画布');
    }
    context.setTransform(renderScale, 0, 0, renderScale, 0, 0);

    const renderState = {
        textLayoutCache: new Map()
    };
    // 导出前一次性快照：减少逐帧 getComputedStyle/getBoundingClientRect 成本
    renderState.exportBarrageCache = buildExportBarrageRenderCache(previewPanel);
    renderState.exportTableCache = buildExportTableRenderCache(previewPanel);
    drawPreviewPanelToCanvas(previewPanel, context, width, height, renderState, true);

    let stream = null;
    let recorder = null;
    let mimeType = '';
    let fileExtension = 'webm';
    const chunks = [];
    let recorderStopped = Promise.resolve();

    if (shouldRecord) {
        if (typeof MediaRecorder === 'undefined') {
            throw new Error('当前浏览器不支持视频录制');
        }
        const fps = 60;
        stream = recordCanvas.captureStream(fps);
        const recorderProfile = getSupportedRecorderProfile();
        mimeType = recorderProfile.mimeType;
        fileExtension = recorderProfile.fileExtension;
        const recorderOptions = { videoBitsPerSecond: 12_000_000 };
        if (mimeType) {
            recorderOptions.mimeType = mimeType;
        }
        try {
            recorder = new MediaRecorder(stream, recorderOptions);
        } catch (e) {
            recorder = mimeType ? new MediaRecorder(stream, { mimeType }) : new MediaRecorder(stream);
        }

        recorder.ondataavailable = event => {
            if (event.data && event.data.size > 0) {
                chunks.push(event.data);
            }
        };

        recorderStopped = new Promise((resolve, reject) => {
            recorder.onstop = resolve;
            recorder.onerror = event => reject(event.error || new Error('录制失败'));
        });
    }

    let rafId = null;
    const drawFrame = () => {
        try {
            drawPreviewPanelToCanvas(previewPanel, context, width, height, renderState, true);
        } catch (e) {
            console.error('Error drawing recording frame:', e);
        }
        rafId = requestAnimationFrame(drawFrame);
    };
    rafId = requestAnimationFrame(drawFrame);

    if (recorder) {
        recorder.start(100);
    }

    return {
        mimeType,
        fileExtension,
        chunks,
        recorderStopped,
        stop: async () => {
            if (rafId) {
                cancelAnimationFrame(rafId);
                rafId = null;
            }
            if (recorder && recorder.state !== 'inactive') {
                recorder.stop();
            }
            await recorderStopped;
            if (stream) {
                stream.getTracks().forEach(track => track.stop());
            }
        }
    };
}

function runTableAnimation() {
    return new Promise((resolve, reject) => {
        try {
            const rows = document.querySelectorAll('#tableScrollWrapper tbody tr');
            if (rows.length === 0) {
                resolve();
                return;
            }

            tableScrollY = 0;

            const interval = parseInt(document.getElementById('animationInterval').value) || 300;
            const animationClass = getSelectedAnimationClass();
            const scrollWrapper = document.getElementById('tableScrollWrapper');
            const table = scrollWrapper.querySelector('table');
            const tbody = table.querySelector('tbody');
            const containerHeight = scrollWrapper.clientHeight;

            const thead = table.querySelector('thead');
            const headerHeight = thead ? thead.getBoundingClientRect().height : 0;
            const totalRows = rows.length;
            const visibleBodyHeight = containerHeight - headerHeight;

            rows.forEach(row => {
                rowAnimationClasses.forEach(typeClass => row.classList.remove(typeClass));
                row.classList.remove('expanded-row');
                row.classList.add(animationClass);
                row.classList.add('collapsed-row');
            });

            // 使用 CSS 里稳定的行高，避免在动画/transform 期间测量 DOM 导致亚像素抖动
            const tableContainer = document.getElementById('tableContainer');
            const tableContainerStyle = tableContainer ? getComputedStyle(tableContainer) : null;

            let bodyRowHeight = tableContainerStyle
                ? parseFloat(tableContainerStyle.getPropertyValue('--row-expand-max'))
                : NaN;
            let lastRowHeight = tableContainerStyle
                ? parseFloat(tableContainerStyle.getPropertyValue('--row-expand-max-last'))
                : NaN;

            if (!Number.isFinite(bodyRowHeight) || bodyRowHeight <= 0 || !Number.isFinite(lastRowHeight) || lastRowHeight <= 0) {
                const firstRowRect = rows[0]?.getBoundingClientRect?.();
                const lastRowRect = rows[rows.length - 1]?.getBoundingClientRect?.();
                if (firstRowRect && Number.isFinite(firstRowRect.height)) {
                    bodyRowHeight = firstRowRect.height;
                } else {
                    bodyRowHeight = Math.max(1, Math.floor(visibleBodyHeight / Math.max(1, totalRows)));
                }
                lastRowHeight = lastRowRect && Number.isFinite(lastRowRect.height) ? lastRowRect.height : bodyRowHeight;
            }

            const totalBodyHeight = Math.max(0, (totalRows - 1) * bodyRowHeight + lastRowHeight);
            const maxVisibleRows = Math.max(1, Math.floor(visibleBodyHeight / bodyRowHeight));
            const totalScrollNeeded = Math.max(0, totalBodyHeight - visibleBodyHeight);

            tbody.style.transform = 'translate3d(0, 0, 0)';
            tbody.style.willChange = 'transform';

            let currentIndex = 0;
            let resolveCalled = false;
            let animationStartTime = null;
            let scrollStartTime = null;
            let isScrolling = false;

            if (animationInterval) {
                clearInterval(animationInterval);
                animationInterval = null;
            }

            const animate = (timestamp) => {
                if (!animationStartTime) {
                    animationStartTime = timestamp;
                }
                
                const elapsed = timestamp - animationStartTime;
                const expectedIndex = Math.floor(elapsed / interval);
                
                // 展开行
                while (currentIndex < rows.length && currentIndex <= expectedIndex) {
                    rows[currentIndex].classList.remove('collapsed-row');
                    rows[currentIndex].classList.add('expanded-row');
                    currentIndex++;
                }
                
                // 当展开的行数超过可见区域时，开始连续滚动
                if (currentIndex > maxVisibleRows && !isScrolling) {
                    isScrolling = true;
                    scrollStartTime = timestamp;
                }
                
                // 连续平滑滚动
                if (isScrolling && scrollStartTime) {
                    const scrollElapsed = timestamp - scrollStartTime;
                    // 滚动速度：每 interval 毫秒滚动一行高度
                    const scrollSpeed = bodyRowHeight / interval;
                    const scrollYRaw = Math.min(scrollElapsed * scrollSpeed, totalScrollNeeded);
                    // 量化到整数像素，避免亚像素渲染导致“整体抖动”
                    const currentScrollY = Math.max(0, Math.min(Math.round(scrollYRaw), totalScrollNeeded));
                    tableScrollY = currentScrollY;
                    tbody.style.transform = `translate3d(0, -${currentScrollY}px, 0)`;
                }
                
                // 检查是否完成
                if (currentIndex < rows.length) {
                    requestAnimationFrame(animate);
                } else if (isScrolling) {
                    // 所有行已展开，继续滚动直到完成
                    const scrollElapsed = timestamp - scrollStartTime;
                    const scrollSpeed = bodyRowHeight / interval;
                    const scrollYRaw = Math.min(scrollElapsed * scrollSpeed, totalScrollNeeded);
                    const currentScrollY = Math.max(0, Math.min(Math.round(scrollYRaw), totalScrollNeeded));
                    tableScrollY = currentScrollY;
                    
                    if (currentScrollY < totalScrollNeeded) {
                        tbody.style.transform = `translate3d(0, -${currentScrollY}px, 0)`;
                        requestAnimationFrame(animate);
                    } else {
                        // 动画完成，等待后重置
                        setTimeout(() => {
                            tbody.style.transition = 'transform 0.8s cubic-bezier(0.25, 0.1, 0.25, 1)';
                            tbody.style.transform = 'translate3d(0, 0, 0)';
                            tableScrollY = 0;
                            setTimeout(() => {
                                tbody.style.transition = 'none';
                                tbody.style.willChange = '';
                                if (!resolveCalled) {
                                    resolveCalled = true;
                                    resolve();
                                }
                            }, 800);
                        }, 2000);
                    }
                } else {
                    // 没有滚动需求，直接完成
                    setTimeout(() => {
                        tbody.style.transition = 'transform 0.8s cubic-bezier(0.25, 0.1, 0.25, 1)';
                        tbody.style.transform = 'translate3d(0, 0, 0)';
                        tableScrollY = 0;
                        setTimeout(() => {
                            tbody.style.transition = 'none';
                            tbody.style.willChange = '';
                            if (!resolveCalled) {
                                resolveCalled = true;
                                resolve();
                            }
                        }, 800);
                    }, 2000);
                }
            };
            
            requestAnimationFrame(animate);
        } catch (e) {
            if (animationInterval) {
                clearInterval(animationInterval);
                animationInterval = null;
            }
            reject(e);
        }
    });
}

async function playAnimation() {
    if (isAnimating || isExporting) return;

    const playBtn = document.getElementById('playAnimationBtn');
    const exportBtn = document.getElementById('exportAnimationBtn');
    playBtn.disabled = true;
    playBtn.textContent = '播放中...';
    if (exportBtn) {
        exportBtn.disabled = true;
    }

    try {
        // 添加一个全局类名阻止布局变化
        document.body.classList.add('animation-running');
        isAnimating = true; // 提前设置状态，以便后续流程正确判断
        
        renderTable();
        updatePreview();
        await wait(120);
        adjustTableSizes();
        await wait(80);
        await runTableAnimation();
        await wait(120);
        
        // 动画结束，清理弹幕
        const barrageLayer = document.getElementById('barrageLayer');
        if (barrageLayer) barrageLayer.innerHTML = '';
    } catch (e) {
        console.error('Error playing animation:', e);
    } finally {
        isAnimating = false; // 重置状态
        document.body.classList.remove('animation-running');
        playBtn.disabled = false;
        playBtn.textContent = '动画预览';
        if (exportBtn && !isExporting) {
            exportBtn.disabled = false;
        }
        // 重置表格位置
        const tableScrollWrapper = document.getElementById('tableScrollWrapper');
        const table = tableScrollWrapper?.querySelector('table');
        if (table) {
            table.style.transition = '';
            table.style.transform = 'translateY(0)';
        }
        tableScrollY = 0;
    }
}

async function exportAnimationVideo() {
    if (isAnimating || isExporting) return;

    const playBtn = document.getElementById('playAnimationBtn');
    const exportBtn = document.getElementById('exportAnimationBtn');

    isExporting = true;
    playBtn.disabled = true;
    exportBtn.disabled = true;
    exportBtn.textContent = '准备录制...';

    let recordingSession = null;

    try {
        // 添加一个全局类名阻止布局变化
        document.body.classList.add('animation-running');
        isExporting = true; // 提前设置状态
        
        renderTable();
        updatePreview();
        await wait(120);
        adjustTableSizes();

        const previewPanel = document.getElementById('previewPanel');
        recordingSession = await startAutomaticRecorder(previewPanel, true);
        exportBtn.textContent = '录制中...';

        await wait(80);
        await runTableAnimation();
        await wait(220);

        const recordedChunks = recordingSession.chunks;
        const recordedMimeType = recordingSession.mimeType;
        const recordedFileExtension = recordingSession.fileExtension || 'webm';
        await recordingSession.stop();
        recordingSession = null;
        if (recordedChunks.length === 0) {
            throw new Error('录制结果为空，请重试');
        }

        const videoBlob = new Blob(recordedChunks, { type: recordedMimeType || 'video/webm' });
        downloadVideo(videoBlob, recordedFileExtension);
        if (recordedFileExtension !== 'mp4') {
            alert('当前浏览器不支持MP4编码，已自动导出为WebM格式。');
        }

        exportBtn.textContent = '导出成功';
        await wait(600);
        
        // 导出结束，清理弹幕
        const barrageLayer = document.getElementById('barrageLayer');
        if (barrageLayer) barrageLayer.innerHTML = '';
    } catch (e) {
        console.error('Error exporting animation video:', e);
        alert(`视频导出失败：${e.message || '请重试'}`);
    } finally {
        document.body.classList.remove('animation-running');
        if (recordingSession) {
            try {
                await recordingSession.stop();
            } catch (stopError) {
                console.error('Error stopping recording session:', stopError);
            }
        }

        isExporting = false;
        isAnimating = false;
        playBtn.disabled = false;
        playBtn.textContent = '动画预览';
        exportBtn.disabled = false;
        exportBtn.textContent = '导出视频';
        // 重置表格位置
        const tableScrollWrapper = document.getElementById('tableScrollWrapper');
        const table = tableScrollWrapper?.querySelector('table');
        if (table) {
            table.style.transition = '';
            table.style.transform = 'translateY(0)';
        }
        tableScrollY = 0;
    }
}

document.addEventListener('DOMContentLoaded', init);
