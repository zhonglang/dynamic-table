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
let isAnimating = false;
let isExporting = false;
const customBackgroundImage = {
    dataUrl: '',
    image: null,
    opacity: 0.45
};
const rowAnimationClasses = [
    'anim-type-expand',
    'anim-type-slide-left',
    'anim-type-slide-right',
    'anim-type-rise-up',
    'anim-type-zoom-in'
];

function init() {
    setupEventListeners();
    renderTable();
    updateBackground();
    updatePreview();
    setTimeout(adjustTableSizes, 300);
    setTimeout(adjustTableSizes, 600);
}

function setupEventListeners() {
    document.getElementById('titleContent').addEventListener('input', updatePreview);
    document.getElementById('titleSize').addEventListener('input', updatePreview);
    document.getElementById('titleColor').addEventListener('input', updatePreview);
    document.getElementById('tableHeaderSize').addEventListener('input', () => setTimeout(adjustTableSizes, 50));
    
    const dataInput = document.getElementById('dataInput');
    dataInput.addEventListener('input', handleDataChange);
    dataInput.addEventListener('change', handleDataChange);

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
}

function handleDataChange() {
    if (renderTimeout) {
        clearTimeout(renderTimeout);
    }
    renderTimeout = setTimeout(() => {
        renderTable();
    }, 100);
}

function updatePreview() {
    const title = document.getElementById('titleContent').value;
    const size = document.getElementById('titleSize').value;
    const color = document.getElementById('titleColor').value;

    const previewTitle = document.getElementById('previewTitle');
    previewTitle.textContent = title;
    previewTitle.style.fontSize = size + 'px';
    previewTitle.style.color = color;
    
    setTimeout(adjustTableSizes, 200);
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

function adjustTableSizes() {
    try {
        const data = JSON.parse(document.getElementById('dataInput').value);
        if (!Array.isArray(data) || data.length === 0) return;

        const rowCount = data.length;
        const tableContainer = document.getElementById('tableContainer');
        const table = tableContainer.querySelector('table');
        
        if (!tableContainer || !table) return;
        
        const containerHeight = tableContainer.clientHeight;
        
        if (containerHeight <= 0) {
            setTimeout(adjustTableSizes, 50);
            return;
        }

        const thead = table.querySelector('thead');
        const headerHeight = thead ? Math.ceil(thead.getBoundingClientRect().height) : 0;
        const rowDividerTotal = Math.max(0, rowCount - 1);
        const availableBodyHeight = Math.max(0, containerHeight - headerHeight - rowDividerTotal);
        const bodyRowHeight = Math.max(1, Math.floor(availableBodyHeight / rowCount));
        const bodyHeightUsed = bodyRowHeight * rowCount;
        const lastRowHeight = Math.max(1, bodyRowHeight + (availableBodyHeight - bodyHeightUsed));
        
        let fontSize, paddingVertical, thFontSize;
        
        if (rowCount <= 8) {
            fontSize = 13;
            paddingVertical = 6;
            thFontSize = 14;
        } else if (rowCount <= 10) {
            fontSize = 12;
            paddingVertical = 5;
            thFontSize = 13;
        } else if (rowCount <= 15) {
            fontSize = 11;
            paddingVertical = 4;
            thFontSize = 12;
        } else if (rowCount <= 20) {
            fontSize = 10;
            paddingVertical = 3;
            thFontSize = 11;
        } else {
            fontSize = 9;
            paddingVertical = 2;
            thFontSize = 10;
        }

        const customHeaderSize = parseInt(document.getElementById('tableHeaderSize')?.value || '0', 10);
        if (Number.isFinite(customHeaderSize) && customHeaderSize > 0) {
            thFontSize = customHeaderSize;
        }
        
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
                height: 100%;
                table-layout: fixed;
                border-collapse: collapse;
            }
            tbody tr.static-row,
            tbody tr.expanded-row {
                height: ${bodyRowHeight}px;
                max-height: ${bodyRowHeight}px;
                box-sizing: border-box;
            }
            tbody tr.static-row td,
            tbody tr.expanded-row td {
                height: ${bodyRowHeight}px;
                max-height: ${bodyRowHeight}px;
                box-sizing: border-box;
                overflow: hidden;
            }
            tbody tr.static-row:last-child,
            tbody tr.expanded-row:last-child {
                height: ${lastRowHeight}px;
                max-height: ${lastRowHeight}px;
            }
            tbody tr.static-row:last-child td,
            tbody tr.expanded-row:last-child td {
                height: ${lastRowHeight}px;
                max-height: ${lastRowHeight}px;
            }
            th {
                font-size: ${thFontSize}px !important;
                padding: 6px 4px !important;
                line-height: 1.2;
                box-sizing: border-box;
                vertical-align: middle;
            }
            td {
                font-size: ${fontSize}px !important;
                padding: ${paddingVertical}px 4px !important;
                line-height: 1.2;
                box-sizing: border-box;
                vertical-align: middle;
            }
            .rank-cell {
                font-size: ${fontSize + 2}px !important;
            }
            .wealth-cell {
                font-size: ${fontSize + 1}px !important;
            }
            .change-up, .change-down {
                font-size: ${Math.max(8, fontSize)}px !important;
            }
        `;
        
        const oldStyle = document.getElementById('dynamic-table-styles');
        if (oldStyle) {
            oldStyle.remove();
        }
        document.head.appendChild(style);
    } catch (e) {
        console.error('Error adjusting table sizes:', e);
    }
}

function renderTable() {
    try {
        const data = JSON.parse(document.getElementById('dataInput').value);
        if (!Array.isArray(data) || data.length === 0) return;

        const keys = Object.keys(data[0]);
        const tableContainer = document.getElementById('tableContainer');
        
        let html = '<table><thead><tr>';
        
        keys.forEach(key => {
            if (key !== '排名变动') {
                html += `<th>${key}</th>`;
            }
        });
        
        html += '</tr></thead><tbody>';
        
        data.forEach((row, index) => {
            html += '<tr class="static-row">';
            
            keys.forEach(key => {
                if (key === '排名变动') return;
                
                let cellClass = '';
                let cellContent = row[key];
                
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
                
                html += `<td class="${cellClass}">${cellContent}</td>`;
            });
            
            html += '</tr>';
        });
        
        html += '</tbody></table>';
        tableContainer.innerHTML = html;
        
        setTimeout(adjustTableSizes, 200);
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
    const tbody = tableContainer.querySelector('tbody');
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

    const rows = Array.from(tableContainer.querySelectorAll('tr'));
    for (const row of rows) {
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

            const isBodyRow = tbody ? row.parentElement === tbody : false;
            const isLastBodyRow = isBodyRow && tbody.lastElementChild === row;
            const stableHeight = isBodyRow ? (isLastBodyRow ? rowHeightLast : rowHeightDefault) : null;
            const rowClassSeed = `${row.rowIndex || 0}`;
            const cellSeed = `${rowClassSeed}-${cell.cellIndex || 0}`;
            drawTextInRect(context, cell.textContent || '', localRect, cellStyle, 'left', renderState, cellSeed, stableHeight);
            context.restore();
        }
    }
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
            const rows = document.querySelectorAll('#tableContainer tbody tr');
            if (rows.length === 0) {
                resolve();
                return;
            }

            const interval = parseInt(document.getElementById('animationInterval').value) || 300;
            const animationClass = getSelectedAnimationClass();
            isAnimating = true;

            rows.forEach(row => {
                rowAnimationClasses.forEach(typeClass => row.classList.remove(typeClass));
                row.classList.remove('expanded-row');
                row.classList.add(animationClass);
                row.classList.add('collapsed-row');
            });

            let currentIndex = 0;

            if (animationInterval) {
                clearInterval(animationInterval);
            }

            animationInterval = setInterval(() => {
                if (currentIndex < rows.length) {
                    rows[currentIndex].classList.remove('collapsed-row');
                    rows[currentIndex].classList.add('expanded-row');
                    currentIndex++;
                    return;
                }

                clearInterval(animationInterval);
                animationInterval = null;
                isAnimating = false;
                resolve();
            }, interval);
        } catch (e) {
            if (animationInterval) {
                clearInterval(animationInterval);
                animationInterval = null;
            }
            isAnimating = false;
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
        renderTable();
        updatePreview();
        await wait(120);
        adjustTableSizes();
        await wait(80);
        await runTableAnimation();
        await wait(120);
    } catch (e) {
        console.error('Error playing animation:', e);
    } finally {
        playBtn.disabled = false;
        playBtn.textContent = '动画预览';
        if (exportBtn && !isExporting) {
            exportBtn.disabled = false;
        }
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
    } catch (e) {
        console.error('Error exporting animation video:', e);
        alert(`视频导出失败：${e.message || '请重试'}`);
    } finally {
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
    }
}

document.addEventListener('DOMContentLoaded', init);
