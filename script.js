document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const previewSection = document.getElementById('preview-section');
    const labelContainer = document.getElementById('label-container');
    const printArea = document.getElementById('print-area');

    // Drag and Drop listeners
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => dropZone.classList.add('dragover'), false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => dropZone.classList.remove('dragover'), false);
    });

    dropZone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        handleFiles(files);
    });

    fileInput.addEventListener('change', (e) => {
        handleFiles(e.target.files);
    });

    function handleFiles(files) {
        if (files.length > 0) {
            const file = files[0];
            const reader = new FileReader();

            reader.onload = (e) => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { raw: false });

                processData(jsonData);
            };

            reader.readAsArrayBuffer(file);
        }
    }

    function formatDate(dateVal) {
        if (!dateVal) return '-';
        
        // Convert to string and clean up
        let dateStr = String(dateVal).trim();
        if (dateStr === '') return '-';

        // 1. Handle Excel serial numbers if they come in as numbers
        if (!isNaN(dateStr) && dateStr.length < 8) {
            // XLS dates are handled by XLSX.read({cellDates: true}), 
            // but just in case raw numbers slip through
            const d = new Date((dateVal - 25569) * 86400 * 1000);
            if (!isNaN(d.getTime())) return formatOutput(d);
        }

        // 2. Handle 8-digit strings: 20240313 -> 2024-03-13
        if (/^\d{8}$/.test(dateStr)) {
            return `${dateStr.substring(0, 4)}-${dateStr.substring(4, 6)}-${dateStr.substring(6, 8)}`;
        }

        // 3. Handle 6-digit strings: 240313 -> 2024-03-13
        if (/^\d{6}$/.test(dateStr)) {
            return `20${dateStr.substring(0, 2)}-${dateStr.substring(2, 4)}-${dateStr.substring(4, 6)}`;
        }

        // 4. Standardize delimiters: dots/slashes to dashes
        dateStr = dateStr.replace(/[\.\/]/g, '-');

        // 5. Try standard Date parsing
        const d = new Date(dateStr);
        if (!isNaN(d.getTime())) {
            return formatOutput(d);
        }

        return dateStr; // Return as is if all attempts fail

        function formatOutput(dateObj) {
            const year = dateObj.getFullYear();
            const month = String(dateObj.getMonth() + 1).padStart(2, '0');
            const day = String(dateObj.getDate()).padStart(2, '0');
            return `${year}-${month}-${day}`;
        }
    }

    function processData(data) {
        labelContainer.innerHTML = '';
        printArea.innerHTML = '';
        previewSection.classList.remove('hidden');

        if (!data || data.length === 0) {
            labelContainer.innerHTML = '<div style="color:red; padding:20px;">엑셀 파일에서 데이터를 읽지 못했습니다.</div>';
            return;
        }

        let addedCount = 0;

        data.forEach((row, index) => {
            const productType = findColumn(row, ['품명', '품목']) || '-';
            const assetName = findColumn(row, ['자산명', '자산번호', '자산']) || '-';
            const acquisitionDateStr = findColumn(row, ['취득일자', '취득일', '날짜']);
            const acquisitionDate = acquisitionDateStr ? formatDate(acquisitionDateStr) : '-';
            const barcodeNumber = String(findColumn(row, ['바코드', 'barcode', 'qr', '일련번호']) || '').trim();

            if (barcodeNumber || assetName !== '-') {
                createLabel(productType, assetName, acquisitionDate, barcodeNumber, index);
                addedCount++;
            }
        });

        if (addedCount === 0) {
            const headers = Object.keys(data[0] || {}).join(', ');
            labelContainer.innerHTML = `<div style="color:#ef4444; padding:20px; background:#fee2e2; border-radius:8px;">
                <strong>라벨 생성 실패</strong><br><br>
                방금 엑셀에서 인식된 열(컬럼) 이름: <strong>[ ${headers || '없음'} ]</strong><br><br>
                엑셀 파일의 <strong>첫 번째 줄(1행)</strong>에 '자산번호', '품명', '바코드' 등의 제목이 제대로 적혀 있는지 확인해 주세요.<br>
                (예: 표 위에 제목이나 빈 칸이 있으면 지우고 다시 저장해 주세요.)
            </div>`;
        }
    }

    function findColumn(row, possibleNames) {
        const keys = Object.keys(row);
        for (let key of keys) {
            const cleanKey = key.replace(/\s+/g, '').toLowerCase();
            for (let name of possibleNames) {
                if (cleanKey.includes(name.replace(/\s+/g, '').toLowerCase())) {
                    return row[key];
                }
            }
        }
        return null;
    }

    function createLabel(productType, assetName, acquisitionDate, barcodeNumber, index) {
        const id = `qr-${index}`;
        const printId = `qr-print-${index}`;
        const cardId = `card-${index}`;
        const printCardId = `print-card-${index}`;

        // UI Preview Card (Clean, only values)
        const card = document.createElement('div');
        card.className = 'label-card';
        card.id = cardId;
        card.innerHTML = `
            <div style="display: flex; flex-direction: column; width: 100%; font-family: '맑은 고딕', 'Malgun Gothic', sans-serif; font-size: 13px; font-weight: normal; color: #000;">
                <div style="display: flex; justify-content: space-between; align-items: flex-start; border-bottom: 1px solid #eee; padding-bottom: 8px; margin-bottom: 8px;">
                    <div class="label-info" style="flex: 1;">
                        <div style="display: flex; gap: 15px; margin-bottom: 4px;">
                            <div class="label-title" style="font-weight: normal; color: #000; font-size: 13px;">${assetName}</div>
                            <div style="font-size: 13px; color: #000; font-weight: normal;">${acquisitionDate}</div>
                        </div>
                        <div style="font-size: 13px; font-weight: normal; color: #000;">${productType}</div>
                    </div>
                    <button class="btn-remove" onclick="removeLabel('${cardId}', '${printCardId}')" title="이 라벨 제외하기">
                        <svg viewBox="0 0 24 24" width="20" height="20" stroke="currentColor" stroke-width="2" fill="none"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>
                    </button>
                </div>
                <div style="display: flex; justify-content: flex-end; align-items: center; padding-top: 8px;">
                    <div style="display: flex; flex-direction: row; align-items: center; gap: 12px;">
                        <div style="font-size: 13px; font-weight: normal; color: #000;">${barcodeNumber}</div>
                        <div id="${id}" class="label-qr"></div>
                    </div>
                </div>
            </div>
        `;
        labelContainer.appendChild(card);

        // Print Card (Blank spaces for pre-printed label alignment)
        const printCard = document.createElement('div');
        printCard.className = 'label-print-wrapper print-only'; 
        printCard.id = printCardId;
        printCard.innerHTML = `
            <div class="label-print-container">
                <div class="top-spacer"></div>
                <table class="label-table">
                    <tr class="row-6mm">
                        <td class="cell-label w-18mm"></td> <!-- 자산번호 pre-printed -->
                        <td class="cell-value w-28mm">${assetName}</td>
                        <td class="cell-label w-18mm"></td> <!-- 취득일자 pre-printed -->
                        <td class="cell-value w-19mm">${acquisitionDate}</td>
                    </tr>
                    <tr class="row-6mm">
                        <td class="cell-label w-18mm"></td> <!-- 품명 pre-printed -->
                        <td class="cell-value" colspan="3">${productType}</td>
                    </tr>
                    <tr class="row-9mm">
                        <td colspan="4" style="position: relative; padding: 0;">
                            <!-- Logo is pre-printed on the label paper -->
                            <!-- Absolute positioning for exact bottom-right alignment of barcode and QR -->
                            <div style="position: absolute; right: 0mm; bottom: -1.5mm; display: flex; align-items: center; gap: 4mm;">
                                <div class="barcode-text" style="font-size: 13px !important; font-weight: normal !important; color: #000 !important;">${barcodeNumber}</div>
                                <div id="${printId}" class="qr-code-box"></div>
                            </div>
                        </td>
                    </tr>
                </table>
            </div>
        `;
        printArea.appendChild(printCard);

        // Generate QRs
        new QRCode(document.getElementById(id), {
            text: barcodeNumber || 'N/A',
            width: 80,
            height: 80,
            correctLevel: QRCode.CorrectLevel.H
        });

        // For print (Fixed 15.5x15.5mm area, generate at 256px for sharp scaling)
        // Using Level L (7%) to ensure the absolute largest module size for hardware scanners
        new QRCode(document.getElementById(printId), {
            text: barcodeNumber || 'N/A',
            width: 256,
            height: 256,
            colorDark : "#000000",
            colorLight : "#ffffff",
            correctLevel : QRCode.CorrectLevel.L
        });
    }

    window.removeLabel = function(cardId, printCardId) {
        const card = document.getElementById(cardId);
        const printCard = document.getElementById(printCardId);
        
        if (card) {
            // Add a quick fade out animation before removing
            card.style.opacity = '0';
            card.style.transform = 'scale(0.95)';
            card.style.transition = 'all 0.3s ease';
            setTimeout(() => card.remove(), 300);
        }
        if (printCard) {
            printCard.remove();
        }
    };

    // Kiosk Mode Command Copy Function
    window.copyKioskCommand = function() {
        const command = "chrome.exe --kiosk-printing";
        navigator.clipboard.writeText(command).then(() => {
            const btn = document.querySelector('#kiosk-modal .btn-primary');
            const originalText = btn.innerText;
            btn.innerText = "명령어가 복사되었습니다!";
            btn.style.background = "#059669";
            setTimeout(() => {
                btn.innerText = originalText;
                btn.style.background = "#1e293b";
            }, 2000);
        });
    };

    // Show Kiosk Tooltip Modal
    document.querySelector('.kiosk-tip')?.addEventListener('click', () => {
        document.getElementById('kiosk-modal').style.display = 'flex';
    });
    document.querySelector('.kiosk-tip')?.style.setProperty('cursor', 'pointer');

    // Sequential Print Function (Fix for skipped pages via Single-Label Hub)
    window.sequentialPrint = async function() {
        const labels = document.querySelectorAll('.label-print-wrapper');
        const overlay = document.getElementById('print-overlay');
        const statusText = document.getElementById('print-status-text');
        const progressBar = document.getElementById('print-progress-bar');
        const printArea = document.getElementById('print-area');

        if (labels.length === 0) {
            alert('인쇄할 라벨이 없습니다.');
            return;
        }

        const confirmMsg = `${labels.length}개의 라벨을 약 3초 간격으로 '자동 순차 인쇄'합니다.\n\n* 중요: 브라우저가 '키오스크 프린팅' 모드여야 클릭 없이 진행됩니다. 계속하시겠습니까?`;
        if (!confirm(confirmMsg)) {
            return;
        }

        // Show Progress Overlay
        overlay.style.display = 'flex';
        document.body.classList.add('sequential-print-active');

        for (let i = 0; i < labels.length; i++) {
            const originalLabel = labels[i];
            const currentCount = i + 1;
            const progress = (currentCount / labels.length) * 100;

            // Update UI
            statusText.innerText = `라벨 인쇄 중 (${currentCount} / ${labels.length})`;
            progressBar.style.width = `${progress}%`;

            // CRITICAL: Clean injection method
            // 1. Clear previous print job
            printArea.innerHTML = '';
            
            // 2. Clone the label to avoid moving original DOM elements from the layout
            const labelClone = originalLabel.cloneNode(true);
            
            // Remove the problematic class that might trigger hiding, or style it directly
            labelClone.classList.remove('print-only'); 
            labelClone.style.display = 'block'; 
            labelClone.style.visibility = 'visible';
            labelClone.style.opacity = '1';
            
            printArea.appendChild(labelClone);
            
            // 3. Wait for rendering to settle (QR, Table, Fonts)
            void labelClone.offsetHeight; 
            await new Promise(r => setTimeout(r, 1000)); // 1s Rendering Buffer (Crucial for 2nd+ pages)
            
            // 4. Trigger print
            window.print();
            
            // 5. Short delay before clearing to ensure print engine captured it
            await new Promise(r => setTimeout(r, 500)); 
            printArea.innerHTML = ''; 

            // 6. Inter-label gap for printer buffer safety
            if (i < labels.length - 1) {
                await new Promise(r => setTimeout(r, 3000)); // 3s Gap
            }
        }

        // Clean up
        document.body.classList.remove('sequential-print-active');
        overlay.style.display = 'none';
        alert('모든 라벨의 순차 인쇄가 완료되었습니다.');
    };
});

