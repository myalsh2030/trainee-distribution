// ============================================
// Export Functions - PDF & Excel
// ============================================

/**
 * تصدير PDF
 * @param {Object} config - إعدادات التصدير
 * @param {string} config.tableId - معرف الجدول (افتراضي: 'dataTable')
 * @param {string} config.theadId - معرف thead (افتراضي: 'tableHead')
 * @param {string} config.tbodyId - معرف tbody (افتراضي: 'tableBody')
 * @param {string} config.title - عنوان التقرير
 * @param {string} config.subtitle - معلومات إضافية (اختياري)
 * @param {string} config.orientation - 'landscape' أو 'portrait' (افتراضي: 'landscape')
 * @param {string} config.headerColor - لون رأس الجدول (افتراضي: '#2563eb')
 * @param {number} config.fontSize - حجم الخط (افتراضي: 8)
 * @param {number} config.zoom - نسبة التكبير للطباعة (افتراضي: 90)
 */
function exportToPDF(config = {}) {
    // الإعدادات الافتراضية
    const settings = {
        tableId: config.tableId || 'dataTable',
        theadId: config.theadId || 'tableHead',
        tbodyId: config.tbodyId || 'tableBody',
        title: config.title || 'تقرير',
        subtitle: config.subtitle || '',
        orientation: config.orientation || 'landscape',
        headerColor: config.headerColor || '#2563eb',
        fontSize: config.fontSize || 8,
        zoom: config.zoom || 90
    };

    if (typeof showToast === 'function') {
        showToast('جاري إنشاء ملف PDF...', 'info');
    }

    const printWindow = window.open('', '_blank');
    const theadContent = document.getElementById(settings.theadId)?.innerHTML || '';
    const tbodyContent = document.getElementById(settings.tbodyId)?.innerHTML || '';
    const currentDate = new Date().toLocaleDateString('ar-SA');

    printWindow.document.write(`
        <!DOCTYPE html>
        <html dir="rtl" lang="ar">
        <head>
            <meta charset="UTF-8">
            <title>${settings.title}</title>
            <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap" rel="stylesheet">
            <style>
                /* Page Setup */
                @page { 
                    size: ${settings.orientation}; 
                    margin: 10mm 10mm 25mm 10mm;
                    
                    @bottom-center {
                        content: "صفحة " counter(page) " من " counter(pages);
                        font-family: 'Cairo', sans-serif;
                        font-size: 9pt;
                        color: #555;
                    }
                }
                
                body { 
                    font-family: 'Cairo', sans-serif; 
                    direction: rtl; 
                    margin: 0;
                    padding: 5mm 15mm 10mm 15mm;
                    box-sizing: border-box;
                }
                
                /* Table */
                table { 
                    border-collapse: collapse; 
                    width: 100%; 
                    font-size: ${settings.fontSize}px; 
                    page-break-inside: auto;
                }
                
                thead { display: table-header-group; }
                tfoot { display: table-footer-group; }
                
                /* Report Title Row */
                .report-title-row td {
                    background: white !important;
                    color: #1f2937 !important;
                    border: none !important;
                    border-bottom: 2px solid ${settings.headerColor} !important;
                    padding: 8px 10px !important;
                    font-size: 10px;
                    -webkit-print-color-adjust: exact;
                    print-color-adjust: exact;
                }
                
                .report-title-row .title-right {
                    text-align: right;
                    font-weight: bold;
                    font-size: 13px;
                }
                
                .report-title-row .title-center {
                    text-align: center;
                    font-weight: bold;
                    font-size: 11px;
                    color: #444;
                }
                
                .report-title-row .title-left {
                    text-align: left;
                    font-size: 9px;
                    color: #666;
                }
                
                /* Column Headers */
                thead tr:not(.report-title-row) th {
                    background: ${settings.headerColor} !important;
                    color: white !important;
                    border: 0.5px solid ${settings.headerColor};
                    padding: 4px 3px;
                    font-weight: 600;
                    -webkit-print-color-adjust: exact;
                    print-color-adjust: exact;
                }
                
                /* Data Cells */
                td { 
                    border: 0.5px solid #333; 
                    padding: 2px 3px; 
                    text-align: center; 
                    white-space: nowrap; 
                }
                
                tbody tr:nth-child(even) { 
                    background: #f3f4f6 !important; 
                    -webkit-print-color-adjust: exact; 
                    print-color-adjust: exact; 
                }
                
                @media print { 
                    body { zoom: ${settings.zoom}%; }
                }
            </style>
        </head>
        <body>
            <table>
                <thead>
                    <tr class="report-title-row">
                        <td class="title-right" colspan="10">
                            📊 ${settings.title}
                        </td>
                        <td class="title-center" colspan="10">
                            ${settings.subtitle}
                        </td>
                        <td class="title-left" colspan="10">
                            التاريخ: ${currentDate}
                        </td>
                    </tr>
                    ${theadContent}
                </thead>
                <tbody>
                    ${tbodyContent}
                </tbody>
            </table>

            <script>
                window.onload = () => {
                    setTimeout(() => {
                        window.print();
                        window.close();
                    }, 500);
                };
            <\/script>
        </body>
        </html>
    `);

    printWindow.document.close();

    if (typeof showToast === 'function') {
        showToast('تم فتح نافذة الطباعة', 'success');
    }
}


/**
 * تصدير Excel (XLSX مع تنسيقات)
 * @param {Object} config - إعدادات التصدير
 * @param {string} config.tableId - معرف الجدول (افتراضي: 'dataTable')
 * @param {string} config.fileName - اسم الملف
 * @param {string} config.sheetName - اسم الورقة
 * @param {Object} config.rowColors - ألوان الصفوف حسب الـ class
 * @param {string} config.headerColor - لون رأس الجدول (افتراضي: '1E3A5F')
 * @param {number} config.colWidth - عرض الأعمدة (افتراضي: 12)
 */
function exportToExcel(config = {}) {
    // الإعدادات الافتراضية
    const settings = {
        tableId: config.tableId || 'dataTable',
        fileName: config.fileName || 'تقرير.xlsx',
        sheetName: config.sheetName || 'البيانات',
        rowColors: config.rowColors || {
            'row-theory': 'FFFFFF',
            'row-practical': 'F0FDF4',
            'row-coop': 'FFFBEB',
            'row-mixed': 'F5F3FF',
            'row-self': 'FCE7F3'
        },
        headerColor: config.headerColor || '1E3A5F',
        colWidth: config.colWidth || 12
    };

    if (typeof showToast === 'function') {
        showToast('جاري تصدير ملف Excel...', 'info');
    }

    const table = document.getElementById(settings.tableId);
    if (!table) {
        console.error('Table not found:', settings.tableId);
        return;
    }

    // تحميل المكتبة ديناميكياً
    if (typeof XLSX === 'undefined') {
        const script = document.createElement('script');
        script.src = 'https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js';
        script.onload = () => generateXLSX(table, settings);
        script.onerror = () => {
            if (typeof showToast === 'function') {
                showToast('فشل تحميل مكتبة Excel، جاري استخدام الطريقة البديلة...', 'warning');
            }
            exportToExcelFallback(table, settings);
        };
        document.head.appendChild(script);
    } else {
        generateXLSX(table, settings);
    }
}

function generateXLSX(table, settings) {
    try {
        const ws = XLSX.utils.table_to_sheet(table, { raw: true });
        const range = XLSX.utils.decode_range(ws['!ref']);

        const headerRowCount = table.querySelectorAll('thead tr').length;

        const bodyRows = table.querySelectorAll('tbody tr');
        const rowClasses = {};
        bodyRows.forEach((row, index) => {
            rowClasses[headerRowCount + index] = row.className;
        });

        const baseBorder = {
            top: { style: 'thin', color: { rgb: '333333' } },
            bottom: { style: 'thin', color: { rgb: '333333' } },
            left: { style: 'thin', color: { rgb: '333333' } },
            right: { style: 'thin', color: { rgb: '333333' } }
        };

        for (let R = range.s.r; R <= range.e.r; R++) {
            const isHeaderRow = R < headerRowCount;
            const rowClassName = rowClasses[R] || '';

            for (let C = range.s.c; C <= range.e.c; C++) {
                const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                if (!ws[cellRef]) continue;

                if (isHeaderRow) {
                    ws[cellRef].s = {
                        fill: { fgColor: { rgb: settings.headerColor } },
                        font: { color: { rgb: 'FFFFFF' }, bold: true, sz: 11 },
                        alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                        border: baseBorder
                    };
                } else {
                    let bgColor = 'FFFFFF';
                    for (const [className, color] of Object.entries(settings.rowColors)) {
                        if (rowClassName.includes(className)) {
                            bgColor = color;
                            break;
                        }
                    }

                    if (bgColor === 'FFFFFF' && (R - headerRowCount) % 2 === 1) {
                        bgColor = 'F9FAFB';
                    }

                    ws[cellRef].s = {
                        fill: { fgColor: { rgb: bgColor } },
                        font: { sz: 10 },
                        alignment: { horizontal: 'center', vertical: 'center' },
                        border: baseBorder
                    };
                }
            }
        }

        const colWidths = [];
        for (let C = range.s.c; C <= range.e.c; C++) {
            colWidths.push({ wch: settings.colWidth });
        }
        ws['!cols'] = colWidths;

        const wb = XLSX.utils.book_new();
        if (!wb.Workbook) wb.Workbook = {};
        if (!wb.Workbook.Views) wb.Workbook.Views = [];
        wb.Workbook.Views[0] = { RTL: true };

        XLSX.utils.book_append_sheet(wb, ws, settings.sheetName);
        XLSX.writeFile(wb, settings.fileName);

        if (typeof showToast === 'function') {
            showToast('تم تصدير ملف Excel بنجاح', 'success');
        }
    } catch (err) {
        console.error('XLSX export error:', err);
        if (typeof showToast === 'function') {
            showToast('جاري استخدام الطريقة البديلة...', 'warning');
        }
        exportToExcelFallback(table, settings);
    }
}

function exportToExcelFallback(table, settings) {
    const uri = 'data:application/vnd.ms-excel;base64,';
    const template = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
    <head>
        <!--[if gte mso 9]>
        <xml>
            <x:ExcelWorkbook>
                <x:ExcelWorksheets>
                    <x:ExcelWorksheet>
                        <x:Name>${settings.sheetName}</x:Name>
                        <x:WorksheetOptions>
                            <x:DisplayGridlines/>
                            <x:RightToLeft/>
                        </x:WorksheetOptions>
                    </x:ExcelWorksheet>
                </x:ExcelWorksheets>
            </x:ExcelWorkbook>
        </xml>
        <![endif]-->
        <meta http-equiv="content-type" content="text/plain; charset=UTF-8"/>
        <style>
            body { font-family: 'Cairo', sans-serif; direction: rtl; }
            table { border-collapse: collapse; width: 100%; }
            th, td { border: 0.5pt solid #000000; padding: 5px; text-align: center; vertical-align: middle; }
            th { background-color: #${settings.headerColor}; color: #ffffff; font-weight: bold; }
        </style>
    </head>
    <body>
        <table>${table.innerHTML}</table>
    </body>
    </html>`;

    const base64 = (s) => window.btoa(unescape(encodeURIComponent(s)));
    const link = document.createElement('a');
    link.href = uri + base64(template);
    link.download = settings.fileName.replace('.xlsx', '.xls');
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    if (typeof showToast === 'function') {
        showToast('تم تصدير ملف Excel (صيغة قديمة)', 'success');
    }
}
