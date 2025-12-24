/**
 * UI Components & Helpers
 */

// Badges & Classes
function getTypeBadge(type) {
    const badges = {
        'theory': '<span class="type-badge type-badge-theory" title="نظري">ن</span>',
        'practical': '<span class="type-badge type-badge-practical" title="عملي">ع</span>',
        'coop': '<span class="type-badge type-badge-coop" title="تعاوني">ت</span>',
        'mixed': '<span class="type-badge type-badge-mixed" title="مدمج">م</span>',
        'self': '<span class="type-badge type-badge-self" title="ذاتي">ذ</span>'
    };
    return badges[type] || badges['theory'];
}

function getDensityClass(value) {
    if (value > 20) return 'density-high';
    if (value >= 10) return 'density-medium';
    if (value >= 5) return 'density-low';
    return 'density-critical';
}

function getPassClass(value) {
    const v = parseFloat(value);
    if (v >= 70) return 'pass-high';
    if (v >= 50) return 'pass-medium';
    return 'pass-low';
}

// Tooltips
function showInstructorTooltip(event, courseCode) {
    // Requires global coursesData
    if (typeof coursesData === 'undefined') return;

    const course = coursesData[courseCode];
    if (!course || Object.keys(course.instructors).length === 0) return;

    let tooltip = document.getElementById('instructorTooltip');
    if (!tooltip) {
        tooltip = document.createElement('div');
        tooltip.id = 'instructorTooltip';
        tooltip.className = 'instructor-tooltip';
        document.body.appendChild(tooltip);
    }

    const instructors = Object.entries(course.instructors)
        .sort((a, b) => b[1] - a[1]);

    let html = `<h4>مدربو المقرر: ${courseCode}</h4>`;
    html += '<table><tr><th>المدرب</th><th>عدد الشعب</th></tr>';
    instructors.forEach(([name, count]) => {
        html += `<tr><td>${name}</td><td style="text-align:center;">${count}</td></tr>`;
    });
    html += '</table>';

    tooltip.innerHTML = html;
    tooltip.style.display = 'block';

    // Position
    const x = event.clientX + 15;
    const y = event.clientY + 15;
    tooltip.style.left = x + 'px';
    tooltip.style.top = y + 'px';

    // Adjust if off screen
    setTimeout(() => {
        const rect = tooltip.getBoundingClientRect();
        if (rect.right > window.innerWidth) {
            tooltip.style.left = (event.clientX - rect.width - 10) + 'px';
        }
        if (rect.bottom > window.innerHeight) {
            tooltip.style.top = (event.clientY - rect.height - 10) + 'px';
        }
    }, 0);
}

function hideInstructorTooltip() {
    const tooltip = document.getElementById('instructorTooltip');
    if (tooltip) tooltip.style.display = 'none';
}

function showSectionsTooltip(event, compositeKey) {
    // Requires global coursesData
    if (typeof coursesData === 'undefined') return;

    const course = coursesData[compositeKey];
    if (!course || !course.sectionDetails) return;

    let tooltip = document.getElementById('sectionsTooltip');
    if (!tooltip) {
        tooltip = document.createElement('div');
        tooltip.id = 'sectionsTooltip';
        tooltip.className = 'sections-tooltip';
        document.body.appendChild(tooltip);
    }

    const sections = Object.values(course.sectionDetails);
    let html = `<h4>تفاصيل الشعب - ${course.code} (${course.scheduleType})</h4>`;
    html += '<table><thead><tr>';
    html += '<th>الرقم المرجعي</th>';
    html += '<th>القسم</th>';
    html += '<th>التخصص</th>';
    html += '<th>نوع الجدولة</th>';
    html += '<th>المدرب</th>';
    html += '<th>المتدربين</th>';
    html += '</tr></thead><tbody>';

    sections.forEach(sec => {
        const count = sec.trainees ? sec.trainees.size : 0;
        html += `<tr>
            <td>${sec.refNum}</td>
            <td>${sec.department}</td>
            <td>${sec.specialty}</td>
            <td>${sec.scheduleType}</td>
            <td>${sec.instructor || '-'}</td>
            <td>${count}</td>
        </tr>`;
    });

    html += '</tbody></table>';
    tooltip.innerHTML = html;
    tooltip.style.display = 'block';

    // Position tooltip
    const x = event.clientX;
    const y = event.clientY;
    tooltip.style.left = Math.min(x - 300, window.innerWidth - 620) + 'px';
    tooltip.style.top = Math.min(y + 20, window.innerHeight - 420) + 'px';
}

function hideSectionsTooltip() {
    const tooltip = document.getElementById('sectionsTooltip');
    if (tooltip) {
        tooltip.style.display = 'none';
    }
}

// Copy Table - Uses HTML format for Excel compatibility
function copyTable(tableId = 'dataTable') {
    const table = document.getElementById(tableId);
    if (!table) {
        console.error('Table not found:', tableId);
        return;
    }

    // Build clean HTML table for clipboard
    let htmlContent = '<table border="1">';

    for (const row of table.rows) {
        htmlContent += '<tr>';
        for (const cell of row.cells) {
            const tag = cell.tagName.toLowerCase();
            let cellText = cell.innerText || cell.textContent || '';
            cellText = cellText.trim().replace(/[\r\n]+/g, ' ');

            // Preserve colspan/rowspan
            const colspan = cell.getAttribute('colspan');
            const rowspan = cell.getAttribute('rowspan');
            let attrs = '';
            if (colspan) attrs += ` colspan="${colspan}"`;
            if (rowspan) attrs += ` rowspan="${rowspan}"`;

            htmlContent += `<${tag}${attrs}>${cellText}</${tag}>`;
        }
        htmlContent += '</tr>';
    }
    htmlContent += '</table>';

    // Also build plain text version with tabs
    const lines = [];
    for (const row of table.rows) {
        const rowData = [];
        for (const cell of row.cells) {
            let cellText = cell.innerText || cell.textContent || '';
            cellText = cellText.trim().replace(/[\r\n\t]+/g, ' ');
            const colspan = parseInt(cell.getAttribute('colspan')) || 1;
            rowData.push(cellText);
            for (let i = 1; i < colspan; i++) rowData.push('');
        }
        lines.push(rowData.join('\t'));
    }
    const textContent = lines.join('\r\n');

    // Use ClipboardItem to write both formats
    try {
        const htmlBlob = new Blob([htmlContent], { type: 'text/html' });
        const textBlob = new Blob([textContent], { type: 'text/plain' });

        navigator.clipboard.write([
            new ClipboardItem({
                'text/html': htmlBlob,
                'text/plain': textBlob
            })
        ]).then(() => {
            const btn = document.querySelector('.copy-btn');
            if (btn) {
                btn.classList.add('copied');
                setTimeout(() => btn.classList.remove('copied'), 1500);
            }
            if (typeof showToast === 'function') {
                showToast('تم نسخ الجدول بنجاح', 'success');
            }
        }).catch(err => {
            console.error('Clipboard write failed:', err);
            // Fallback to text only
            navigator.clipboard.writeText(textContent);
            if (typeof showToast === 'function') {
                showToast('تم نسخ الجدول (نص فقط)', 'success');
            }
        });
    } catch (err) {
        // Fallback for browsers that don't support ClipboardItem
        navigator.clipboard.writeText(textContent).then(() => {
            if (typeof showToast === 'function') {
                showToast('تم نسخ الجدول', 'success');
            }
        });
    }
}
