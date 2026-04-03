// ----------------------------------------------------
// ระบบจัดซื้อจัดจ้าง - Client Side Logic (เชื่อมต่อ Google Apps Script)
// ----------------------------------------------------

// คุณต้องนำ URL ของ Web App เปลี่ยนตรงนี้ หลังจาก Deploy ใน Google Apps Script
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbzl1hh0siZl5fAUBlG-_7BGzsC7s992Q35j1iVNBgANrXiiAZ8y03XC1NOt7UyS3cyO/exec';

let projects = [];
let selectedProjectId = null;

const statusSteps = ['s-tor-date', 's-announce-date', 's-consideration-date', 's-appeal-date', 's-waitsign-date', 's-signed-date'];

function initStepValidation() {
    statusSteps.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('input', checkStatusStepsSequence);
    });
}

function checkStatusStepsSequence() {
    let allowNext = true;
    for (let i = 0; i < statusSteps.length; i++) {
        const id = statusSteps[i];
        const dateInput = document.getElementById(id);
        // Note inputs are typically next to date inputs, derive id
        const noteInput = document.getElementById(id.replace('-date', '-note'));

        if (!dateInput) continue;

        if (i === 0) {
            dateInput.disabled = false;
            if (noteInput) noteInput.disabled = false;
        } else {
            dateInput.disabled = !allowNext;
            if (noteInput) noteInput.disabled = !allowNext;

            if (dateInput.disabled) {
                dateInput.value = "";
                if (noteInput) noteInput.value = "";
            }
        }

        const val = dateInput.value.trim();
        allowNext = allowNext && val !== "" && !val.includes("รอ");
    }
}

document.addEventListener('DOMContentLoaded', () => {
    initStepValidation();
    document.getElementById('nav-dashboard').addEventListener('click', () => switchView('dashboard'));
    document.getElementById('nav-report').addEventListener('click', () => switchView('report'));

    const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    document.getElementById('current-date').innerText = new Date().toLocaleDateString('th-TH', options);

    fetchProjects();
});

// API Functions
async function fetchProjects() {
    if (WEB_APP_URL === 'YOUR_DEPLOYED_WEB_APP_URL_HERE') {
        Swal.fire('คำเตือน', 'กรุณานำ Web App URL ของ Google Apps Script มาใส่ในไฟล์ script.js บรรทัดที่ 5<br><br>(ระหว่างนี้จะใช้ Mock data ให้ดูชั่วคราว)', 'info');
        projects = [];
        renderDashboard();
        return;
    }

    try {
        Swal.fire({ title: 'กำลังโหลดข้อมูล...', text: 'กรุณารอสักครู่', allowOutsideClick: false, didOpen: () => { Swal.showLoading() } });
        const response = await fetch(`${WEB_APP_URL}?action=getProjects`);
        const json = await response.json();

        if (json.status === 'success') {
            projects = json.data.map(p => {
                p.dates.inspection = (p.notes && p.notes.date_inspection) ? p.notes.date_inspection : '';
                p.dates.payment = (p.notes && p.notes.date_payment) ? p.notes.date_payment : '';
                return p;
            });
            renderDashboard();
            if (document.getElementById('view-report').classList.contains('active')) renderReport();
            Swal.close();
        }
    } catch (e) {
        Swal.fire('เกิดข้อผิดพลาด', 'ไม่สามารถเชื่อมต่อฐานข้อมูลได้ ' + e.message, 'error');
    }
}

async function apiSaveProject(projectData) {
    if (WEB_APP_URL === 'YOUR_DEPLOYED_WEB_APP_URL_HERE') return true;
    try {
        Swal.fire({ title: 'กำลังบันทึกข้อมูล...', allowOutsideClick: false, didOpen: () => { Swal.showLoading() } });

        const response = await fetch(WEB_APP_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'saveProject', project: projectData })
        });
        const json = await response.json();
        return json.status === 'success';
    } catch (e) {
        Swal.fire('เกิดข้อผิดพลาด', 'ไม่สามารถบันทึกข้อมูลได้', 'error');
        return false;
    }
}

async function apiDeleteProject(id) {
    if (WEB_APP_URL === 'YOUR_DEPLOYED_WEB_APP_URL_HERE') return true;
    try {
        Swal.fire({ title: 'กำลังลบข้อมูล...', allowOutsideClick: false, didOpen: () => { Swal.showLoading() } });
        const response = await fetch(WEB_APP_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'deleteProject', id: id })
        });
        const json = await response.json();
        return json.status === 'success';
    } catch (e) {
        Swal.fire('เกิดข้อผิดพลาด', 'ไม่สามารถลบข้อมูลได้', 'error');
        return false;
    }
}


function switchView(viewName) {
    document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
    document.querySelectorAll('.view-section').forEach(section => section.classList.remove('active'));

    const headerActionsContainer = document.getElementById('header-actions');
    const pageTitle = document.getElementById('page-title');

    if (viewName === 'dashboard') {
        document.getElementById('nav-dashboard').classList.add('active');
        document.getElementById('view-dashboard').classList.add('active');
        pageTitle.innerText = "ระบบจัดซื้อจัดจ้าง กองบริหารการบินเกษตร";
        headerActionsContainer.innerHTML = '';
        renderDashboard();
    } else if (viewName === 'report') {
        document.getElementById('nav-report').classList.add('active');
        document.getElementById('view-report').classList.add('active');
        pageTitle.innerText = "แบบรายงานแผนการดำเนินงานงบลงทุน";
        headerActionsContainer.innerHTML = '<button class="btn btn-teal" onclick="window.print()"><i class="fa-solid fa-print"></i> พิมพ์รายงาน</button>';
        renderReport();
    }
}

const formatCurrency = (num) => Number(num).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

function getCurrentProjectStatusHTML(p) {
    const stepNames = { payment: 'เบิกจ่ายเงินเสร็จสิ้น', inspection: 'ตรวจรับแล้ว', signed: 'ลงนามสัญญาแล้ว', waitsign: 'รอลงนามผูกพัน', appeal: 'ช่วงอุทธรณ์', consideration: 'พิจารณาผลผู้ชนะ', announce: 'ประกาศจัดซื้อจัดจ้าง', tor: 'ร่าง TOR' };
    const order = ['payment', 'inspection', 'signed', 'waitsign', 'appeal', 'consideration', 'announce', 'tor'];

    let currentStepText = '<span style="color:#888;">ยังไม่เริ่มดำเนินการ</span>';
    let isSigned = false;

    for (let k of order) {
        if (p.dates && p.dates[k] && p.dates[k] !== '-' && !p.dates[k].includes('รอ')) {
            currentStepText = `<span style="color:var(--teal); font-weight:600;"><i class="fa-regular fa-circle-check"></i> ${stepNames[k]}</span><br><span style="font-size:12px; color:#666;">อัปเดต: ${formatThaiShort(p.dates[k])}</span>`;
            if (k === 'signed') isSigned = true;
            break;
        }
    }

    let expiryHtml = '';
    if (isSigned && p.duration > 0) {
        const expStr = calculateExpiryDate(p.dates.signed, p.duration);
        if (expStr) {
            const expDate = parseThaiDate(expStr);
            if (expDate) {
                const now = new Date();
                now.setHours(0, 0, 0, 0);
                expDate.setHours(0, 0, 0, 0);

                const diffTime = expDate - now;
                const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

                let daysText = '';
                let bColor = '#f0fdfa', tColor = 'var(--teal)';

                if (diffDays > 0) {
                    daysText = `เหลือ ${diffDays} วัน`;
                    if (diffDays <= 30) { bColor = '#fff8e1'; tColor = '#d4a529'; }
                } else if (diffDays === 0) {
                    daysText = 'ครบกำหนดวันนี้พอดี';
                    bColor = '#fdf0cc'; tColor = '#e67e3a';
                } else {
                    daysText = `เลยกำหนด ${Math.abs(diffDays)} วัน!`;
                    bColor = '#fce4ec'; tColor = '#c2185b';
                }

                expiryHtml = `<div style="margin-top:8px; font-size:11.5px; border: 1px solid ${tColor}; background-color: ${bColor}; color: ${tColor}; padding: 4px 8px; border-radius: 4px; display:inline-block; font-weight:500; min-width: 140px;">
                    <div><i class="fa-regular fa-clock"></i> สิ้นสุด: ${formatThaiShort(expStr)}</div>
                    <div style="margin-top:2px;">&#10148; ${daysText}</div>
                </div>`;
            }
        }
    }

    return `<div style="line-height:1.4;">${currentStepText}</div>${expiryHtml}`;
}

function handleSearch() {
    renderDashboard();
}

function renderDashboard() {
    const searchInputEl = document.getElementById('search-input');
    const filterPlanEl = document.getElementById('filter-plan');
    const filterItemEl = document.getElementById('filter-item-type');

    const searchTerm = searchInputEl ? searchInputEl.value : '';
    const filterPlan = filterPlanEl ? filterPlanEl.value : 'all';
    const filterItem = filterItemEl ? filterItemEl.value : 'all';

    let totalBudget = 0;
    let totalPO = 0;
    let totalDisbursed = 0;

    const filteredProjects = projects.filter(p => {
        const c_no = (p.notes && p.notes.contractNo) ? p.notes.contractNo.toLowerCase() : '';
        const search = searchTerm.toLowerCase();
        const matchSearch = p.name.toLowerCase().includes(search) ||
            (p.contractor && p.contractor.toLowerCase().includes(search)) ||
            c_no.includes(search);

        const pPlan = (p.notes && p.notes.plan) ? p.notes.plan : 'การปฏิบัติการฝนหลวง';
        const pItem = (p.notes && p.notes.itemType) ? p.notes.itemType : 'งบรายจ่ายอื่น';

        const matchPlan = filterPlan === 'all' || pPlan === filterPlan;
        const matchItem = filterItem === 'all' || pItem === filterItem;

        return matchSearch && matchPlan && matchItem;
    });

    const tbody = document.getElementById('project-table-body');
    if (!tbody) return;
    tbody.innerHTML = '';

    filteredProjects.forEach((p) => {
        totalBudget += Number(p.budget);
        totalPO += Number(p.po || 0);
        totalDisbursed += Number(p.disbursed);
    });

    const netBalance = totalBudget - totalDisbursed;

    document.getElementById('stat-total-projects').innerText = filteredProjects.length;
    document.getElementById('stat-total-budget').innerText = formatCurrency(totalBudget);
    document.getElementById('stat-total-po').innerText = formatCurrency(totalPO);
    document.getElementById('stat-total-disbursed').innerText = formatCurrency(totalDisbursed);
    document.getElementById('stat-net-balance').innerText = formatCurrency(netBalance);

    filteredProjects.forEach((p, index) => {
        const tr = document.createElement('tr');
        tr.onclick = (e) => {
            if (!e.target.closest('.icon-btn')) openProjectReport(p.id);
        };

        const contractBadge = (p.notes && p.notes.contractNo) ? `<div style="margin-top:6px; font-size:12px; color:#c2185b; border: 1px solid #f48fb1; padding: 2px 6px; border-radius: 4px; display: inline-block; background-color: #fce4ec;"><i class="fa-solid fa-file-signature"></i> สัญญาเลขที่: ${p.notes.contractNo}</div>` : '';
        const badgesHtml = `
            <div style="margin-top: 5px;">
                <span class="badge outline-badge" style="font-size:10px;">${(p.notes && p.notes.plan) ? p.notes.plan : 'การปฏิบัติการฝนหลวง'}</span>
                <span class="badge outline-badge" style="font-size:10px; border-color:#d4a529; color:#d4a529;">${(p.notes && p.notes.itemType) ? p.notes.itemType : 'งบรายจ่ายอื่น'}</span>
            </div>
        `;
        tr.innerHTML = `
            <td>${index + 1}</td>
            <td>
                ${p.name.substring(0, 80) + (p.name.length > 80 ? '...' : '')}<br>
                ${badgesHtml}
                ${contractBadge}
            </td>
            <td>
                <div class="sub-text bold">${p.contractor || '-'}</div>
                <div class="sub-text">ผู้ควบคุมงาน: ${p.supervisor || '-'}</div>
            </td>
            <td>${getCurrentProjectStatusHTML(p)}</td>
            <td>฿${formatCurrency(p.budget)}</td>
            <td class="text-yellow">฿${formatCurrency(p.po || 0)}</td>
            <td class="text-teal">฿${formatCurrency(p.disbursed)}</td>
            <td>
                <button class="icon-btn" onclick="openEditModal(${p.id})" title="แก้ไข"><i class="fa-solid fa-pen"></i></button>
                <button class="icon-btn" style="color: #e74c3c; margin-left: 10px;" onclick="deleteProject(event, ${p.id})" title="ลบ"><i class="fa-solid fa-trash"></i></button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function openProjectReport(id) {
    selectedProjectId = id;
    switchView('report');
}

function printProjectSummary() {
    // Get the current active filter data
    const filterPlan = document.getElementById('filter-plan').value;
    const filterItem = document.getElementById('filter-item-type').value;
    const searchVal = document.getElementById('search-input').value;

    const filtered = projects.filter(p => {
        const c_no = (p.notes && p.notes.contractNo) ? p.notes.contractNo.toLowerCase() : '';
        const search = searchVal.toLowerCase();
        const matchSearch = p.name.toLowerCase().includes(search) || (p.contractor && p.contractor.toLowerCase().includes(search)) || c_no.includes(search);
        const pPlan = (p.notes && p.notes.plan) ? p.notes.plan : 'การปฏิบัติการฝนหลวง';
        const pItem = (p.notes && p.notes.itemType) ? p.notes.itemType : 'งบรายจ่ายอื่น';
        const matchPlan = filterPlan === 'all' || pPlan === filterPlan;
        const matchItem = filterItem === 'all' || pItem === filterItem;
        return matchSearch && matchPlan && matchItem;
    });

    if (filtered.length === 0) {
        Swal.fire('ไม่มีข้อมูล', 'ไม่พบรายการโครงการที่จะพิมพ์', 'info');
        return;
    }

    const printWindow = window.open('', '_blank');
    const thaiDate = new Date().toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric' });

    let rowsHtml = filtered.map((p, index) => {
        const statusHtml = getCurrentProjectStatusHTML(p).replace(/<[^>]*>?/gm, ''); // Strip tags for print
        return `
            <tr>
                <td style="text-align: center;">${index + 1}</td>
                <td>${p.name} <br> <small style="color:#666;">(${(p.notes && p.notes.plan) || '-'} / ${(p.notes && p.notes.itemType) || '-'})</small></td>
                <td>${p.contractor || '-'} <br> <small>เลขที่สัญญา: ${(p.notes && p.notes.contractNo) || '-'}</small></td>
                <td>${statusHtml}</td>
                <td style="text-align: right;">${formatCurrency(p.budget)}</td>
                <td style="text-align: right;">${formatCurrency(p.po || 0)}</td>
                <td style="text-align: right;">${formatCurrency(p.disbursed)}</td>
            </tr>
        `;
    }).join('');

    const totalBudget = filtered.reduce((sum, p) => sum + Number(p.budget), 0);
    const totalPO = filtered.reduce((sum, p) => sum + Number(p.po || 0), 0);
    const totalDisbursed = filtered.reduce((sum, p) => sum + Number(p.disbursed), 0);

    printWindow.document.write(`
        <html>
        <head>
            <title>รายงานสรุปโครงการ - กองบริหารการบินเกษตร</title>
            <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700&display=swap" rel="stylesheet">
            <style>
                body { font-family: 'Sarabun', sans-serif; padding: 30px; color: #333; }
                h1 { text-align: center; color: #1e564d; margin-bottom: 5px; }
                .subtitle { text-align: center; margin-bottom: 30px; font-size: 14px; color: #666; }
                table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 13px; }
                th, td { border: 1px solid #ccc; padding: 8px 10px; text-align: left; vertical-align: top; }
                th { background-color: #f2f2f2; font-weight: bold; }
                .footer { margin-top: 40px; text-align: right; font-size: 12px; border-top: 1px solid #eee; padding-top: 10px; }
                .summary { margin-top: 20px; float: right; width: 40%; }
                .summary table { border: none; }
                .summary td { border: none; padding: 3px 0; }
                .summary .total { font-weight: bold; font-size: 16px; border-top: 2px solid #1e564d; padding-top: 10px; }
            </style>
        </head>
        <body>
            <h1>รายงานสรุปโครงการจัดซื้อจัดจ้าง</h1>
            <div class="subtitle">กองบริหารการบินเกษตร | ข้อมูล ณ วันที่ ${thaiDate}</div>
            
            <table>
                <thead>
                    <tr>
                        <th style="width: 40px;">ลำดับ</th>
                        <th>ชื่อโครงการ / แผนงาน</th>
                        <th>ผู้รับจ้าง / เลขที่สัญญา</th>
                        <th>สถานะปัจจุบัน</th>
                        <th>งบประมาณ (บาท)</th>
                        <th>PO / ก่อหนี้ (บาท)</th>
                        <th>เบิกจ่าย (บาท)</th>
                    </tr>
                </thead>
                <tbody>
                    ${rowsHtml}
                </tbody>
            </table>

            <div class="summary">
                <table>
                    <tr><td>รวมวงเงินงบประมาณทั้งสิ้น:</td><td style="text-align: right;">${formatCurrency(totalBudget)} บาท</td></tr>
                    <tr><td>รวมวงเงิน PO / ก่อหนี้:</td><td style="text-align: right;">${formatCurrency(totalPO)} บาท</td></tr>
                    <tr><td>รวมเบิกจ่ายแล้วทั้งสิ้น:</td><td style="text-align: right; color: #1bb295;">${formatCurrency(totalDisbursed)} บาท</td></tr>
                    <tr class="total"><td>คงเหลือสุทธิ:</td><td style="text-align: right; color: #f5965b;">${formatCurrency(totalBudget - totalDisbursed)} บาท</td></tr>
                </table>
            </div>

            <div style="clear: both;"></div>
            <div class="footer">จัดทำโดย: ระบบบริหารจัดการสัญญา กองบริหารการบินเกษตร</div>
            <script>
                window.onload = function() { window.print(); window.close(); }
            </script>
        </body>
        </html>
    `);
    printWindow.document.close();
}

function parseThaiDate(thaiDateStr) {
    if (!thaiDateStr) return null;

    // Check if it's an auto-corrupted Date object from Google Sheets (e.g. "Thu Apr 03 1969")
    if (thaiDateStr.includes('GMT') || thaiDateStr.includes('T') || thaiDateStr.length > 20) {
        let dt = new Date(thaiDateStr);
        if (!isNaN(dt.getTime())) {
            let y = dt.getFullYear();
            if (y >= 1900 && y < 2000) {
                y += 57; // 1969 -> 2026, recovers the data
                dt.setFullYear(y);
            } else if (y > 2400) {
                // Google auto-parsed "03/04/2569" as 2569 AD
                dt.setFullYear(y - 543);
            }
            return dt;
        }
    }

    if (!thaiDateStr.includes('/')) return null;
    const p = thaiDateStr.split('/');
    if (p.length !== 3) return null;
    let year = parseInt(p[2], 10);
    if (year < 100) year += 2500;
    if (year > 2400) year -= 543;
    return new Date(year, parseInt(p[1], 10) - 1, parseInt(p[0], 10));
}

const shortMonths = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];

function formatThaiShort(dateStr) {
    if (!dateStr || dateStr === '-' || dateStr.includes('รอ')) return dateStr;
    const dt = parseThaiDate(dateStr);
    if (dt) {
        return `${dt.getDate()} ${shortMonths[dt.getMonth()]} ${String(dt.getFullYear() + 543).slice(-2)}`;
    }
    return dateStr;
}

function calculateExpiryDate(signedDateStr, durationDays) {
    if (!signedDateStr || signedDateStr === '-' || signedDateStr.includes('รอ') || !durationDays || durationDays <= 0) return null;

    const startDate = parseThaiDate(signedDateStr);
    if (!startDate || isNaN(startDate.getTime())) return null;

    startDate.setDate(startDate.getDate() + parseInt(durationDays, 10));

    const eDay = String(startDate.getDate()).padStart(2, '0');
    const eMonth = String(startDate.getMonth() + 1).padStart(2, '0');
    const eYear = startDate.getFullYear() + 543;

    return `${eDay}/${eMonth}/${eYear}`;
}

function renderReport() {
    const reportContent = document.getElementById('report-content');
    const noSelection = document.getElementById('no-report-selection');

    if (!selectedProjectId) {
        if (reportContent) reportContent.style.display = 'none';
        if (noSelection) noSelection.style.display = 'block';
        return;
    }

    const project = projects.find(p => p.id === selectedProjectId);
    if (!project) return;

    if (reportContent) reportContent.style.display = 'block';
    if (noSelection) noSelection.style.display = 'none';

    document.getElementById('r-id').innerText = `ลำดับที่ ${project.id}`;
    document.getElementById('r-type').innerText = project.contractType;

    const rContractNoBox = document.getElementById('r-contract-no');
    if (project.notes && project.notes.contractNo) {
        rContractNoBox.style.display = 'inline-block';
        rContractNoBox.innerText = `สัญญาเลขที่: ${project.notes.contractNo}`;
    } else {
        rContractNoBox.style.display = 'none';
        rContractNoBox.innerText = '';
    }

    document.getElementById('r-name').innerText = project.name;
    document.getElementById('r-contractor').innerText = project.contractor || '-';
    document.getElementById('r-supervisor').innerText = project.supervisor || '-';

    document.getElementById('r-comm-tor').innerText = (project.notes && project.notes.comm_tor) ? project.notes.comm_tor : '-';
    document.getElementById('r-comm-eval').innerText = (project.notes && project.notes.comm_eval) ? project.notes.comm_eval : '-';
    document.getElementById('r-comm-inspect').innerText = (project.notes && project.notes.comm_inspect) ? project.notes.comm_inspect : '-';

    const balance = Number(project.budget) - Number(project.po) - Number(project.disbursed);

    document.getElementById('r-budget').innerText = formatCurrency(project.budget);
    document.getElementById('r-po').innerText = formatCurrency(project.po);
    document.getElementById('r-disbursed').innerText = formatCurrency(project.disbursed);
    document.getElementById('r-balance').innerText = formatCurrency(balance);

    const progressPercent = project.budget > 0 ? (Number(project.disbursed) / Number(project.budget)) * 100 : 0;
    document.getElementById('r-progress-bar').style.width = `${progressPercent}%`;
    document.getElementById('r-progress-text').innerText = `${progressPercent.toFixed(1)}% เบิกจ่ายแล้ว`;

    const notes = project.notes || {};

    // Calculate days between steps
    const dOrder = ['tor', 'announce', 'consideration', 'appeal', 'waitsign', 'signed', 'inspection', 'payment'];
    let daysDiffText = {};
    let earliestDate = null;

    for (let i = 0; i < dOrder.length; i++) {
        const key = dOrder[i];
        const currentDt = parseThaiDate(project.dates[key]);
        if (currentDt) {
            if (!earliestDate || currentDt < earliestDate) earliestDate = currentDt;

            // Look backward for previous valid date to diff against
            let prevDt = null;
            for (let j = i - 1; j >= 0; j--) {
                const checkDt = parseThaiDate(project.dates[dOrder[j]]);
                if (checkDt) { prevDt = checkDt; break; }
            }

            if (prevDt) {
                const diffTime = Math.abs(currentDt - prevDt);
                const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
                daysDiffText[key] = `(ใช้เวลา ${diffDays} วัน)`;
            } else {
                daysDiffText[key] = '';
            }
        } else {
            daysDiffText[key] = '';
        }
    }

    updateTimelineStep('tor', '1', project.dates.tor, notes.tor, daysDiffText['tor']);
    updateTimelineStep('announce', '2', project.dates.announce, notes.announce, daysDiffText['announce']);
    updateTimelineStep('consideration', '3', project.dates.consideration, notes.consideration, daysDiffText['consideration']);
    updateTimelineStep('appeal', '4', project.dates.appeal, notes.appeal, daysDiffText['appeal']);
    updateTimelineStep('waitsign', '5', project.dates.waitsign, notes.waitsign, daysDiffText['waitsign']);
    updateTimelineStep('signed', '6', project.dates.signed, notes.signed, daysDiffText['signed']);
    updateTimelineStep('inspection', '7', project.dates.inspection, notes.inspection, daysDiffText['inspection']);
    updateTimelineStep('payment', '8', project.dates.payment, notes.payment, daysDiffText['payment']);

    // Contract End Calculation
    const endAlert = document.getElementById('contract-end-alert');
    const endText = document.getElementById('contract-end-text');
    const expiryDate = calculateExpiryDate(project.dates.signed, project.duration);

    if (expiryDate) {
        endAlert.style.display = 'block';
        endText.innerText = `วันสิ้นสุดสัญญา: ${formatThaiShort(expiryDate)} (ระยะเวลา ${project.duration} วัน)`;
    } else {
        endAlert.style.display = 'none';
        endText.innerText = '';
    }

    renderTimelineMonths(earliestDate, expiryDate, dOrder.map(k => project.dates[k]));
}

function renderTimelineMonths(earliestObj, expiryStr, allDates) {
    const container = document.getElementById('dynamic-months-grid');
    if (!container) return;

    container.innerHTML = '';

    let latestStr = null;
    for (let d of allDates) {
        if (d && d !== '-' && !d.includes('รอ')) {
            const dt = parseThaiDate(d);
            if (dt) latestStr = d;
        }
    }

    const end = expiryStr ? parseThaiDate(expiryStr) : (latestStr ? parseThaiDate(latestStr) : null);
    const start = earliestObj;

    if (!start || !end || start > end) {
        container.innerHTML = '<div style="grid-column: 1 / -1; text-align: center; color: #888; padding: 20px;">ระบุกำหนดวันที่ในแต่ละขั้นตอนใดๆ ก่อน เพื่อเริ่มแสดงไทม์ไลน์</div>';
        return;
    }

    const thaiMonths = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
    let current = new Date(start.getFullYear(), start.getMonth(), 1);
    const endBound = new Date(end.getFullYear(), end.getMonth(), 1);

    while (current <= endBound) {
        let m = current.getMonth();
        let y = String(current.getFullYear() + 543).slice(-2);

        let label = (current.getTime() === endBound.getTime()) ? (expiryStr ? 'สิ้นสุดสัญญา' : 'ดำเนินงานอยู่') : 'ดำเนินงาน';
        if (current.getFullYear() === start.getFullYear() && current.getMonth() === start.getMonth()) label = 'เริ่ม (ร่าง TOR)';

        let boxClass = 'month-box';

        if (current.getTime() === endBound.getTime() && expiryStr) {
            boxClass += ' end';
        } else if (current.getFullYear() === start.getFullYear() && current.getMonth() === start.getMonth()) {
            boxClass += ' active';
        } else {
            boxClass += ' active-work';
        }

        const div = document.createElement('div');
        div.className = boxClass;
        div.innerHTML = `
            <div class="month-name">${thaiMonths[m]}</div>
            <div class="month-year">${y}</div>
            <div class="month-status">${label}</div>
        `;
        container.appendChild(div);

        current.setMonth(current.getMonth() + 1);
        if (container.children.length > 70) break; // Hard limit roughly 5 years
    }
}

function updateTimelineStep(stepId, iconNum, dateValue, noteValue, daysTaken) {
    const elDate = document.getElementById(`date-${stepId}`);
    const elStep = document.getElementById(`step-${stepId}`);
    const elLine = document.getElementById(`line-${parseInt(iconNum) - 1}`);

    if (!elDate || !elStep) return;

    let displayTxt = formatThaiShort(dateValue) || '-';
    if (daysTaken && daysTaken !== '') {
        displayTxt += `<br><span style="color:#e67e3a; font-size:11px; font-weight:bold;">${daysTaken}</span>`;
    }
    if (noteValue && noteValue.trim() !== '') {
        displayTxt += `<br><small style="color:#888; font-size: 11px;">(${noteValue})</small>`;
    }
    elDate.innerHTML = displayTxt;

    const iconContainer = elStep.querySelector('.step-icon');

    if (dateValue && dateValue !== '-' && !dateValue.includes('รอ')) {
        elStep.classList.remove('pending');
        elStep.classList.add('completed');
        if (iconContainer) iconContainer.innerHTML = '<i class="fa-solid fa-check"></i>';
        if (elLine) elLine.classList.add('completed');
    } else {
        elStep.classList.remove('completed');
        elStep.classList.add('pending');
        if (iconContainer) iconContainer.innerHTML = iconNum;
        if (elLine) elLine.classList.remove('completed');
    }
}


// --- MAIN PROJECT MODAL ---
function openModal() {
    document.getElementById('projectForm').reset();
    document.getElementById('p-id').value = '';
    document.getElementById('modal-title').innerText = 'เพิ่มโครงการใหม่';
    document.getElementById('projectModal').classList.add('show');
}

function openEditModal(id) {
    if (!id) id = selectedProjectId;
    const project = projects.find(p => p.id === id);
    if (!project) return;

    document.getElementById('p-id').value = project.id;
    document.getElementById('p-name').value = project.name;
    document.getElementById('p-plan').value = (project.notes && project.notes.plan) ? project.notes.plan : 'การปฏิบัติการฝนหลวง';
    document.getElementById('p-item-type').value = (project.notes && project.notes.itemType) ? project.notes.itemType : 'งบรายจ่ายอื่น';
    document.getElementById('p-type').value = project.contractType;
    document.getElementById('p-contractor').value = project.contractor;
    document.getElementById('p-supervisor').value = project.supervisor;
    document.getElementById('p-budget').value = project.budget;
    document.getElementById('p-po').value = project.po;
    document.getElementById('p-disbursed').value = project.disbursed;
    if (document.getElementById('p-duration')) document.getElementById('p-duration').value = project.duration || 0;

    document.getElementById('modal-title').innerText = 'อัปเดตโครงการ';
    document.getElementById('projectModal').classList.add('show');
}

function closeModal() {
    document.getElementById('projectModal').classList.remove('show');
}

async function saveProject(e) {
    e.preventDefault();

    const idVal = document.getElementById('p-id').value;
    // We fetch existing dates/notes to preserve them if we only edit main info
    let existingDates = { tor: '', announce: '', consideration: '', appeal: '', waitsign: '', signed: '' };
    let existingNotes = {};
    if (idVal) {
        const found = projects.find(p => p.id === parseInt(idVal));
        if (found) {
            existingDates = found.dates;
            existingNotes = found.notes;
        }
    }

    const projectData = {
        name: document.getElementById('p-name').value,
        contractType: document.getElementById('p-type').value,
        contractor: document.getElementById('p-contractor').value,
        supervisor: document.getElementById('p-supervisor').value,
        budget: parseFloat(document.getElementById('p-budget').value) || 0,
        po: parseFloat(document.getElementById('p-po').value) || 0,
        disbursed: parseFloat(document.getElementById('p-disbursed').value) || 0,
        duration: parseInt(document.getElementById('p-duration')?.value || 0, 10),
        dates: existingDates,
        notes: existingNotes
    };

    projectData.notes.plan = document.getElementById('p-plan').value;
    projectData.notes.itemType = document.getElementById('p-item-type').value;
    projectData.notes.date_inspection = projectData.dates.inspection || '';
    projectData.notes.date_payment = projectData.dates.payment || '';

    if (idVal) projectData.id = parseInt(idVal);

    closeModal();
    const success = await apiSaveProject(projectData);
    if (success) {
        await fetchProjects();
        Swal.fire({ toast: true, position: 'top-end', icon: 'success', title: 'บันทึกข้อมูลสำเร็จ', showConfirmButton: false, timer: 2000 });
    }
}


function toISODate(thaiDateStr) {
    const dt = parseThaiDate(thaiDateStr);
    if (!dt) return '';
    let day = String(dt.getDate()).padStart(2, '0');
    let month = String(dt.getMonth() + 1).padStart(2, '0');
    let year = dt.getFullYear();
    return `${year}-${month}-${day}`;
}

function toThaiDate(isoDateStr) {
    if (!isoDateStr) return '';
    const parts = isoDateStr.split('-');
    if (parts.length !== 3) return isoDateStr;
    const year = parseInt(parts[0], 10) + 543;
    return `${parts[2]}/${parts[1]}/${year}`;
}

// --- STATUS UPDATE ---
async function updateSingleStep(stepKey) {
    if (!selectedProjectId) return;
    const project = projects.find(p => p.id === selectedProjectId);
    if (!project) return;

    const stepNames = {
        tor: 'ร่าง TOR',
        announce: 'ประกาศจัดซื้อจัดจ้าง',
        consideration: 'พิจารณาผลผู้ชนะ',
        appeal: 'อยู่ระหว่างช่วงอุทธรณ์',
        waitsign: 'รอลงนามผูกพัน',
        signed: 'ลงนามสัญญาเสร็จสิ้น',
        inspection: 'ตรวจรับงาน/พัสดุ',
        payment: 'เบิกจ่ายเงิน'
    };

    const order = ['tor', 'announce', 'consideration', 'appeal', 'waitsign', 'signed', 'inspection', 'payment'];
    const idx = order.indexOf(stepKey);

    if (idx > 0) {
        const prevKey = order[idx - 1];
        const prevVal = project.dates[prevKey];
        if (!prevVal || prevVal === '-' || prevVal.includes('รอ')) {
            Swal.fire('แจ้งเตือน', `กรุณาอัปเดตขั้นตอน "${stepNames[prevKey]}" ให้เรียบร้อยก่อน`, 'warning');
            return;
        }
    }

    const currentIsoDate = toISODate(project.dates[stepKey]);

    let extraInputHtml = '';
    let currentContractNo = (project.notes && project.notes.contractNo) ? project.notes.contractNo : '';
    if (stepKey === 'signed') {
        extraInputHtml = `
            <label style="display:block; margin-bottom:5px; font-weight:600; color:#555; font-size:14px; margin-top: 15px;">เลขที่สัญญา (เพื่อให้ค้นหาง่ายขึ้น)</label>
            <input type="text" id="swal-contract" class="swal2-input" placeholder="ตัวอย่าง: 11/2569" value="${currentContractNo}" style="width: 100%; margin: 0; font-size:14px; height: 40px;">
        `;
    }

    const { value: formValues } = await Swal.fire({
        title: `อัปเดต: ${stepNames[stepKey]}`,
        html: `
            <div style="text-align: left; padding: 0 10px; font-family: Sarabun;">
                <label style="display:block; margin-bottom:5px; font-weight:600; color:#555; font-size:14px;">วันที่</label>
                <input type="text" id="swal-date" class="swal2-input" placeholder="เลือกวันที่จากปฏิทิน" value="${currentIsoDate}" style="width: 100%; margin: 0 0 15px 0; font-size:14px; height: 40px; cursor: pointer; background: white;">
                <label style="display:block; margin-bottom:5px; font-weight:600; color:#555; font-size:14px;">หมายเหตุ</label>
                <input type="text" id="swal-note" class="swal2-input" placeholder="เว้นว่างได้ หรือระบุสถานะ..." value="${(project.notes && project.notes[stepKey]) || ''}" style="width: 100%; margin: 0; font-size:14px; height: 40px;">
                ${extraInputHtml}
                <div style="margin-top:15px; font-size:12px; color:#888;">* หากทำการล้างวันที่ออก ขั้นตอนถัดไปทั้งหมดจะถูกรีเซ็ตไปด้วย</div>
            </div>
        `,
        focusConfirm: false,
        showCancelButton: true,
        confirmButtonText: 'บันทึก',
        cancelButtonText: 'ยกเลิก',
        didOpen: () => {
            flatpickr("#swal-date", {
                locale: "th",
                dateFormat: "Y-m-d",
                altInput: true,
                altFormat: "d M Y",
                allowInput: false,
                formatDate: (date, formatStr) => {
                    if (formatStr === "d M Y") {
                        const thaiMonths = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
                        let y = date.getFullYear() + 543;
                        let m = thaiMonths[date.getMonth()];
                        let d = String(date.getDate()).padStart(2, '0');
                        return `${d} ${m} ${y}`;
                    }
                    return flatpickr.formatDate(date, formatStr);
                }
            });
        },
        preConfirm: () => {
            const rawDate = document.getElementById('swal-date').value;
            return {
                date: toThaiDate(rawDate),
                note: document.getElementById('swal-note').value,
                contractNo: document.getElementById('swal-contract') ? document.getElementById('swal-contract').value : null
            }
        }
    });

    if (formValues) {
        if (formValues.date.trim() === '') {
            for (let i = idx; i < order.length; i++) {
                project.dates[order[i]] = '';
                if (project.notes) project.notes[order[i]] = '';
            }
            if (stepKey === 'signed' && project.notes) project.notes.contractNo = '';
        } else {
            project.dates[stepKey] = formValues.date;
            if (!project.notes) project.notes = {};
            project.notes[stepKey] = formValues.note;
            if (formValues.contractNo !== null) project.notes.contractNo = formValues.contractNo;
        }

        project.notes.date_inspection = project.dates.inspection || '';
        project.notes.date_payment = project.dates.payment || '';

        const success = await apiSaveProject(project);
        if (success) {
            await fetchProjects();
            Swal.fire({ toast: true, position: 'top-end', icon: 'success', title: 'อัปเดตสถานะสำเร็จ', showConfirmButton: false, timer: 1500 });
        }
    }
}

// Global Deletions
async function deleteProject(e, id) {
    if (e) e.stopPropagation();

    const result = await Swal.fire({
        title: 'ยืนยันการลบ?',
        text: "คุณแน่ใจหรือไม่ว่าต้องการลบโครงการนี้?",
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#e74c3c',
        cancelButtonColor: '#bdc3c7',
        confirmButtonText: 'ใช่, ฉันต้องการลบ',
        cancelButtonText: 'ยกเลิก'
    });

    if (result.isConfirmed) {
        const success = await apiDeleteProject(id);
        if (success) {
            if (selectedProjectId === id) selectedProjectId = null;
            await fetchProjects();
            Swal.fire({ toast: true, position: 'top-end', icon: 'success', title: 'ลบโครงการเรียบร้อย', showConfirmButton: false, timer: 2000 });
        }
    }
}

function deleteCurrentProject() {
    if (selectedProjectId) deleteProject(null, selectedProjectId);
}

// Global Committee Updates
async function openCommitteeModal() {
    if (!selectedProjectId) return;
    const project = projects.find(p => p.id === selectedProjectId);
    if (!project) return;

    const comm_tor = (project.notes && project.notes.comm_tor) ? project.notes.comm_tor : '';
    const comm_eval = (project.notes && project.notes.comm_eval) ? project.notes.comm_eval : '';
    const comm_inspect = (project.notes && project.notes.comm_inspect) ? project.notes.comm_inspect : '';

    const { value: formValues } = await Swal.fire({
        title: 'อัปเดตรายชื่อคณะกรรมการ',
        width: '600px',
        html: `
            <div style="text-align: left; padding: 0 10px; font-family: Sarabun;">
                <label style="display:block; margin-bottom:5px; font-weight:600; color:#555; font-size:14px;">1. คณะกรรมการร่าง TOR</label>
                <textarea id="swal-comm-tor" class="swal2-textarea" style="width: 100%; margin: 0 0 15px 0; font-size:14px; height: 80px;" placeholder="ระบุรายชื่อ (ขึ้นบรรทัดใหม่ได้)">${comm_tor}</textarea>
                
                <label style="display:block; margin-bottom:5px; font-weight:600; color:#555; font-size:14px;">2. คณะกรรมการพิจารณาผล</label>
                <textarea id="swal-comm-eval" class="swal2-textarea" style="width: 100%; margin: 0 0 15px 0; font-size:14px; height: 80px;" placeholder="ระบุรายชื่อ (ขึ้นบรรทัดใหม่ได้)">${comm_eval}</textarea>
                
                <label style="display:block; margin-bottom:5px; font-weight:600; color:#555; font-size:14px;">3. คณะกรรมการตรวจรับพัสดุ</label>
                <textarea id="swal-comm-inspect" class="swal2-textarea" style="width: 100%; margin: 0 0 15px 0; font-size:14px; height: 80px;" placeholder="ระบุรายชื่อ (ขึ้นบรรทัดใหม่ได้)">${comm_inspect}</textarea>
            </div>
        `,
        focusConfirm: false,
        showCancelButton: true,
        confirmButtonText: 'บันทึก',
        cancelButtonText: 'ยกเลิก',
        preConfirm: () => {
            return {
                tor: document.getElementById('swal-comm-tor').value,
                eval: document.getElementById('swal-comm-eval').value,
                inspect: document.getElementById('swal-comm-inspect').value
            };
        }
    });

    if (formValues) {
        if (!project.notes) project.notes = {};
        project.notes.comm_tor = formValues.tor;
        project.notes.comm_eval = formValues.eval;
        project.notes.comm_inspect = formValues.inspect;

        const success = await apiSaveProject(project);
        if (success) {
            await fetchProjects();
            Swal.fire({ toast: true, position: 'top-end', icon: 'success', title: 'อัปเดตรายชื่อเรียบร้อย', showConfirmButton: false, timer: 1500 });
        }
    }
}
