const express = require('express');
const { chromium } = require('playwright');
const ExcelJS = require('exceljs');
const app = express();

// 🛠️ สำคัญมาก: ต้องไปสมัครที่ browserless.io (ฟรี) แล้วเอา Token ยาวๆ มาแปะตรงนี้ครับ
const BROWSERLESS_TOKEN = "วาง_TOKEN_ของคุณตรงนี้"; 

app.get('/', (req, res) => {
    res.send(`
    <!DOCTYPE html>
    <html lang="th">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Astro Daily Table Exporter</title>
        <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-slate-100 min-h-screen flex items-center justify-center p-4">
        <div class="bg-white p-10 rounded-3xl shadow-2xl w-full max-w-xl border border-slate-50 text-center">
            <h1 class="text-3xl font-bold text-indigo-600 mb-2 tracking-tight">ระบบดึงข้อมูลปฏิทินดาราศาสตร์</h1>
            <h1 class="text-2xl font-bold text-slate-700 mb-6 tracking-tight">พิกัดดาวรายวัน (เวอร์ชัน Vercel Fast Fast)</h1>
            
            <div class="text-left mb-8 space-y-4">
                <div>
                    <label class="block text-slate-800 font-bold text-base mb-2">เลือกปี พ.ศ. (2021 - 2027)</label>
                    <input type="number" id="targetYear" value="2024" min="2021" max="2027" 
                           class="w-full px-4 py-3 bg-white border border-slate-200 rounded-xl text-lg font-medium text-slate-700 focus:outline-none focus:ring-2 focus:ring-indigo-500 transition shadow-sm">
                </div>
                
                <div>
                    <label class="block text-slate-800 font-bold text-base mb-2">เลือกเดือนที่ต้องการข้อมูล</label>
                    <select id="targetMonth" class="w-full px-4 py-3 bg-white border border-slate-200 rounded-xl text-lg font-medium text-slate-700 focus:outline-none focus:ring-2 focus:ring-indigo-500 transition shadow-sm">
                        <option value="1">มกราคม</option>
                        <option value="2">กุมภาพันธ์</option>
                        <option value="3">มีนาคม</option>
                        <option value="4">เมษายน</option>
                        <option value="5">พฤษภาคม</option>
                        <option value="6">มิถุนายน</option>
                        <option value="7">กรกฎาคม</option>
                        <option value="8">สิงหาคม</option>
                        <option value="9">กันยายน</option>
                        <option value="10">ตุลาคม</option>
                        <option value="11">พฤศจิกายน</option>
                        <option value="12">ธันวาคม</option>
                    </select>
                </div>
            </div>

            <button id="exportBtn" onclick="downloadExcel()" 
                    class="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-semibold py-4 rounded-xl transition duration-200 shadow-lg flex justify-center items-center gap-3 text-lg">
                📁 ดาวน์โหลดโครงสร้าง Excel เรฟตัวเต็ม
            </button>
            
            <p id="statusMessage" class="text-center text-sm text-indigo-500 font-medium hidden mt-6 bg-indigo-50 p-3 rounded-xl border border-indigo-100">
                ⏳ บอทคลาวด์กำลังกวาดพิกัดดาวรายวันให้แบบด่วนพิเศษ กรุณารอสักครู่นะคะ^^
            </p>
        </div>

        <script>
            function downloadExcel() {
                const year = document.getElementById('targetYear').value;
                const month = document.getElementById('targetMonth').value;
                const status = document.getElementById('statusMessage');
                const btn = document.getElementById('exportBtn');
                
                status.classList.remove('hidden');
                btn.disabled = true;
                btn.classList.add('opacity-50');

                window.location.href = '/export-excel?year=' + year + '&month=' + month;

                setTimeout(() => {
                    status.innerHTML = "✅ ดึงข้อมูลสำเร็จและดาวน์โหลดแล้ว!";
                    btn.disabled = false;
                    btn.classList.remove('opacity-50');
                }, 10000);
            }
        </script>
    </body>
    </html>
    `);
});

app.get('/export-excel', async (req, res) => {
    const cYear = parseInt(req.query.year) || 2024;
    const month = parseInt(req.query.month) || 1;
    const thYear = cYear + 543;

    let browser;
    try {
        // เชื่อมต่อบอทผ่านเบราว์เซอร์คลาวด์ภายนอกเพื่อหลีกเลี่ยงข้อจำกัด Vercel
        browser = await chromium.connectOverCDP(
            `wss://chrome.browserless.io?token=${BROWSERLESS_TOKEN}`
        );
        
        const context = await browser.newContext();
        const page = await context.newPage();

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('ตารางพิกัดดาวรายวัน');

        worksheet.columns = [
            { header: 'ปี ค.ศ.', key: 'year_en', width: 10 },
            { header: 'ปี พ.ศ.', key: 'year_th', width: 10 },
            { header: 'เดือน', key: 'month', width: 14 },
            { header: 'ที่', key: 'index_no', width: 6 },
            { header: 'วัน', key: 'day_th', width: 6 },
            { header: 'ข-ร', key: 'cho_ro', width: 8 },
            { header: 'ด.', key: 'do_num', width: 6 },
            { header: 'อาทิตย์ (๑)', key: 'sun', width: 12 },
            { header: 'จันทร์ (๒)', key: 'moon', width: 12 },
            { header: 'ยก', key: 'yok_1', width: 10 },
            { header: 'ฤกษ์', key: 'reuk', width: 12 },
            { header: 'เต็ม', key: 'tem_1', width: 10 },
            { header: 'ดิถี', key: 'dithi', width: 12 },
            { header: 'เต็ม', key: 'tem_2', width: 10 },
            { header: 'อังคาร (๓)', key: 'mars', width: 12 },
            { header: 'พุธ (๔)', key: 'mercury', width: 12 },
            { header: 'พฤหัสฯ (๕)', key: 'jupiter', width: 12 },
            { header: 'ศุกร์ (๖)', key: 'venus', width: 12 },
            { header: 'เสาร์ (๗)', key: 'saturn', width: 12 },
            { header: 'ราหู (๘)', key: 'rahu', width: 12 },
            { header: 'เกตุ (๙)', key: 'ketu', width: 12 }
        ];

        const months_th = ['', 'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
        const targetUrl = `https://myhora.com/calendar/astro-suriyayas-${month}-${thYear}.aspx`;

        await page.goto(targetUrl, { waitUntil: 'load', timeout: 8000 });

        const extractedRows = await page.evaluate(() => {
            const rowsData = [];
            const trs = document.querySelectorAll('table tr');
            trs.forEach((tr) => {
                const tds = tr.querySelectorAll('td');
                if (tds.length >= 15) {
                    const cells = Array.from(tds).map(td => td.innerText ? td.innerText.trim().replace(/\s+/g, ' ') : '');
                    const dayNumber = parseInt(cells[0]);
                    if (!isNaN(dayNumber) && dayNumber >= 1 && dayNumber <= 31) {
                        rowsData.push(cells);
                    }
                }
            });
            return rowsData;
        });

        if (extractedRows && extractedRows.length > 0) {
            extractedRows.forEach(cells => {
                worksheet.addRow({
                    year_en: cYear, year_th: thYear, month: months_th[month],
                    index_no: cells[0] || '', day_th: cells[1] || '', cho_ro: cells[2] || '', do_num: cells[3] || '',
                    sun: cells[4] || '', moon: cells[5] || '', yok_1: cells[6] || '', reuk: cells[7] || '', tem_1: cells[8] || '',
                    dithi: cells[9] || '', tem_2: cells[10] || '', mars: cells[11] || '', mercury: cells[12] || '',
                    jupiter: cells[13] || '', venus: cells[14] || '', saturn: cells[15] || '', rahu: cells[16] || '', ketu: cells[17] || ''
                });
            });
        } else {
            throw new Error("โครงสร้างตารางบนเว็บปลายทางเปลี่ยนไปหรือไม่พบข้อมูล");
        }

        // ตกแต่งสไตล์หัวตารางสีน้ำเงิน Indigo
        worksheet.getRow(1).eachCell((cell) => {
            cell.font = { bold: true, color: { argb: 'FFFFFF' }, name: 'Arial', size: 11 };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '4F46E5' } };
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=Astro_Daily_Ref_${cYear}_${month}.xlsx`);

        await workbook.xlsx.write(res);
        res.end();

    } catch (err) {
        console.error("Vercel Function Error: ", err.message);
        res.status(500).send(`ระบบแครชเนื่องจาก: ${err.message}`);
    } finally {
        if (browser) await browser.close();
    }
});

module.exports = app;
