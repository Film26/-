const express = require('express');
const { chromium } = require('playwright');
const app = express();

// 🛠️ สำคัญมาก: ต้องไปสมัครที่ browserless.io แล้วเอา Token ยาวๆ มาแปะตรงนี้ให้ถูกต้องนะคะ
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
            <h1 class="text-2xl font-bold text-slate-700 mb-6 tracking-tight">พิกัดดาวรายวัน (เวอร์ชันปลอดภัย)</h1>
            
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

            <button id="exportBtn" onclick="downloadData()" 
                    class="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-semibold py-4 rounded-xl transition duration-200 shadow-lg flex justify-center items-center gap-3 text-lg">
                📁 ดึงข้อมูลพิกัดดาวรายวัน
            </button>
            
            <div id="resultBox" class="mt-6 p-4 bg-slate-50 rounded-xl border border-slate-200 text-left hidden">
                <p class="text-sm font-bold text-slate-700 mb-2">💡 คัดลอกข้อมูลด้านล่างนี้ไปวางใน Excel ได้เลยค่ะ:</p>
                <textarea id="csvOutput" class="w-full h-48 p-2 bg-white border rounded-lg font-mono text-xs focus:outline-none" readonly></textarea>
            </div>
        </div>

        <script>
            async function downloadData() {
                const year = document.getElementById('targetYear').value;
                const month = document.getElementById('targetMonth').value;
                const btn = document.getElementById('exportBtn');
                const resultBox = document.getElementById('resultBox');
                const csvOutput = document.getElementById('csvOutput');
                
                btn.disabled = true;
                btn.innerHTML = "⏳ กำลังดึงข้อมูล...";

                try {
                    const response = await fetch('/get-astro-data?year=' + year + '&month=' + month);
                    const text = await response.text();
                    
                    resultBox.classList.remove('hidden');
                    csvOutput.value = text;
                    btn.innerHTML = "📁 ดึงข้อมูลสำเร็จ!";
                } catch (err) {
                    alert("เกิดข้อผิดพลาดในการดึงข้อมูลค่ะ");
                } finally {
                    btn.disabled = false;
                }
            }
        </script>
    </body>
    </html>
    `);
});

app.get('/get-astro-data', async (req, res) => {
    const cYear = parseInt(req.query.year) || 2024;
    const month = parseInt(req.query.month) || 1;
    const thYear = cYear + 543;

    let browser;
    try {
        // ต่อเข้าคลาวด์ภายนอกตรงๆ ทำให้ไม่เรียกหา playwright-core ในตัว Vercel เอง
        browser = await chromium.connectOverCDP(
            `wss://chrome.browserless.io?token=${BROWSERLESS_TOKEN}`
        );
        
        const context = await browser.newContext();
        const page = await context.newPage();
        
        const months_th = ['', 'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
        const targetUrl = `https://myhora.com/calendar/astro-suriyayas-${month}-${thYear}.aspx`;

        await page.goto(targetUrl, { waitUntil: 'load', timeout: 8000 });

        const dataText = await page.evaluate((info) => {
            let output = "ปี ค.ศ.,ปี พ.ศ.,เดือน,ที่,วัน,ข-ร,ด.,อาทิตย์ (๑),จันทร์ (๒),ยก,ฤกษ์,เต็ม,ดิถี,เต็ม,อังคาร (๓),พุธ (๔),พฤหัสฯ (๕),ศุกร์ (๖),เสาร์ (๗),ราหู (๘),เกตุ (๙)\\n";
            const trs = document.querySelectorAll('table tr');
            
            trs.forEach((tr) => {
                const tds = tr.querySelectorAll('td');
                if (tds.length >= 15) {
                    const cells = Array.from(tds).map(td => td.innerText ? td.innerText.trim().replace(/\\s+/g, ' ') : '');
                    const dayNumber = parseInt(cells[0]);
                    if (!isNaN(dayNumber) && dayNumber >= 1 && dayNumber <= 31) {
                        output += `\${info.cYear},\${info.thYear},\${info.mName},\${cells.join(',')}\\n`;
                    }
                }
            });
            return output;
        }, { cYear, thYear, mName: months_th[month] });

        res.setHeader('Content-Type', 'text/plain; charset=utf-8');
        res.send(dataText);

    } catch (err) {
        res.status(500).send(`Error: ${err.message}`);
    } finally {
        if (browser) await browser.close();
    }
});

module.exports = app;
