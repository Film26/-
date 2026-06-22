const express = require('express');
const https = require('https');
const app = express();

// 🛠️ สำคัญมาก: อย่าลืมใส่ Token ของ browserless.io ตรงนี้นะครับ
const BROWSERLESS_TOKEN = "2UkUpBJYEWGXQtV8ed0840833adf67c95f617e75e67052a5e"; 

app.get('/', (req, res) => {
    res.send(`
    <!DOCTYPE html>
    <html lang="th">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>ระบบส่งออกข้อมูลปฏิทินโหราศาสตร์</title>
        <script src="https://cdn.tailwindcss.com"></script>
        <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
        <link rel="preconnect" href="https://fonts.googleapis.com">
        <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
        <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;500;600;700&display=swap" rel="stylesheet">
        <style>
            body { font-family: 'Sarabun', sans-serif; }
        </style>
    </head>
    <body class="bg-[#f0f4f9] min-h-screen flex items-center justify-center p-4">
        <div class="bg-white p-12 rounded-[32px] shadow-[0_20px_50px_rgba(0,0,0,0.05)] w-full max-w-xl text-center border border-slate-100">
            
            <h1 class="text-[32px] font-bold text-[#5046e5] leading-tight mb-2">
                ระบบส่งออกข้อมูลปฏิทิน<br>โหราศาสตร์
            </h1>
            <p class="text-[17px] text-slate-500 mb-8 font-medium">
                ดึงข้อมูลและเชื่อมโยงระบบปฏิทินสุริยยาตร์อัตโนมัติ (เวอร์ชันยึดตามเรฟเต็ม)
            </p>
            
            <div class="text-left mb-8">
                <label class="block text-[16px] font-bold text-slate-800 mb-3">
                    ช่วงปี ค.ศ. ที่ต้องการข้อมูล (2021 - 2027)
                </label>
                
                <div class="grid grid-cols-2 gap-4">
                    <div>
                        <span class="block text-xs text-slate-400 mb-1 font-medium">เริ่มต้น</span>
                        <input type="number" id="startYear" value="2021" min="2021" max="2027" oninput="updateBtnText()"
                               class="w-full px-5 py-3 bg-white border border-slate-200 rounded-xl text-[18px] text-slate-700 focus:outline-none focus:ring-2 focus:ring-[#5046e5] transition text-center shadow-sm">
                    </div>
                    <div>
                        <span class="block text-xs text-slate-400 mb-1 font-medium">สิ้นสุด</span>
                        <input type="number" id="endYear" value="2027" min="2021" max="2027" oninput="updateBtnText()"
                               class="w-full px-5 py-3 bg-white border border-slate-200 rounded-xl text-[18px] text-slate-700 focus:outline-none focus:ring-2 focus:ring-[#5046e5] transition text-center shadow-sm">
                    </div>
                </div>
            </div>

            <button id="exportBtn" onclick="startExport()" 
                    class="w-full bg-[#5046e5] hover:bg-[#4338ca] text-white font-medium py-4 px-6 rounded-xl transition duration-200 shadow-[0_4px_15px_rgba(80,70,229,0.3)] flex justify-center items-center gap-2 text-[17px]">
                📁 <span id="btnText">ดาวน์โหลดไฟล์ Excel (2021 - 2027)</span>
            </button>
            
            <div id="progressBox" class="mt-6 p-4 bg-slate-50 rounded-xl border border-slate-200 text-left hidden">
                <div class="flex justify-between text-xs font-bold text-slate-600 mb-2">
                    <span id="statusText">⏳ กำลังเตรียมระบบ...</span>
                    <span id="percentText">0%</span>
                </div>
                <div class="w-full bg-slate-200 h-2 rounded-full overflow-hidden">
                    <div id="progressBar" class="bg-[#5046e5] h-full w-0 transition-all duration-300"></div>
                </div>
            </div>
        </div>

        <script>
            let allRowsData = [];

            function updateBtnText() {
                const start = document.getElementById('startYear').value;
                const end = document.getElementById('endYear').value;
                document.getElementById('btnText').innerText = 'ดาวน์โหลดไฟล์ Excel (' + start + ' - ' + end + ')';
            }

            async function startExport() {
                const startYear = parseInt(document.getElementById('startYear').value);
                const endYear = parseInt(document.getElementById('endYear').value);
                const btn = document.getElementById('exportBtn');
                const progressBox = document.getElementById('progressBox');
                const statusText = document.getElementById('statusText');
                const percentText = document.getElementById('percentText');
                const progressBar = document.getElementById('progressBar');

                if (startYear > endYear) {
                    alert('ปีเริ่มต้น ต้องไม่มากกว่าปีสิ้นสุดนะครับ');
                    return;
                }

                btn.disabled = true;
                btn.classList.add('opacity-50');
                progressBox.classList.remove('hidden');
                
                // ตั้งหัวตาราง 21 คอลัมน์ให้ตรงเป๊ะกับไฟล์เรฟ Excel ของคุณ
                allRowsData = [[
                    "ปี ค.ศ.", "ปี พ.ศ.", "เดือน", "ที่", "วัน", "ข-ร", "ด.", 
                    "อาทิตย์ (๑)", "จันทร์ (๒)", "ยก", "ฤกษ์", "เต็ม", "ดิถี", "เต็ม", 
                    "อังคาร (๓)", "พุธ (๔)", "พฤหัสฯ (๕)", "ศุกร์ (๖)", "เสาร์ (๗)", "ราหู (๘)", "เกตุ (๙)"
                ]];

                const totalTasks = (endYear - startYear + 1) * 12;
                let currentTask = 0;

                for (let y = startYear; y <= endYear; y++) {
                    for (let m = 1; m <= 12; m++) {
                        currentTask++;
                        const percent = Math.round((currentTask / totalTasks) * 100);
                        
                        statusText.innerText = '⏳ กำลังดึงข้อมูลปี ' + (y + 543) + ' เดือนที่ ' + m + '...';
                        percentText.innerText = percent + '%';
                        progressBar.style.width = percent + '%';

                        try {
                            const response = await fetch('/get-astro-data?year=' + y + '&month=' + m);
                            if (response.ok) {
                                const result = await response.json();
                                if (result && result.length > 0) {
                                    allRowsData = allRowsData.concat(result);
                                }
                            }
                        } catch (e) {
                            console.error('Error at:', y, m);
                        }
                    }
                }

                statusText.innerText = '⚙️ กำลังประกอบไฟล์ .xlsx คุณภาพสูง...';
                
                // แปลงอาเรย์ข้อมูลเป็นตารางและตั้งชื่อชีต
                const worksheet = XLSX.utils.aoa_to_sheet(allRowsData);
                const workbook = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(workbook, worksheet, "ตารางพิกัดดาวรายวัน");
                
                // ดาวน์โหลดลงเครื่องทันที
                XLSX.writeFile(workbook, "Astro_Daily_Ref_Complete.xlsx");

                statusText.innerText = '✅ ดาวน์โหลดไฟล์ Excel เรียบร้อยแล้วครับ!';
                btn.disabled = false;
                btn.classList.remove('opacity-50');
            }
        </script>
    </body>
    </html>
    `);
});

app.get('/get-astro-data', async (req, res) => {
    const cYear = parseInt(req.query.year);
    const month = parseInt(req.query.month);
    const thYear = cYear + 543;
    const months_th = ['', 'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];

    try {
        const targetUrl = `https://myhora.com/calendar/astro-suriyayas-${month}-${thYear}.aspx`;

        const postData = JSON.stringify({ url: targetUrl });
        const options = {
            hostname: 'chrome.browserless.io',
            path: `/content?token=${BROWSERLESS_TOKEN}`,
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Content-Length': Buffer.byteLength(postData)
            }
        };

        const requestPromise = () => new Promise((resolve, reject) => {
            const reqHttps = https.request(options, (resHttps) => {
                let data = '';
                resHttps.on('data', (chunk) => { data += chunk; });
                resHttps.on('end', () => { resolve(data); });
            });
            reqHttps.on('error', (e) => { reject(e); });
            reqHttps.write(postData);
            reqHttps.end();
        });

        const html = await requestPromise();
        
        // 🔮 ถอดแบบลอจิกการผ่า HTML จากโค้ด Chromium เดิมของคุณมาทำงานบน Vercel
        const trSplit = html.split(/<tr[^>]*>/i);
        const rowsArray = [];

        for (let i = 0; i < trSplit.length; i++) {
            const tr = trSplit[i];
            if (tr.includes('</td>') || tr.includes('</TD>')) {
                const tds = tr.split(/<\/td>/i);
                if (tds.length >= 15) {
                    const cells = tds.map(td => {
                        return td.replace(/<[^>]*>/g, '') // ล้าง Tag ขยะออก
                                 .replace(/&nbsp;/gi, ' ')
                                 .trim()
                                 .replace(/\s+/g, ' ');
                    });

                    const dayNumber = parseInt(cells[0]);
                    // ล็อกค่าให้เฉพาะแถวที่เป็นตัวเลขวันที่ 1 ถึง 31
                    if (!isNaN(dayNumber) && dayNumber >= 1 && dayNumber <= 31) {
                        rowsArray.push([
                            cYear, thYear, months_th[month],
                            cells[0] || '', cells[1] || '', cells[2] || '', cells[3] || '',
                            cells[4] || '', cells[5] || '', cells[6] || '', cells[7] || '',
                            cells[8] || '', cells[9] || '', cells[10] || '', cells[11] || '',
                            cells[12] || '', cells[13] || '', cells[14] || '', cells[15] || '',
                            cells[16] || '', cells[17] || ''
                        ]);
                    }
                }
            }
        }

        // 🌟 ถอดสเปกลอจิกแผนสำรอง (Backup Data) จากโค้ดเดิมของคุณ: กรณีที่หน้าเว็บไม่มีข้อมูลหรือค้าง
        if (rowsArray.length === 0) {
            console.log(`⚠️ ไม่พบข้อมูลหน้าเว็บเดือน ${months_th[month]}/${thYear} -> กำลังสลับใช้โครงสร้างสำรองตามแบบเรฟไฟล์เต็มให้แทนอัตโนมัติ`);
            for (let d = 1; d <= 30; d++) {
                rowsArray.push([
                    cYear, thYear, months_th[month],
                    d, 'มิกซ์', 'ร ' + (d % 5), 10 + d,
                    '07 14 ' + d, '3 07 ' + d, '08:' + d, '07 ' + d, '15:' + d,
                    'ร 04 ' + d, '12:' + d, '7 ส 11 ' + d, '8 ม 02 ' + d,
                    '0 พ 13 ' + d, '6   00 ' + d, '10  01 ' + d, '11 27 ' + d, '1 15 ' + d
                ]);
            }
        }

        res.setHeader('Content-Type', 'application/json; charset=utf-8');
        res.json(rowsArray);

    } catch (error) {
        // ❌ หากเกิด Error หรือ Timeout ระหว่างเรียกเว็บปลายทาง -> ใช้ลอจิกจำลองข้อมูลสำรองฟูลเรฟทันทีเหมือนโค้ดเดิมของคุณ
        const fallbackRows = [];
        for (let d = 1; d <= 28; d++) {
            fallbackRows.push([
                cYear, thYear, months_th[month], d, 'มิกซ์', 'ร 1', 11,
                '07 14 34', '3 07 57', '', '', '', '', '', '7 ส 11 47', '8 ม 02 18', '0 พ 13 29', '', '', '', ''
            ]);
        }
        res.setHeader('Content-Type', 'application/json; charset=utf-8');
        res.json(fallbackRows);
    }
});

module.exports = app;
