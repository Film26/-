const express = require('express');
const https = require('https');
const app = express();

// 🛠️ สำคัญมาก: ใส่ Token ของ browserless.io ตรงนี้เหมือนเดิมนะครับ
const BROWSERLESS_TOKEN = "วาง_TOKEN_ของคุณตรงนี้"; 

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
                ดึงข้อมูลและเชื่อมโยงระบบปฏิทินสุริยยาตร์อัตโนมัติ
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

                statusText.innerText = '⚙️ กำลังประกอบไฟล์ .xlsx...';
                
                const worksheet = XLSX.utils.aoa_to_sheet(allRowsData);
                const workbook = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(workbook, worksheet, "ตารางพิกัดดาวรายวัน");
                
                XLSX.writeFile(workbook, "Astro_Daily_Data_" + startYear + "_" + endYear + ".xlsx");

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

    try {
        const months_th = ['', 'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
        const targetUrl = `https://myhora.com/calendar/astro-suriyayas-${month}-${thYear}.aspx`;

        const postData = JSON.stringify({ url: targetUrl });
        const options = {
            hostname: 'chrome.browserless.io',
            path: `/content?token=${BROWSERLESS_TOKEN}`, // กลับมาใช้ตัวจับเว็บดั้งเดิมแบบเสถียร
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
        
        // ใช้การสปลิตตารางแบบยืดหยุ่นสูง เพื่อป้องกันกรณีโครงสร้างคลาสเปลี่ยน
        const trSplit = html.split(/<tr[^>]*>/i);
        const rowsArray = [];

        for (let i = 0; i < trSplit.length; i++) {
            const tr = trSplit[i];
            if (tr.includes('</td>') || tr.includes('</TD>')) {
                // สปลิตแต่ละคอลัมน์ด้วย td ปิด
                const tds = tr.split(/<\/td>/i);
                if (tds.length >= 15) {
                    // ทำความสะอาดเอาโค้ด HTML Tag ออกทั้งหมดให้เหลือแค่คำเดี่ยวๆ
                    const cells = tds.map(td => {
                        return td.replace(/<[^>]*>/g, '') // ลบ Tag ทั้งหมด
                                 .replace(/&nbsp;/gi, ' ') // แปลง Space ขยะ
                                 .trim()
                                 .replace(/\s+/g, ' ');
                    });

                    const dayNumber = parseInt(cells[0]);
                    // ล็อกค่าให้เฉพาะแถวที่ขึ้นต้นด้วยวันที่ 1 ถึง 31 ผ่านเท่านั้น
                    if (!isNaN(dayNumber) && dayNumber >= 1 && dayNumber <= 31) {
                        rowsArray.push([
                            cYear, 
                            thYear, 
                            months_th[month], 
                            cells[0] || '', // ที่
                            cells[1] || '', // วัน
                            cells[2] || '', // ข-ร
                            cells[3] || '', // ด.
                            cells[4] || '', // อาทิตย์ (๑)
                            cells[5] || '', // จันทร์ (๒)
                            cells[6] || '', // ยก
                            cells[7] || '', // ฤกษ์
                            cells[8] || '', // เต็ม
                            cells[9] || '', // ดิถี
                            cells[10] || '', // เต็ม
                            cells[11] || '', // อังคาร (๓)
                            cells[12] || '', // พุธ (๔)
                            cells[13] || '', // พฤหัสฯ (๕)
                            cells[14] || '', // ศุกร์ (๖)
                            cells[15] || '', // เสาร์ (๗)
                            cells[16] || '', // ราหู (๘)
                            cells[17] || ''  // เกตุ (๙)
                        ]);
                    }
                }
            }
        }

        res.setHeader('Content-Type', 'application/json; charset=utf-8');
        res.json(rowsArray);

    } catch (err) {
        res.status(500).json([]);
    }
});

module.exports = app;
