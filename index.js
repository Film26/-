const express = require('express');
const https = require('https');
const app = express();

// 🛠️ สำคัญมาก: ใส่ Token ของ browserless.io ตรงนี้เหมือนเดิมนะคะ
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
        <link rel="preconnect" href="https://fonts.googleapis.com">
        <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
        <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;500;600;700&display=swap" rel="stylesheet">
        <style>
            body { font-family: 'Sarabun', sans-serif; }
        </style>
    </head>
    <body class="bg-[#f0f4f9] min-h-screen flex items-center justify-center p-4">
        <!-- การ์ดดีไซน์ตามเรฟภาพ image_554e00.png -->
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

            <!-- ปุ่มสีน้ำเงิน/ม่วงตามสไตล์ในภาพเรฟ -->
            <button id="exportBtn" onclick="startExport()" 
                    class="w-full bg-[#5046e5] hover:bg-[#4338ca] text-white font-medium py-4 px-6 rounded-xl transition duration-200 shadow-[0_4px_15px_rgba(80,70,229,0.3)] flex justify-center items-center gap-2 text-[17px]">
                📁 <span id="btnText">ดาวน์โหลดไฟล์ Excel (2021 - 2027)</span>
            </button>
            
            <!-- กล่องแสดงสถานะและความก้าวหน้า -->
            <div id="progressBox" class="mt-6 p-4 bg-slate-50 rounded-xl border border-slate-200 text-left hidden">
                <div class="flex justify-between text-xs font-bold text-slate-600 mb-2">
                    <span id="statusText">⏳ กำลังเตรียมระบบ...</span>
                    <span id="percentText">0%</span>
                </div>
                <div class="w-full bg-slate-200 h-2 rounded-full overflow-hidden">
                    <div id="progressBar" class="bg-[#5046e5] h-full w-0 transition-all duration-300"></div>
                </div>
                <textarea id="csvAccumulator" class="w-full h-32 p-2 bg-white border rounded-lg font-mono text-[10px] mt-3 focus:outline-none hidden" readonly></textarea>
            </div>
        </div>

        <script>
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
                const csvAccumulator = document.getElementById('csvAccumulator');

                if (startYear > endYear) {
                    alert('ปีเริ่มต้น ต้องไม่มากกว่าปีสิ้นสุดนะคะ');
                    return;
                }

                btn.disabled = true;
                btn.classList.add('opacity-50');
                progressBox.classList.remove('hidden');
                csvAccumulator.value = "ปี ค.ศ.,ปี พ.ศ.,เดือน,ที่,วัน,ข-ร,ด.,อาทิตย์ (๑),จันทร์ (๒),ยก,ฤกษ์,เต็ม,ดิถี,เต็ม,อังคาร (๓),พุธ (๔),พฤหัสฯ (๕),ศุกร์ (๖),เสาร์ (๗),ราหู (๘),เกตุ (๙)\\n";

                // คำนวณจำนวนรอบงานทั้งหมด (ปี x 12 เดือน)
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
                                const chunkData = await response.text();
                                csvAccumulator.value += chunkData;
                            }
                        } catch (e) {
                            console.error('Skip error at:', y, m);
                        }
                    }
                }

                statusText.innerText = '✅ รวมไฟล์ CSV สำเร็จ! กำลังเริ่มดาวน์โหลด...';
                
                // แปลงข้อความเป็นไฟล์ CSV ให้กดดาวน์โหลดอัตโนมัติ
                const blob = new Blob([csvAccumulator.value], { type: 'text/csv;charset=utf-8;' });
                const link = document.createElement("a");
                link.href = URL.createObjectURL(blob);
                link.setAttribute("download", "Astro_Data_" + startYear + "_" + endYear + ".csv");
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);

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
            path: `/content?token=${BROWSERLESS_TOKEN}`,
            method: 'POST',
            headers: {
                'Content-Type': 'application/vnd.api+json',
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
        const trSplit = html.split('<tr');
        let chunkOutput = "";

        for (let i = 0; i < trSplit.length; i++) {
            const tr = trSplit[i];
            if (tr.includes('</td>')) {
                const tds = tr.split('</td>');
                if (tds.length >= 15) {
                    const cells = tds.map(td => td.replace(/<[^>]*>/g, '').trim().replace(/\\s+/g, ' '));
                    const dayNumber = parseInt(cells[0]);
                    if (!isNaN(dayNumber) && dayNumber >= 1 && dayNumber <= 31) {
                        chunkOutput += `${cYear},${thYear},${months_th[month]},${cells.slice(0, 18).join(',')}\n`;
                    }
                }
            }
        }

        res.setHeader('Content-Type', 'text/plain; charset=utf-8');
        res.send(chunkOutput);

    } catch (err) {
        res.status(500).send("");
    }
});

module.exports = app;
