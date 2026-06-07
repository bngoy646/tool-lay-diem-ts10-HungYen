const { chromium } = require('playwright');
const xlsx = require('xlsx');
const fs = require('fs');
const DAU_SO_LIST = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 25, 26, 27, 28, 29, 30, 31, 32, 33, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 54, 55, 56, 57, 58, 211, 212, 231, 232, 341, 342, 521, 522, 591, 592, 611, 612];
const GIOI_HAN_SBD_MOI_TINH = 70000; 
const KHOANG_DUNG_KHI_HET_DIEM = 50; 
const SO_LUONG_TAB = 20; 
const MAX_RETRY = 2; 

let hangDoiSBD = [];
let trangThaiDauSo = {}; 
let tatCaHocSinh = [];
let tongThanhCong = 0;
let tongLoi = 0;
let isStopping = false;

DAU_SO_LIST.forEach(ds => {
    const prefix = String(ds).padStart(2, '0');
    trangThaiDauSo[prefix] = { countLoi: 0, hoanThanh: false };
    const padLength = 8 - prefix.length; 
    for (let i = 1; i <= GIOI_HAN_SBD_MOI_TINH; i++) {
        hangDoiSBD.push({ sbd: prefix + String(i).padStart(padLength, '0'), prefix: prefix });
    }
});
function ghiFileExcel() {
    if (tatCaHocSinh.length === 0) return;
    const headers = [
        "STT", "SBD", "Trường", 
        "KHTN (Hóa học chuyên)", "KHTN (Sinh học chuyên)", "KHTN (Vật lý chuyên)", 
        "LSDL (Lịch sử chuyên)", "LSDL (Địa lý chuyên)", 
        "Ngữ văn chuyên", "Tin học chuyên", "Tiếng anh chuyên", "Toán học chuyên", // Nhóm chuyên
        "Ngữ văn chung", "Tiếng anh chung", "Toán học chung" // Đã đẩy xuống cuối
    ];

    const getScore = (hs, subjectKey, isSpecialized = null) => {
        const keys = Object.keys(hs);
        const match = keys.find(k => {
            const kLow = k.toLowerCase();
            const sLow = subjectKey.toLowerCase();
            if (!kLow.includes(sLow)) return false;
            if (isSpecialized === true) return kLow.includes("chuyên");
            if (isSpecialized === false) return !kLow.includes("chuyên");
            return true;
        });
        if (match) {
            const val = parseFloat(hs[match]);
            return isNaN(val) ? "" : val;
        }
        return "";
    };

    tatCaHocSinh.sort((a, b) => a["Số báo danh"].localeCompare(b["Số báo danh"]));

    const exportData = [headers];

    tatCaHocSinh.forEach((hs, index) => {
        // 2. Cập nhật thứ tự nạp dữ liệu tương ứng với Header phía trên
        const row = [
            index + 1,
            hs["Số báo danh"],
            hs["Trường dự thi"] || "",
            getScore(hs, "Hóa học"),         // KHTN (Hóa học chuyên)
            getScore(hs, "Sinh học"),        // KHTN (Sinh học chuyên)
            getScore(hs, "Vật lý"),          // KHTN (Vật lý chuyên)
            getScore(hs, "Lịch sử"),         // LSDL (Lịch sử chuyên)
            getScore(hs, "Địa lý"),          // LSDL (Địa lý chuyên)
            getScore(hs, "Ngữ văn", true),   // Ngữ văn chuyên
            getScore(hs, "Tin học"),         // Tin học chuyên
            getScore(hs, "Tiếng anh", true), // Tiếng anh chuyên
            getScore(hs, "Toán", true),      // Toán học chuyên
            getScore(hs, "Ngữ văn", false),  // Ngữ văn chung
            getScore(hs, "Tiếng anh", false),// Tiếng anh chung
            getScore(hs, "Toán", false)       // Toán học chung
        ];
        exportData.push(row);
    });

    const wb = xlsx.utils.book_new();
    const ws = xlsx.utils.aoa_to_sheet(exportData);
    xlsx.utils.book_append_sheet(wb, ws, "KetQua");
    
    const fileName = `KetQua_TuyenSinh_Full.xlsx`;
    try {
        xlsx.writeFile(wb, fileName);
        console.log(`\nĐã lưu file: ${fileName}`);
    } catch (e) {
        xlsx.writeFile(wb, `Backup_${Date.now()}.xlsx`);
    }
}

process.on('SIGINT', () => {
    isStopping = true;
    console.log("\nStopping");
    ghiFileExcel();
    process.exit(0);
});
async function cauHinhGiaoDien(page) {
    try {
        await page.goto('https://tsdc.edu.vn/hung-yen/tra-cuu-diem-thi', { waitUntil: 'domcontentloaded', timeout: 45000 });
        await page.evaluate(async () => {
            const delay = (ms) => new Promise(res => setTimeout(res, ms));
            const dropdowns = Array.from(document.querySelectorAll(".el-input__inner")).filter(i => i.placeholder.includes("Chọn"));
            if(dropdowns.length > 0) {
                dropdowns[0].click(); await delay(600);
                const thpt = Array.from(document.querySelectorAll(".el-select-dropdown__item")).find(el => el.textContent.includes("THPT"));
                if (thpt) thpt.click(); await delay(600);
                const kythi = dropdowns[dropdowns.length - 1];
                kythi.click(); await delay(600);
                const lop10 = Array.from(document.querySelectorAll(".el-select-dropdown__item")).find(el => el.textContent.includes("lớp 10"));
                if (lop10) lop10.click(); await delay(800);
            }
        });
        return true;
    } catch (e) { return false; }
}

async function chayWorker(browser, tabIndex) {
    const context = await browser.newContext();
    const page = await context.newPage();
    await page.route('**/*', (r) => ['image', 'font'].includes(r.request().resourceType()) ? r.abort() : r.continue());

    let ready = await cauHinhGiaoDien(page);
    if (!ready) { await context.close(); return; }

    while (hangDoiSBD.length > 0 && !isStopping) {
        const task = hangDoiSBD.shift();
        if (!task) break;

        const { sbd, prefix } = task;
        if (trangThaiDauSo[prefix].hoanThanh) continue;

        let retryCount = 0;
        let success = false;

        while (retryCount <= MAX_RETRY && !success) {
            try {
                await page.evaluate(() => {
                    const table = document.querySelector(".el-table");
                    if (table && table.__vue__) table.__vue__.store.states.data = []; 
                    document.querySelectorAll('.el-message, .el-notification').forEach(el => el.remove());
                    const input = document.querySelector("input[placeholder='Nhập']");
                    if (input) { input.value = ''; input.dispatchEvent(new Event('input', { bubbles: true })); }
                });

                const apiPromise = page.waitForResponse(res => res.url().includes('-api'), { timeout: 6000 }).catch(() => null);

                await page.evaluate((val) => {
                    const input = document.querySelector("input[placeholder='Nhập']");
                    if (input) {
                        input.value = val;
                        input.dispatchEvent(new Event('input', { bubbles: true }));
                        const btn = Array.from(document.querySelectorAll("button")).find(el => el.textContent.includes("Tra cứu"));
                        if (btn) btn.click();
                    }
                }, sbd);

                await apiPromise;
                await page.waitForTimeout(250); 

                const result = await page.evaluate((currentSbd) => {
                    const table = document.querySelector(".el-table");
                    if (!table || !table.__vue__ || table.__vue__.store.states.data.length === 0) return null;
                    
                    let school = "";
                    const target = Array.from(document.querySelectorAll('div, p, span')).find(el => el.innerText && el.innerText.includes("Trường dự thi:"));
                    if (target) school = target.innerText.split("Trường dự thi:")[1]?.split("\n")[0]?.trim() || "";

                    let res = { "Số báo danh": currentSbd, "Trường dự thi": school };
                    table.__vue__.store.states.data.forEach(row => {
                        const ten = row.tenMonThi || row.ten_bai_thi || "Môn khác";
                        res[ten] = String(row.diem || row.diemThi || "0").trim();
                    });
                    return res;
                }, sbd);

                if (result) {
                    tatCaHocSinh.push(result);
                    tongThanhCong++;
                    trangThaiDauSo[prefix].countLoi = 0; 
                    success = true;
                    process.stdout.write(`\r[Tab ${tabIndex}] 🟢 OK: ${sbd} | Tổng: ${tongThanhCong} `);
                } else {
                    retryCount++;
                }
            } catch (e) {
                retryCount++;
                if (e.message.includes('closed')) await cauHinhGiaoDien(page);
            }
        }

        if (!success) {
            tongLoi++;
            trangThaiDauSo[prefix].countLoi++;
            if (trangThaiDauSo[prefix].countLoi >= KHOANG_DUNG_KHI_HET_DIEM) {
                if (!trangThaiDauSo[prefix].hoanThanh) {
                    trangThaiDauSo[prefix].hoanThanh = true;
                    console.log(`\n⏩ Hoàn tất đầu số ${prefix}.`);
                    hangDoiSBD = hangDoiSBD.filter(item => item.prefix !== prefix);
                }
            }
        }
    }
    await context.close();
}

(async () => {
    console.clear();
    console.log("Lấy điểm thi tuyển sinh vào 10 tỉnh Hưng Yên");
    const browser = await chromium.launch({ headless: true, channel: 'chrome' }); 

    const workers = [];
    for (let i = 0; i < SO_LUONG_TAB; i++) {
        workers.push(chayWorker(browser, i + 1));
        await new Promise(r => setTimeout(r, 800));
    }
    
    await Promise.all(workers);
    await browser.close();
    ghiFileExcel();
    console.log(`\nDone.`);
})();