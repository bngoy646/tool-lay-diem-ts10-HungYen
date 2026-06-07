import asyncio
import sys
import time
from collections import deque
import pandas as pd
from playwright.async_api import async_playwright

DAU_SO_LIST = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 25, 26, 27, 28, 29, 30, 31, 32, 33, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 54, 55, 56, 57, 58, 211, 212, 231, 232, 341, 342, 521, 522, 591, 592, 611, 612]
GIOI_HAN_SBD_MOI_TINH = 70000
KHOANG_DUNG_KHI_HET_DIEM = 50
SO_LUONG_TAB = 20
MAX_RETRY = 2

hangDoiSBD = deque()
trangThaiDauSo = {}
tatCaHocSinh = []
tongThanhCong = 0
tongLoi = 0
isStopping = False

# Tạo hàng đợi SBD chuẩn 8 chữ số
for ds in DAU_SO_LIST:
    prefix = str(ds).zfill(2)
    trangThaiDauSo[prefix] = {'countLoi': 0, 'hoanThanh': False}
    padLength = 8 - len(prefix)
    for i in range(1, GIOI_HAN_SBD_MOI_TINH + 1):
        hangDoiSBD.append({'sbd': prefix + str(i).zfill(padLength), 'prefix': prefix})

def ghiFileExcel():
    if not tatCaHocSinh:
        return

    headers = [
        "STT", "SBD", "Trường", 
        "KHTN (Hóa học chuyên)", "KHTN (Sinh học chuyên)", "KHTN (Vật lý chuyên)", 
        "LSDL (Lịch sử chuyên)", "LSDL (Địa lý chuyên)", 
        "Ngữ văn chuyên", "Tin học chuyên", "Tiếng anh chuyên", "Toán học chuyên", 
        "Ngữ văn chung", "Tiếng anh chung", "Toán học chung"
    ]

    def getScore(hs, subjectKey, isSpecialized=None):
        for key in hs.keys():
            k_low = key.lower()
            s_low = subjectKey.lower()
            if s_low not in k_low:
                continue
            if isSpecialized is True and "chuyên" not in k_low:
                continue
            if isSpecialized is False and "chuyên" in k_low:
                continue
            try:
                return float(hs[key])
            except ValueError:
                return ""
        return ""

    tatCaHocSinh.sort(key=lambda x: x["Số báo danh"])

    exportData = []
    for index, hs in enumerate(tatCaHocSinh):
        row = [
            index + 1,
            hs["Số báo danh"],
            hs.get("Trường dự thi", ""),
            getScore(hs, "Hóa học"),         
            getScore(hs, "Sinh học"),        
            getScore(hs, "Vật lý"),          
            getScore(hs, "Lịch sử"),         
            getScore(hs, "Địa lý"),          
            getScore(hs, "Ngữ văn", True),   
            getScore(hs, "Tin học"),         
            getScore(hs, "Tiếng anh", True), 
            getScore(hs, "Toán", True),      
            getScore(hs, "Ngữ văn", False),  
            getScore(hs, "Tiếng anh", False),
            getScore(hs, "Toán", False)      
        ]
        exportData.append(row)

    df = pd.DataFrame(exportData, columns=headers)
    fileName = "KetQua_TuyenSinh_Full.xlsx"
    
    try:
        df.to_excel(fileName, index=False)
        print(f"\nĐã lưu file: {fileName}")
    except Exception:
        backup_name = f"Backup_{int(time.time())}.xlsx"
        df.to_excel(backup_name, index=False)
        print(f"\nLỗi lưu file chính, đã lưu backup: {backup_name}")

async def cauHinhGiaoDien(page):
    try:
        await page.goto('https://tsdc.edu.vn/hung-yen/tra-cuu-diem-thi', wait_until='domcontentloaded', timeout=45000)
        js_config = """
        async () => {
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
        }
        """
        await page.evaluate(js_config)
        return True
    except Exception:
        return False

async def chayWorker(browser, tabIndex):
    global tongThanhCong, tongLoi, isStopping

    context = await browser.new_context()
    page = await context.new_page()
    
    await page.route('**/*', lambda r: r.abort() if r.request.resource_type in ['image', 'font'] else r.continue_())

    ready = await cauHinhGiaoDien(page)
    if not ready:
        await context.close()
        return

    while hangDoiSBD and not isStopping:
        task = hangDoiSBD.popleft()
        sbd = task['sbd']
        prefix = task['prefix']

        if trangThaiDauSo[prefix]['hoanThanh']:
            continue

        retryCount = 0
        success = False

        while retryCount <= MAX_RETRY and not success:
            try:
                await page.evaluate("""
                () => {
                    const table = document.querySelector(".el-table");
                    if (table && table.__vue__) table.__vue__.store.states.data = []; 
                    document.querySelectorAll('.el-message, .el-notification').forEach(el => el.remove());
                    const input = document.querySelector("input[placeholder='Nhập']");
                    if (input) { input.value = ''; input.dispatchEvent(new Event('input', { bubbles: true })); }
                }
                """)

                js_submit = """
                (val) => {
                    const input = document.querySelector("input[placeholder='Nhập']");
                    if (input) {
                        input.value = val;
                        input.dispatchEvent(new Event('input', { bubbles: true }));
                        const btn = Array.from(document.querySelectorAll("button")).find(el => el.textContent.includes("Tra cứu"));
                        if (btn) btn.click();
                    }
                }
                """
                
                try:
                    async with page.expect_response(lambda response: '-api' in response.url, timeout=6000):
                        await page.evaluate(js_submit, sbd)
                except Exception:
                    pass

                await page.wait_for_timeout(250)

                result = await page.evaluate("""
                (currentSbd) => {
                    const table = document.querySelector(".el-table");
                    if (!table || !table.__vue__ || table.__vue__.store.states.data.length === 0) return null;
                    
                    let school = "";
                    const target = Array.from(document.querySelectorAll('div, p, span')).find(el => el.innerText && el.innerText.includes("Trường dự thi:"));
                    if (target) school = target.innerText.split("Trường dự thi:")[1]?.split("\\n")[0]?.trim() || "";

                    let res = { "Số báo danh": currentSbd, "Trường dự thi": school };
                    table.__vue__.store.states.data.forEach(row => {
                        const ten = row.tenMonThi || row.ten_bai_thi || "Môn khác";
                        res[ten] = String(row.diem || row.diemThi || "0").trim();
                    });
                    return res;
                }
                """, sbd)

                if result:
                    tatCaHocSinh.append(result)
                    tongThanhCong += 1
                    trangThaiDauSo[prefix]['countLoi'] = 0
                    success = True
                    sys.stdout.write(f"\r[Tab {tabIndex}] 🟢 OK: {sbd} | Tổng: {tongThanhCong} ")
                    sys.stdout.flush()
                else:
                    retryCount += 1

            except Exception as e:
                retryCount += 1
                if 'closed' in str(e).lower():
                    await cauHinhGiaoDien(page)

        if not success:
            tongLoi += 1
            trangThaiDauSo[prefix]['countLoi'] += 1
            if trangThaiDauSo[prefix]['countLoi'] >= KHOANG_DUNG_KHI_HET_DIEM:
                if not trangThaiDauSo[prefix]['hoanThanh']:
                    trangThaiDauSo[prefix]['hoanThanh'] = True
                    print(f"\nHoàn tất đầu số {prefix}.")

    await context.close()

async def main():
    print("Lấy điểm thi tuyển sinh vào 10 tỉnh Hưng Yên)
    async with async_playwright() as p:
        # channel="chrome" ép hệ thống nhận diện Chrome gốc của bạn làm môi trường chạy chính thức
        browser = await p.chromium.launch(headless=True, channel="chrome")
        
        workers = []
        for i in range(SO_LUONG_TAB):
            workers.append(asyncio.create_task(chayWorker(browser, i + 1)))
            await asyncio.sleep(0.8)
            
        await asyncio.gather(*workers)
        await browser.close()
        ghiFileExcel()
        print("\nDone.")

if __name__ == "__main__":
    try:
        print("\033[H\033[J", end="")
        asyncio.run(main())
    except KeyboardInterrupt:
        isStopping = True
        print("\nStopping...")
        ghiFileExcel()
        sys.exit(0)