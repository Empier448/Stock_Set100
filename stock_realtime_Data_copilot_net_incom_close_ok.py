from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from pathlib import Path
import yfinance as yf
from openpyxl import Workbook
from datetime import datetime
from io import StringIO

# เรียกใช้ ChromeDriver โดยให้ WebDriver Manager จัดการเวอร์ชันให้อัตโนมัติ
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))


# ไปยังหน้าหลักของ SET
driver.get('http://siamchart.com/stock')
# ดึงข้อมูลหน้าเว็บ
data = driver.page_source
# อ่านข้อมูลตารางจากหน้าเว็บ
data_df = pd.read_html(data)[2]
# ทำความสะอาดชื่อคอลัมน์
data_df.columns = [c.replace(' (Click to sort Ascending)', '') for c in data_df.columns]
# ลบแถวแรกที่ไม่จำเป็น
data_df.drop([0], inplace=True)
# ตั้งค่า index
data_df.set_index('Name', inplace=True)

# ตรวจสอบว่าไดเรกทอรีที่ต้องการบันทึกไฟล์มีอยู่จริง
output_dir = Path(r'C:/Users/plaifa/Downloads/python/data')
output_dir.mkdir(parents=True, exist_ok=True)

# ฟังก์ชันสำหรับดึงข้อมูลหุ้น
def get_stock_data(symbol):
    symbol += '.BK'  # เพิ่มส่วนขยาย .BK หลังรายชื่อหุ้น

    try:

        stock = yf.Ticker(symbol)
        
        # ดึงข้อมูลราคาปิดล่าสุดจาก history()
        try:
            stock_history = stock.history(period="1d")
            if stock_history.empty:
                stock_history = stock.history(period="1mo")
            close = stock_history['Close'].iloc[-1] if not stock_history.empty else 'N/A'
        except Exception as e:
            close = 'N/A'
            print(f"ข้อผิดพลาดในการดึงข้อมูล Close: {e}")

        # ดึงข้อมูลอื่น ๆ จาก stock.info
 
        stock_info = yf.Ticker(symbol).info
        open = stock_info.get('open', 'N/A')
        high = stock_info.get('dayHigh', 'N/A')
        low = stock_info.get('dayLow', 'N/A')
        #close = stock_info.get('regularMarketPreviousClose', 'N/A')
        price = stock_info.get('currentPrice', 'N/A')  # Adj Close = Price
        volume = stock_info.get('volume', 'N/A')
        eps = stock_info.get('trailingEps', 'N/A')
        pe = stock_info.get('trailingPE', 'N/A')
        roa = stock_info.get('returnOnAssets', 'N/A')
        roe = stock_info.get('returnOnEquity', 'N/A')
        
        dividend_yield = stock_info.get('dividendYield', 0) * 100 if stock_info.get('dividendYield') is not None else 'N/A'
        shares_outstanding = stock_info.get('sharesOutstanding', 'N/A')
        book_value = stock_info.get('bookValue', 'N/A')
        dividend_rate = stock_info.get('dividendRate', 'N/A')
        revenue = stock_info.get('totalRevenue', 'N/A')
        net_income = stock_info.get('netIncomeToCommon', 'N/A')
        assets = stock_info.get('totalAssets', 'N/A')

        # คำนวณ Dividend Yield
        div_yield = (float(dividend_rate) / float(price)) * 100 if dividend_rate != 'N/A' and price != 'N/A' else 'N/A'
        # คำนวณ BVPS
        bvps = float(book_value) / float(shares_outstanding) if shares_outstanding != 'N/A' and book_value != 'N/A' else 'N/A'

        # คำนวณ P/BV
        pbv = float(close) / bvps if bvps != 'N/A' else 'N/A'#(Price-to-Book Value):ที่นี่, BVPS (Book Value Per Share) คำนวณโดยการหาร Book Value (มูลค่าตามบัญชี) ด้วยจำนวนหุ้นที่ออกจำหน่าย (Shares Outstanding).
        #ใช้ค่าของ close (ราคาปิด) และ BVPS (Book Value Per Share) ที่คำนวณจากข้อมูล book_value และ shares_outstanding.
        pbv2 = close / book_value if book_value != 'N/A' else 'N/A'# (Price-to-Book Value, Alternative Calculation)

        # เปลี่ยนลำดับข้อมูลโดยให้ Ticker มาก่อน Date
        return [symbol.replace('.BK', ''), current_date, open, high, low, close, volume, eps, pbv2, dividend_yield, pe, roa, roe, shares_outstanding, book_value, dividend_rate, div_yield, bvps, pbv, revenue, net_income, assets]
    except Exception as e:
        print(f"ไม่สามารถดึงข้อมูลสำหรับ {symbol} ได้: {e}")
        return [symbol.replace('.BK', '')] + ['N/A'] * 22

# สร้าง workbook ใหม่
wb = Workbook()
ws = wb.active
ws.append(["Ticker", "Date", "Open", "High", "Low", "Close", "Volume", "EPS","P/BV","Yield","P/E", "ROA", "ROE",  "ShareOut", "BookValue", "dividend_rate", "div_yield", "BVPS", "P/BV2",  "Revenue", "Net Income", "Assets"])

# รับวันที่ปัจจุบัน
current_date = datetime.now().strftime('%Y-%m-%d')

# รวบรวมข้อมูลจากทุกหุ้นและบันทึกลงใน Excel
for stock in data_df.index:
    stock_data = get_stock_data(stock)
    ws.append(stock_data)

# ตั้งชื่อไฟล์พร้อมวันที่สำหรับไฟล์ Excel
excel_file_name = f'history_EOD_{current_date}.xlsx'

# บันทึก workbook ลงในไฟล์ Excel
wb.save(output_dir / excel_file_name)

# โหลดไฟล์ Excel ที่บันทึกไว้
df = pd.read_excel(output_dir / excel_file_name)

# ตรวจสอบและคำนวณ Revenue Growth (YoY) และ Net Income Growth (YoY) ถ้ามีข้อมูลเพียงพอ
if df.shape[0] > 1:
    df['Revenue'] = pd.to_numeric(df['Revenue'], errors='coerce')
    df['Net Income'] = pd.to_numeric(df['Net Income'], errors='coerce')
    df['Revenue Growth (YoY)'] = df['Revenue'].pct_change() * 100
    df['Net Income Growth (YoY)'] = df['Net Income'].pct_change() * 100

# ตั้งชื่อไฟล์พร้อมวันที่สำหรับไฟล์ CSV
csv_file_name = f'stocks_data_{current_date}.csv'
# บันทึกข้อมูลลงในไฟล์ CSV
df.to_csv(output_dir / csv_file_name, index=False)
df.to_csv(output_dir / 'stocks_data.csv', index=False)

print(f"ข้อมูลได้ถูกบันทึกไว้ในไฟล์ {csv_file_name}")
