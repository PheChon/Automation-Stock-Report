import os
import pandas as pd
import numpy as np
from openpyxl.utils import get_column_letter  # [ส่วนที่เพิ่มมา] นำเข้าคำสั่งสำหรับแปลงตัวเลขเป็นตัวอักษรคอลัมน์ เช่น 1 -> A

# 1. กำหนดตำแหน่งโฟลเดอร์ตามโครงสร้างเครื่อง Mac
INPUT_DIR = "/Users/phachon/Documents/DKSH/auto-stock-report/input"
OUTPUT_DIR = "/Users/phachon/Documents/DKSH/auto-stock-report/output"

print("--- เริ่มต้นขั้นตอนที่ 1: โหลดไฟล์อัจฉริยะและสร้างคีย์อ้างอิง ---")

def smart_load_file(folder, base_name):
    """ ฟังก์ชันดึงข้อมูลอัจฉริยะ รองรับทั้ง CSV และ Excel โดยไม่สนใจนามสกุลที่ซ้อนกัน """
    for file_name in os.listdir(folder):
        if file_name.startswith(base_name) and not file_name.startswith("~$"):
            full_path = os.path.join(folder, file_name)
            if file_name.lower().endswith('.csv'):
                print(f"-> กำลังอ่านไฟล์ CSV: {file_name}")
                return pd.read_csv(full_path)
            elif file_name.lower().endswith(('.xlsx', '.xls')):
                print(f"-> กำลังอ่านไฟล์ Excel: {file_name}")
                return pd.read_excel(full_path)
    raise FileNotFoundError(f"ไม่พบไฟล์ที่ขึ้นต้นด้วย '{base_name}' ในโฟลเดอร์ {folder}")

# เรียกใช้งาน Smart Loader ดึงข้อมูลเข้าสู่ระบบ
mb52_th40 = smart_load_file(INPUT_DIR, "MB52_TH40")
mb52_th44 = smart_load_file(INPUT_DIR, "MB52_TH44")
r138_th40 = smart_load_file(INPUT_DIR, "R138_TH40")
r138_th44 = smart_load_file(INPUT_DIR, "R138_TH44")
product_group = smart_load_file(INPUT_DIR, "Product Group")

# สร้างคีย์อ้างอิงเชื่อมโยงข้อมูล (Link Key) สำหรับเก็บไว้ใช้งานและส่งออกภายนอก
mb52_th40['Link_Key'] = mb52_th40['Material'].astype(str) + mb52_th40['Unrestricted'].astype(str) + mb52_th40['Batch'].astype(str)
mb52_th44['Link_Key'] = mb52_th44['Material'].astype(str) + mb52_th44['Unrestricted'].astype(str) + mb52_th44['Batch'].astype(str)

r138_th40['Link_Key'] = r138_th40['Material No.'].astype(str) + r138_th40['Quantity'].astype(str) + r138_th40['Batch no.'].astype(str)
r138_th44['Link_Key'] = r138_th44['Material No.'].astype(str) + r138_th44['Quantity'].astype(str) + r138_th44['Batch no.'].astype(str)

print("โหลดข้อมูลและสร้างคีย์เสร็จสิ้น")


print("\n--- เริ่มต้นขั้นตอนที่ 2: รวมตารางและทำ Conditional Lookup (XLOOKUP) ---")

# รวมตาราง MB52
df_data = pd.concat([mb52_th40, mb52_th44], ignore_index=True)

# ลบคอลัมน์คำอธิบายเดิมออกเพื่อเตรียมแทรกแบบเป็นระบบ
if 'Material Description' in df_data.columns:
    df_data = df_data.drop(columns=['Material Description'])

# ดึงข้อมูลกลุ่มผลิตภัณฑ์ (Product Group)
df_pg_lookup = product_group[['Material', 'Product Group']].drop_duplicates()
df_data = pd.merge(df_data, df_pg_lookup, on='Material', how='left')

# เตรียม Data Mapping จาก R138 ทั้งสองโรงงาน
r130_th40_lookup = r138_th40[['Link_Key', 'Level 4 Product Group', 'Profit center', 'Last GR']].rename(
    columns={'Level 4 Product Group': 'Shipper', 'Last GR': 'GR Date'}
).drop_duplicates(subset=['Link_Key'])

r130_th44_lookup = r138_th44[['Link_Key', 'Level 4 Product Group', 'Profit center', 'Last GR']].rename(
    columns={'Level 4 Product Group': 'Shipper', 'Last GR': 'GR Date'}
).drop_duplicates(subset=['Link_Key'])

# คัดแยกสายข้อมูลเพื่อดึงค่าตามเงื่อนไขของคลังสินค้า (Plant)
df_th40 = df_data[df_data['Plant'] != 'TH44'].copy()
df_th44 = df_data[df_data['Plant'] == 'TH44'].copy()

df_th40 = pd.merge(df_th40, r130_th40_lookup, on='Link_Key', how='left')
df_th44 = pd.merge(df_th44, r130_th44_lookup, on='Link_Key', how='left')

# มัดรวมกลับมาเป็นตารางหลักตารางเดียว
df_data = pd.concat([df_th40, df_th44], ignore_index=True)

# จัดโครงสร้างลำดับคอลัมน์ของชีท DATA
columns_order = [
    'Plant', 'Storage location', 'Material', 'Unrestricted', 'Value Unrestricted', 
    'Material type', 'Material Group', 'Product Group', 'Shipper', 'Profit center', 
    'GR Date', 'Batch'
]
df_data = df_data[columns_order]
print("เชื่อมโยงความสัมพันธ์ข้อมูลเสร็จสิ้น")


print("\n--- เริ่มต้นขั้นตอนที่ 3: การคำนวณอายุสินทรัพย์ (Ageing) และการจัดกลุ่ม (Bucketing) ---")

# เปลี่ยนคอลัมน์ GR Date เป็นชนิดข้อมูลวันที่เพื่อใช้คำนวณคณิตศาสตร์
df_data['GR Date'] = pd.to_datetime(df_data['GR Date'], errors='coerce')

# อ้างอิงเวลาปัจจุบันของระบบ
current_date = pd.Timestamp.today().normalize()

# สูตรคำนวณอายุสินค้าคงคลัง: TODAY + 1 - GR Date
df_data['Ageing'] = (current_date + pd.Timedelta(days=1) - df_data['GR Date']).dt.days

# ทำการแปลงข้อมูลรูปแบบวันที่ให้เหลือเฉพาะ 'YYYY-MM-DD' ตัดส่วนของเวลา 00:00 ออกไปจากระบบผลลัพธ์
df_data['GR Date'] = df_data['GR Date'].dt.strftime('%Y-%m-%d')

# จำแนกกลุ่มตามช่วงอายุขอบเขต (0-30 ไปจนถึง >365)
bins = [-np.inf, 30, 90, 180, 365, np.inf]
labels = ["0-30", "31-90", "91-180", "181-365", ">365"]
df_data['Bucket'] = pd.cut(df_data['Ageing'], bins=bins, labels=labels)

# บันทึกลำดับคอลัมน์เวอร์ชันสมบูรณ์ของชีท DATA
final_columns_order = [
    'Plant', 'Storage location', 'Material', 'Unrestricted', 'Value Unrestricted', 
    'Material type', 'Material Group', 'Product Group', 'Shipper', 'Profit center', 
    'GR Date', 'Ageing', 'Bucket', 'Batch'
]
df_data = df_data[final_columns_order]
print("คำนวณและจำแนก Bucket สำเร็จ")


print("\n--- เริ่มต้นขั้นตอนที่ 4: การสรุปข้อมูลทุก Bucket และสร้างไฟล์ Excel ผลลัพธ์ ---")

# คำนวณมูลค่ารวมทั้งระบบเพื่อใช้คิดสัดส่วนร้อยละ (%)
grand_total_value = df_data['Value Unrestricted'].sum()

# หมุนมิติข้อมูลด้วย Pivot Table
pivot_df = pd.pivot_table(
    df_data,
    index='Shipper',
    columns='Bucket',
    values=['Unrestricted', 'Value Unrestricted'],
    aggfunc='sum',
    fill_value=0
)

summary_data = []
buckets_list = [">365", "181-365", "91-180", "31-90", "0-30"]
clients = df_data['Shipper'].dropna().unique()

for client in clients:
    row_data = {'Client': client}
    total_qty = 0
    total_val = 0
    
    for b in buckets_list:
        try:
            qty = pivot_df.loc[client, ('Unrestricted', b)]
            val = pivot_df.loc[client, ('Value Unrestricted', b)]
        except KeyError:
            qty = 0
            val = 0
            
        pct = (val / grand_total_value) * 100 if grand_total_value > 0 else 0
        
        row_data[f'Quantity {b}'] = qty
        row_data[f'Stock Value THB. {b}'] = val
        row_data[f'% {b}'] = pct
        
        total_qty += qty
        total_val += val
        
    row_data['Total Quantity'] = total_qty
    row_data['Total Stock Value THB.'] = total_val
    summary_data.append(row_data)

df_summary = pd.DataFrame(summary_data)

# บันทึกข้อมูลลงสู่ระบบดิสก์ของเครื่อง
os.makedirs(OUTPUT_DIR, exist_ok=True)
output_path = os.path.join(OUTPUT_DIR, "Result_Report.xlsx")

# ทำการบันทึก DataFrame ทั้งหมดลงในไฟล์เดี่ยวแยกแท็บแผ่นงานย่อยให้ครบถ้วน
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df_data.to_excel(writer, sheet_name='DATA', index=False)
    df_summary.to_excel(writer, sheet_name='Ageing > 365 D', index=False)
    
    # บันทึกไฟล์ต้นทาง 5 ชุด
    mb52_th40.to_excel(writer, sheet_name='MB52_TH40', index=False)
    mb52_th44.to_excel(writer, sheet_name='MB52_TH44', index=False)
    r138_th40.to_excel(writer, sheet_name='R138_TH40', index=False)
    r138_th44.to_excel(writer, sheet_name='R138_TH44', index=False)
    product_group.to_excel(writer, sheet_name='Product Group', index=False)

    # --- [ส่วนที่เพิ่มเข้ามาใหม่]: ปรับความกว้างคอลัมน์อัตโนมัติ (Auto-fit) ---
    print("กำลังจัดรูปแบบหน้าตา Excel และปรับขนาดความกว้างคอลัมน์อัตโนมัติ...")
    for sheet_name in writer.sheets:
        ws = writer.sheets[sheet_name]
        for col in ws.columns:
            max_len = 0
            # ดึงตัวอักษรของคอลัมน์ (เช่น A, B, C)
            col_letter = get_column_letter(col[0].column)
            
            # วนลูปหาข้อมูลที่มีความยาวมากที่สุดในแต่ละคอลัมน์ (รวมชื่อหัวตารางด้วย)
            for cell in col:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            
            # กำหนดขนาดความกว้าง = ความยาวที่มากที่สุด + 4 (บวกเผื่อพื้นที่ว่างให้อ่านง่าย)
            # ตั้งค่าความกว้างต่ำสุดไว้ที่ 12 เพื่อป้องกันไม่ให้ช่องที่มีข้อมูลสั้นแคบจนเกินไป
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

print(f"\n[สำเร็จ] ประมวลผลและจัดรูปแบบหน้าตาเสร็จสิ้นแบบ 100%!")
print(f"-> ไฟล์รายงานพร้อมใช้งานถูกสร้างขึ้นแล้วที่: {output_path}")