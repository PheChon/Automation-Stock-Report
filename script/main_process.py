import os
import pandas as pd
import numpy as np
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Border, Side

# ==========================================
# ⚠️ กำหนดตำแหน่ง Path โฟลเดอร์ของ Windows
# ==========================================
INPUT_DIR = r"C:\Users\thiraphop.chantra\Downloads\Auto-report\input"
OUTPUT_DIR = r"C:\Users\thiraphop.chantra\Downloads\Auto-report\output"

try:
    print("--- เริ่มต้นขั้นตอนที่ 1: โหลดไฟล์อัจฉริยะและสร้างคีย์อ้างอิง ---")

    def smart_load_file(folder, base_name):
        for file_name in os.listdir(folder):
            if file_name.startswith(base_name) and not file_name.startswith("~$"):
                full_path = os.path.join(folder, file_name)
                if file_name.lower().endswith('.csv'):
                    print(f"-> กำลังอ่านไฟล์ CSV: {file_name}")
                    return pd.read_csv(full_path, dtype=str)
                elif file_name.lower().endswith(('.xlsx', '.xls')):
                    print(f"-> กำลังอ่านไฟล์ Excel: {file_name}")
                    return pd.read_excel(full_path)
        raise FileNotFoundError(f"ไม่พบไฟล์ที่ขึ้นต้นด้วย '{base_name}' ในโฟลเดอร์ {folder}")

    mb52_th40 = smart_load_file(INPUT_DIR, "MB52_TH40")
    mb52_th44 = smart_load_file(INPUT_DIR, "MB52_TH44")
    r138_th40 = smart_load_file(INPUT_DIR, "R138_TH40")
    r138_th44 = smart_load_file(INPUT_DIR, "R138_TH44")
    product_group = smart_load_file(INPUT_DIR, "Product Group")

    for df in [mb52_th40, mb52_th44]:
        df['Unrestricted'] = pd.to_numeric(df['Unrestricted'], errors='coerce')
        df['Value Unrestricted'] = pd.to_numeric(df['Value Unrestricted'], errors='coerce')

    def clean_key_col(series):
        s = series.fillna('').astype(str).str.strip().str.upper()
        s = s.str.replace(r'\.0$', '', regex=True)
        s = s.replace('NAN', '')
        return s

    if 'Last GR' in r138_th40.columns:
        r138_th40['Last GR'] = pd.to_datetime(r138_th40['Last GR'], errors='coerce').dt.strftime('%Y-%m-%d')
    if 'Last GR' in r138_th44.columns:
        r138_th44['Last GR'] = pd.to_datetime(r138_th44['Last GR'], errors='coerce').dt.strftime('%Y-%m-%d')

    mb52_th40['Link_Key'] = clean_key_col(mb52_th40['Material']) + clean_key_col(mb52_th40['Unrestricted']) + clean_key_col(mb52_th40['Batch'])
    mb52_th44['Link_Key'] = clean_key_col(mb52_th44['Material']) + clean_key_col(mb52_th44['Unrestricted']) + clean_key_col(mb52_th44['Batch'])
    
    r138_th40['Link_Key'] = clean_key_col(r138_th40['Material No.']) + clean_key_col(r138_th40['Quantity']) + clean_key_col(r138_th40['Batch no.'])
    r138_th44['Link_Key'] = clean_key_col(r138_th44['Material No.']) + clean_key_col(r138_th44['Quantity']) + clean_key_col(r138_th44['Batch no.'])

    print("\n--- เริ่มต้นขั้นตอนที่ 2: รวมตารางและทำ Conditional Lookup ---")
    df_data = pd.concat([mb52_th40, mb52_th44], ignore_index=True)

    # กรองขยะและบรรทัด Grand Total จาก SAP ออก
    df_data = df_data.dropna(subset=['Material', 'Plant'])
    df_data = df_data[df_data['Plant'].isin(['TH40', 'TH44'])].copy()

    if 'Material Description' in df_data.columns:
        df_data = df_data.drop(columns=['Material Description'])

    df_data['Material_Key'] = clean_key_col(df_data['Material'])
    df_pg_lookup = product_group[['Material', 'Product Group']].copy()
    df_pg_lookup['Material_Key'] = clean_key_col(df_pg_lookup['Material'])
    df_pg_lookup = df_pg_lookup.drop_duplicates(subset=['Material_Key'])
    
    df_data = pd.merge(df_data, df_pg_lookup[['Material_Key', 'Product Group']], on='Material_Key', how='left')
    df_data = df_data.drop(columns=['Material_Key'])

    r130_th40_lookup = r138_th40[['Link_Key', 'Material Group Desc', 'Profit center', 'Last GR']].rename(
        columns={'Material Group Desc': 'Shipper', 'Last GR': 'GR Date'}
    ).drop_duplicates(subset=['Link_Key'])

    r130_th44_lookup = r138_th44[['Link_Key', 'Material Group Desc', 'Profit center', 'Last GR']].rename(
        columns={'Material Group Desc': 'Shipper', 'Last GR': 'GR Date'}
    ).drop_duplicates(subset=['Link_Key'])

    df_th40 = df_data[df_data['Plant'] == 'TH40'].copy()
    df_th44 = df_data[df_data['Plant'] == 'TH44'].copy()

    df_th40 = pd.merge(df_th40, r130_th40_lookup, on='Link_Key', how='left')
    df_th44 = pd.merge(df_th44, r130_th44_lookup, on='Link_Key', how='left')

    df_data = pd.concat([df_th40, df_th44], ignore_index=True)

    # เติมค่าว่างด้วย Unassigned
    df_data['Shipper'] = df_data['Shipper'].fillna('Unassigned Shipper')
    df_data['Product Group'] = df_data['Product Group'].fillna('Unassigned Group')

    columns_order = [
        'Plant', 'Storage location', 'Material', 'Unrestricted', 'Value Unrestricted',
        'Material type', 'Material Group', 'Product Group', 'Shipper', 'Profit center',
        'GR Date', 'Batch'
    ]
    df_data = df_data[columns_order]

    print("\n--- เริ่มต้นขั้นตอนที่ 3: การคำนวณอายุสินทรัพย์ ---")
    df_data['GR Date'] = pd.to_datetime(df_data['GR Date'], errors='coerce')
    current_date = pd.Timestamp.today().normalize()
    df_data['Ageing'] = (current_date + pd.Timedelta(days=1) - df_data['GR Date']).dt.days
    
    df_data['GR Date'] = df_data['GR Date'].dt.strftime('%Y-%m-%d')

    bins = [-np.inf, 30, 90, 180, 365, np.inf]
    labels = ["0-30", "31-90", "91-180", "181-365", ">365"]
    df_data['Bucket'] = pd.cut(df_data['Ageing'], bins=bins, labels=labels)

    final_columns_order = [
        'Plant', 'Storage location', 'Material', 'Unrestricted', 'Value Unrestricted',
        'Material type', 'Material Group', 'Product Group', 'Shipper', 'Profit center',
        'GR Date', 'Ageing', 'Bucket', 'Batch'
    ]
    df_data = df_data[final_columns_order]

    print("\n--- เริ่มต้นขั้นตอนที่ 4: การสรุปข้อมูลแนวกว้าง (Ageing > 365 D) ---")
    grand_total_value = df_data['Value Unrestricted'].sum()
    
    # [แก้ไข] คำนวณยอดรวมของแต่ละช่วงอายุ (Bucket Totals) เพื่อดึงไปใช้เป็นฐาน %
    bucket_totals = df_data.groupby('Bucket', observed=False)['Value Unrestricted'].sum().to_dict()

    pivot_df = pd.pivot_table(
        df_data, index='Shipper', columns='Bucket',
        values=['Unrestricted', 'Value Unrestricted'], aggfunc='sum', fill_value=0
    )

    summary_data = []
    buckets_list = [">365", "181-365", "91-180", "31-90", "0-30"]
    clients = df_data['Shipper'].dropna().unique()

    for client in clients:
        row_data = {'Client': client}
        total_qty = total_val = 0
        for b in buckets_list:
            try:
                qty = pivot_df.loc[client, ('Unrestricted', b)]
                val = pivot_df.loc[client, ('Value Unrestricted', b)]
            except KeyError:
                qty = val = 0
            
            # [แก้ไข] ใช้ยอดรวมของ Bucket ตัวเองเป็นตัวหารในการหา %
            b_total = bucket_totals.get(b, 0)
            pct = round((val / b_total) * 100, 2) if b_total > 0 else 0
            
            row_data[f'Quantity {b}'] = qty
            row_data[f'Stock Value THB. {b}'] = val
            row_data[f'% {b}'] = pct
            total_qty += qty
            total_val += val
            
        row_data['Total Quantity'] = total_qty
        row_data['Total Stock Value THB.'] = total_val
        summary_data.append(row_data)
    df_summary = pd.DataFrame(summary_data)

    print("\n--- เริ่มต้นขั้นตอนที่ 5: การสร้าง Dashboard สรุปผล 5 ตาราง (PV DATA) ---")
    
    # [แก้ไข] เพิ่ม Parameter denominator เพื่อรับฐานตัวเลขที่จะเอามาหาร
    def make_client_table(df, denominator):
        if df.empty:
            return pd.DataFrame({'Client': ['Grand Total'], 'Quantity': [0], 'Stock Value THB.': [0], '%': [0]})
        grouped = df.groupby('Shipper')[['Unrestricted', 'Value Unrestricted']].sum().reset_index()
        grouped.rename(columns={'Shipper': 'Client', 'Unrestricted': 'Quantity', 'Value Unrestricted': 'Stock Value THB.'}, inplace=True)
        grouped['%'] = (grouped['Stock Value THB.'] / denominator * 100).round(2) if denominator else 0
        grouped = grouped.sort_values(by='Stock Value THB.', ascending=False)
        
        gt_row = pd.DataFrame({
            'Client': ['Grand Total'], 'Quantity': [grouped['Quantity'].sum()],
            'Stock Value THB.': [grouped['Stock Value THB.'].sum()], '%': [grouped['%'].sum().round(2)]
        })
        return pd.concat([grouped, gt_row], ignore_index=True)

    # ตารางหลัก หารด้วยยอด Grand Total ปกติ
    df_inv_all = make_client_table(df_data, grand_total_value)
    
    # ตาราง > 365 หารด้วย ยอดรวมของเฉพาะกลุ่ม > 365
    total_365_val = bucket_totals.get('>365', 0)
    df_inv_365 = make_client_table(df_data[df_data['Bucket'] == '>365'], total_365_val)

    def make_group_report(df, group_col):
        rows = []
        for p in ['TH40', 'TH44']:
            p_data = df[df['Plant'] == p]
            if p_data.empty: continue
            
            p_qty, p_val = p_data['Unrestricted'].sum(), p_data['Value Unrestricted'].sum()
            p_pct = round((p_val / grand_total_value) * 100, 2) if grand_total_value else 0
            # แถวบรรทัดยอดรวมของแต่ละ Plant
            rows.append({'Category': p, 'Quantity': p_qty, 'Stock Value THB.': p_val, '%': p_pct})
            
            sub_groups = ["0-30", "31-90", "91-180", "181-365", ">365"] if group_col == 'Bucket' else sorted(p_data[group_col].dropna().unique())
            for sg in sub_groups:
                sg_data = p_data[p_data[group_col] == sg]
                if sg_data.empty: continue
                sg_qty, sg_val = sg_data['Unrestricted'].sum(), sg_data['Value Unrestricted'].sum()
                
                # [แก้ไข] แถวบรรทัดซอยย่อย ถ้าเป็นตาราง Ageing ให้หารด้วยยอดรวมของ Ageing นั้นๆ
                if group_col == 'Bucket':
                    denominator = bucket_totals.get(sg, 0)
                else:
                    denominator = grand_total_value
                    
                sg_pct = round((sg_val / denominator) * 100, 2) if denominator > 0 else 0
                rows.append({'Category': f"   {sg}", 'Quantity': sg_qty, 'Stock Value THB.': sg_val, '%': sg_pct})
                
        # แถวสรุปสุดท้าย
        rows.append({
            'Category': 'Grand Total', 'Quantity': df['Unrestricted'].sum(),
            'Stock Value THB.': df['Value Unrestricted'].sum(), 
            '%': round((df['Value Unrestricted'].sum() / grand_total_value) * 100, 2) if grand_total_value else 0
        })
        return pd.DataFrame(rows)

    df_pg_report = make_group_report(df_data, 'Product Group').rename(columns={'Category': 'Plant / Product Group'})
    df_ageing_report = make_group_report(df_data, 'Bucket').rename(columns={'Category': 'Plant / Ageing'})

    def make_plant_table(df):
        grouped = df.groupby('Plant')[['Unrestricted', 'Value Unrestricted']].sum().reset_index()
        grouped.rename(columns={'Unrestricted': 'Quantity', 'Value Unrestricted': 'Stock Value THB.'}, inplace=True)
        grouped['%'] = (grouped['Stock Value THB.'] / grand_total_value * 100).round(2) if grand_total_value else 0
        gt_row = pd.DataFrame({'Plant': ['Grand Total'], 'Quantity': [grouped['Quantity'].sum()],
                               'Stock Value THB.': [grouped['Stock Value THB.'].sum()], '%': [grouped['%'].sum().round(2)]})
        return pd.concat([grouped, gt_row], ignore_index=True)

    df_plant_report = make_plant_table(df_data)

    dash_layouts = [
        (df_plant_report, "Plant Report", 1, 1),
        (df_pg_report, "Product Group Report", 1 + len(df_plant_report) + 3, 1),
        (df_ageing_report, "Ageing Report", 1 + len(df_plant_report) + 3 + len(df_pg_report) + 3, 1),
        (df_inv_all, "Inventory Clients (THB)", 1, 6),
        (df_inv_365, "Inventory Clients (THB) (>365)", 1, 11)
    ]

    print("\n--- กำลังบันทึกและจัดรูปแบบเอกสารขั้นสุดท้าย ---")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    output_path = os.path.join(OUTPUT_DIR, "Result_Report.xlsx")

    header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    header_font = Font(bold=True)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        dash_sheet_name = 'PV DATA (Dashboard)'
        pd.DataFrame().to_excel(writer, sheet_name=dash_sheet_name, index=False)
        
        for t_df, title, r, c in dash_layouts:
            t_df.to_excel(writer, sheet_name=dash_sheet_name, startrow=r, startcol=c, index=False)
            ws = writer.sheets[dash_sheet_name]
            
            cell = ws.cell(row=r, column=c+1, value=title)
            cell.font = Font(bold=True, color="000080", size=12)
            
            for col_num in range(c + 1, c + 1 + len(t_df.columns)):
                header_cell = ws.cell(row=r+1, column=col_num)
                header_cell.fill = header_fill
                header_cell.font = header_font
                header_cell.border = thin_border
                
            for row_num in range(r + 2, r + 2 + len(t_df)):
                for col_num in range(c + 1, c + 1 + len(t_df.columns)):
                    ws.cell(row=row_num, column=col_num).border = thin_border

        df_data.to_excel(writer, sheet_name='DATA', index=False)
        df_summary.to_excel(writer, sheet_name='Ageing > 365 D', index=False)
        
        mb52_th40.to_excel(writer, sheet_name='MB52_TH40', index=False)
        mb52_th44.to_excel(writer, sheet_name='MB52_TH44', index=False)
        r138_th40.to_excel(writer, sheet_name='R138_TH40', index=False)
        r138_th44.to_excel(writer, sheet_name='R138_TH44', index=False)
        product_group.to_excel(writer, sheet_name='Product Group', index=False)

        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.value is not None:
                        max_len = max(max_len, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
                
            if sheet_name != dash_sheet_name:
                for col_num in range(1, ws.max_column + 1):
                    header_cell = ws.cell(row=1, column=col_num)
                    header_cell.fill = header_fill
                    header_cell.font = header_font
                    header_cell.border = thin_border
                
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                    for cell in row:
                        cell.border = thin_border

    print(f"\n[สำเร็จ] ประมวลผลและสร้าง Dashboard เสร็จสิ้นแบบ 100%!")
    print(f"-> ไฟล์รายงานพร้อมใช้งานถูกสร้างขึ้นแล้วที่: {output_path}")

except Exception as e:
    print("\n" + "="*60)
    print("❌ ระบบพบข้อผิดพลาด (ERROR) ไม่สามารถประมวลผลต่อได้ ❌")
    print("รายละเอียด:")
    print(e)
    print("="*60 + "\n")
finally:
    input("\nกด Enter เพื่อปิดหน้าต่างนี้...")