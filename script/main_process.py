import os
import pandas as pd
import numpy as np
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# ==========================================
# ⚠️ กำหนดตำแหน่ง Path โฟลเดอร์ของ Windows
# ==========================================
INPUT_DIR = r"/Users/phachon/Documents/DKSH/auto-stock-report/input"
OUTPUT_DIR = r"/Users/phachon/Documents/DKSH/auto-stock-report/output"

try:
    print("--- เริ่มต้นขั้นตอนที่ 1: โหลดไฟล์อัจฉริยะและทำความสะอาดข้อมูล ---")

    def smart_load_file(folder, base_name):
        for file_name in os.listdir(folder):
            if file_name.startswith(base_name) and not file_name.startswith("~$"):
                full_path = os.path.join(folder, file_name)
                if file_name.lower().endswith('.csv'):
                    print(f"-> กำลังอ่านไฟล์ CSV: {file_name}")
                    return pd.read_csv(full_path, dtype=str)
                elif file_name.lower().endswith(('.xlsx', '.xls')):
                    print(f"-> กำลังอ่านไฟล์ Excel: {file_name}")
                    return pd.read_excel(full_path, dtype=str)
        raise FileNotFoundError(f"ไม่พบไฟล์ที่ขึ้นต้นด้วย '{base_name}' ในโฟลเดอร์ {folder}")

    mb52_th40 = smart_load_file(INPUT_DIR, "MB52_TH40")
    mb52_th44 = smart_load_file(INPUT_DIR, "MB52_TH44")
    r138_th40 = smart_load_file(INPUT_DIR, "R138_TH40")
    r138_th44 = smart_load_file(INPUT_DIR, "R138_TH44")
    product_group = smart_load_file(INPUT_DIR, "Product Group")

    for df in [mb52_th40, mb52_th44]:
        df['Unrestricted'] = pd.to_numeric(df['Unrestricted'].astype(str).str.replace(',', ''), errors='coerce')
        df['Value Unrestricted'] = pd.to_numeric(df['Value Unrestricted'].astype(str).str.replace(',', ''), errors='coerce')

    mb52_th40 = mb52_th40[(mb52_th40['Unrestricted'] != 0) & (mb52_th40['Unrestricted'].notna())].copy()
    mb52_th44 = mb52_th44[(mb52_th44['Unrestricted'] != 0) & (mb52_th44['Unrestricted'].notna())].copy()

    def clean_col(series):
        s = series.fillna('').astype(str).str.strip().str.upper()
        s = s.str.replace(',', '', regex=False)
        s = s.str.replace(r'\.0$', '', regex=True)
        s = s.replace('NAN', '')
        return s

    r138_all = pd.concat([r138_th40, r138_th44], ignore_index=True)
    r138_all['Last GR'] = pd.to_datetime(r138_all['Last GR'], errors='coerce')
    r138_all = r138_all.sort_values('Last GR', na_position='last') # เรียงเอาวันที่เก่าสุดขึ้นก่อน

    print("\n--- เริ่มต้นขั้นตอนที่ 2: จำลองระบบ VLOOKUP เสมือนมนุษย์ (Multi-Level Mapping) ---")
    
    # ---------------------------------------------------------
    # สร้าง Dictionaries สำหรับการแมป (Mapping) แบบหลายชั้น
    # ---------------------------------------------------------
    # 1. Dict สำหรับ GR Date (4 ชั้น)
    dict_gr_full = dict(zip(clean_col(r138_all['Material No.']) + "_" + clean_col(r138_all['Quantity']) + "_" + clean_col(r138_all['Batch no.']), r138_all['Last GR']))
    dict_gr_mb = dict(zip(clean_col(r138_all['Material No.']) + "_" + clean_col(r138_all['Batch no.']), r138_all['Last GR']))
    dict_gr_m = dict(zip(clean_col(r138_all['Material No.']), r138_all['Last GR']))

    # 2. Dict สำหรับ Product Group (3 ชั้น)
    dict_pg_1 = dict(zip(clean_col(product_group['Material']), product_group['Product Group']))
    dict_pg_2 = dict(zip(clean_col(r138_all['Material No.']), r138_all['Level 4 Product Group']))
    dict_pg_3 = dict(zip(clean_col(r138_all['Material Group']), r138_all['Level 4 Product Group']))

    # 3. Dict สำหรับ Shipper
    dict_shipper = dict(zip(clean_col(r138_all['Material Group']), r138_all['Material Group Desc']))

    # 4. Dict สำหรับ Profit Center
    dict_pc = dict(zip(clean_col(r138_all['Plant']) + "_" + clean_col(r138_all['Material No.']), r138_all['Profit center']))

    # ---------------------------------------------------------
    # นำตาราง MB52 มารวมกันและยิง VLOOKUP แบบลึกล้ำ
    # ---------------------------------------------------------
    df_data = pd.concat([mb52_th40, mb52_th44], ignore_index=True)
    df_data = df_data.dropna(subset=['Material', 'Plant'])
    df_data = df_data[df_data['Plant'].isin(['TH40', 'TH44'])].copy()

    k_full = clean_col(df_data['Material']) + "_" + clean_col(df_data['Unrestricted']) + "_" + clean_col(df_data['Batch'])
    k_mb = clean_col(df_data['Material']) + "_" + clean_col(df_data['Batch'])
    k_m = clean_col(df_data['Material'])
    k_grp = clean_col(df_data['Material Group'])
    k_plant_mat = clean_col(df_data['Plant']) + "_" + k_m

    # VLOOKUP: GR Date (4 ชั้น ตามที่วินิจฉัย)
    df_data['GR Date'] = k_full.map(dict_gr_full)
    df_data['GR Date'] = df_data['GR Date'].fillna(k_mb.map(dict_gr_mb))
    df_data['GR Date'] = df_data['GR Date'].fillna(k_m.map(dict_gr_m))
    df_data['GR Date'] = df_data['GR Date'].fillna(pd.to_datetime(df_data['Batch'], format='%Y%m%d', errors='coerce')) # ชั้นที่ 4: ดึงจากชื่อ Batch

    # VLOOKUP: Product Group (3 ชั้น ตามที่วินิจฉัย)
    df_data['Product Group'] = k_m.map(dict_pg_1)
    df_data['Product Group'] = df_data['Product Group'].fillna(k_m.map(dict_pg_2))
    df_data['Product Group'] = df_data['Product Group'].fillna(k_grp.map(dict_pg_3))
    df_data['Product Group'] = df_data['Product Group'].fillna('Unassigned Group') # เผื่อเหนียว

    # VLOOKUP: Shipper & Profit Center
    df_data['Shipper'] = k_grp.map(dict_shipper).fillna('Unassigned Shipper')
    df_data['Profit center'] = k_plant_mat.map(dict_pc)

    columns_order = [
        'Plant', 'Storage location', 'Material', 'Unrestricted', 'Value Unrestricted',
        'Material type', 'Material Group', 'Product Group', 'Shipper', 'Profit center',
        'GR Date', 'Batch'
    ]
    df_data = df_data[columns_order]

    print("\n--- เริ่มต้นขั้นตอนที่ 3: การคำนวณอายุสินทรัพย์และการจัดกลุ่ม ---")
    current_date = pd.Timestamp.today().normalize()
    # คำนวณ Ageing (หากไม่มี GR Date ท้ายที่สุดจริงๆ จะให้เป็น 0 เพื่อจัดกลุ่ม 0-30 ป้องกัน Error)
    df_data['GR Date'] = df_data['GR Date'].fillna(current_date)
    
    df_data['Ageing'] = (current_date + pd.Timedelta(days=1) - df_data['GR Date']).dt.days
    df_data['GR Date'] = df_data['GR Date'].dt.strftime('%Y-%m-%d')

    bins = [-np.inf, 30, 90, 180, 365, np.inf]
    labels = ["0-30", "31-90", "91-180", "181-365", ">365"]
    df_data['Bucket'] = pd.cut(df_data['Ageing'], bins=bins, labels=labels).astype(str)

    final_columns_order = [
        'Plant', 'Storage location', 'Material', 'Unrestricted', 'Value Unrestricted',
        'Material type', 'Material Group', 'Product Group', 'Shipper', 'Profit center',
        'GR Date', 'Ageing', 'Bucket', 'Batch'
    ]
    df_data = df_data[final_columns_order]

    print("\n--- เริ่มต้นขั้นตอนที่ 4: การสรุปข้อมูลแนวกว้าง (Ageing > 365 D) ---")
    grand_total_value = df_data['Value Unrestricted'].sum()
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
            
            b_total = bucket_totals.get(b, 0)
            pct = (val / b_total) * 100 if b_total > 0 else 0
            
            row_data[f'Quantity {b}'] = qty
            row_data[f'Stock Value THB. {b}'] = val
            row_data[f'% {b}'] = f"{pct:.2f}%" 
            
            total_qty += qty
            total_val += val
            
        row_data['Total Quantity'] = total_qty
        row_data['Total Stock Value THB.'] = total_val
        summary_data.append(row_data)
    df_summary = pd.DataFrame(summary_data)

    print("\n--- เริ่มต้นขั้นตอนที่ 5: การสร้าง Dashboard สรุปผล 5 ตาราง (PV DATA) ---")
    
    def make_client_table(df, denominator):
        if df.empty:
            return pd.DataFrame({'No.': [''], 'Client': ['Grand Total'], 'Quantity': [0], 'Stock Value THB.': [0], '%': ['0.00%']})
        grouped = df.groupby('Shipper')[['Unrestricted', 'Value Unrestricted']].sum().reset_index()
        grouped.rename(columns={'Shipper': 'Client', 'Unrestricted': 'Quantity', 'Value Unrestricted': 'Stock Value THB.'}, inplace=True)
        
        grouped['%_num'] = (grouped['Stock Value THB.'] / denominator * 100) if denominator else 0
        grouped = grouped.sort_values(by='Stock Value THB.', ascending=False).reset_index(drop=True)
        grouped.insert(0, 'No.', range(1, len(grouped) + 1))
        
        total_val = grouped['Stock Value THB.'].sum()
        sum_pct = (total_val / denominator * 100) if denominator else 0
        
        grouped['%'] = grouped['%_num'].apply(lambda x: f"{x:.2f}%")
        grouped.drop(columns=['%_num'], inplace=True)
        
        gt_row = pd.DataFrame({
            'No.': [''],
            'Client': ['Grand Total'], 
            'Quantity': [grouped['Quantity'].sum()],
            'Stock Value THB.': [total_val], 
            '%': [f"{sum_pct:.2f}%"]
        })
        return pd.concat([grouped, gt_row], ignore_index=True)

    df_inv_all = make_client_table(df_data, grand_total_value)
    total_365_val = bucket_totals.get('>365', 0)
    df_inv_365 = make_client_table(df_data[df_data['Bucket'] == '>365'], total_365_val)

    def make_group_report(df, group_col):
        rows = []
        for p in ['TH40', 'TH44']:
            p_data = df[df['Plant'] == p]
            if p_data.empty: continue
            
            p_qty, p_val = p_data['Unrestricted'].sum(), p_data['Value Unrestricted'].sum()
            p_pct = (p_val / grand_total_value) * 100 if grand_total_value else 0
            rows.append({'Category': p, 'Quantity': p_qty, 'Stock Value THB.': p_val, '%': f"{p_pct:.2f}%"})
            
            sub_groups = ["0-30", "31-90", "91-180", "181-365", ">365"] if group_col == 'Bucket' else sorted(p_data[group_col].dropna().unique())
            for sg in sub_groups:
                sg_data = p_data[p_data[group_col] == sg]
                if sg_data.empty: continue
                sg_qty, sg_val = sg_data['Unrestricted'].sum(), sg_data['Value Unrestricted'].sum()
                
                if group_col == 'Bucket':
                    denominator = bucket_totals.get(sg, 0)
                else:
                    denominator = grand_total_value
                    
                sg_pct = (sg_val / denominator) * 100 if denominator > 0 else 0
                rows.append({'Category': f"   {sg}", 'Quantity': sg_qty, 'Stock Value THB.': sg_val, '%': f"{sg_pct:.2f}%"})
                
        gt_pct = (df['Value Unrestricted'].sum() / grand_total_value) * 100 if grand_total_value else 0
        rows.append({
            'Category': 'Grand Total', 'Quantity': df['Unrestricted'].sum(),
            'Stock Value THB.': df['Value Unrestricted'].sum(), 
            '%': f"{gt_pct:.2f}%"
        })
        return pd.DataFrame(rows)

    df_pg_report = make_group_report(df_data, 'Product Group').rename(columns={'Category': 'Plant / Product Group'})
    df_ageing_report = make_group_report(df_data, 'Bucket').rename(columns={'Category': 'Plant / Ageing'})

    def make_plant_table(df):
        grouped = df.groupby('Plant')[['Unrestricted', 'Value Unrestricted']].sum().reset_index()
        grouped.rename(columns={'Unrestricted': 'Quantity', 'Value Unrestricted': 'Stock Value THB.'}, inplace=True)
        
        grouped['%_num'] = (grouped['Stock Value THB.'] / grand_total_value * 100) if grand_total_value else 0
        total_val = grouped['Stock Value THB.'].sum()
        sum_pct = (total_val / grand_total_value * 100) if grand_total_value else 0
        
        grouped['%'] = grouped['%_num'].apply(lambda x: f"{x:.2f}%")
        grouped.drop(columns=['%_num'], inplace=True)
        
        gt_row = pd.DataFrame({'Plant': ['Grand Total'], 'Quantity': [grouped['Quantity'].sum()],
                               'Stock Value THB.': [total_val], '%': [f"{sum_pct:.2f}%"]})
        return pd.concat([grouped, gt_row], ignore_index=True)

    df_plant_report = make_plant_table(df_data)

    dash_layouts = [
        (df_plant_report, "Plant Report", 5, 1),
        (df_pg_report, "Product Group Report", 5 + len(df_plant_report) + 3, 1),
        (df_ageing_report, "Ageing Report", 5 + len(df_plant_report) + 3 + len(df_pg_report) + 3, 1),
        (df_inv_all, "Inventory Clients (THB)", 5, 6),
        (df_inv_365, "Inventory Clients (THB) (>365)", 5, 12)
    ]

    print("\n--- กำลังบันทึกและจัดรูปแบบเอกสารขั้นสุดท้าย (ปรับฟอนต์และตัวเลข) ---")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    output_path = os.path.join(OUTPUT_DIR, "Result_Report.xlsx")

    FONT_NAME = "Segoe UI"
    
    header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid") 
    top10_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") 
    
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    header_font = Font(name=FONT_NAME, bold=True)
    total_font = Font(name=FONT_NAME, bold=True, color="000000")
    data_font = Font(name=FONT_NAME)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        dash_sheet_name = 'PV DATA (Dashboard)'
        pd.DataFrame().to_excel(writer, sheet_name=dash_sheet_name, index=False)
        ws = writer.sheets[dash_sheet_name]
        
        pct_365 = round((total_365_val / grand_total_value * 100), 2) if grand_total_value else 0
        top_client_365 = df_inv_365.iloc[0]['Client'] if len(df_inv_365) > 1 else "N/A"
        
        summary_text = (
            f"Executive Summary: Total Inventory Value is {grand_total_value:,.2f} THB. "
            f"Dead Stock (>365 Days) accounts for {total_365_val:,.2f} THB ({pct_365}% of total inventory). "
            f"The primary client driver for dead stock is '{top_client_365}'."
        )
        ws.merge_cells("B2:P4")
        summary_cell = ws["B2"]
        summary_cell.value = summary_text
        summary_cell.font = Font(name=FONT_NAME, bold=True, color="002060", size=11)
        summary_cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        summary_cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
        
        for t_df, title, r, c in dash_layouts:
            t_df.to_excel(writer, sheet_name=dash_sheet_name, startrow=r, startcol=c, index=False)
            
            cell = ws.cell(row=r, column=c+1, value=title)
            cell.font = Font(name=FONT_NAME, bold=True, color="000080", size=12)
            
            for col_num in range(c + 1, c + 1 + len(t_df.columns)):
                header_cell = ws.cell(row=r+1, column=col_num)
                header_cell.fill = header_fill
                header_cell.font = header_font
                header_cell.border = thin_border
                
            for row_num in range(r + 2, r + 2 + len(t_df)):
                val_col1 = str(ws.cell(row=row_num, column=c+1).value)
                val_col2 = str(ws.cell(row=row_num, column=c+2).value)
                is_total_row = (val_col1 in ["Grand Total", "TH40", "TH44"]) or (val_col2 in ["Grand Total"])
                
                is_client_table = "Clients" in title
                is_top_10 = is_client_table and (row_num < r + 2 + 10) and not is_total_row

                for col_num in range(c + 1, c + 1 + len(t_df.columns)):
                    data_cell = ws.cell(row=row_num, column=col_num)
                    data_cell.border = thin_border
                    data_cell.font = data_font
                    
                    if is_total_row:
                        data_cell.fill = total_fill
                        data_cell.font = total_font
                    elif is_top_10:
                        data_cell.fill = top10_fill
                        
                    header_val = str(ws.cell(row=r+1, column=col_num).value)
                    if isinstance(data_cell.value, (int, float)):
                        if "Stock Value" in header_val or "Value" in header_val:
                            data_cell.number_format = '#,##0.00'
                        elif "Quantity" in header_val or "Unrestricted" in header_val:
                            data_cell.number_format = '#,##0'

        df_data.to_excel(writer, sheet_name='DATA', index=False)
        df_summary.to_excel(writer, sheet_name='Ageing > 365 D', index=False)
        
        mb52_th40.to_excel(writer, sheet_name='MB52_TH40', index=False)
        mb52_th44.to_excel(writer, sheet_name='MB52_TH44', index=False)
        r138_th40.to_excel(writer, sheet_name='R138_TH40', index=False)
        r138_th44.to_excel(writer, sheet_name='R138_TH44', index=False)
        product_group.to_excel(writer, sheet_name='Product Group', index=False)

        for sheet_name in writer.sheets:
            target_ws = writer.sheets[sheet_name]
            
            if sheet_name != dash_sheet_name:
                col_formats = {}
                for col_num in range(1, target_ws.max_column + 1):
                    h_cell = target_ws.cell(row=1, column=col_num)
                    h_cell.fill = header_fill
                    h_cell.font = header_font
                    h_cell.border = thin_border
                    
                    h_val = str(h_cell.value)
                    if "Stock Value" in h_val or "Value" in h_val:
                        col_formats[col_num] = '#,##0.00'
                    elif "Quantity" in h_val or "Unrestricted" in h_val:
                        col_formats[col_num] = '#,##0'
                
                for row in target_ws.iter_rows(min_row=2, max_row=target_ws.max_row, min_col=1, max_col=target_ws.max_column):
                    for cell in row:
                        cell.border = thin_border
                        cell.font = data_font
                        if cell.column in col_formats and isinstance(cell.value, (int, float)):
                            cell.number_format = col_formats[cell.column]

            for col in target_ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.value is not None:
                        if sheet_name == dash_sheet_name and cell.row < 5:
                            continue
                        max_len = max(max_len, len(str(cell.value)))
                
                target_ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 45)

    print(f"\n[สำเร็จ] ประมวลผลแบบ Multi-Level VLOOKUP เสร็จสมบูรณ์ 100%!")
    print(f"-> ไฟล์รายงานพร้อมใช้งานถูกสร้างขึ้นแล้วที่: {output_path}")

except Exception as e:
    print("\n" + "="*60)
    print("❌ ระบบพบข้อผิดพลาด (ERROR) ไม่สามารถประมวลผลต่อได้ ❌")
    print("รายละเอียด:")
    print(e)
    print("="*60 + "\n")
finally:
    input("\nกด Enter เพื่อปิดหน้าต่างนี้...")