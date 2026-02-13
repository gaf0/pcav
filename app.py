import streamlit as st
import pandas as pd
import random
from datetime import datetime, timedelta
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# --- CONFIGURATION ---
COLORS = {
    "Sous - sol": "DDEBF7", 
    "RDC": "FCE4D6",      
    "1st floor": "E2EFDA", 
    "Header": "F2F2F2"    
}

st.set_page_config(page_title="Lab Schedule Generator", layout="wide")
st.title("🧪 Lab Task Draw Generator")

uploaded_file = st.file_uploader("Choose your Excel input file", type="xlsx")

if uploaded_file:
    with st.sidebar:
        st.header("Parameters")
        start_date = st.date_input("Start Date", datetime.now())
        end_date = st.date_input("End Date", datetime.now() + timedelta(weeks=12))

    if st.button("Generate & Style Schedule"):
        # 1. Process Data
        df = pd.read_excel(uploaded_file, header=[0, 1])
        
        raw_cols = df.columns.to_frame()
        raw_cols[0] = raw_cols[0].mask(raw_cols[0].str.contains('Unnamed')).ffill()
        df.columns = [f"{col[0]} - {col[1]}" if "Unnamed" not in str(col[1]) else col[0] for col in raw_cols.values]
        
        name_col, contract_col, task_cols = df.columns[0], df.columns[2], df.columns[3:]

        def parse_contract(x):
            try: return datetime.strptime(str(x), "%b.%y").date()
            except: return datetime.max.date()
        df['expiry'] = df[contract_col].apply(parse_contract)

        # 2. Draw Logic
        buckets = {task: [] for task in task_cols}
        schedule_data = []
        curr = start_date - timedelta(days=start_date.weekday()) # Monday anchor

        while curr <= end_date:
            year, week_num, _ = curr.isocalendar()
            week_label = f"W{week_num} ({curr.strftime('%d/%m/%Y')})"
            week_row = {"Week": week_label}
            picked_this_week = set()

            for task in task_cols:
                eligible = df[(df[task].notna()) & (df[task] != False) & (df['expiry'] >= curr)][name_col].tolist()
                if not eligible:
                    week_row[task] = "N/A"
                    continue
                if not any(p in eligible for p in buckets[task]):
                    buckets[task] = list(eligible)
                    random.shuffle(buckets[task])
                
                candidates = [p for p in buckets[task] if p in eligible]
                fresh = [p for p in candidates if p not in picked_this_week]
                winner = random.choice(fresh if fresh else candidates)
                
                buckets[task].remove(winner)
                picked_this_week.add(winner)
                week_row[task] = winner
                
            schedule_data.append(week_row)
            curr += timedelta(days=7)

        # 3. Create Styled Excel in Memory
        output = BytesIO()
        final_df = pd.DataFrame(schedule_data)
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Schedule')
            workbook = writer.book
            worksheet = writer.sheets['Schedule']
            
            header_font = Font(bold=True)
            border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

            for col_num, column_title in enumerate(final_df.columns, 1):
                cell = worksheet.cell(row=1, column=col_num)
                cell.font = header_font
                cell.border = border
                cell.alignment = Alignment(horizontal='center')
                
                # Apply Colors
                fill_color = COLORS["Header"]
                if "Sous - sol" in column_title: fill_color = COLORS["Sous - sol"]
                elif "RDC" in column_title: fill_color = COLORS["RDC"]
                elif "1st floor" in column_title: fill_color = COLORS["1st floor"]
                
                cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
                worksheet.column_dimensions[cell.column_letter].width = 22

        st.success("Drawing complete!")
        st.download_button(
            label="📥 Download Styled Excel Schedule",
            data=output.getvalue(),
            file_name=f"Lab_Schedule_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )