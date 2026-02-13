import streamlit as st
import pandas as pd
import random
from datetime import datetime, timedelta
from io import BytesIO

# --- STYLING CONSTANTS ---
COLORS = {"Sous - sol": "DDEBF7", "RDC": "FCE4D6", "1st floor": "E2EFDA", "Header": "F2F2F2"}

st.set_page_config(page_title="Lab Draw Generator", page_icon="🧪")
st.title("🧪 Lab Task Draw Generator")
st.write("Upload your Excel file and select the date range to generate the schedule.")

# 1. File Upload
uploaded_file = st.file_uploader("Upload your Input Excel File", type=["xlsx"])

if uploaded_file:
    # 2. Date Inputs in Sidebar
    with st.sidebar:
        st.header("Settings")
        start_date = st.date_input("Start Date", datetime.now())
        end_date = st.date_input("End Date", datetime.now() + timedelta(weeks=8))
        
    if st.button("Generate Schedule"):
        # Snap start date to Monday
        monday_start = start_date - timedelta(days=start_date.weekday())
        
        # Load and process data
        df = pd.read_excel(uploaded_file, header=[0, 1])
        
        # Header Reconstruction
        raw_cols = df.columns.to_frame()
        raw_cols[0] = raw_cols[0].mask(raw_cols[0].str.contains('Unnamed')).ffill()
        df.columns = [f"{col[0]} - {col[1]}" if "Unnamed" not in str(col[1]) else col[0] for col in raw_cols.values]
        
        name_col, contract_col, task_cols = df.columns[0], df.columns[2], df.columns[3:]

        def parse_contract(x):
            try: return datetime.strptime(str(x), "%b.%y").date()
            except: return datetime.max.date()
        df['expiry'] = df[contract_col].apply(parse_contract)

        # Draw Logic
        buckets = {task: [] for task in task_cols}
        schedule_data = []
        curr = monday_start

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

        # Create Styled Excel in Memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Schedule')
            workbook = writer.book
            worksheet = writer.sheets['Schedule']
            
            # Add your styling logic here (Fonts, Fills, Borders)
            # Example: 
            for cell in worksheet[1]: # Style the header row
                cell.font = Font(bold=True)
                # ... etc ...

        st.success("Schedule Generated!")
        st.download_button(
            label="📥 Download Excel Schedule",
            data=output.getvalue(),
            file_name="lab_schedule.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )