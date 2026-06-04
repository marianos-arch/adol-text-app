import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import re
import requests
from datetime import datetime


st.title("Quick ADOL Template Guide")

st.write("Insert numbers into the pre-formatted ADOL Scoresheet on Excel and Print it!")

# Step 1: Get Mentee's Name
st.subheader("1. Enter Mentee's name and date")
col1, col2 = st.columns(2)

with col1:
    mentee_name = st.text_input(
        "Enter name:",
        placeholder="John Doe"
    )

with col2:
    date_input = st.text_input(
        "Enter date:",
        placeholder="MM/DD/YYYY"
    )

# Combine name and date
combined_name = f"{mentee_name} {date_input}".strip()

# Step 2: Get user input
st.subheader("2. Enter Your Numbers (From Question 1 to Question 33)")
user_input = st.text_input(
    "Enter numbers (e.g., '44442145' or '4 4 4 4 2 1 4 5'):",
    placeholder="44442145"
)

# Step 3: Process and validate input
if user_input:
    # Remove spaces and extract only digits
    cleaned_input = re.sub(r'\s+', '', user_input)
    numbers = [int(char) for char in cleaned_input if char.isdigit()]
    
    # Validate: Should have exactly 33 numbers (1-5 scale)
    if len(numbers) != 33:
        st.error(f"❌ Please enter exactly 33 numbers. You entered {len(numbers)}.")
    elif not all(1 <= num <= 5 for num in numbers):
        st.error("❌ All numbers must be between 1 and 5.")
    else:
        st.success(f"✅ Valid input! You entered: {numbers}")

        try:
            # Download template from GitHub
            github_url = "https://raw.githubusercontent.com/marianos-arch/adol-text-app/main/template.xlsx"
            response = requests.get(github_url)
            
            if response.status_code != 200:
                st.error(f"❌ Failed to download template. Status code: {response.status_code}")
            else:
                template_bytes = BytesIO(response.content)
                
                # Load workbook
                wb = load_workbook(template_bytes)
                wb.calculation.calcMode = 'auto' 
                ws = wb.active  # use the active sheet
                
                
                # Step 4: Map each question number to its specific cell
                # Format: question_number -> (row, col)
                question_cell_mapping = {
                    1: (48, 4),   # D48
                    2: (38, 4),   # D38
                    3: (39, 4),   # D39
                    4: (49, 4),   # D49
                    5: (40, 4),   # D40
                    6: (41, 4),   # D41
                    7: (50, 4),   # D50
                    8: (51, 4),   # D51
                    9: (42, 4),   # D42
                    10: (52, 4),  # D52
                    11: (43, 4),  # D43
                    12: (4, 4),   # D4
                    13: (29, 4),  # D29
                    14: (20, 4),  # D20
                    15: (5, 4),   # D5
                    16: (6, 4),   # D6
                    17: (7, 4),   # D7
                    18: (21, 4),  # D21
                    19: (30, 4),  # D30
                    20: (8, 4),   # D8
                    21: (9, 4),   # D9
                    22: (22, 4),  # D22
                    23: (10, 4),  # D10
                    24: (11, 4),  # D11
                    25: (31, 4),  # D31
                    26: (12, 4),  # D12
                    27: (32, 4),  # D32
                    28: (23, 4),  # D23
                    29: (13, 4),  # D13
                    30: (14, 4),  # D14
                    31: (24, 4),  # D24
                    32: (15, 4),  # D15
                    33: (33, 4),  # D33
                }
                
                # Insert combined mentee name and date at C1
                ws.cell(row=1, column=3).value = combined_name
                
                # Insert each number into its corresponding cell
                for question_num, number in enumerate(numbers, start=1):
                    if question_num in question_cell_mapping:
                        row, col = question_cell_mapping[question_num]
                        cell = ws.cell(row=row, column=col)
                        cell.value = number

                #force Excel to recalculate formulas
                try:
                    wb.calculate()
                except: 
                        pass # some older excel versions may not support this
                
                # Step 5: Save to BytesIO for download
                output = BytesIO()
                wb.save(output)
                output.seek(0)
                
                # Step 6: Provide download button
                # Generate filename from mentee name and date
                safe_name = re.sub(r'[^a-zA-Z0-9_-]', '', mentee_name.replace(' ', '_'))
                safe_date = re.sub(r'[^0-9_-]', '', date_input.replace('/', '_'))
                file_name = f"ADOL_Scoresheet_{safe_name}_{safe_date}.xlsx" if safe_name else "ADOL_Scoresheet_Filled.xlsx"
                
                st.download_button(
                    label="📥 Download Filled ADOL Sheet",
                    data=output.getvalue(),
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # Display the numbers for verification
                st.subheader("3. Your Entries:")
                st.write(f"Mentee Name and Date: {combined_name}")
                st.write(f"Numbers entered: {', '.join(map(str, numbers))}")
                
        except Exception as e:
            st.error(f"❌ Error processing template: {e}")
            st.info("Make sure the template.xlsx file exists in your GitHub repository.")
