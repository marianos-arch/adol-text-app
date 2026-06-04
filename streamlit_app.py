import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import re
import requests


st.title("Quick ADOL Template Guide")

st.write("Insert numbers into the pre-formatted ADOL Scoresheet on Excel and Print it!")

# Step 1: Get Mentee's Name
st.subheader("1. Enter Mentee's name")
user_input = st.text_name(
    "Enter name:)",
    placeholder="John Doe"
)
# Step 2: Get user input
st.subheader("2. Enter Your Numbers (From Question 1 to Question 33")
user_input = st.text_input(
    "Enter numbers (e.g., '44442145' or '4 4 4 4 2 1 4 5'):",
    placeholder="44442145"
)

# Step 2: Process and validate input
if user_input:
    # Remove spaces and extract only digits
    cleaned_input = re.sub(r'\s+', '', user_input)
    numbers = [int(char) for char in cleaned_input if char.isdigit()]
    
    # Validate: Should have exactly 32 numbers (1-5 scale)
    if len(numbers) != 33:
        st.error(f"❌ Please enter exactly 32 numbers. You entered {len(numbers)}.")
    elif not all(1 <= num <= 5 for num in numbers):
        st.error("❌ All numbers must be between 1 and 5.")
    else:
        st.success(f"✅ Valid input! You entered: {numbers}")

        try 
            # Download template from GitHub
            github_url = "https://raw.githubusercontent.com/marianos-arch/adol-text-app/main/template.xlsx"
            response = requests.get(github_url)
            
            if response.status_code != 200:
                st.error(f"❌ Failed to download template. Status code: {response.status_code}")
            else:
                template_bytes = BytesIO(response.content)
                
                # Load workbook
                wb = load_workbook(template_bytes)
                ws = wb.active
                
                # Step 4: Insert numbers into specific cells
                # Adjust the starting cell based on your template layout
                # Example: Starting at cell B2 (row 2, column 2)
                start_row = 2
                start_col = 2
                
                for idx, number in enumerate(numbers):
                    row = start_row + idx
                    cell = ws.cell(row=row, column=start_col)
                    cell.value = number
                
                # Step 5: Save to BytesIO for download
                output = BytesIO()
                wb.save(output)
                output.seek(0)
                
                # Step 6: Provide download button
                st.download_button(
                    label="📥 Download Filled ADOL Sheet",
                    data=output.getvalue(),
                    file_name="ADOL_Scoresheet_Filled.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # Display the numbers for verification
                st.subheader("3. Your Entries:")
                st.write(f"Numbers entered: {', '.join(map(str, numbers))}")
                
        except Exception as e:
            st.error(f"❌ Error processing template: {e}")
            st.info("Make sure the template.xlsx file exists in your GitHub repository.")



        

        # Download template from GitHub
        # github_url = "https://raw.githubusercontent.com/marianos-arch/adol-text-app/main/template.xlsx"
        # response = requests.get(github_url)
