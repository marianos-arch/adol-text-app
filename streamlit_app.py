import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import re
import requests
from PIL import Image, ImageDraw
try:
    from streamlit_image_coordinates import streamlit_image_coordinates
except ImportError:
    streamlit_image_coordinates = None

MAX_IMAGE_WIDTH = 1000

st.title("Quick ADOL Template Guide")
st.write("Insert numbers into the pre-formatted ADOL Scoresheet on Excel and Print it!")

st.subheader("1. Enter Mentee's name and today's date")
col1, col2 = st.columns(2)

with col1:
    mentee_name = st.text_input("Enter name:", placeholder="John Doe")

with col2:
    date_input = st.text_input("Enter Today's date:", placeholder="MM/DD/YYYY")

combined_name = f"{mentee_name} {date_input}".strip()

st.subheader("2. Choose Your Input Method")
input_method = st.radio("How would you like to enter your scores?", options=["Type Numbers", "Click on Image"], key="input_method")

PAGE_1_COORDINATES = {
    1:  [(798, 774),  (948, 774),  (1100, 774),  (1248, 774),  (1400, 774)],
    2:  [(798, 874),  (948, 874),  (1100, 874),  (1248, 874),  (1400, 874)],
    3:  [(798, 974),  (948, 974),  (1100, 974),  (1248, 974),  (1400, 974)],
    4:  [(798, 1041), (948, 1041), (1100, 1041), (1248, 1041), (1400, 1041)],
    5:  [(798, 1141), (948, 1141), (1100, 1141), (1248, 1141), (1400, 1141)],
    6:  [(798, 1241), (948, 1241), (1100, 1241), (1248, 1241), (1400, 1241)],
    7:  [(798, 1341), (948, 1341), (1100, 1341), (1248, 1341), (1400, 1341)],
    8:  [(798, 1441), (948, 1441), (1100, 1441), (1248, 1441), (1400, 1441)],
    9:  [(798, 1574), (948, 1574), (1100, 1574), (1248, 1574), (1400, 1574)],
    10: [(798, 1674), (948, 1674), (1100, 1674), (1248, 1674), (1400, 1674)],
    11: [(798, 1774), (948, 1774), (1100, 1774), (1248, 1774), (1400, 1774)],
}

PAGE_2_COORDINATES = {
    12: [(798, 466),  (948, 466),  (1098, 466),  (1248, 466),  (1400, 466)],
    13: [(798, 566),  (948, 566),  (1098, 566),  (1248, 566),  (1400, 566)],
    14: [(798, 666),  (948, 666),  (1098, 666),  (1248, 666),  (1400, 666)],
    15: [(798, 733),  (948, 733),  (1098, 733),  (1248, 733),  (1400, 733)],
    16: [(798, 800),  (948, 800),  (1098, 800),  (1248, 800),  (1400, 800)],
    17: [(798, 900),  (948, 900),  (1098, 900),  (1248, 900),  (1400, 900)],
    18: [(798, 1000), (948, 1000), (1098, 1000), (1248, 1000), (1400, 1000)],
    19: [(798, 1100), (948, 1100), (1098, 1100), (1248, 1100), (1400, 1100)],
    20: [(798, 1166), (948, 1166), (1098, 1166), (1248, 1166), (1400, 1166)],
    21: [(798, 1266), (948, 1266), (1098, 1266), (1248, 1266), (1400, 1266)],
    22: [(798, 1365), (948, 1365), (1098, 1365), (1248, 1365), (1400, 1365)],
    23: [(798, 1433), (948, 1433), (1098, 1433), (1248, 1433), (1400, 1433)],
    24: [(798, 1500), (948, 1500), (1098, 1500), (1248, 1500), (1400, 1500)],
    25: [(798, 1600), (948, 1600), (1098, 1600), (1248, 1600), (1400, 1600)],
    26: [(798, 1700), (948, 1700), (1098, 1700), (1248, 1700), (1400, 1700)],
}

PAGE_3_COORDINATES = {
    27: [(798, 466),  (948, 466),  (1100, 466),  (1248, 466),  (1400, 466)],
    28: [(798, 566),  (948, 566),  (1100, 566),  (1248, 566),  (1400, 566)],
    29: [(798, 666),  (948, 666),  (1100, 666),  (1248, 666),  (1400, 666)],
    30: [(798, 761),  (948, 761),  (1100, 761),  (1248, 761),  (1400, 761)],
    31: [(798, 861),  (948, 861),  (1100, 861),  (1248, 861),  (1400, 861)],
    32: [(798, 961),  (948, 961),  (1100, 961),  (1248, 961),  (1400, 961)],
    33: [(798, 1028), (948, 1028), (1100, 1028), (1248, 1028), (1400, 1028)],
}

ALL_COORDINATES = {**PAGE_1_COORDINATES, **PAGE_2_COORDINATES, **PAGE_3_COORDINATES}

if 'numbers' not in st.session_state:
    st.session_state.numbers = [0] * 33
if 'current_page' not in st.session_state:
    st.session_state.current_page = 1

def find_question_at_click(x, y, page_num, scale_factor):
    if page_num == 1:
        question_range = range(1, 12)
    elif page_num == 2:
        question_range = range(12, 27)
    elif page_num == 3:
        question_range = range(27, 34)
    else:
        return None, None
    x_orig = x / scale_factor
    y_orig = y / scale_factor
    for question_num in question_range:
        if question_num in ALL_COORDINATES:
            coords = ALL_COORDINATES[question_num]
            for option_idx, (coord_x, coord_y) in enumerate(coords):
                distance = ((x_orig - coord_x) ** 2 + (y_orig - coord_y) ** 2) ** 0.5
                if distance < 80:
                    return question_num, option_idx + 1
    return None, None

def create_annotated_page_image(image_bytes, page_num, numbers, max_width=1000):
    img = Image.open(image_bytes).convert("RGB")
    original_width = img.width
    scale_factor = max_width / original_width
    new_width = max_width
    new_height = int(img.height * scale_factor)
    img_display = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
    draw = ImageDraw.Draw(img_display, 'RGBA')
    marker_color = (255, 0, 0)
    marker_radius = int(40 * scale_factor)
    if page_num == 1:
        question_range = range(1, 12)
    elif page_num == 2:
        question_range = range(12, 27)
    elif page_num == 3:
        question_range = range(27, 34)
    else:
        return img_display
    for question_num in question_range:
        selected_option = numbers[question_num - 1]
        if question_num in ALL_COORDINATES and 1 <= selected_option <= 5:
            coords = ALL_COORDINATES[question_num]
            x_orig, y_orig = coords[selected_option - 1]
            x = int(x_orig * scale_factor)
            y = int(y_orig * scale_factor)
            draw.ellipse([(x - marker_radius, y - marker_radius), (x + marker_radius, y + marker_radius)], outline=marker_color, width=5, fill=marker_color + (100,))
    return img_display

def process_excel_file(numbers, combined_name):
    try:
        github_url = "https://raw.githubusercontent.com/marianos-arch/adol-text-app/main/template.xlsx"
        response = requests.get(github_url)
        if response.status_code != 200:
            st.error(f"Failed to download template. Status code: {response.status_code}")
            return None
        template_bytes = BytesIO(response.content)
        wb = load_workbook(template_bytes)
        wb.calculation.calcMode = 'auto'
        ws = wb.active
        question_cell_mapping = {1: (48, 4), 2: (38, 4), 3: (39, 4), 4: (49, 4), 5: (40, 4), 6: (41, 4), 7: (50, 4), 8: (51, 4), 9: (42, 4), 10: (52, 4), 11: (43, 4), 12: (4, 4), 13: (29, 4), 14: (20, 4), 15: (5, 4), 16: (6, 4), 17: (7, 4), 18: (21, 4), 19: (30, 4), 20: (8, 4), 21: (9, 4), 22: (22, 4), 23: (10, 4), 24: (11, 4), 25: (31, 4), 26: (12, 4), 27: (32, 4), 28: (23, 4), 29: (13, 4), 30: (14, 4), 31: (24, 4), 32: (15, 4), 33: (33, 4)}
        ws.cell(row=1, column=3).value = combined_name
        for question_num, number in enumerate(numbers, start=1):
            if question_num in question_cell_mapping:
                row, col = question_cell_mapping[question_num]
                ws.cell(row=row, column=col).value = number
        try:
            wb.calculate()
        except:
            pass
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error processing template: {e}")
        return None

if input_method == "Type Numbers":
    st.subheader("2. Enter Your Numbers (From Question 1 to Question 33)")
    user_input = st.text_input("Enter numbers (e.g., '44442145' or '4 4 4 4 2 1 4 5'):", placeholder="44442145")
    st.caption("Each number must be between 1 and 5. Strong Disagree (1) - Strongly Agree (5)")
    if user_input:
        cleaned_input = re.sub(r'\s+', '', user_input)
        numbers = [int(char) for char in cleaned_input if char.isdigit()]
        if len(numbers) != 33:
            st.error(f"Please enter exactly 33 numbers. You entered {len(numbers)}.")
        elif not all(1 <= num <= 5 for num in numbers):
            st.error("All numbers must be between 1 and 5.")
        else:
            st.success(f"Valid input! You entered: {numbers}")
            st.session_state.numbers = numbers
            output = process_excel_file(numbers, combined_name)
            if output:
                safe_name = re.sub(r'[^a-zA-Z0-9_-]', '', mentee_name.replace(' ', '_'))
                safe_date = re.sub(r'[^0-9_-]', '', date_input.replace('/', '_'))
                file_name = f"ADOL_Scoresheet_{safe_name}_{safe_date}.xlsx" if safe_name else "ADOL_Scoresheet_Filled.xlsx"
                st.download_button(label="Download Filled ADOL Sheet", data=output.getvalue(), file_name=file_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.write("NOTE: Hit [Enable Editing] on Excel to reveal the scores")
                st.subheader("3. Your Entries - Visual Confirmation on Images:")
                st.write(f"Mentee Name and Today's Date: {combined_name}")
                for page_num in [1, 2, 3]:
                    image_filename = f"adol_blank-{page_num}.png"
                    github_image_url = f"https://raw.githubusercontent.com/marianos-arch/adol-text-app/main/{image_filename}"
                    try:
                        img_response = requests.get(github_image_url)
                        if img_response.status_code == 200:
                            annotated_img = create_annotated_page_image(BytesIO(img_response.content), page_num, numbers, max_width=MAX_IMAGE_WIDTH)
                            if page_num == 1:
                                st.markdown(f"**Page {page_num} - Questions 1-11:**")
                            elif page_num == 2:
                                st.markdown(f"**Page {page_num} - Questions 12-26:**")
                            else:
                                st.markdown(f"**Page {page_num} - Questions 27-33:**")
                            st.image(annotated_img, use_column_width=False, width=MAX_IMAGE_WIDTH)
                    except:
                        pass
else:
    st.subheader("2. Click on the ADOL Scoresheet to Enter Your Scores")
    st.write("Navigate through each page and click on your response for each question (1-5 scale)")
    try:
        image_filename = f"adol_blank-{st.session_state.current_page}.png"
        github_image_url = f"https://raw.githubusercontent.com/marianos-arch/adol-text-app/main/{image_filename}"
        img_response = requests.get(github_image_url)
        if img_response.status_code == 200:
            annotated_img = create_annotated_page_image(BytesIO(img_response.content), st.session_state.current_page, st.session_state.numbers, max_width=MAX_IMAGE_WIDTH)
            page_num = st.session_state.current_page
            if page_num == 1:
                st.markdown(f"**Page {page_num} - Questions 1-11:**")
            elif page_num == 2:
                st.markdown(f"**Page {page_num} - Questions 12-26:**")
            else:
                st.markdown(f"**Page {page_num} - Questions 27-33:**")
            if streamlit_image_coordinates is not None:
                coords = streamlit_image_coordinates(annotated_img, key=f"coords_page_{page_num}")
                if coords:
                    x = coords["x"]
                    y = coords["y"]
                    display_scale = MAX_IMAGE_WIDTH / Image.open(BytesIO(img_response.content)).width
                    question_num, option_value = find_question_at_click(x, y, page_num, display_scale)
                    if question_num and option_value:
                        st.session_state.numbers[question_num - 1] = option_value
                        st.success(f"Question {question_num} set to {option_value}")
                        st.rerun()
            else:
                st.image(annotated_img, use_column_width=False, width=MAX_IMAGE_WIDTH)
            if page_num == 1:
                question_range = range(1, 12)
            elif page_num == 2:
                question_range = range(12, 27)
            else:
                question_range = range(27, 34)
            st.write("---")
            st.write("**Or enter scores manually for this page:**")
            cols = st.columns(len(list(question_range)))
            for idx, question_num in enumerate(question_range):
                with cols[idx]:
                    st.session_state.numbers[question_num - 1] = st.selectbox(f"Q{question_num}", options=[0, 1, 2, 3, 4, 5], index=st.session_state.numbers[question_num - 1], key=f"question_{question_num}")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                if st.button("Previous Page", disabled=(st.session_state.current_page == 1)):
                    st.session_state.current_page -= 1
                    st.rerun()
            with col2:
                st.write(f"Page {st.session_state.current_page} of 3")
            with col3:
                if st.button("Next Page", disabled=(st.session_state.current_page == 3)):
                    st.session_state.current_page += 1
                    st.rerun()
            with col4:
                filled_count = sum(1 for n in st.session_state.numbers if n > 0)
                st.metric("Progress", f"{filled_count}/33")
            st.write("---")
            if any(n > 0 for n in st.session_state.numbers):
                st.subheader("Your Current Entries:")
                col1, col2 = st.columns(2)
                with col1:
                    st.write("**Summary:**")
                    entries_str = ''.join(str(n) if n > 0 else '_' for n in st.session_state.numbers)
                    st.code(entries_str, language="text")
                with col2:
                    st.write("**Detailed View:**")
                    filled_answers = {i+1: n for i, n in enumerate(st.session_state.numbers) if n > 0}
                    st.json(filled_answers)
                if all(n > 0 for n in st.session_state.numbers):
                    st.success("All questions answered!")
                    output = process_excel_file(st.session_state.numbers, combined_name)
                    if output:
                        safe_name = re.sub(r'[^a-zA-Z0-9_-]', '', mentee_name.replace(' ', '_'))
                        safe_date = re.sub(r'[^0-9_-]', '', date_input.replace('/', '_'))
                        file_name = f"ADOL_Scoresheet_{safe_name}_{safe_date}.xlsx" if safe_name else "ADOL_Scoresheet_Filled.xlsx"
                        st.download_button(label="Download Filled ADOL Sheet", data=output.getvalue(), file_name=file_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        st.write("NOTE: Hit [Enable Editing] on Excel to reveal the scores")
    except:
        pass
