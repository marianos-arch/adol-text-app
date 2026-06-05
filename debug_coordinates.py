import streamlit as st
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO
import requests

st.set_page_config(layout="wide")
st.title("🔍 ADOL Coordinates Debug Tool")
st.write("This tool displays ALL coordinates for every question on each page so you can verify positioning.")

# Coordinate mappings for all three pages
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
    12: [(798, 466),  (948, 466),  (1098, 466),  (1258, 466),  (1400, 466)],
    13: [(798, 566),  (948, 566),  (1098, 566),  (1258, 566),  (1400, 566)],
    14: [(798, 666),  (948, 666),  (1098, 666),  (1258, 666),  (1400, 666)],
    15: [(798, 733),  (948, 733),  (1098, 733),  (1258, 733),  (1400, 733)],
    16: [(798, 800),  (948, 800),  (1098, 800),  (1258, 800),  (1400, 800)],
    17: [(798, 900),  (948, 900),  (1098, 900),  (1258, 900),  (1400, 900)],
    18: [(798, 1000), (948, 1000), (1098, 1000), (1258, 1000), (1400, 1000)],
    19: [(798, 1100), (948, 1100), (1098, 1100), (1258, 1100), (1400, 1100)],
    20: [(798, 1166), (948, 1166), (1098, 1166), (1258, 1166), (1400, 1166)],
    21: [(798, 1266), (948, 1266), (1098, 1266), (1258, 1266), (1400, 1266)],
    22: [(798, 1365), (948, 1365), (1098, 1365), (1258, 1365), (1400, 1365)],
    23: [(798, 1433), (948, 1433), (1098, 1433), (1258, 1433), (1400, 1433)],
    24: [(798, 1500), (948, 1500), (1098, 1500), (1258, 1500), (1400, 1500)],
    25: [(798, 1600), (948, 1600), (1098, 1600), (1258, 1600), (1400, 1600)],
    26: [(798, 1700), (948, 1700), (1098, 1700), (1258, 1700), (1400, 1700)],
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


def create_debug_page_image(image_bytes, page_num):
    """
    Create a debug version of a page with ALL markers and labels for every option.
    
    Args:
        image_bytes: BytesIO object containing the image data
        page_num: Page number (1, 2, or 3)
    
    Returns:
        PIL Image object with all markers and labels
    """
    img = Image.open(image_bytes).convert("RGB")
    draw = ImageDraw.Draw(img, 'RGBA')
    
    # Colors for each option (1-5)
    colors = [
        (255, 0, 0),      # Red for option 1
        (0, 255, 0),      # Green for option 2
        (0, 0, 255),      # Blue for option 3
        (255, 255, 0),    # Yellow for option 4
        (255, 0, 255),    # Magenta for option 5
    ]
    
    marker_radius = 40
    
    # Determine question range for this page
    if page_num == 1:
        question_range = range(1, 12)
    elif page_num == 2:
        question_range = range(12, 27)
    elif page_num == 3:
        question_range = range(27, 34)
    else:
        return img
    
    # Draw all markers for this page
    for question_num in question_range:
        if question_num in ALL_COORDINATES:
            coords = ALL_COORDINATES[question_num]
            
            # Draw all 5 options for this question
            for option_idx, (x, y) in enumerate(coords):
                color = colors[option_idx]
                
                # Draw circle
                draw.ellipse(
                    [(x - marker_radius, y - marker_radius), (x + marker_radius, y + marker_radius)],
                    outline=color,
                    width=3,
                    fill=color + (80,)
                )
                
                # Draw option number inside the circle
                text = str(option_idx + 1)
                try:
                    # Try to use a larger font
                    font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 30)
                except:
                    font = ImageFont.load_default()
                
                # Get text bounding box to center it
                bbox = draw.textbbox((0, 0), text, font=font)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]
                text_x = x - text_width // 2
                text_y = y - text_height // 2
                
                draw.text((text_x, text_y), text, font=font, fill=(255, 255, 255))
                
                # Add small label with question and option number
                label = f"Q{question_num}-{option_idx + 1}"
                try:
                    small_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 12)
                except:
                    small_font = ImageFont.load_default()
                
                draw.text((x - marker_radius - 5, y - marker_radius - 20), label, font=small_font, fill=(0, 0, 0))
    
    return img


# Select page to view
st.subheader("Select a Page to Debug")
page_selection = st.radio("Choose page:", [1, 2, 3], horizontal=True)

try:
    image_filename = f"adol_blank-{page_selection}.png"
    github_image_url = f"https://raw.githubusercontent.com/marianos-arch/adol-text-app/main/{image_filename}"
    
    st.info(f"Loading {image_filename}...")
    img_response = requests.get(github_image_url)
    
    if img_response.status_code == 200:
        # Create debug version
        debug_img = create_debug_page_image(BytesIO(img_response.content), page_selection)
        
        # Display the debug image
        st.markdown(f"### Page {page_selection} - All Coordinates Visible")
        st.image(debug_img, use_column_width=True)
        
        # Show coordinate data in a table
        st.subheader("Coordinate Details")
        
        if page_selection == 1:
            coords_dict = PAGE_1_COORDINATES
            question_range = range(1, 12)
        elif page_selection == 2:
            coords_dict = PAGE_2_COORDINATES
            question_range = range(12, 27)
        else:
            coords_dict = PAGE_3_COORDINATES
            question_range = range(27, 34)
        
        # Display coordinates in a readable format
        for question_num in question_range:
            if question_num in coords_dict:
                coords = coords_dict[question_num]
                st.write(f"**Question {question_num}:**")
                cols = st.columns(5)
                for option_idx, (x, y) in enumerate(coords):
                    with cols[option_idx]:
                        st.caption(f"Option {option_idx + 1}\n({x}, {y})")
    else:
        st.error(f"❌ Could not load {image_filename} from GitHub (Status: {img_response.status_code})")
        st.info("Make sure the PNG files exist in your repository.")

except Exception as e:
    st.error(f"❌ Error: {e}")

# Show color legend
st.subheader("🎨 Color Legend")
colors_legend = {
    "Option 1": "🔴 Red",
    "Option 2": "🟢 Green",
    "Option 3": "🔵 Blue",
    "Option 4": "🟡 Yellow",
    "Option 5": "🟣 Magenta",
}

for option, color in colors_legend.items():
    st.write(f"{color} = {option}")

st.info("💡 **Tip:** Each circle is labeled with Q[question_number]-[option_number]. Use these coordinates to verify and adjust your template positions.")
