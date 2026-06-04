import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import re
import requests


st.title("Quick ADOL Template Guide")

st.write("Insert numbers into the pre-formatted ADOL Scoresheet on Excel and Print it!")

# Step 1: Get user input
st.subheader("1. Enter Your Numbers (From Question 1 to Question 32")
user_input = st.text_input(
    "Enter numbers (e.g., '44442145' or '4 4 4 4 2 1 4 5'):",
    placeholder="44442145"
)

        # Download template from GitHub
        # github_url = "https://raw.githubusercontent.com/marianos-arch/adol-text-app/main/template.xlsx"
        # response = requests.get(github_url)
