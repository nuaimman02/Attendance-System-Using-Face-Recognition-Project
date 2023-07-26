import streamlit as st
import pandas as pd

from streamlit_autorefresh import st_autorefresh
from datetime import datetime

now = datetime.now()
dtString = now.strftime("%d/%m/%Y,%H:%M:%S")

count = st_autorefresh(interval=2000, limit=100, key="fizzbuzzcounter")

if count == 0:
    st.write("Count is zero")
elif count % 3 == 0 and count % 5 == 0:
    st.write("FizzBuzz")
elif count % 3 == 0:
    st.write("Fizz")
elif count % 5 == 0:
    st.write("Buzz")
else:
    st.write(f"Count: {count}")

attendance_path = "C:\\Users\\User\\Documents\\Programming Codes\\Python\\PyCharmProject\\pythonProject\\CSC3600_AttendanceSystemProject\\attendance.csv"

df = pd.read_csv(attendance_path)

st.dataframe(df.style.highlight_max(axis=0))
