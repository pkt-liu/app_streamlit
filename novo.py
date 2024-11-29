import pandas as pd
import numpy as np
import os
import time
import streamlit as st

df_franqueados = pd.read_excel(r'C:\Users\thiagoferreira\Desktop\df_franqueados.xlsx')

st.write("Giro - Verão 25:")
st.dataframe(df_franqueados,hide_index=True)





