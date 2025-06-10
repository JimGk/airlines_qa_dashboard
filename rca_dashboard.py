import os
import sys
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

# Modern, professional style
graphics_size = (10, 4)
plt.style.use('fivethirtyeight')

# Default input path (override via upload)
DEFAULT_INPUT = './input/data.xlsx'

@st.cache_data
def load_data(path):
    df = pd.read_excel(path, sheet_name='Sheet1', parse_dates=['Date'])
    return df

def identify_mistaken_policies(row):
    correct = eval(
        row['Correct Email Policy List'] if row['Method'] == 'Call'
        else row['Correct Call Policy List']
    )
    sim = eval(row['Policy List'])
    fp = list(set(sim) - set(correct))
    fn = list(set(correct) - set(sim))
    return fp, fn

# Streamlit setup
st.set_page_config(page_title='QA Dashboard', layout='wide')
st.title('Airlines QA Dashboard')

# Sidebar: File upload & date filter
uploaded = st.sidebar.file_uploader('Upload Excel file', type=['xlsx'])
if uploaded:
    df = load_data(uploaded)
else:
    df = load_data(DEFAULT_INPUT)

# Validate Date column
if 'Date' not in df.columns:
    st.error('No "Date" column found in the data!')
    st.stop()

# Date range filter
min_date, max_date = df['Date'].min(), df['Date'].max()
start, end = st.sidebar.date_input(
    'Select date range', [min_date, max_date], min_value=min_date, max_value=max_date
)
mask = (df['Date'] >= pd.to_datetime(start)) & (df['Date'] <= pd.to_datetime(end))
filtered = df.loc[mask].copy()

# Compute FP/FN, correctness, mistake count
filtered[['FP', 'FN']] = filtered.apply(
    identify_mistaken_policies, axis=1, result_type='expand'
)
filtered['Is_Correct'] = filtered['Is_Task_Correct'] == 'correct'
filtered['Mistake_Count'] = filtered['FP'].apply(len) + filtered['FN'].apply(len)

# 1) Accuracy over time & 2) AHT over time
col1, col2 = st.columns(2)
with col1:
    st.subheader('Accuracy Over Time')
    acc = filtered.groupby('Date')['Is_Correct'].mean() * 100
    fig1, ax1 = plt.subplots(figsize=graphics_size)
    ax1.plot(acc.index, acc.values, marker='o')
    ax1.set_ylabel('Accuracy (%)')
    ax1.set_xlabel('Date')
    ax1.set_ylim(0, 100)
    ax1.tick_params(axis='x', rotation=45)
    st.pyplot(fig1)

with col2:
    st.subheader('Average Handling Time (m)')
    if 'Handling time (m)' in filtered.columns:
        aht = filtered.groupby('Date')['Handling time (m)'].mean()
        fig2, ax2 = plt.subplots(figsize=graphics_size)
        ax2.plot(aht.index, aht.values, marker='o')
        ax2.set_ylabel('AHT (m)')
        ax2.set_xlabel('Date')
        ax2.set_ylim(bottom=0)
        ax2.tick_params(axis='x', rotation=45)
        st.pyplot(fig2)
    else:
        st.warning('Column "Handling time (m)" not found.')

# Shared policy filter for both charts
st.subheader('Filter Policies')
all_policies = pd.concat([
    filtered.explode('FP')['FP'],
    filtered.explode('FN')['FN']
]).value_counts().index.tolist()
filter_mode = st.selectbox('Filter mode', ['Top 5', 'Bottom 5', 'Manual'])
if filter_mode == 'Top 5':
    selected = all_policies[:5]
elif filter_mode == 'Bottom 5':
    selected = all_policies[-5:]
else:
    selected = st.multiselect('Select policies', all_policies)

# 3 & 4) Side-by-side: Mistaken Policy & RCA counts
col3, col4 = st.columns(2)

with col3:
    st.subheader('Mistaken Policies')
    mistakes = filtered[~filtered['Is_Correct']]
    policy_counts = pd.concat([
        mistakes.explode('FP')['FP'],
        mistakes.explode('FN')['FN']
    ]).value_counts()
    if selected:
        sel_counts = policy_counts.reindex(selected).fillna(0)
        fig3, ax3 = plt.subplots(figsize=graphics_size)
        sel_counts.plot(kind='bar', ax=ax3)
        ax3.set_ylabel('Count')
        ax3.set_ylim(bottom=0)
        ax3.tick_params(axis='x', rotation=45)
        st.pyplot(fig3)
    else:
        st.info('No policies selected.')

with col4:
    st.subheader('RCA')
    mistakes = filtered[~filtered['Is_Correct']]
    rca_df = pd.concat([
        mistakes.explode('FP')[['FP', 'Controllable', 'RCA']].rename(columns={'FP': 'Policy'}),
        mistakes.explode('FN')[['FN', 'Controllable', 'RCA']].rename(columns={'FN': 'Policy'})
    ])
    if selected:
        rca_filt = rca_df.loc[rca_df['Policy'].isin(selected)]
        rca_counts = rca_filt.groupby(['RCA','Controllable']).size().unstack(fill_value=0)
        fig4, ax4 = plt.subplots(figsize=graphics_size)
        rca_counts.plot(kind='bar', stacked=True, ax=ax4)
        ax4.set_ylabel('Count')
        ax4.set_ylim(bottom=0)
        ax4.tick_params(axis='x', rotation=45)
        st.pyplot(fig4)
    else:
        st.info('No policies selected.')