import os
import sys
import pandas as pd
import streamlit as st

# Default input path (override via upload)
DEFAULT_INPUT = './input/data.xlsx'

@st.cache_data
def load_data(path):
    try:
        df = pd.read_excel(path, sheet_name='Sheet1', parse_dates=['Date'], engine='openpyxl')
    except ImportError:
        st.error("Missing dependency: please install openpyxl (`pip install openpyxl`)")
        st.stop()
    except Exception as e:
        st.error(f"Error loading data: {e}")
        st.stop()
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

# Date range filter (handles single date or range)
min_date, max_date = df['Date'].min(), df['Date'].max()
date_range = st.sidebar.date_input(
    'Select date range',
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date
)
if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
    start, end = date_range
else:
    start = end = date_range

mask = (df['Date'] >= pd.to_datetime(start)) & (df['Date'] <= pd.to_datetime(end))
filtered = df.loc[mask].copy()

# Compute FP/FN, correctness, mistake count
filtered[['FP', 'FN']] = filtered.apply(identify_mistaken_policies, axis=1, result_type='expand')
filtered['Is_Correct'] = filtered['Is_Task_Correct'] == 'correct'
filtered['Mistake_Count'] = filtered['FP'].apply(len) + filtered['FN'].apply(len)

# Top row: Accuracy & AHT over time
col1, col2 = st.columns(2)
with col1:
    st.subheader('Accuracy Over Time')
    acc = filtered.groupby('Date')['Is_Correct'].mean() * 100
    st.line_chart(acc)
with col2:
    st.subheader('Average Handling Time (m)')
    if 'Handling time (m)' in filtered.columns:
        aht = filtered.groupby('Date')['Handling time (m)'].mean()
        st.line_chart(aht)
    else:
        st.warning('Column "Handling time (m)" not found.')

# Shared policy filter
st.subheader('Filter Policies')
all_policies = pd.concat([filtered.explode('FP')['FP'], filtered.explode('FN')['FN']]).value_counts().index.tolist()
filter_mode = st.selectbox('Filter mode', ['Top 5', 'Bottom 5', 'Manual'])
if filter_mode == 'Top 5':
    selected = all_policies[:5]
elif filter_mode == 'Bottom 5':
    selected = all_policies[-5:]
else:
    selected = st.multiselect('Select policies', all_policies)

# Mistaken Policies (full-width, rotated x-axis labels)
st.subheader('Mistaken Policies')
mistakes = filtered[~filtered['Is_Correct']]
policy_counts = pd.concat([mistakes.explode('FP')['FP'], mistakes.explode('FN')['FN']]).value_counts()
if selected:
    sel_counts = policy_counts.reindex(selected).fillna(0)
    # prepare for Vega-Lite
    chart_data = sel_counts.reset_index()
    chart_data.columns = ['Policy', 'Count']
    st.vega_lite_chart(
        chart_data,
        {
            "width": 700,
            "height": 350,
            "mark": "bar",
            "encoding": {
                "x": {"field": "Policy", "type": "nominal", "axis": {"labelAngle": -45}},
                "y": {"field": "Count", "type": "quantitative"}
            }
        }
    )
else:
    st.info('No policies selected.')

# RCA (full-width, below Mistaken Policies)
st.subheader('RCA')
rca_df = pd.concat([
    mistakes.explode('FP')[['FP', 'Controllable', 'RCA']].rename(columns={'FP': 'Policy'}),
    mistakes.explode('FN')[['FN', 'Controllable', 'RCA']].rename(columns={'FN': 'Policy'})
])
if selected:
    rca_filt = rca_df.loc[rca_df['Policy'].isin(selected)]
    rca_counts = rca_filt.groupby(['RCA', 'Controllable']).size().unstack(fill_value=0)
    chart_data = rca_counts.reset_index().melt(id_vars='RCA', var_name='Controllable', value_name='Count')
    st.vega_lite_chart(
        chart_data,
        {
            "width": 700,
            "height": 350,
            "mark": "bar",
            "encoding": {
                "x": {"field": "RCA", "type": "ordinal", "axis": {"labelAngle": -45}},
                "y": {"field": "Count", "type": "quantitative"},
                "color": {"field": "Controllable", "type": "nominal"}
            }
        }
    )
else:
    st.info('No policies selected.')