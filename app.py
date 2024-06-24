import streamlit as st
import pandas as pd
import plotly.express as px
import datetime

thisyear = datetime.date.today().year

# Load the Excel file
excel_path = '秉儒保單.xlsx'
excel_sheets = pd.read_excel(excel_path, sheet_name=None)
# Function to read columns from D onwards
def load_relevant_columns(sheet):
    return sheet.iloc[:, 4:]

# Read and process each sheet into a dictionary
sheets_dict = {sheet_name: load_relevant_columns(sheet).assign(Policy=sheet_name) for sheet_name, sheet in excel_sheets.items()}

# Combine all sheets into one dataframe with the sheet name as the 'Policy' column
combined_df = pd.concat([load_relevant_columns(sheet).assign(Policy=sheet_name) for sheet_name, sheet in excel_sheets.items()], ignore_index=True)

combined_df['Year'] = combined_df['西元年']

# List of columns to exclude from the item selection
exclude_columns = ['Year', 'Policy', '年齡', '西元年']

# Streamlit App
st.title('秉儒的保險')

# Sidebar for filters
st.sidebar.header('Filter Options')

# Determine the range of years in the data
min_year = int(combined_df['Year'].min())
max_year = int(combined_df['Year'].max())

selected_years = st.sidebar.slider('Select Year Range', min_year, max_year, (thisyear, thisyear +100))
selected_policies = st.sidebar.multiselect('Select Policies', options=combined_df['Policy'].unique(), default=combined_df['Policy'].unique())
selected_item = st.sidebar.selectbox('Select Item', options=[col for col in combined_df.columns if col not in exclude_columns])

# Filter data based on selections
filtered_df = combined_df[(combined_df['Year'] >= selected_years[0]) & (combined_df['Year'] <= selected_years[1]) & (combined_df['Policy'].isin(selected_policies))]

# Exclude policies with all zero values for the selected item
policies_to_include = filtered_df.groupby('Policy')[selected_item].sum()
policies_to_include = policies_to_include[policies_to_include != 0].index
filtered_df = filtered_df[filtered_df['Policy'].isin(policies_to_include)]

# Plot the stacked area chart
fig = px.area(filtered_df, x='Year', y=selected_item, color='Policy', 
              labels={'value': 'Amount', 'Year': 'Year'},
              title=f'{selected_item}')

# Display the plot in Streamlit
st.plotly_chart(fig)


# Filter the dictionary to include only the selected policies
filtered_sheets_dict = {policy: sheets_dict[policy] for policy in selected_policies}

unique_items = []
for df in filtered_sheets_dict.values():
    for i in df.columns:
        if i not in unique_items and i not in ['西元年', '年齡', 'Policy']:
            unique_items.append(i)
            
# Initialize a new DataFrame with '西元年'
combined_years = pd.DataFrame({'西元年': pd.concat([df['西元年'] for df in filtered_sheets_dict.values()]).unique()})
combined_years.sort_values('西元年', inplace=True)
combined_years.reset_index(drop=True, inplace=True)

# Add columns for each unique item and fill with zeros
for item in unique_items:
    combined_years[item] = 0

# Sum the values for each item by year across all selected policies
for item in unique_items:
    for policy, df in filtered_sheets_dict.items():
        if item in df.columns:
            combined_years = combined_years.merge(df[['西元年', item]], on='西元年', how='left', suffixes=('', f'_{policy}'))
            combined_years[item] += combined_years[f'{item}_{policy}'].fillna(0)
            combined_years.drop(columns=[f'{item}_{policy}'], inplace=True)

# Rename columns for clarity
combined_years.rename(columns={'西元年': 'Year'}, inplace=True)

# Filter to only show rows where the sum of values is not zero
filtered_grouped_df = combined_years.loc[:, (combined_years != 0).any(axis=0)]

# Display the table in Streamlit
st.header('保單整理表')
st.dataframe(filtered_grouped_df)
