import pandas as pd
import streamlit as st
import openpyxl
#E:\padoshi\myvenv\Scripts>activate
#https://www.geeksforgeeks.org/create-virtual-environment-using-venv-python/
# define a function to create a pivot table
import streamlit as st
import pandas as pd

# Function to read excel file and return selected sheet
def read_excel(file, sheet_name):
    xl = pd.ExcelFile(file)
    df = xl.parse(sheet_name)
    return df

def main():
    st.title("Excel File Analyzer")
    # Customizing sidebar color and style
    st.markdown(
        """
        <style>
        .sidebar .sidebar-content {
            background-color: #f0f0f0;
        }
        </style>
        """,
        unsafe_allow_html=True
    )
    st.sidebar.title("Excel File Viewer")
    st.sidebar.write("Upload your Excel file below:")
    
    uploaded_file = st.sidebar.file_uploader("Choose a file", type=['xls', 'xlsx'])
    
    if uploaded_file is not None:
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names
        selected_sheet = st.sidebar.selectbox("Select a sheet to display", sheet_names)
        df = read_excel(uploaded_file, selected_sheet)
        with st.expander("Selected Sheet data"):
            st.write(df)
        st.sidebar.title("Create Pivot Table")
        
        columns = st.sidebar.multiselect("Select columns for pivot table", df.columns.tolist())
        index = st.sidebar.selectbox("Select index column", df.columns.tolist())
        agg_functions = st.sidebar.multiselect("Select aggregation functions", ['sum', 'mean', 'median', 'min', 'max'])
        
        if st.sidebar.button("Create Pivot Table"):
            if columns and index and agg_functions:
                pivot_table = pd.pivot_table(df, index=index, values=columns, aggfunc={col: agg_functions for col in columns})
                st.write("Pivot Table:")
                st.write(pivot_table)
            else:
                st.warning("Please select at least one column, one index column, and one aggregation function.")

if __name__ == "__main__":
    main()
