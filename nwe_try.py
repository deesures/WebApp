import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder
from io import BytesIO


# Step 1: Read the original Excel files (monthly and weekly)
def load_data():
    # Load the original monthly data (the file paths can be adjusted)
    monthly_data = pd.read_excel("nwe.xlsx")
    
    # Load the updated weekly data (the file paths can be adjusted)
    weekly_data = pd.read_excel("weekly.xlsx")
    
    return monthly_data, weekly_data

st.set_page_config(page_title="Grid Example", layout="wide")

# Step 2: Process the user inputs in the original report
def process_user_input(monthly_data):
    # Configure the editable grid for user input (only "User Input" column)
    gb = GridOptionsBuilder.from_dataframe(monthly_data)
    gb.configure_column("SKU", editable=False)  # Make 'Parameter' non-editable
    gb.configure_column("Rebalancing", editable=True)  # Make 'User Input' editable
    grid_options = gb.build()

    # Display the editable grid
    response = AgGrid(
        monthly_data,
        gridOptions=grid_options,
        fit_columns_on_grid_load=True,
        height=300,
    )

    # Return the updated monthly data (with user inputs)
    return pd.DataFrame(response["data"])

def calculate_weights(monthly_data):
    monthly_data['Rebalanced'] = monthly_data['Rebalancing'] + monthly_data['Neptune_Allocation']
    monthly_data['country_weights'] = monthly_data.groupby(['SKU','PL','Plant','Market_level','Month'])['Rebalanced'].transform(lambda x: x/x.sum()).fillna(0)

    return monthly_data

def updated_monthly(monthly_data):
    st.subheader("Updated monthly Report with User Inputs")
    st.dataframe(monthly_data)

# Step 3: Propagate user input from monthly data to weekly data
def propagate_user_input_to_weekly(monthly_data, weekly_data):
    # Ensure the user input is propagated across the weeks of the month
    weekly_data['country_weights'] = weekly_data['Country'].map(
        lambda sku: monthly_data[monthly_data['Country'] == sku]['country_weights'].values[0] 
        if not monthly_data[monthly_data['Country'] == sku].empty else None
    )
    weekly_data['Final_alloc'] = weekly_data['Hood'] * weekly_data['country_weights']
    return weekly_data

# Step 4: Display the updated weekly report with user inputs
def display_updated_weekly_report(weekly_data):
    # Display the updated weekly report
    st.subheader("Updated Weekly Report with User Inputs")
    st.dataframe(weekly_data)

    # Allow the user to download the updated weekly report as an Excel file
    st.subheader("Download Updated Weekly Report")

    # Save the DataFrame to a BytesIO buffer to avoid the "missing excel_writer" error
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        weekly_data.to_excel(writer, index=False, sheet_name="Weekly_Report")
    
    # Rewind the buffer to the beginning so Streamlit can read from it
    excel_buffer.seek(0)

    # Provide the file for download
    st.download_button(
        label="Download Updated Report",
        data=excel_buffer,
        file_name="updated_weekly_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# Step 5: Run the Streamlit app
def main():
    st.title("Monthly to Weekly Report Conversion")

    # Load the data (both monthly and weekly)
    monthly_data, weekly_data = load_data()

    # Show the monthly report and allow user input
    st.subheader("Original Monthly Report")
    updated_monthly_data = process_user_input(monthly_data)

    updated_monthly_data_w = calculate_weights(updated_monthly_data)
    updated_monthly(updated_monthly_data_w)

    # Once the user inputs are done, propagate the input to the weekly data
    updated_weekly_data = propagate_user_input_to_weekly(updated_monthly_data_w, weekly_data)

    # Display the updated weekly report
    display_updated_weekly_report(updated_weekly_data)

# Run the main function
if __name__ == "__main__":
    main()
