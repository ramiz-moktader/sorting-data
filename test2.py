import streamlit as st
import pandas as pd
import io

# Define main function
def main():
    # Add custom CSS for styling
    st.markdown(
        """
        <style>
        body, .stSelectbox, .stTextInput, .stButton, .stMarkdown, .stSmall, .stText, .stError, .stHeader, .stImage, .stProgressBar {
            font-size: 16px; /* Adjust the font size as needed */
            font-family: 'Times New Roman', Times, serif; /* Use Times New Roman font */
        }
        .stMarkdown a {
            font-size: 16px; /* Adjust the font size for links as needed */
        }
        .css-1v3fvcr {
            font-size: 16px; /* Adjust the font size for Streamlit data frame text as needed */
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    # Title and markdown examples
    st.title("Excel File Reader")

    # Upload Excel file through Streamlit
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

    if uploaded_file is not None:
        # Use pandas to read the Excel file
        df = pd.read_excel(uploaded_file, engine='openpyxl', sheet_name=None)

        # Display sheet options
        sheet_option = st.selectbox("Select a sheet", list(df.keys()))

        # Display the selected sheet
        st.write("Selected Sheet:")
        st.write(df[sheet_option])

        # Option to choose multiple columns for describing in the first column
        st.markdown("""
    <style>
        .stMultiSelect [data-baseweb=select] span{
            max-width: 700px;
            font-size: 0.9rem;
        }
    </style>
    """, unsafe_allow_html=True)

        selected_columns = st.multiselect("Select columns for value counts", list(df[sheet_option].columns))
           
        if selected_columns:
            # Calculate value counts and percentage for each selected column
            main_value_counts_df = pd.DataFrame(columns=['Column', 'Value', 'Percentage', 'Count'])
            for col in selected_columns:
                value_counts = df[sheet_option][col].value_counts()
                total_rows = len(df[sheet_option])
                percentage = (value_counts / total_rows) * 100

                # Create a DataFrame for the current column
                col_value_counts_df = pd.DataFrame({'Value': value_counts.index, 'Percentage': percentage.values, 'Count': value_counts.values})

                # Add column name as the first row
                col_value_counts_df = pd.concat([pd.DataFrame({'Column': [col]}), col_value_counts_df], axis=1)

                # Add a row for total sum
                total_count = col_value_counts_df['Count'].sum()
                total_percentage = col_value_counts_df['Percentage'].sum()
                total_row = pd.DataFrame({'Column': ['Total'], 'Value': ['Total'], 'Percentage': [total_percentage], 'Count': [total_count]})

                # Concatenate total row to the current column's DataFrame
                col_value_counts_df = pd.concat([col_value_counts_df, total_row])

                # Append the current column's DataFrame to the main DataFrame
                main_value_counts_df = pd.concat([main_value_counts_df, col_value_counts_df])

                # Display value counts and percentage for the selected columns
                st.write(f"Value Counts and Percentage for '{col}':")
                st.write(col_value_counts_df)

        # Display value counts and percentage for all selected columns in a single DataFrame
        if not main_value_counts_df.empty:
            st.write("Value Counts and Percentage for All Selected Columns:")
            st.write(main_value_counts_df)

            # Add a download button for the main DataFrame as Excel
            if st.button("Download as Excel"):
                # Prepare the Excel file as binary data
                excel_binary = io.BytesIO()
                with pd.ExcelWriter(excel_binary, engine='openpyxl', mode='w') as writer:
                    main_value_counts_df.to_excel(writer, sheet_name=sheet_option, index=False)

                # Trigger the download of the Excel file
                st.download_button(label="Download Now", data=excel_binary, file_name=f"{uploaded_file.name}_{sheet_option}_output.xlsx", key="download_button")

# Call the main function
if __name__ == "__main__":
    main()
