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
            font-size: 16px;  /* Adjust the font size as needed */
            font-family: 'Times New Roman', Times, serif; /* Use Times New Roman font */
        }
        .stMarkdown a {
            font-size: 16px;  /* Adjust the font size for links as needed */
        }
        .css-1v3fvcr {
            font-size: 16px;  /* Adjust the font size for Streamlit data frame text as needed */
        }
        </style>
        """,
        unsafe_allow_html=True
    )
    
    # Slider example
    num = st.slider("Choose a number", 1, 10, key="slider")
    st.write(st.session_state)

    # Title and markdown examples
    st.title("Excel File Reader")
    st.markdown("*Streamlit* is **really** ***cool***.")
    st.markdown(':red[Streamlit] :orange[can] :green[write] :blue[text] :violet[in] :gray[pretty] :rainbow[colors].')
    st.markdown("Here's a bouquet &mdash; :tulip::cherry_blossom::rose::hibiscus::sunflower::blossom:")

    # Multi-line markdown example
    multi = '''
        ### Hello welcome 
        Today we will talk about 
        2023-12-18
        Expanded access to AI coding has arrived in Colab across 175 locales for all tiers of Colab users
        Improvements to display of ML-based inline completions (for eligible Pro/Pro+ users)
        Started a series of notebooks highlighting Gemini API capabilities
        Enable ⌘/Ctrl+L to select the full line in an editor
        Fixed bug where we weren't correctly formatting output from multiple execution results
        * List item
        ```
        # This is formatted as code
        ```
        * List item
    '''
    st.markdown(multi)

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

        # Option to choose a column for describing
        describe_column_option = st.selectbox("Select a column for describe", list(df[sheet_option].columns))

        # Describe the selected column
        st.write("Description of Selected Column:")
        st.write(df[sheet_option][describe_column_option].describe())

        # Option to choose a column for value counts
        value_counts_column_option = st.selectbox("Select a column for value counts", list(df[sheet_option].columns))

        # Calculate value counts for the selected column
        value_counts = df[sheet_option][value_counts_column_option].value_counts()

        # Calculate percentage for each row
        total_rows = len(df[sheet_option])
        percentage = (value_counts / total_rows) * 100

        # Create a DataFrame with value counts and percentage
        value_counts_df = pd.DataFrame({'Value': value_counts.index, 'Count': value_counts.values, 'Percentage': percentage.values})

        # Add column names as the first row
        columns_row = pd.DataFrame({'Value': [value_counts_column_option,]
                                     })
        
        # Concatenate column names row to the DataFrame
        value_counts_df = pd.concat([columns_row, value_counts_df])

        # Add a row for total sum
        total_count = value_counts_df['Count'].sum()
        total_percentage = value_counts_df['Percentage'].sum()
        total_row = pd.DataFrame({'Value': ['Total'], 'Count': [total_count], 'Percentage': [total_percentage]})

        # Concatenate total row to the DataFrame
        value_counts_df = pd.concat([value_counts_df, total_row])

        # Display value counts and percentage for the selected column
        st.write("Value Counts and Percentage of Selected Column:")
        st.write(value_counts_df)

        # Add a download button for the selected sheet as Excel
        if st.button("Download as Excel"):
            # Prepare the Excel file as binary data
            excel_binary = io.BytesIO()
            with pd.ExcelWriter(excel_binary, engine='openpyxl', mode='w') as writer:
                value_counts_df.to_excel(writer, sheet_name=sheet_option, index=False)

            # Trigger the download of the Excel file
            st.download_button(label="Download Now", data=excel_binary, file_name=f"{uploaded_file.name}_{sheet_option}_output.xlsx", key="download_button")

# Call the main function
if __name__ == "__main__":
    main()
