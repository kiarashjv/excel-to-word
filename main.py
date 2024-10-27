import pandas as pd
import win32com.client

# Read the Excel file
df = pd.read_excel("./persons.xlsx")

# Define the column names
column_names = ["نام ونام خانوادگی", "شخصی/سازمانی", "کد ثبت نام"]

# Calculate the number of PDFs to create
num_pdfs = len(df) // 4

# Start Word
word = win32com.client.Dispatch("Word.Application")
word.Visible = True
# Initialize the last non-empty value for each column
last_values = [""] * len(column_names)
# Create each PDF
for pdf_num in range(num_pdfs):
    # Open the Word document
    doc = word.Documents.Open(
        "C:\\Users\\Kiarash\\Desktop\\Project\\ExcelToWord\\template.docx"
    )
    # Initialize the current row and column
    current_row = pdf_num * 4
    current_column = 0

    # Initialize the row counter
    row_counter = 0

    # Iterate over all shapes
    for i in range(1, doc.Shapes.Count + 1):
        shape = doc.Shapes.Item(i)
        # Check if the shape is a textbox
        if shape.Type == 17:  # Use the value directly
            # Get the text from the current row and column
            text = df.at[current_row, column_names[current_column]]
            # If the text is NaN, use the last non-empty value
            if pd.isna(text):
                text = last_values[current_column]
            else:
                # Update the last non-empty value
                last_values[current_column] = text
            # Change the text
            shape.TextFrame.TextRange.Text = text
            # Move to the next column
            current_column += 1
            # If we've processed three columns, move to the next row
            if current_column == 3:
                current_row += 1
                current_column = 0
                row_counter += 1

    # Export as PDF
    output_file_name = (
        f"C:\\Users\\Kiarash\\Desktop\\Project\\ExcelToWord\\output\\{pdf_num}.pdf"
    )
    doc.ExportAsFixedFormat(
        output_file_name, 17
    )  # 17 is the value for wdExportFormatPDF

    # Close the document
    doc.Close(False)  # Don't save the changes to the original document

# We're not quitting Word at the end, so it will stay open
