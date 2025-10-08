from docx import Document

def populate_word(template_path, df, date, output_path):
    doc = Document(template_path)
    table = doc.tables[0]

    # Insert date in row 0, col 1
    table.rows[0].cells[1].text = date

    # Insert data starting from row 5
    data_start_row = 5
    for i, row in df.iterrows():
        row_index = data_start_row + i
        if row_index >= len(table.rows):
            table.add_row()
        table_row = table.rows[row_index]
        table_row.cells[0].text = str(row['first_name'])
        table_row.cells[1].text = str(row['last_name'])
        table_row.cells[2].text = str(row['date_of_birth'])
        table_row.cells[3].text = str(row['UN_first_name'])
        table_row.cells[4].text = str(row['UN_last_name'])
        table_row.cells[5].text = str(row['UN_date_of_birth'])
        table_row.cells[6].text = str(row['name_match_score'])

    doc.save(output_path)
