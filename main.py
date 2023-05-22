import pandas as pd
import numpy as np
import os

# Directory containing the input files
input_directory = "./sheets"

# Directory for the output files
output_directory = "./sheets-vobi"
os.makedirs(output_directory, exist_ok=True)  # create output directory if it does not exist

# Loop through all files in the input directory
for filename in os.listdir(input_directory):
    # Only process .xls files
    if filename.endswith(".xls"):
        # Full path to the input file
        input_filepath = os.path.join(input_directory, filename)

        # Load spreadsheet skipping the first 5 rows as they do not contain column names
        xls_file = pd.read_excel(input_filepath, skiprows=5)

        # Prepare a dictionary to apply the mapping
        mapping = {
            "Código Composição": xls_file['CODIGO DA COMPOSICAO'],
            "Código Item": xls_file['CODIGO ITEM'],
            "Tipo": np.where(xls_file['TIPO ITEM'].isnull() | xls_file['TIPO ITEM'].eq(''), 'Composição',
                            np.where(xls_file['TIPO ITEM'] == 'INSUMO', 'Produto', 
                                    np.where(xls_file['TIPO ITEM'] == 'COMPOSICAO', 'Serviço', xls_file['TIPO ITEM']))),
            "Nome": np.where(xls_file['TIPO ITEM'].isnull() | xls_file['TIPO ITEM'].eq(''), xls_file['DESCRICAO DA COMPOSICAO'], xls_file['DESCRIÇÃO ITEM']),
            "Quantidade": pd.to_numeric(xls_file['COEFICIENTE'], errors='coerce'), # convert to numeric
            "Un": np.where(xls_file['UNIDADE ITEM'].isnull() | xls_file['UNIDADE ITEM'].eq(''), xls_file['UNIDADE'], xls_file['UNIDADE ITEM']),
            "Custo unitário": xls_file['PRECO UNITARIO'],
            "Classe": np.where(xls_file['TIPO ITEM'].isnull() | xls_file['TIPO ITEM'].eq(''), xls_file['SIGLA DA CLASSE'], ''),
        }

        # Create a new dataframe from the mapping
        xlsx_file = pd.DataFrame(mapping)

        # Create an empty row
        empty_row = pd.DataFrame([['']*len(xlsx_file.columns)], columns=xlsx_file.columns)

        # Insert an empty row after each group of rows with the same 'Código Composição'
        indices = xlsx_file[xlsx_file['Código Composição'] != xlsx_file['Código Composição'].shift(-1)].index
        for index in reversed(indices):
            xlsx_file = pd.concat([xlsx_file.iloc[:index+1], empty_row, xlsx_file.iloc[index+1:]], ignore_index=True)

        # Full path to the output file
        output_filepath = os.path.join(output_directory, filename.replace(".xls", "-vobi.xlsx"))

        # Save it as .xlsx
        xlsx_file.to_excel(output_filepath, index=False)
