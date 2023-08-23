import pandas as pd
import os

# Replace these with your actual paths
input_folder = "."
output_file = "output_excel_file.xlsx"

def combine_text_to_excel(input_folder, output_file):
    combined_data = {}

    for filename in os.listdir(input_folder):
        if filename.endswith(".txt"):
            file_path = os.path.join(input_folder, filename)
            sheet_name = os.path.splitext(filename)[0]

            with open(file_path, 'r') as file:
                lines = file.readlines()

            combined_data[sheet_name] = lines

    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        for sheet_name, lines in combined_data.items():
            df = pd.DataFrame({'Text': lines})
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"Combined text files into {output_file}")

if __name__ == '__main__':
    combine_text_to_excel(input_folder, output_file)
