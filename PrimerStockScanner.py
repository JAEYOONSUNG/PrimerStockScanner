import os
import re
import pandas as pd
from Bio.SeqUtils import MeltingTemp as mt
from Bio.Seq import Seq
from your_module import snapgene_file_to_seqrecord  # Replace with the correct import for SnapGene parser

# Prompt the user for the Excel file path
excel_file_path = input("Enter the path to the Excel file (e.g., /Users/JaeYoon/Documents/Lab/Primer stock list.xlsx): ")
default_excel_path = '/Users/JaeYoon/Documents/Lab/Primer stock list.xlsx'

# Use the default path if no input is provided
if not excel_file_path:
    excel_file_path = default_excel_path

# Read all sheets from the Excel file into a dictionary of DataFrames
excel_sheets = pd.read_excel(excel_file_path, sheet_name=None)

# Initialize a list to store DataFrames for each sheet
all_data = []

# Iterate through each sheet, adding the sheet name as a new column, and store the DataFrames in a list
for sheet_name, df in excel_sheets.items():
    df['Sheet Name'] = sheet_name  # Add the sheet name as a new column
    all_data.append(df)

# Combine all DataFrames into a single DataFrame
combined_df = pd.concat(all_data, ignore_index=True)

# Remove rows where the 'Sequence (5\'-3\')' column is empty
combined_df.dropna(subset=['Sequence (5\'-3\')'], inplace=True)

# Clean up the 'Sequence (5\'-3\')' column: remove non-alphabetic characters, remove spaces, and convert to uppercase
combined_df['Sequence (5\'-3\')'] = combined_df['Sequence (5\'-3\')'].apply(lambda x: re.sub(r'[^A-Za-z]', '', str(x)).upper())

# Prompt the user for the SnapGene file path
snapgene_file_path = input("Enter the path to the SnapGene file: ")

# Read the DNA sequence from the SnapGene file
seq_record = snapgene_file_to_seqrecord(snapgene_file_path)
dna_sequence = str(seq_record.seq).upper()

# Generate the reverse complement of the DNA sequence
reverse_complement_sequence = str(seq_record.seq.reverse_complement()).upper()

# Find matches and their positions in the sequence
results = []
sequence_length = len(dna_sequence)
for index, row in combined_df.iterrows():
    primer_seq = row['Sequence (5\'-3\')']
    try:
        tm_value = mt.Tm_NN(Seq(primer_seq))
    except ValueError as e:
        print(f"Error calculating Tm for {primer_seq}: {e}")
        tm_value = None
    if primer_seq in dna_sequence:
        start_pos = dna_sequence.index(primer_seq)
        end_pos = start_pos + len(primer_seq) - 1
        results.append((primer_seq, start_pos + 1, end_pos + 1, row['Location'], row['Sheet Name'], row.get('Primer Label', ''), '+', tm_value))
    elif primer_seq in reverse_complement_sequence:
        start_pos = reverse_complement_sequence.index(primer_seq)
        end_pos = start_pos + len(primer_seq) - 1
        original_start_pos = sequence_length - end_pos - 1
        original_end_pos = sequence_length - start_pos - 1
        results.append((primer_seq, original_start_pos + 1, original_end_pos + 1, row['Location'], row['Sheet Name'], row.get('Primer Label', ''), '-', tm_value))

# Create a DataFrame for the results
results_df = pd.DataFrame(results, columns=['Primer Sequence', 'Start Position', 'End Position', 'Location', 'Sheet Name', 'Primer Label', 'Strand', 'Tm'])

# Display the results
print(results_df)

# Prompt the user for the output Excel file path
output_file_path = input("Enter the path to save the results as an Excel file (e.g., ~/Desktop/matched_primers.xlsx): ")

# Use a default path if no input is provided
if not output_file_path:
    output_file_path = '~/Desktop/matched_primers.xlsx'

# Ensure the file has the correct extension
if not output_file_path.endswith('.xlsx'):
    output_file_path += '.xlsx'

# Convert '~' to the full path
output_file_path = os.path.expanduser(output_file_path)

# Save the results to an Excel file
results_df.to_excel(output_file_path, index=False, engine='openpyxl')

print(f"File successfully saved: {output_file_path}")