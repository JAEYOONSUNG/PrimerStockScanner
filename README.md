# PrimerStockScanner
## Primer Matching Tool for SnapGene DNA Files
This program is designed to scan a laboratory primer stock list (Excel format) and identify primers that match DNA sequences from SnapGene `.dna` files. It simplifies the process of cross-referencing primer stock with DNA sequences, ensuring efficient primer selection for experimental workflows.

---

## Features
- **Excel Integration**: Reads a primer stock list formatted as `Primer stock list.xlsx`.
- **DNA Matching**: Matches primers to sequences in SnapGene `.dna` files, including reverse complements.
- **Tm Calculation**: Calculates melting temperatures (Tm) for each primer using the Nearest Neighbor method.
- **Customizable Input**: Allows user-defined file paths for both the primer stock list and DNA sequence files.
- **Excel Output**: Saves matching results into a new Excel file for easy review and documentation.

---

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/JAEYOONSUNG/PrimerStockScanner.git
   cd PrimerStockScanner
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Ensure the following tools/libraries are installed:
   - **Python 3.8+**
   - **Biopython**: For DNA sequence handling.
   - **Pandas**: For working with Excel files.
   - **openpyxl**: For Excel file output.

---

## Usage

1. **Prepare Your Excel File**:
   - Format your primer stock list as an Excel file (`.xlsx`).
   - Use the provided example file, `Primer stock list.xlsx`, as a reference for structure:
     - Column **"Sequence (5'-3')"**: Primer sequences.
     - Column **"Location"**: Primer locations in the lab.
     - Additional columns (optional): Primer labels or notes.

2. **Run the Program**:
   - Execute the script:
     ```bash
     python primer_matching_tool.py
     ```
   - Enter the file path to your `Primer stock list.xlsx` when prompted.
   - Enter the file path to your SnapGene `.dna` file when prompted.

3. **View Results**:
   - The program will generate an Excel file containing:
     - Primer sequences.
     - Start and end positions in the DNA sequence.
     - Matched strand (`+` or `-`).
     - Melting temperature (Tm) for each primer.
   - Example output file name: `matched_primers.xlsx`.

---

## Example Workflow

1. Prepare your primer stock list:
   ```text
   Primer stock list.xlsx
   ├── Sequence (5'-3')   # Primer sequences (required)
   ├── Location           # Lab location (required)
   ├── Primer Label       # Optional primer labels
   ```

2. Run the script:
   ```bash
   python primer_matching_tool.py
   ```

3. Enter file paths when prompted:
   - **Excel file**: `/path/to/Primer stock list.xlsx`
   - **DNA file**: `/path/to/snapgene_file.dna`

4. Review the results in the output file:
   - **Output file**: `matched_primers.xlsx`

---

## Example Input and Output

### Input Example: Primer stock list (`Primer stock list.xlsx`)
| Sequence (5'-3') | Location | Primer Label |
|------------------|----------|--------------|
| ATGCGTACGTAG     | Lab A    | Primer 1     |
| TACGGTACGCTT     | Lab B    | Primer 2     |

### Output Example: Matched primers (`matched_primers.xlsx`)
| Primer Sequence | Start Position | End Position | Location | Sheet Name | Primer Label | Strand | Tm   |
|-----------------|----------------|--------------|----------|------------|--------------|-------|------|
| ATGCGTACGTAG    | 101            | 113          | Lab A    | Sheet1     | Primer 1     | +     | 62.5 |
| TACGGTACGCTT    | 401            | 413          | Lab B    | Sheet1     | Primer 2     | -     | 58.3 |

---

## Notes

- Ensure your Excel file follows the format outlined in the `Primer stock list.xlsx` example.
- DNA files must be in SnapGene `.dna` format for proper sequence extraction.
- The tool is case-insensitive and handles non-alphabetic characters in primer sequences automatically.

---

## License
This project is licensed under the MIT License. See the `LICENSE` file for details.
