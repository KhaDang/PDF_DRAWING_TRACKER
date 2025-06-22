# Drawing Revision Checker

This Python script automates the comparison between PDF drawing filenames (containing revision letters) and a Bill of Materials (BOM) stored in an Excel file. It is useful for tracking drawing revisions and identifying discrepancies between what's in your folder and what's documented in your BOM.

## 📂 What It Does

- Scans a folder for `.pdf` drawing files.
- Extracts the **drawing number** and **revision letter** from each filename.
- Reads a BOM Excel file that contains drawing numbers and their corresponding revisions.
- Compares the revisions from the PDF filenames and the Excel BOM.
- Categorizes the results into:
  - ✅ Drawings with matching revisions.
  - ⚠️ Drawings with different revisions.
  - 📁 Drawings found only in the folder.
  - 📄 Drawings found only in the Excel file.
- Exports all results into a multi-sheet Excel file.

## 🧪 Example Filename Format

The script expects drawing filenames like:  
- ABC123A.pdf
- XYZ789B.pdf

Where:
- `ABC123` is the drawing number
- `A` is the revision

## 🛠 How It Works

### 1. Folder Scanning
Scans the folder:
folder_path = r'D:\path\to\your\folder'
### 2. Excel BOM
Loads the Excel file:
existing_excel = r'D:\path\to\your\BOM.xlsx'
### 3. Output
Generates an output Excel file:
comparison_with_drawing_revision.xlsx
With these sheets:
- Same Revision
- Different Revision
- Only in Folder
- Only in Excel

📋 Requirements
Install required Python packages:
- pip install pandas openpyxl

🚀 How to Run
- python main.py

Make sure to adjust the folder_path and existing_excel paths in the script to your local directories before running.

📦 Output Example
An Excel file like this:

Sheet: Different Revision

Drawing Number	Folder Revision	Excel Revision
ABC123	A	B

💡 Use Cases
- Engineering change tracking
- Revision consistency check between documents and BOM
- Pre-manufacturing quality control

📄 License
MIT License ()

👤 Author
Kha Dang


Feel free to submit issues or enhancements via pull requests.


