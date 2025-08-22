# staff-folder-mover


Moves staff folders (top-level only) from multiple source directories to a destination (“quarantine/archive”) folder by matching folder names to an input list of Surname, Given, and Preferred names. Generates an Excel log with a Matches sheet (what was moved or would be moved) and a Not Found sheet (input rows that didn’t match any folder). Supports dry-run.
Why this exists

When records are spread across several staff folders (terminated, contractors, interns, etc.), it’s tedious and risky to manually consolidate or quarantine them. This script automates the move with robust name matching and a full audit trail.
How it works (high level)

1.	Read input list from .csv, .xlsx, or .xls (e.g., exported HR list).

2.	Resolve columns for Surname / Given / Preferred via fuzzy header matching (e.g., surname, last name, preferred_name, etc.).

3.	Generate name variants per person (e.g., Given Surname, Surname, Given, Preferred Surname, plus versions without accents or punctuation).

4.	Index top-level folders across all configured source roots (no recursion).

5.	Match folder names to name variants (case-insensitive; accent/punctuation tolerant).

6.	Move or simulate each matched folder to the destination, ensuring a unique destination name if a collision occurs (Folder, Folder_1, …).

7.	Log outcomes to an Excel workbook with Matches and Not Found sheets.
Features

•	✅ Dry-run mode – see exactly what would be moved before making changes.

•	✅ Robust matching – tolerates punctuation differences (Smith, John vs Smith John) and accents (José vs Jose).

•	✅ Duplicate-safe – auto-suffixes destination names to avoid overwriting.

•	✅ Two-sheet Excel log – auditable output for both matches and misses.

•	✅ Flexible headers – finds Surname/Given/Preferred columns by common aliases.
Configuration

Edit the constants near the top of the script:

•	INPUT_PATH: Path to your input list (.csv, .xlsx, or .xls).

•	DRY_RUN: True to simulate; False to actually move folders.

•	FOLDERS_TO_SEARCH: List of source root paths (top-level folders under each root are considered).

•	DESTINATION_FOLDER: Where matched folders are moved.
Column name detection

The script will try to find headers using these candidate sets (case-insensitive; spaces/underscores ignored):

•	Surname: {"surname","last name","last_name","family name","family_name","lastname"}

•	Given: {"given names","given name","first names","first name","first_name","first","firstname"}

•	Preferred: {"preferred name","preferred","preferred_name","preferred given name","nickname"}

If your headers are unusual, rename the columns in your file or add new candidates in the script.
Input & Output

•	Input file (INPUT_PATH): A table with at least one of Surname, Given, or Preferred columns.

•	Source folders: Only top-level directories under each path in FOLDERS_TO_SEARCH are indexed.

•	Output log: An Excel file saved to DESTINATION_FOLDER named: 
 move_name_folders_log_YYYYMMDD_HHMMSS.xlsx
 
 Matches sheet columns:
• Matched Variants, Folder Name, Source Folder, Destination Folder, Action (MOVE or DRY-RUN), Result, Timestamp

• Not Found sheet:
		• Original input columns that existed (subset of Surname/Given/Preferred) for rows that matched nothing.
Safety & Limits

• Dry-run first. Set DRY_RUN = True to validate matching before moving
	• The script does not recurse into subfolders when indexing—only the first level under each source root.
	• Ensure you have permissions on source and destination paths.
	• Locked folders/files may fail to move (these will be logged as errors).
	• Moving large folders can take time; the script logs progress to the console.
 
Usage
1. Edit the config block at the top of the script:
• Set INPUT_PATH
• Set FOLDERS_TO_SEARCH
• Set DESTINATION_FOLDER
• Start with DRY_RUN = True
 
2.	Run: "move_name_folders.py"

3.	Review the console output and the Excel log in the destination folder.

	4.	If correct, set DRY_RUN = False and run again to move folders.
Matching details

For each person, the script generates candidates from any of the available fields:

•	Standalone: Preferred, Given, Surname

•	Combos: Given Surname, Surname Given, Surname, Given, and the same for Preferred + Surname

•	Normalizations for each candidate:

•	Lowercased

•	Accents removed (e.g., José → Jose)

•	Common punctuation removed (,.-’'()`)

•	Collapsed whitespace

Each top-level folder name is normalized the same way and compared against all variants.
Troubleshooting

•	“Could not read input …” – Check the path/extension and that the file isn’t open/locked.

•	“Could not find surname/given/preferred columns.” – Rename headers or extend candidate sets.

•	Nothing moved – Verify that folder names resemble your name variants; try dry-run and inspect the Matched Variants column.

•	Permission denied – Run with sufficient privileges or choose a destination you can write to.
Example

•	Input row: Surname = "González", Given = "Joaquín", Preferred = ""

•	Generated variants (examples): joaquín gonzález, gonzález joaquín, gonzalez joaquin, gonzalez, joaquin

	•	A folder named Gonzalez, Joaquin will match and be moved.
 
