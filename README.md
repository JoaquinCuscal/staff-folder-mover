# staff-folder-mover


Moves staff folders (top-level only) from multiple source directories to a quarantine/archive location by matching folder names to an input list of Surname/Given/Preferred names, with dry-run support and an Excel log of matches and misses.

⸻
README.md (paste into your repo)
Name-based Staff Folder Mover

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
 
staff-folder-mover
 
