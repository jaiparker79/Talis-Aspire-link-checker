This PowerShell script checks URLs and DOIs exported in CSV format from the Talis Aspire Reading List system.  It was made by Jai Parker, Information Access Librarian at the Queensland University of Technology with help from Microsoft Copilot.  As per the GPL license, caveat emptor.

To run this script:
1. Download a csv of items for checking from Talis. The CSV file name format must start with all_list_items_ for the script to pick them up.
2. Open the CSV file and do a cleanup of the DOIs in Column O. The script automatically prepends https://doi.org/ to anything starting with 10. in this field so all the duplicated DOIs, mostly delimited by a ; need to be removed first. 
