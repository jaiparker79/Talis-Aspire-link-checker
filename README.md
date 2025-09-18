# Talis Aspire link checking tool

This PowerShell script checks URLs and DOIs exported in CSV format from the
Talis Aspire Reading List system.  It was made by Jai Parker, Information
Access Librarian at the Queensland University of Technology with help from
Microsoft Copilot.  As per the [license](./LICENSE), caveat emptor.

To run this script:

1. Download a csv of items for checking from Talis. The CSV file name format
   must start with all_list_items_ for the script to pick them up.
2. Open the CSV file and do a cleanup of the DOIs in Column O. The script
   automatically prepends https://doi.org/ to anything starting with 10. in
   this field so all the duplicated DOIs, mostly delimited by a ; need to be
   removed first.
3. If you are running a search for links of cancelled Alma items add a file named cancelled.xlsx containing the direct export of the Alma Portfolios which have been deleted from the catalogue.

### Input

URLs are fed into the script via CSV files. Files must be named to match the
pattern `all_list_items_*.csv` (for example `all_list_items_2025.csv`). When
it runs, the script will list all the files in the current directory with
matching names and prompt you for which one to load.

The CSV itself must contain column headers, and the script will search for
columns names either:

* `Online Resource Web Address` -- for the exact URLs; or
* `DOI` -- for raw DOIs (these will have `"https://doi.org/"` prepended
   to them if necessary)

and

* `Item Link` -- the link in your readings that points to the URL

### Errors

The script contains a hard-coded list of URL patterns that can be flagged as
broken. This can and should be customised for your specific needs.

Otherwise, for every URL in the list, the script will flag it as broken if:

* the webserver returns a 300 - 399 redirect response AND the URL redirected to is a domain. This test picks up where a deep link to a page or file redirects to the homepage of an organisation
* the webserver returns a "400 Bad Request" response
* the webserver returns a "404 Not Found" response
* the webserver returns a 500 - 599 range server error response.
* the hostname in the URL cannot be resolved (i.e. DNS error)
* a connection timeout occurs (by default this is 90 seconds)
* the connection terminates incorrectly

Other error values (e.g. "401 Not authorised", "403 Forbidden") are _not_ flagged by this script.

### Output

In addition to the CSV menu, the script will output each URL it is checking,
and the result.

Finally, it will produce a CSV file of the broken links, with `broken-links-`
prepended to the filename. For example `broken-links-all_list_items_2025.csv`

This report file contains two columns, "Item Link" for the readings link in
question, and "HTTP Error Code" which is a brief description of what was
wrong with it.
