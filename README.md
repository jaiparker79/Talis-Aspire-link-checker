# Talis Aspire link checking tool

This PowerShell script iterates over a list of URLs and checks for any that
are identified "broken" in various ways.

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

* the hostname in the URL cannot be resolved (i.e. DNS error)
* a connection timeout occurs (by default this is 30 seconds)
* the connection terminates incorrectly, or
* the webserver returns a "404 Not Found" response

Other error values (e.g. "400 Bad Request", "403 Forbidden", "500 Internal
Service Error", etc.) are _not_ flagged by this script.

### Output

In addition to the CSV menu, the script will output each URL it is checking,
and the result.

Finally, it will produce a CSV file of the broken links, with `broken-links-`
prepended to the filename. For example `broken-links-all_list_items_2025.csv`

This report file contains two columns, "Item Link" for the readings link in
question, and "HTTP Error Code" which is a brief description of what was
wrong with it.

