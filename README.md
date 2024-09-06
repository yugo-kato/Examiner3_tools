# Examiner3_tools
 Related tools for Examiner™software for 3D GPR data 

## csv_cut

CSV data cutter for CSV files exported by the Examiner3 "Geo5ExportTimeData" process.

- **Input File:** The program uses a file dialog to select an Excel file (`.xlsx`) as input.
It expects the Excel file to contain a table where each row defines a file name, a name , a start index, and an end index for filtering.

| filename | name | startX | endX |
| -------------- | -- | --- | --- |
| 2024-04-01-005 | Left | 0 | 400 |
| 2024-04-02-006 | Center | 0 | 403 |
| 2024-04-03-007 | Right  | 0 | 412 |

- **Filter Criteria:** The program selects rows in the input CSV files based on the following:
  - The value in the first column (`x`) must be between the startX and endX index.
  - The value in the third column (`z`) must be less than or equal to 93.

### Output:
The output CSV files will be named in the following format:

```
<filename>_X<startX>-<endX>.csv
```


## txt_to_latloncsv

Extract latitude and longitude from Examiner3 processed ASCII txt files.

- **Input File:** The program uses a file dialog to select an Excel file (`.xlsx`) as input.
This program uses an Excel file in the same format as csv_cut.


### CSV Output: The processed data is saved into a .csv file, where each row represents:

X (Scan ID)
Y (Channel No.)
Latitude
Longitude

