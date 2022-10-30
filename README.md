# Starting

## with Python

Run `python main.py` to start the GUI. To start in the terminal, run `python calculator.py`.

## as an Executable

To compile into an executable with PyInstaller on Windows 10, run

`pyinstaller.exe --onedir --windowed --name "Grand Champion Calculator" .\src\main.py`

Once compiled, the application folder will be under `dist\Grand Champion Calculator\`. Run
`Grand Champion Calculator.exe` to start the GUI.

Note: copy `trip.ico` and `icons8-xlsx-160.png` to `dist\Grand Champion Calculator\` for the icons to display after
compiling.

# General Usage

Currently, the calculator only takes .xlsx files as input. In the future, other input filetypes might be added. Columns
"number", "handler", "call name", "class", and "score" are required for both competition types. The "obedience"
competition also requires the "group" column, and optionally takes a "champion" column. 

A settings file, `settings.json` is generated upon first usage. You can change how the calculator breaks ties by
changing the order of classes under the "class hierarchy" key.

## GUI Usage

### Required Fields

- Input File: the input .xlsx file. You can either specify a path and filename in the text field, or use the "Open"
button and select a file.
- Output File: the output file. By default, this is the same as the input file and will write the output to a new
worksheet. To write to a new file, check the "Write to a new file" checkbox in "Options".
- Competition: the type of competition.

### Options
- Break ties by class: break ties by class if possible, according to the class hierarchy specified in `settings.json`.
- Write to new file: writes the output to a new .xlsx file.

## Command Line Usage

`usage: calculator.py [-h] [-n] input_file output_file {obedience,rally}`

### Positional Arguments

- `input_file`: the input .xlsx file.
- `output_file`: the output file. If identical to `input_file`, writes to a new worksheet.
- `{obedience,rally}`: the type of competition.

### Optional Arguments

- `-h`, `--help`: show the help message and exit.
- `-n`, `--no_ties`: break ties by class if possible

# Version History

## v1.1

- Optimized find_feature_names function.
- Renamed some variables for clarity.
- Added more documentation.
- Fixed some typos.
- Added yaml.

## v1.0

- Initial release.