# text_extract_tools
A GUI to paste text in, which can extract certain patterns and put them in a template excel file.

Only works with very specific texts at the moment.
See Sample.txt for an example of this.

# Installing dependencies

### Windows

Press Win+R, type: "powershell" and press Enter.
In Powershell check if you have Python 3.10 installed.
`python3 --version` (can also start with `py` instead of `python3`)
If version is >= 3.10 continue at next paragraph.
Otherwise install a python version >= 3.10 from https://www.python.org/downloads/
Make sure you check the checkbox signifying to add Python to Path.

### OS Indifferent steps
- Download ZIP by pressing Code and selecting Download as ZIP
- Unzip the file
- Open your terminal (Powershell etc...)
- Navigate to unzipped project folder. (`cd [directoryname] or [full path]`, example: `cd ~/Downloads/text_extract_tools`)
- Run the following command: `python3 -m pip install -r requirements.txt`
- Done

# How to use
- Open the extract.py file
- For a demo paste sample.txt inside opened program
- Press Extract
- Inspect the two new excel sheets in the folder (Warning: Sheets will be overwritten if not renamed or moved to another folder)
- Optional: If you don\'t want the terminal to pop up when opening the program then rename extract.py to extract.pyw

- If permissions are blocked on *windows* a workaround could be:
- Open Powershell
- Navigate to folder (use `cd [directory]` and `ls` commands)
- Run: `py extractor.py`, `python3 extract.py` or with a 'w' added to the extension.
