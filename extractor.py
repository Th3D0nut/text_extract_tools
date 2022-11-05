import re
import tkinter as tk
from openpyxl import load_workbook


def extract(text_box):
    text = text_box.get("1.0", tk.END)
    # define patterns
    name_pattern = re.compile(r"(?P<last_name>[A-Z\u00C0-\u00DD][a-z\u00E0-\u00FF]+(?:\s?[A-Za-z\u00C0-\u00DD][a-z\u00E0-\u00FF]+)*), (?P<first_name>[A-Z\u00C0-\u00DD][a-z\u00E0-\u00FF]+(?:\s?[A-Z\u00C0-\u00DD][a-z\u00E0-\u00FF]+)*)")
    mail_pattern = re.compile(r"[0-9a-zA-Z.]+@[a-zA-Z.]+\.(?:nl|com|org)")
    username_pattern = re.compile(r"[a-z]+\d{3}")
    # match in text
    names = [
        f"{name.group('first_name')} {name.group('last_name')}"
        for name in name_pattern.finditer(text)
    ]
    emails = mail_pattern.findall(text)
    usernames = username_pattern.findall(text)
    return names, emails, usernames


def write_to_bulkformulier(names, emails, usernames):
    wb = load_workbook(filename='template_mutatie_medewerker.xlsx')
    sheet = wb.worksheets[0]
    for index, name in enumerate(names, 2):
        sheet.cell(index, 3).value = name
        sheet.cell(index, 2).value = emails[index - 2]
        sheet.cell(index, 1).value = usernames[index - 2]
    wb.save("test.xlsx")


def extract_write(text_box):
    names, emails, usernames = extract(text_box)
    write_to_bulkformulier(names, emails, usernames)


def main():
    window = tk.Tk()
    lbl_title = tk.Label(text="Extract contents of text box to excel")
    lbl_title.pack()
    txt_box = tk.Text()
    txt_box.pack()
    btn_extract = tk.Button(text="Extract", command=lambda: extract_write(txt_box))
    btn_extract.pack()
    window.mainloop()


if __name__ == "__main__":
    main()

