import openpyxl
import tkinter as tk
from tkinter import filedialog, Text, Scrollbar

# Function to extract hyperlinks from the selected Excel file
def extract_hyperlinks():
    excel_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if excel_file:
        workbook = openpyxl.load_workbook(excel_file)
        sheet = workbook.active

        # Create a list to store the hyperlinks
        hyperlinks = []

        for row in sheet.iter_rows():
            for cell in row:
                if cell.hyperlink:
                    hyperlink = cell.hyperlink.target
                    hyperlinks.append(hyperlink)

        # Display the hyperlinks in the result text widget
        result_text.config(state=tk.NORMAL)
        result_text.delete(1.0, tk.END)
        for link in hyperlinks:
            result_text.insert(tk.END, link + "\n")
        result_text.config(state=tk.DISABLED)

# Function to exit the application
def exit_application():
    root.destroy()

# Create a tkinter window
root = tk.Tk()
root.title("Excel Hyperlink Extractor")

# Button to select an Excel file and extract hyperlinks
import_button = tk.Button(root, text="Import Excel", command=extract_hyperlinks)
import_button.pack()

# Text widget to display extracted hyperlinks
result_text = Text(root, wrap=tk.WORD, width=80, height=20)
result_text.pack()
result_text.config(state=tk.DISABLED)

# Scrollbar for the text widget
scrollbar = Scrollbar(root, command=result_text.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
result_text.config(yscrollcommand=scrollbar.set)

# Button to exit the application
exit_button = tk.Button(root, text="Exit", command=exit_application)
exit_button.pack()

root.mainloop()
