from tkinter import *
from tkinter import filedialog, messagebox
from datetime import datetime
import pandas as pd

window = Tk()
window.title("Upload File")
window.minsize(width=550, height=300)
window.configure(bg="#f0f8ff")

header_label = Label(window, text="Upload Your Excel File", font=("Arial", 16, "bold"), bg="#f0f8ff", fg="#333")
header_label.pack(pady=20)

def upload_file():
    file_path = filedialog.askopenfilename(
        title="Select a Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if file_path:
        df = pd.read_excel(file_path)
        duplicate_rows = df.duplicated().sum()
        copy_df = df.drop_duplicates()
        # To Hide the Previous Buttons
        # text_label.forget()
        header_label.forget()
        button.forget()
        success_label = Label(
            text="File Uploaded Successfully",
            fg="green"
        )
        success_label.pack()

        text_area = Text(window, height=10, width=50)
        text_area.pack()
        text_area.insert(END, f"Duplicate Rows Count: {duplicate_rows}\n\n")
        text_area.insert(END, f"Uploaded File:\n{copy_df}\n")
        save_button = Button(text="Save a Copy", command=lambda : download_file(copy_df))
        save_button.pack(pady=5)

    else:
        print("No Valid File Format Selected")


def download_file(df):
    today = datetime.now().strftime("%Y-%m-%d")
    output_file_name = f"Cleaned_data_{today}.xlsx"
    df.to_excel(output_file_name, index=False)
    messagebox.showinfo("Success",f"File Saved Successfully as {output_file_name}")



# label
# text_label = Label(text="Please upload a Excel format File", font=("Arial", 12), bg="#f0f8ff")
# text_label.pack(pady=10)
# button
button = Button(text="Upload File", command=upload_file,bg="#2196F3", fg="white")
button.pack(pady=10)

window.mainloop()
