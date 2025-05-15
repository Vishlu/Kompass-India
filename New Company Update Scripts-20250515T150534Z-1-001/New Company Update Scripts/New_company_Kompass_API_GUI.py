import tkinter as tk
from tkinter import filedialog
# Now you can import the script
from New_company_Kompass_API_script import 


from tkinter import messagebox

class NewCompany(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Database Adder (DA)")

        self.file_label = tk.Label(self, text="Select Excel File:",width=30)
        self.file_label.pack()

        self.file_entry = tk.Entry(self,width=60)
        self.file_entry.pack()

        self.browse_button = tk.Button(self, text="Browse", command=self.browse_file)
        self.browse_button.pack()

        self.process_button = tk.Button(self, text="Process", command=self.process_excel)
        self.process_button.pack()

        self.output_label = tk.Label(self, text="Output:")
        self.output_label.pack()

        self.output_text = tk.Text(self, height=10, width=50)
        self.output_text.pack()

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        self.file_entry.delete(0, tk.END)
        self.file_entry.insert(0, filename)

    def process_excel(self):
        excel_file = self.file_entry.get()

        if not excel_file:
            messagebox.showerror("Error", "No file selected. Please select an Excel file.")
        else:

            main(excel_file)  # Call main function with the selected Excel file path
            self.output_text.insert(tk.END, "Processing Complete.")

if __name__ == "__main__":
    app = NewCompany()
    app.mainloop()