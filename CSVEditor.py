import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import csv
import pandas as pd

class RowEditor(tk.Toplevel):
    def __init__(self, parent, row_data, columns):
        super().__init__(parent)
        self.title("Edit Row")
        self.row_data = row_data
        self.columns = columns
        self.result = None
        self.delete_flag = False

        self.geometry("600x400")  # Set initial size
        self.minsize(300, 200)    # Set minimum size

        self.create_widgets()

    def create_widgets(self):
        # Create a canvas with scrollbar
        self.canvas = tk.Canvas(self)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Pack canvas and scrollbar
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Create widgets inside the scrollable frame
        for i, (column, value) in enumerate(zip(self.columns, self.row_data)):
            label = ttk.Label(self.scrollable_frame, text=column)
            label.grid(row=i, column=0, padx=5, pady=5, sticky='ne')

            text = tk.Text(self.scrollable_frame, wrap=tk.WORD, width=40, height=3)
            text.insert(tk.END, value)
            text.grid(row=i, column=1, padx=5, pady=5, sticky='nsew')
            setattr(self, f'text_{i}', text)

        # Configure grid to make text widgets expandable
        self.scrollable_frame.grid_columnconfigure(1, weight=1)

        # Add Save, Delete, and Cancel buttons
        button_frame = ttk.Frame(self.scrollable_frame)
        button_frame.grid(row=len(self.columns), column=0, columnspan=2, pady=10)

        save_button = ttk.Button(button_frame, text="Save", command=self.save)
        save_button.pack(side=tk.LEFT, padx=5)

        delete_button = ttk.Button(button_frame, text="Delete", command=self.delete)
        delete_button.pack(side=tk.LEFT, padx=5)

        cancel_button = ttk.Button(button_frame, text="Cancel", command=self.cancel)
        cancel_button.pack(side=tk.LEFT, padx=5)

    def save(self):
        self.result = [getattr(self, f'text_{i}').get("1.0", tk.END).strip() for i in range(len(self.columns))]
        self.destroy()

    def delete(self):
        if messagebox.askyesno("Confirm Deletion", "Are you sure you want to delete this row?"):
            self.delete_flag = True
            self.destroy()

    def cancel(self):
        self.destroy()
    
class CSVEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV Editor")
        self.data = []
        self.sort_column = None
        self.sort_reverse = False

        # Create menu
        self.create_menu()

        # Create table
        self.create_table()

        # Bind double-click event
        self.tree.bind('<Double-1>', self.on_row_double_click)

    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open", command=self.load_csv)
        file_menu.add_command(label="Save", command=self.save_csv)
        file_menu.add_command(label="Merge", command=self.merge_csv)

        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        edit_menu.add_command(label="Insert Row", command=self.insert_row)
        edit_menu.add_command(label="Paste Data", command=self.paste_data)

    def create_table(self):
        # Create a frame for the treeview and scrollbars
        frame = ttk.Frame(self.root)
        frame.pack(expand=True, fill='both')

        # Create the treeview with a dummy column for the index
        self.tree = ttk.Treeview(frame, show='headings')
        self.tree.pack(side='left', expand=True, fill='both')

        # Add vertical scrollbar
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=vsb.set)

        # Add horizontal scrollbar
        hsb = ttk.Scrollbar(self.root, orient="horizontal", command=self.tree.xview)
        hsb.pack(fill='x')
        self.tree.configure(xscrollcommand=hsb.set)

        # Bind the header click event
        self.tree.bind('<Button-1>', self.on_header_click)

    def on_header_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == 'heading':
            column = self.tree.identify_column(event.x)
            column_index = int(column[1:]) - 1  # Convert column identifier to index
            if column_index == 0:  # Index column, do not sort
                return
            column_name = self.tree['columns'][column_index]

            # Toggle sort order if clicking on the same column
            if self.sort_column == column_name:
                self.sort_reverse = not self.sort_reverse
            else:
                self.sort_reverse = False

            self.sort_column = column_name
            self.sort_data()

    def sort_data(self):
        # Convert data to DataFrame for easier sorting
        df = pd.DataFrame(self.data[1:], columns=self.data[0])

        # Sort the DataFrame
        df = df.sort_values(by=self.sort_column, ascending=not self.sort_reverse)

        # Update the data attribute
        self.data = [self.data[0]] + df.values.tolist()

        # Redisplay the sorted data
        self.display_data()

    def load_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if file_path:
            with open(file_path, 'r') as file:
                csv_reader = csv.reader(file)
                self.data = list(csv_reader)
            self.display_data()

    def display_data(self):
        # Clear the existing data
        self.tree.delete(*self.tree.get_children())
        
        # Configure the columns
        self.tree['columns'] = ['Index'] + self.data[0]
        self.tree.heading('Index', text='#')
        self.tree.column('Index', width=50, anchor='center')
        
        for col in self.data[0]:
            self.tree.heading(col, text=col, command=lambda c=col: self.on_header_click(c))
            self.tree.column(col, width=100)  # Adjust the width as needed

        # Insert the data rows
        for i, row in enumerate(self.data[1:], start=1):
            self.tree.insert('', 'end', values=[i] + row)

        # Highlight the sorted column
        if self.sort_column:
            self.tree.heading(self.sort_column, text=f"{self.sort_column} {'↑' if not self.sort_reverse else '↓'}")

    def insert_row(self):
        new_row = []
        for col in self.tree['columns']:
            value = tk.simpledialog.askstring("Input", f"Enter value for {col}")
            new_row.append(value if value else "")
        self.tree.insert('', 'end', values=new_row)
        self.data.append(new_row)

    def paste_data(self):
        clipboard = self.root.clipboard_get()
        rows = clipboard.split('\n')
        for row in rows:
            values = row.split('\t')
            if len(values) == len(self.tree['columns']):
                self.tree.insert('', 'end', values=values)
                self.data.append(values)

    def merge_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if file_path:
            # Read the current data and the new CSV file
            current_df = pd.DataFrame(self.data[1:], columns=self.data[0])
            new_df = pd.read_csv(file_path)

            # Check if the columns match
            if list(current_df.columns) != list(new_df.columns):
                messagebox.showerror("Error", "The columns of the two CSV files do not match.")
                return

            # Function to convert numeric-like columns to numeric
            def convert_numeric(df):
                for col in df.columns:
                    if df[col].dtype == 'object':
                        try:
                            df[col] = pd.to_numeric(df[col], errors='raise')
                        except ValueError:
                            pass  # Keep as object if conversion fails
                return df

            # Convert numeric-like columns in both DataFrames
            current_df = convert_numeric(current_df)
            new_df = convert_numeric(new_df)

            # Concatenate the DataFrames
            merged_df = pd.concat([current_df, new_df], ignore_index=True)

            # Remove duplicate rows, keeping the first occurrence
            merged_df.drop_duplicates(keep='first', inplace=True)

            # Update the data and display
            self.data = [list(merged_df.columns)] + merged_df.values.tolist()
            self.display_data()

            # Show info about the merge
            skipped_rows = len(new_df) - (len(merged_df) - len(current_df))
            messagebox.showinfo("Merge Complete", f"Merged successfully. Skipped {skipped_rows} duplicate row(s).")

    def save_csv(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        if file_path:
            with open(file_path, 'w', newline='') as file:
                csv_writer = csv.writer(file)
                csv_writer.writerows(self.data)  # Write only the actual data, excluding the index

    def on_row_double_click(self, event):
        item = self.tree.identify('item', event.x, event.y)
        if item:
            # Get the values of the clicked row
            values = self.tree.item(item, 'values')
            
            # Open the RowEditor
            editor = RowEditor(self.root, values[1:], self.data[0])  # Exclude the index column
            self.root.wait_window(editor)
            
            if editor.delete_flag:
                # Delete the row
                index = int(values[0]) - 1  # Get the actual index in self.data
                del self.data[index + 1]  # +1 because self.data[0] is headers
                self.tree.delete(item)
                self.update_row_numbers()
            elif editor.result:
                # Update the data
                index = int(values[0]) - 1  # Get the actual index in self.data
                self.data[index + 1] = editor.result  # +1 because self.data[0] is headers
                
                # Update the treeview
                self.tree.item(item, values=[values[0]] + editor.result)

    def update_row_numbers(self):
        # Update the row numbers in the treeview
        for i, item in enumerate(self.tree.get_children(), start=1):
            values = self.tree.item(item, 'values')
            self.tree.item(item, values=[i] + values[1:])


if __name__ == "__main__":
    root = tk.Tk()
    app = CSVEditor(root)
    root.mainloop()
