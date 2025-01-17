import tkinter as tk
from tkinter import ttk
import os
from tkinter import filedialog
import csv

class NCGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("NC File Generator")
        
        # List to store all rows
        self.rows = []
        
        # Create main frame
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Top frame for process number and import
        self.top_frame = ttk.LabelFrame(self.main_frame, text="Controls", padding="5")
        self.top_frame.grid(row=0, column=0, columnspan=5, pady=5, sticky=(tk.W, tk.E))
        
        # Process number
        ttk.Label(self.top_frame, text="Process Number:").grid(row=0, column=0, padx=5)
        self.process_number = ttk.Entry(self.top_frame, width=10)
        self.process_number.grid(row=0, column=1, padx=5)
        
        # Import CSV button
        self.import_button = ttk.Button(self.top_frame, text="Import CSV", command=self.import_csv)
        self.import_button.grid(row=0, column=2, padx=20)
        
        # Add tooltip/help text
        ttk.Label(self.top_frame, text="CSV Format: G Code, Distance, Feed Rate, Comment").grid(row=1, column=0, columnspan=3, pady=5)
        
        # Middle frame for G-code input
        self.input_frame = ttk.LabelFrame(self.main_frame, text="G-Code Input", padding="5")
        self.input_frame.grid(row=1, column=0, columnspan=5, pady=5, sticky=(tk.W, tk.E))
        
        # Column labels
        ttk.Label(self.input_frame, text="G Code").grid(row=0, column=0, padx=5)
        ttk.Label(self.input_frame, text="Distance").grid(row=0, column=1, padx=5)
        ttk.Label(self.input_frame, text="Feed Rate").grid(row=0, column=2, padx=5)
        ttk.Label(self.input_frame, text="Comment").grid(row=0, column=3, padx=5)
        
        # Row control buttons at bottom of input frame
        self.row_control_frame = ttk.Frame(self.input_frame)
        self.row_control_frame.grid(row=999, column=0, columnspan=4, pady=5)  # Using high row number to ensure it stays at bottom
        
        ttk.Label(self.row_control_frame, text="Row:").grid(row=0, column=0, padx=5)
        self.add_button = ttk.Button(self.row_control_frame, text="+", width=3, command=self.add_row)
        self.add_button.grid(row=0, column=1, padx=2)
        
        self.subtract_button = ttk.Button(self.row_control_frame, text="-", width=3, command=self.subtract_row)
        self.subtract_button.grid(row=0, column=2, padx=2)
        
        # Bottom frame for generate controls
        self.generate_frame = ttk.LabelFrame(self.main_frame, text="Output", padding="5")
        self.generate_frame.grid(row=2, column=0, columnspan=5, pady=5, sticky=(tk.W, tk.E))
        
        # Folder selection
        ttk.Label(self.generate_frame, text="Output Folder:").grid(row=0, column=0, padx=5)
        self.folder_path = tk.StringVar()
        self.folder_entry = ttk.Entry(self.generate_frame, textvariable=self.folder_path, width=40)
        self.folder_entry.grid(row=0, column=1, padx=5)
        
        self.browse_button = ttk.Button(self.generate_frame, text="Browse", command=self.browse_folder)
        self.browse_button.grid(row=0, column=2, padx=5)
        
        # Generate button at bottom
        self.generate_button = ttk.Button(self.generate_frame, text="Generate", command=self.generate_file)
        self.generate_button.grid(row=1, column=0, columnspan=3, pady=10)
        
        # Define valid G-codes
        self.valid_codes = ["G00", "G01", "M00", "M01", "M30"]
        
        # Add initial row
        self.add_row()

    def add_row(self):
        # Remove M30 row temporarily
        if len(self.rows) > 0 and self.rows[-1][0].get() == "M30":
            m30_row = self.rows.pop()
            for widget in m30_row:
                widget.grid_remove()
        
        # Calculate new row position
        row_num = len(self.rows)
        
        # Create new row entries
        available_codes = self.valid_codes[:-1]  # Exclude M30
        
        g_code = ttk.Combobox(self.input_frame, width=7, values=available_codes, state="readonly")
        g_code.bind('<<ComboboxSelected>>', lambda e, row=row_num: self.on_code_select(row))
        
        distance = ttk.Entry(self.input_frame, width=10)
        feed = ttk.Entry(self.input_frame, width=10)
        comment = ttk.Entry(self.input_frame, width=20)
        
        # Position the new row
        g_code.grid(row=row_num + 1, column=0, padx=5, pady=2)
        distance.grid(row=row_num + 1, column=1, padx=5, pady=2)
        feed.grid(row=row_num + 1, column=2, padx=5, pady=2)
        comment.grid(row=row_num + 1, column=3, padx=5, pady=2)
        
        # Add row to our list
        self.rows.append((g_code, distance, feed, comment))
        
        # Set default value and update field states
        g_code.set(self.valid_codes[0])
        self.on_code_select(row_num)
        
        # Add M30 row if this is the first row or re-add it at the end
        if len(self.rows) == 1 or 'M30' not in [row[0].get() for row in self.rows]:
            self.add_m30_row()
        else:
            # Re-add the M30 row that we removed
            row_num = len(self.rows)
            for i, widget in enumerate(m30_row):
                widget.grid(row=row_num + 1, column=i, padx=5, pady=2)
            self.rows.append(m30_row)

    def add_m30_row(self):
        row_num = len(self.rows)
        
        # Create M30 row
        g_code = ttk.Combobox(self.input_frame, width=7, values=["M30"], state="readonly")
        g_code.bind('<<ComboboxSelected>>', lambda e, row=row_num: self.on_code_select(row))
        
        distance = ttk.Entry(self.input_frame, width=10)
        feed = ttk.Entry(self.input_frame, width=10)
        comment = ttk.Entry(self.input_frame, width=20)
        
        # Position the row
        g_code.grid(row=row_num + 1, column=0, padx=5, pady=2)
        distance.grid(row=row_num + 1, column=1, padx=5, pady=2)
        feed.grid(row=row_num + 1, column=2, padx=5, pady=2)
        comment.grid(row=row_num + 1, column=3, padx=5, pady=2)
        
        # Add row to our list
        self.rows.append((g_code, distance, feed, comment))
        
        # Set values and update field states
        g_code.set("M30")
        comment.insert(0, "End of Program")
        self.on_code_select(row_num)
        
        # Disable fields
        distance.configure(state='disabled')
        feed.configure(state='disabled')
        comment.configure(state='readonly')

    def on_code_select(self, row_num):
        g_code, distance, feed, comment = self.rows[row_num]
        selected_code = g_code.get()
        
        if selected_code.startswith('M'):
            # M codes - disable all except comment
            distance.configure(state='disabled')
            feed.configure(state='disabled')
            comment.configure(state='normal')
            # Clear disabled fields
            distance.delete(0, tk.END)
            feed.delete(0, tk.END)
        else:
            # G codes
            distance.configure(state='normal')
            comment.configure(state='normal')
            
            if selected_code == 'G00':
                # Rapid movement - disable feed
                feed.configure(state='disabled')
                feed.delete(0, tk.END)
            elif selected_code == 'G01':
                # Feed movement - enable feed
                feed.configure(state='normal')

    def subtract_row(self):
        if len(self.rows) > 2:  # Keep M30 and at least one other row
            # Don't allow removing the M30 row
            if self.rows[-1][0].get() == "M30":
                # Remove the second-to-last row instead
                for widget in self.rows[-2]:
                    widget.destroy()
                self.rows.pop(-2)
            else:
                # Remove the last row
                for widget in self.rows[-1]:
                    widget.destroy()
                self.rows.pop()
        else:
            tk.messagebox.showwarning("Warning", "Cannot remove more rows. Minimum is one row plus M30.")

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.set(folder_selected)
            
    def generate_file(self):
        process_num = self.process_number.get()
        output_folder = self.folder_path.get()
        
        # Validate process number
        if not process_num:
            tk.messagebox.showerror("Error", "Please enter a process number!")
            return
        if not process_num.isdigit():
            tk.messagebox.showerror("Error", "Process number must contain only digits!")
            return
            
        # Validate output folder
        if not output_folder:
            tk.messagebox.showerror("Error", "Please select an output folder!")
            return
        if not os.path.exists(output_folder):
            tk.messagebox.showerror("Error", "Selected output folder does not exist!")
            return
            
        nc_content = []
        unwind_content = []
        has_m30 = False
        
        # Generate forward and reverse content
        for row_index, row in enumerate(self.rows, 1):
            g_code = row[0].get()
            distance = row[1].get()
            feed = row[2].get()
            comment = row[3].get()
            
            # Validate required fields based on G/M code
            if not g_code.startswith('M'):  # G codes require distance
                if not distance:
                    tk.messagebox.showerror("Error", f"Distance is required for {g_code} in row {row_index}!")
                    return
                try:
                    float(distance)  # Validate distance is a number
                except ValueError:
                    tk.messagebox.showerror("Error", f"Invalid distance value in row {row_index}. Must be a number!")
                    return
                    
                if g_code == "G01" and not feed:  # G01 requires feed rate
                    tk.messagebox.showerror("Error", f"Feed rate is required for G01 in row {row_index}!")
                    return
                if feed:
                    try:
                        float(feed)  # Validate feed is a number
                    except ValueError:
                        tk.messagebox.showerror("Error", f"Invalid feed rate in row {row_index}. Must be a number!")
                        return
            
            # Create NC line
            line = g_code
            if not g_code.startswith('M'):  # Only add X movement for G codes
                line += f" X{distance}"
                if feed and g_code == "G01":
                    line += f" F{feed}"
            if comment:
                line += f" ({comment})"
            nc_content.append(line)
            
            # Handle unwind content
            if g_code == "M30":
                has_m30 = True
                continue  # Skip M30 for now, we'll add it at the end
            
            # Create unwind line (inverse distance) only for G codes
            if not g_code.startswith('M'):
                if distance:
                    try:
                        inverse_distance = str(-float(distance))
                        unwind_line = f"{g_code} X{inverse_distance}"
                        if feed and g_code == "G01":
                            unwind_line += f" F{feed}"
                        if comment:
                            unwind_line += f" (UNWIND {comment})"
                        unwind_content.insert(0, unwind_line)
                    except ValueError:
                        tk.messagebox.showerror("Error", f"Invalid distance value in row {row_index}!")
                        return
            elif g_code != "M30":  # Add M codes (except M30) to unwind
                unwind_content.insert(0, line)
        
        # Always add M30 at the end of unwind content
        unwind_content.append("M30 (END OF PROGRAM)")
        
        # Check if files already exist
        recipe_path = os.path.join(output_folder, f"{process_num}-recipe.nc")
        unwind_path = os.path.join(output_folder, f"{process_num}-unwind.nc")
        
        if os.path.exists(recipe_path) or os.path.exists(unwind_path):
            if not tk.messagebox.askyesno("Warning", 
                "One or both output files already exist. Do you want to overwrite them?"):
                return
        
        # Write to files
        try:
            with open(recipe_path, "w") as f:
                f.write("\n".join(nc_content))
                
            with open(unwind_path, "w") as f:
                f.write("\n".join(unwind_content))
            
            # Show success message
            tk.messagebox.showinfo("Success", "NC files generated successfully!")
        except Exception as e:
            tk.messagebox.showerror("Error", f"Failed to write files: {str(e)}")

    def import_csv(self):
        # Ask for file
        file_path = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
            
        try:
            # Clear existing rows except M30
            while len(self.rows) > 0:
                if self.rows[0][0].get() == "M30":
                    m30_row = self.rows.pop(0)
                    for widget in m30_row:
                        widget.grid_remove()
                    break
                for widget in self.rows[0]:
                    widget.destroy()
                self.rows.pop(0)
            
            # Read CSV file
            with open(file_path, 'r') as file:
                csv_reader = csv.reader(file)
                next(csv_reader, None)  # Skip header row
                
                for row in csv_reader:
                    if len(row) < 1:  # Skip empty rows
                        continue
                        
                    # Validate G code
                    g_code = row[0].strip().upper()
                    if g_code == "M30":
                        continue  # Skip M30 as we'll add it at the end
                    if g_code not in self.valid_codes:
                        raise ValueError(f"Invalid G code: {g_code}")
                    
                    # Add new row
                    row_num = len(self.rows)
                    g_code_combo = ttk.Combobox(self.input_frame, width=7, values=self.valid_codes[:-1], state="readonly")
                    distance = ttk.Entry(self.input_frame, width=10)
                    feed = ttk.Entry(self.input_frame, width=10)
                    comment = ttk.Entry(self.input_frame, width=20)
                    
                    # Position the new row
                    g_code_combo.grid(row=row_num + 1, column=0, padx=5, pady=2)
                    distance.grid(row=row_num + 1, column=1, padx=5, pady=2)
                    feed.grid(row=row_num + 1, column=2, padx=5, pady=2)
                    comment.grid(row=row_num + 1, column=3, padx=5, pady=2)
                    
                    # Set values
                    g_code_combo.set(g_code)
                    if len(row) > 1:  # Distance
                        distance.insert(0, row[1].strip())
                    if len(row) > 2:  # Feed Rate
                        feed.insert(0, row[2].strip())
                    if len(row) > 3:  # Comment
                        comment.insert(0, row[3].strip())
                    
                    # Add row to list
                    self.rows.append((g_code_combo, distance, feed, comment))
                    
                    # Update field states
                    self.on_code_select(row_num)
            
            # Add M30 row at the end
            self.add_m30_row()
            
            tk.messagebox.showinfo("Success", "CSV file imported successfully!")
            
        except Exception as e:
            tk.messagebox.showerror("Error", f"Failed to import CSV: {str(e)}")
            # Ensure at least one row exists
            if len(self.rows) == 0:
                self.add_row()

if __name__ == "__main__":
    root = tk.Tk()
    app = NCGenerator(root)
    root.mainloop() 