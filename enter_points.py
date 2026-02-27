import csv
import os
import time
import argparse
import customtkinter as ctk
from tkinter import messagebox, ttk



class ExamPointsApp(ctk.CTk):
    def __init__(self, args):
        super().__init__()
        self.title("Exam Points Entry")

        # Theme
        ctk.set_appearance_mode("dark")   # "light" or "dark"
        ctk.set_default_color_theme("blue")  # also: "green", "dark-blue"

        # Config
        self.delimiter = args.delimiter
        self.input_file = args.file
        self.output_file = args.output
        # self.dv_str = args.dv_str
        self.num_exercises = args.n
        self.sound_enabled = not args.no_sound
        self.rwth_strings = {
            'Matrikelnummer': 'REGISTRATION_NUMBER',
            'Vorname': 'FIRST_NAME_OF_STUDENT',
            'Name': 'FAMILY_NAME_OF_STUDENT',
            'Zulassung': 'Zulassung',
            'Drittversuch': 'GUEL_U_AKTUELLE_ANTRITTE_SPO'
        }

        # base_height = 400     # enough for ~3 exercises
        # extra_per_ex = 40     # per additional exercise
        # height = base_height + (self.num_exercises * extra_per_ex)
        # width = 1000 + (self.num_exercises * 60)
        # self.geometry(f"{width}x{height}")
        self.attributes("-fullscreen", True)
        self.is_fullscreen = True

        self.bind("<F11>", self.toggle_fullscreen)

        self.exercise_entries = []
        self.update_mode = False   # default: adding new rows
        self.editing_matric_number = None  # None = insert mode, otherwise update mode

        self._initialize_output_file()
        self._build_ui()
        self._load_csv_into_table()


    # -----------------------------
    # Initialization
    # -----------------------------
    def _initialize_output_file(self):
        if not os.path.exists(self.output_file):
            with open(self.output_file, "w", newline="") as file:
                writer = csv.writer(file, delimiter=self.delimiter)
                writer.writerow(
                    ["Matrikelnummer"]
                    + [f"Exercise {i+1}" for i in range(self.num_exercises)]
                    + ["Total Points"]
                )

        with open(self.output_file, 'r') as file:
            reader = csv.DictReader(file, delimiter=self.delimiter)
            if reader.fieldnames != ['Matrikelnummer'] + [f'Exercise {i+1}' for i in range(self.num_exercises)] + ['Total Points']:
                raise ValueError(f"Output file {self.output_file} has incorrect format. Either correct the column names or delete the file to start fresh.")


    # -----------------------------
    # UI Building
    # -----------------------------
    def _build_ui(self):
        # Layout: left input panel | right table
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)


        # Outer frame
        input_frame = ctk.CTkFrame(main_frame, width=300, corner_radius=15)
        input_frame.pack(side="left", fill="y", padx=(0, 20), pady=10)


        # Inner frame with padding for margins
        input_inner = ctk.CTkFrame(input_frame, fg_color="transparent")
        input_inner.pack(fill="both", expand=True, padx=15, pady=15)

        # Matriculation
        ctk.CTkLabel(input_inner, text="Matriculation Number:").pack(pady=(15, 5), anchor="w")
        self.matriculation_entry = ctk.CTkEntry(input_inner, width=200)
        self.matriculation_entry.pack(pady=5)
        self.matriculation_entry.bind("<Tab>", self.lookup_student_data)
        self.matriculation_entry.focus_set()

        self.name_label = ctk.CTkLabel(input_inner, text="", text_color="lightblue")
        self.name_label.pack(pady=(5, 15))

        # Exercises
        for i in range(self.num_exercises):
            # Create a small frame for each row
            row_frame = ctk.CTkFrame(input_inner, fg_color="transparent")
            row_frame.pack(fill="x", pady=5)

            # Label
            ctk.CTkLabel(row_frame, text=f"Exercise {i+1}:", width=100, anchor="w").pack(side="left")

            # Entry
            entry = ctk.CTkEntry(row_frame)
            if i == self.num_exercises - 1:
                entry.bind("<Tab>", self.compute_sum)
            entry.pack(side="left", fill="x", expand=True)
            self.exercise_entries.append(entry)
            self._disable_entries()

        # Result
        ctk.CTkLabel(input_inner, text="Total Points:").pack(pady=(15, 0))
        self.points_label = ctk.CTkLabel(input_inner, text="", text_color="green", font=("Arial", 16, "bold"))
        self.points_label.pack()

        # Submit button
        self.submit_button = ctk.CTkButton(input_inner, text="Submit", command=self.submit, state="disabled", fg_color="gray")
        self.submit_button.pack(pady=20)

        # Quit button
        quit_button = ctk.CTkButton(
            input_inner,
            text="Exit",
            command=self.quit,   # or self.destroy
            fg_color="red",
            hover_color="darkred"
        )
        quit_button.pack(side="left", anchor="sw", padx=10, pady=10)


        self.bind('<Return>', lambda event: self.submit_button.invoke())
        self.bind('<KP_Enter>', lambda event: self.submit_button.invoke())
        self.bind('<Escape>', lambda event: self._reset_fields())

        # CSV Table
        table_frame = ctk.CTkFrame(main_frame, corner_radius=15)
        table_frame.pack(side="left", fill="both", expand=True, pady=10)

        # Treeview + Scrollbar
        vsb = ttk.Scrollbar(table_frame, orient="vertical")
        vsb.pack(side="right", fill="y")

        self.tree = ttk.Treeview(table_frame, show="headings", yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=15, pady=15)

        vsb.config(command=self.tree.yview)

        self.bind("<Delete>", self._handle_delete_row)
        self.tree.bind("<Double-1>", self._on_row_double_click)



    def toggle_fullscreen(self, event=None):
        self.is_fullscreen = not self.is_fullscreen
        self.attributes("-fullscreen", self.is_fullscreen)


    def _on_row_double_click(self, event):
        item = self.tree.identify_row(event.y)
        if not item:
            return
        values = self.tree.item(item, "values")
        if not values:
            return

        # Same as lookup — enter edit mode
        matric_number = values[0]
        self._reset_fields()
        self.matriculation_entry.insert(0, matric_number)
        self.editing_matric_number = matric_number
        self.name_label.configure(text="Editing existing entry", text_color="yellow")

        self._enable_entries()
        for i, entry in enumerate(self.exercise_entries):
            entry.insert(0, values[i+1])
        self.points_label.configure(text=values[-1])
        self._enable_submit_button()

    # -----------------------------
    # Data Logic
    # -----------------------------
    def _load_csv_into_table(self):
        for col in self.tree.get_children():
            self.tree.delete(col)

        with open(self.output_file, "r") as file:
            reader = csv.reader(file, delimiter=self.delimiter)
            headers = next(reader)

            # Setup columns dynamically
            self.tree["columns"] = headers
            for col in headers:
                self.tree.heading(col, text=col, command=lambda c=col: self._sort_by_column(c, False))
                self.tree.column(col, width=100, anchor="center")

            for row in reader:
                self.tree.insert("", "end", values=row)

    def _sort_by_column(self, col, reverse):
        """Sort Treeview rows by given column."""
        # Get all items and their values for the column
        items = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]

        # Try numeric sort, fallback to string
        try:
            items.sort(key=lambda t: float(t[0]), reverse=reverse)
        except ValueError:
            items.sort(key=lambda t: t[0], reverse=reverse)

        # Reorder rows
        for index, (_, k) in enumerate(items):
            self.tree.move(k, '', index)

        # Toggle sorting order next time
        self.tree.heading(col, command=lambda: self._sort_by_column(col, not reverse))


    def submit(self):
        matric_number = self.matriculation_entry.get().strip()
        points = [e.get().strip() or "0" for e in self.exercise_entries]
        total = sum(float(p) for p in points)

        with open(self.output_file, "r") as f:
            rows = list(csv.reader(f, delimiter=self.delimiter))
            header, data = rows[0], rows[1:]

        if self.editing_matric_number:  # Update existing
            new_data = []
            for row in data:
                if row[0] == self.editing_matric_number:
                    new_data.append([matric_number] + points + [f"{total:.1f}"])
                else:
                    new_data.append(row)
            data = new_data
        else:  # Insert new
            data.append([matric_number] + points + [f"{total:.1f}"])

        with open(self.output_file, "w", newline="") as f:
            writer = csv.writer(f, delimiter=self.delimiter)
            writer.writerow(header)
            writer.writerows(data)

        self._load_csv_into_table()
        self._reset_fields()


    def lookup_student_data(self, event=None):
        matric_number = self.matriculation_entry.get().strip()

        with open(self.output_file, "r") as file:
            reader = csv.DictReader(file, delimiter=self.delimiter)
            for row in reader:
                if row["Matrikelnummer"] == matric_number:
                    self.editing_matric_number = matric_number
                    self.name_label.configure(text="Editing existing entry", text_color="yellow")

                    # Highlight row in Treeview
                    for item in self.tree.get_children():
                        values = self.tree.item(item, "values")
                        if values and values[0] == matric_number:
                            self.tree.selection_set(item)
                            self.tree.see(item)

                            # Fill inputs
                            self._enable_entries()
                            for i, entry in enumerate(self.exercise_entries):
                                entry.insert(0, values[i+1])  # exercises start at col 1
                            self.points_label.configure(text=values[-1])
                            self._enable_submit_button()
                            return

        # Search in input file
        with open(self.input_file, "r") as file:
            reader = csv.DictReader(file, delimiter=self.delimiter)
            for row in reader:
                if row[self.rwth_strings["Matrikelnummer"]] == matric_number:
                    name = f"{row[self.rwth_strings['Vorname']]} {row[self.rwth_strings['Name']]}"
                    if row[self.rwth_strings["Drittversuch"]] == self.dv_str:
                        name += " (Drittversuch)"
                        if self.sound_enabled:
                            print("\a")
                    self.name_label.configure(text=f"Name: {name}", text_color="lightblue")
                    self._enable_entries()
                    return

        self.name_label.configure(text="Not found", text_color="orange")
        self._disable_entries()


    def compute_sum(self, event=None):
        points = [e.get().strip() or "0" for e in self.exercise_entries]
        total = sum(float(p) for p in points)
        self.points_label.configure(text=f"{total:.1f}")
        self._enable_submit_button()
        
    # -----------------------------
    # Helpers
    # -----------------------------
    def _reset_fields(self):
        self.matriculation_entry.delete(0, "end")
        for e in self.exercise_entries:
            e.delete(0, "end")
        self._disable_entries()
        self.points_label.configure(text="")
        self._disable_submit_button()
        self.matriculation_entry.focus_set()
        self.update_mode = False
        self.editing_matric_number = None
        self.name_label.configure(text="", text_color="white")

    def _disable_entries(self):
        for e in self.exercise_entries:
            e.configure(state=ctk.DISABLED, fg_color="gray")

    def _disable_submit_button(self):
        self.submit_button.configure(state="disabled", fg_color="gray")

    def _enable_entries(self):
        for e in self.exercise_entries:
            e.configure(state=ctk.NORMAL, fg_color="#FFFFFF", text_color="#000000")

    def _enable_submit_button(self):
        self.submit_button.configure(state="normal", fg_color="#4CAF50")  # green

    # -----------------------------
    # CSV editing
    # -----------------------------
    def _handle_delete_row(self, event=None):
        """Delete selected row from CSV with confirmation."""
        selection = self.tree.selection()
        if not selection:
            return  # nothing selected

        item = selection[0]
        row_values = self.tree.item(item, "values")

        # Confirm deletion
        confirm = messagebox.askyesno(
            "Delete Row",
            f"Are you sure you want to delete:\n\n{row_values} ?"
        )
        if not confirm:
            return

        # Read all rows, filter out the one to delete
        with open(self.output_file, "r") as f:
            reader = list(csv.reader(f, delimiter=self.delimiter))
            header, rows = reader[0], reader[1:]

        # Remove the first matching row
        rows = [r for r in rows if r != list(row_values)]

        # Rewrite file
        with open(self.output_file, "w", newline="") as f:
            writer = csv.writer(f, delimiter=self.delimiter)
            writer.writerow(header)
            writer.writerows(rows)

        # Refresh table
        self._load_csv_into_table()


# -----------------------------
# Main Entry
# -----------------------------
if __name__ == "__main__":
    parser = argparse.ArgumentParser("python enter_points.py")
    parser.add_argument("file", help="Input CSV file with student data", type=str)
    parser.add_argument("n", type=int, help="Number of exercises")
    parser.add_argument("-d", dest="delimiter", default=";", help="CSV delimiter, default=';'")
    parser.add_argument("-output", dest="output", default="results.csv", help="Output CSV file")
    # parser.add_argument("-dv_str", dest="dv_str", default="3", help="Drittversuch string")
    parser.add_argument("-no_sound", dest="no_sound", action="store_true", help="Disable sound alerts")
    args = parser.parse_args()

    app = ExamPointsApp(args)
    app.mainloop()
