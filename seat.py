import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from tkinter import font as tkfont
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment


class SeatingAllocationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automatic Seating Allocation System")
        self.root.geometry("1200x800")
        self.root.configure(bg="#2B2B2B", bd=10, relief="solid", highlightbackground="white", highlightthickness=2)

        self.college_details = {}
        self.rooms = []
        self.uploaded_files = []
        self.student_data = []
        self.seating_plan = None

        # Defining fonts
        self.title_font = tkfont.Font(family="Segoe UI", size=18, weight="bold")
        self.label_font = tkfont.Font(family="Helvetica", size=12, weight="normal")
        self.button_font = tkfont.Font(family="Helvetica", size=14, weight="bold")

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill="both", padx=20, pady=20)

        self.create_college_details_tab()
        self.create_room_details_tab()
        self.create_excel_upload_tab()
        self.create_generate_seating_tab()
        self.create_find_your_room_tab()
        self.create_export_tab()

        # Tab styling
        style = ttk.Style()
        style.configure("TNotebook", background="#2B2B2B", padding=10)
        style.configure("TNotebook.Tab", font=("Segoe UI", 16, "bold"), padding=[12, 6], background="#2B2B2B", foreground="black")  # Change foreground to black
        style.map("TNotebook.Tab", background=[("selected", "#B3B3B3"), ("active", "#D4D4D4")])
        style.configure("TFrame", background="#2B2B2B")
        style.configure("TLabel", background="#2B2B2B", font=self.label_font, foreground="#FFFFFF")
        style.configure("TButton", font=self.button_font, padding=10, background="#B3B3B3", foreground="#2B2B2B")
        style.map("TButton", background=[("active", "#D4D4D4"), ("pressed", "#B3B3B3")])
        style.configure("TEntry", font=self.label_font, fieldbackground="#D4D4D4", foreground="#2B2B2B")

    def create_college_details_tab(self):
        self.college_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.college_tab, text="College Details")

        ttk.Label(self.college_tab, text="College Name:").grid(row=0, column=0, padx=20, pady=10, sticky="e")
        self.college_name_entry = ttk.Entry(self.college_tab, width=40, font=self.label_font)
        self.college_name_entry.grid(row=0, column=1, padx=20, pady=10)

        ttk.Label(self.college_tab, text="Exam Type:").grid(row=1, column=0, padx=20, pady=10, sticky="e")
        self.exam_type_combo = ttk.Combobox(self.college_tab, values=["Mid1", "Mid2", "Semester"], state="readonly", width=37)
        self.exam_type_combo.grid(row=1, column=1, padx=20, pady=10)

        ttk.Label(self.college_tab, text="Exam Date (DD/MM/YYYY):").grid(row=2, column=0, padx=20, pady=10, sticky="e")
        self.exam_date_entry = ttk.Entry(self.college_tab, width=40)
        self.exam_date_entry.grid(row=2, column=1, padx=20, pady=10)

        ttk.Label(self.college_tab, text="Exam Time (From-To):").grid(row=3, column=0, padx=20, pady=10, sticky="e")
        self.exam_time_entry = ttk.Entry(self.college_tab, width=40)
        self.exam_time_entry.grid(row=3, column=1, padx=20, pady=10)

        ttk.Button(self.college_tab, text="Submit", command=self.save_college_details).grid(row=4, column=0, columnspan=2, pady=20)
        ttk.Button(self.college_tab, text="Next", command=lambda: self.notebook.select(self.room_tab)).grid(row=5, column=0, columnspan=2, pady=20)

    def save_college_details(self):
        college_name = self.college_name_entry.get().strip()
        exam_type = self.exam_type_combo.get()
        exam_date = self.exam_date_entry.get().strip()
        exam_time = self.exam_time_entry.get().strip()

        if not college_name or not exam_type or not exam_date or not exam_time:
            messagebox.showerror("Error", "All fields are required.")
            return

        try:
            datetime.strptime(exam_date, "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Error", "Invalid date format. Use DD/MM/YYYY.")
            return

        self.college_details = {
            "College Name": college_name,
            "Exam Type": exam_type,
            "Exam Date": exam_date,
            "Exam Time": exam_time,
        }
        self.display_data_in_tab(self.college_tab, [self.college_details])
        messagebox.showinfo("Info", "College details saved successfully!")

    def create_room_details_tab(self):
        self.room_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.room_tab, text="Room Details")

        # Static labels and input fields
        ttk.Label(self.room_tab, text="Room Name:").grid(row=0, column=0, padx=20, pady=10, sticky="e")
        self.room_name_entry = ttk.Entry(self.room_tab, width=30)
        self.room_name_entry.grid(row=0, column=1, padx=20, pady=10)

        ttk.Label(self.room_tab, text="Rows:").grid(row=1, column=0, padx=20, pady=10, sticky="e")
        self.row_entry = ttk.Entry(self.room_tab, width=30)
        self.row_entry.grid(row=1, column=1, padx=20, pady=10)

        ttk.Label(self.room_tab, text="Columns:").grid(row=2, column=0, padx=20, pady=10, sticky="e")
        self.column_entry = ttk.Entry(self.room_tab, width=30)
        self.column_entry.grid(row=2, column=1, padx=20, pady=10)

        ttk.Button(self.room_tab, text="Add Room", command=self.add_room).grid(row=3, column=0, columnspan=2, pady=20)
        ttk.Button(self.room_tab, text="Delete Room", command=self.delete_room).grid(row=4, column=0, columnspan=2,pady=20)
        ttk.Button(self.room_tab, text="Next", command=lambda: self.notebook.select(self.upload_tab)).grid(row=5, column=0,columnspan=2,pady=20)
    def add_room(self):
        room_name = self.room_name_entry.get().strip()
        rows = self.row_entry.get().strip()
        columns = self.column_entry.get().strip()

        if not room_name or not rows or not columns:
            messagebox.showerror("Error", "All fields are required.")
            return

        try:
            rows = int(rows)
            columns = int(columns)
            if rows <= 0 or columns <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Rows and columns must be positive integers.")
            return

        room = {
            "Room Name": room_name,
            "Rows": rows,
            "Columns": columns
        }
        self.rooms.append(room)
        self.display_data_in_tab(self.room_tab, self.rooms)
        self.clear_room_entries()

    def delete_room(self):
        if self.rooms:
            self.rooms.pop()
            self.display_data_in_tab(self.room_tab, self.rooms)

    def clear_room_entries(self):
        self.room_name_entry.delete(0, tk.END)
        self.row_entry.delete(0, tk.END)
        self.column_entry.delete(0, tk.END)

    def create_excel_upload_tab(self):
        self.upload_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.upload_tab, text="Excel Upload")

        # Static labels and input fields
        ttk.Label(self.upload_tab, text="Branch:").grid(row=0, column=0, padx=20, pady=10, sticky="e")
        self.branch_entry = ttk.Entry(self.upload_tab, width=30)
        self.branch_entry.grid(row=0, column=1, padx=20, pady=10)

        ttk.Label(self.upload_tab, text="Year:").grid(row=1, column=0, padx=20, pady=10, sticky="e")
        self.year_entry = ttk.Entry(self.upload_tab, width=30)
        self.year_entry.grid(row=1, column=1, padx=20, pady=10)

        self.branch_type = tk.StringVar(value="Odd")
        ttk.Radiobutton(self.upload_tab, text="Odd", variable=self.branch_type, value="Odd").grid(row=2, column=0,
                                                                                                  pady=10)
        ttk.Radiobutton(self.upload_tab, text="Even", variable=self.branch_type, value="Even").grid(row=2, column=1,
                                                                                                    pady=10)

        ttk.Button(self.upload_tab, text="Upload File", command=self.upload_file).grid(row=3, column=0, columnspan=2,
                                                                                       pady=20)
        ttk.Button(self.upload_tab, text="Next", command=lambda: self.notebook.select(self.generate_tab)).grid(row=4,column=0,columnspan=2,pady=20)
    def upload_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            try:
                data = pd.read_excel(file_path)
                if "StudentPIN" not in data.columns:
                    messagebox.showerror("Error", "The file must contain a 'StudentPIN' column.")
                    return
                self.student_data.extend(data.to_dict(orient="records"))
                self.uploaded_files.append({
                    "Branch": self.branch_entry.get().strip(),
                    "Year": self.year_entry.get().strip(),
                    "Branch Type": self.branch_type.get(),
                    "File Path": file_path
                })
                self.display_data_in_tab(self.upload_tab, self.uploaded_files)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read the file: {str(e)}")

    def create_generate_seating_tab(self):
        self.generate_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.generate_tab, text="Generate Seating")

        ttk.Button(self.generate_tab, text="Generate Seating Plan", command=self.generate_seating_plan).grid(pady=20)
        ttk.Button(self.generate_tab, text="Back", command=lambda: self.notebook.select(self.upload_tab)).grid(pady=10)

    def generate_seating_plan(self):
        if not self.rooms or not self.uploaded_files:
            messagebox.showerror("Error", "Rooms and student data are required.")
            return

        # Separate students by branch type (Odd/Even)
        odd_students, even_students = [], []
        for file_info in self.uploaded_files:
            data = pd.read_excel(file_info["File Path"])
            pins = data["StudentPIN"].tolist()
            if file_info["Branch Type"] == "Odd":
                odd_students.extend(pins)
            else:
                even_students.extend(pins)

        seating_plan = []
        odd_idx = even_idx = 0

        for room in self.rooms:
            rows, cols = room["Rows"], room["Columns"]
            seats = [["Empty" for _ in range(cols)] for _ in range(rows)]

            # Allocate odd columns to odd-branch students
            odd_cols = [col for col in range(cols) if (col + 1) % 2 != 0]
            for col in odd_cols:
                for row in range(rows):
                    if odd_idx < len(odd_students):
                        seats[row][col] = odd_students[odd_idx]
                        odd_idx += 1

            # Allocate even columns to even-branch students
            even_cols = [col for col in range(cols) if (col + 1) % 2 == 0]
            for col in even_cols:
                for row in range(rows):
                    if even_idx < len(even_students):
                        seats[row][col] = even_students[even_idx]
                        even_idx += 1

            seating_plan.append({"Room Name": room["Room Name"], "Seats": seats})

        # Dynamic adjustment for final branches to minimize empty spaces
        if odd_idx < len(odd_students) or even_idx < len(even_students):
            for room in seating_plan:
                seats = room["Seats"]
                for row in range(len(seats)):
                    for col in range(len(seats[row])):
                        if seats[row][col] == "Empty":
                            if (col + 1) % 2 != 0 and odd_idx < len(odd_students):
                                seats[row][col] = odd_students[odd_idx]
                                odd_idx += 1
                            elif (col + 1) % 2 == 0 and even_idx < len(even_students):
                                seats[row][col] = even_students[even_idx]
                                even_idx += 1

        # Handle remaining students if any
        if odd_idx < len(odd_students) or even_idx < len(even_students):
            messagebox.showwarning("Warning", "Not all students could be seated due to insufficient room capacity.")

        self.seating_plan = seating_plan
        self.show_seating_plan_popup()

    def create_find_your_room_tab(self):
        self.find_room_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.find_room_tab, text="Find Your Room")

        ttk.Label(self.find_room_tab, text="Enter Student PIN:").grid(row=1, column=0, padx=20, pady=10, sticky="e")
        self.student_pin_entry = ttk.Entry(self.find_room_tab, width=30)
        self.student_pin_entry.grid(row=1, column=1, padx=20, pady=10)

        ttk.Button(self.find_room_tab, text="Search", command=self.find_student_room).grid(row=2, column=0, columnspan=2, pady=20)

        self.summary_label = ttk.Label(self.find_room_tab, text="", font=self.label_font)
        self.summary_label.grid(row=3, column=0, columnspan=2, pady=10)

        self.result_label = ttk.Label(self.find_room_tab, text="", font=self.label_font)
        self.result_label.grid(row=4, column=0, columnspan=2, pady=10)

    def update_branch_summary(self, event=None):
        branch = self.branch_combo.get()
        if not branch or not self.seating_plan:
            return

        branch_students = {}
        for room in self.seating_plan:
            for row in room["Seats"]:
                for pin in row:
                    if isinstance(pin, int) or (isinstance(pin, str) and pin.isdigit()):
                        if pin in self.student_data:
                            student_info = next((s for s in self.student_data if s["StudentPIN"] == pin), None)
                            if student_info and student_info.get("Branch") == branch:
                                if room["Room Name"] not in branch_students:
                                    branch_students[room["Room Name"]] = []
                                branch_students[room["Room Name"]].append(pin)

        summary_text = f"Branch: {branch}\n"
        for room, pins in branch_students.items():
            summary_text += f"{room}: {min(pins)}-{max(pins)}\n"
        self.summary_label.config(text=summary_text)

    def find_student_room(self):
        pin = self.student_pin_entry.get().strip()
        if not pin:
            messagebox.showerror("Error", "Please enter a Student PIN.")
            return

        for room in self.seating_plan:
            for row_idx, row in enumerate(room["Seats"]):
                for col_idx, seat in enumerate(row):
                    if str(seat) == pin:
                        self.result_label.config(text=f"Room: {room['Room Name']}, Row: {row_idx + 1}, Column: {col_idx + 1}")
                        return

        self.result_label.config(text="Student PIN not found.")

    def show_seating_plan_popup(self):
        seating_popup = tk.Toplevel(self.root)
        seating_popup.title("Seating Plan")
        seating_popup.geometry("800x600")
        seating_popup.configure(bg="#2B2B2B")

        canvas = tk.Canvas(seating_popup, bg="#2B2B2B")
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(seating_popup, orient="vertical", command=canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        canvas.config(yscrollcommand=scrollbar.set)

        frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=frame, anchor="nw")

        college_info = (
            f"College Name: {self.college_details['College Name']}\n"
            f"Exam Type: {self.college_details['Exam Type']}\n"
            f"Exam Date: {self.college_details['Exam Date']}\n"
            f"Exam Time: {self.college_details['Exam Time']}\n\n"
        )
        ttk.Label(frame, text=college_info, font=self.title_font, foreground="#FFFFFF").pack(pady=10)

        total_seats = 0
        for room in self.seating_plan:
            room_label = ttk.Label(frame, text=f"Room: {room['Room Name']}", font=self.label_font, foreground="#FFFFFF")
            room_label.pack(pady=10)

            seats = room["Seats"]
            total_seats += len(seats) * len(seats[0])

            # Create a table-like structure for seating arrangement
            table_frame = ttk.Frame(frame)
            table_frame.pack(pady=10)

            for row_idx, row in enumerate(seats):
                for col_idx, seat in enumerate(row):
                    seat_label = ttk.Label(table_frame, text=str(seat), width=12, relief="solid", padding=5, background="#D4D4D4", foreground="#2B2B2B")
                    seat_label.grid(row=row_idx, column=col_idx, padx=5, pady=5)

            # Add labels for total allotted, present, absent, and invigilator signature
            total_allotted = sum(1 for row in seats for seat in row if seat != "Empty")
            total_present = 0  # Placeholder for actual attendance tracking
            total_absent = 0

            ttk.Label(frame, text=f"Total Allotted in {room['Room Name']}: {total_allotted}", font=self.label_font, foreground="#FFFFFF").pack(pady=5)
            ttk.Label(frame, text=f"Total Present: {total_present}", font=self.label_font, foreground="#FFFFFF").pack(pady=5)
            ttk.Label(frame, text=f"Total Absent: {total_absent}", font=self.label_font, foreground="#FFFFFF").pack(pady=5)
            ttk.Label(frame, text="Signature of the Invigilator: ___________________", font=self.label_font, foreground="#FFFFFF").pack(pady=10)

        ttk.Label(frame, text=f"Total Seats: {total_seats}", font=self.label_font, foreground="#FFFFFF").pack(pady=20)

        # Update the canvas scroll region
        frame.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))

    def create_export_tab(self):
        self.export_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.export_tab, text="Export Seating Plan")

        ttk.Button(self.export_tab, text="Export to Excel", command=self.export_to_excel).grid(pady=20)
        ttk.Button(self.export_tab, text="Back", command=lambda: self.notebook.select(self.generate_tab)).grid(pady=10)

    def export_to_excel(self):
        if not self.seating_plan:
            messagebox.showerror("Error", "No seating plan to export.")
            return

        try:
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if file_path:
                wb = Workbook()
                for room in self.seating_plan:
                    ws = wb.create_sheet(title=room["Room Name"])
                    ws.append(["College Name", self.college_details["College Name"]])
                    ws.append(["Exam Type", self.college_details["Exam Type"]])
                    ws.append(["Exam Date", self.college_details["Exam Date"]])
                    ws.append(["Exam Time", self.college_details["Exam Time"]])
                    ws.append([])
                    ws.append(["Room Name", room["Room Name"]])
                    ws.append([])

                    # Write seating plan
                    for row_idx, row in enumerate(room["Seats"]):
                        ws.append([f"Row {row_idx + 1}"] + row)

                    # Add labels
                    total_allotted = sum(1 for row in room["Seats"] for seat in row if seat != "Empty")
                    total_present = 0  # Placeholder for actual attendance tracking
                    total_absent = 0

                    ws.append([])
                    ws.append(["Total Allotted", total_allotted])
                    ws.append(["Total Present", total_present])
                    ws.append(["Total Absent", total_absent])
                    ws.append(["Signature of the Invigilator", "___________________"])

                # Remove the default sheet created by openpyxl
                wb.remove(wb["Sheet"])
                wb.save(file_path)
                messagebox.showinfo("Success", "Seating plan exported successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export: {str(e)}")

    def display_data_in_tab(self, tab, data):
        # Destroy only the dynamic data labels (those added after the input fields)
        for widget in tab.winfo_children():
            if isinstance(widget, ttk.Label) and widget["text"] != "" and widget.grid_info()["row"] >= 6:
                widget.destroy()

        if not data:
            return

        headers = list(data[0].keys())
        for col, header in enumerate(headers):
            ttk.Label(tab, text=header, font=self.label_font, relief="solid", padding=5, background="#D4D4D4",
                      foreground="#2B2B2B").grid(row=6, column=col, padx=5, pady=5)

        for row_idx, item in enumerate(data):
            for col_idx, key in enumerate(headers):
                ttk.Label(tab, text=str(item[key]), font=self.label_font, relief="solid", padding=5,
                          background="#D4D4D4", foreground="#2B2B2B").grid(
                    row=7 + row_idx, column=col_idx, padx=5, pady=5)
if __name__ == "__main__":
    root = tk.Tk()
    app = SeatingAllocationApp(root)
    root.mainloop()