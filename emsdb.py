from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from ems import Database
import csv
from PIL import Image, ImageDraw, ImageFont
import os
from datetime import datetime

# Initialize Database
db = Database("Employee.db")

root = Tk()
root.title("Employee Management System")
root.geometry("1920x1080+0+0")
root.config(bg="#2c3e50")
root.state("zoomed")

# Variables
name = StringVar()
code = StringVar()
email = StringVar()
contact = StringVar()
gender = StringVar()
department = StringVar()

# Entries Frame
entries_frame = Frame(root, bg="#a8e6cf")
entries_frame.pack(side=TOP, fill=X)

# Title
title = Label(entries_frame, text="Employee Management System", font=("Calibri", 18, "bold"), bg="#535c68", fg="white")
title.grid(row=0, columnspan=4, padx=10, pady=20, sticky="w")

# Labels and Entry fields
Label(entries_frame, text="Name", font=("Calibri", 16), bg="#535c68", fg="white").grid(row=1, column=0, padx=10, pady=10, sticky="w")
txtName = Entry(entries_frame, textvariable=name, font=("Calibri", 16), width=30)
txtName.grid(row=1, column=1, padx=10, pady=10, sticky="w")

Label(entries_frame, text="Employee Code", font=("Calibri", 16), bg="#535c68", fg="white").grid(row=1, column=2, padx=10, pady=10, sticky="w")
txtCode = Entry(entries_frame, textvariable=code, font=("Calibri", 16), width=30)
txtCode.grid(row=1, column=3, padx=10, pady=10, sticky="w")

Label(entries_frame, text="Email", font=("Calibri", 16), bg="#535c68", fg="white").grid(row=2, column=0, padx=10, pady=10, sticky="w")
txtEmail = Entry(entries_frame, textvariable=email, font=("Calibri", 16), width=30)
txtEmail.grid(row=2, column=1, padx=10, pady=10, sticky="w")

Label(entries_frame, text="Contact No", font=("Calibri", 16), bg="#535c68", fg="white").grid(row=2, column=2, padx=10, pady=10, sticky="w")
txtContact = Entry(entries_frame, textvariable=contact, font=("Calibri", 16), width=30)
txtContact.grid(row=2, column=3, padx=10, pady=10, sticky="w")

# Gender Field
Label(entries_frame, text="Gender", font=("Calibri", 16), bg="#535c68", fg="white").grid(row=3, column=0, padx=10, pady=10, sticky="w")
frame_gender = Frame(entries_frame, bg="#5dade2")
frame_gender.grid(row=3, column=1, padx=10, pady=10, sticky="w")

Radiobutton(frame_gender, text="Male", variable=gender, value="Male", font=("Calibri", 16), bg="#5dade2", fg="white").pack(side=LEFT)
Radiobutton(frame_gender, text="Female", variable=gender, value="Female", font=("Calibri", 16), bg="#5dade2", fg="white").pack(side=LEFT)
Radiobutton(frame_gender, text="Other", variable=gender, value="Other", font=("Calibri", 16), bg="#5dade2", fg="white").pack(side=LEFT)

# Department Field
Label(entries_frame, text="Department", font=("Calibri", 16), bg="#535c68", fg="white").grid(row=3, column=2, padx=10, pady=10, sticky="w")
txtDepartment = Entry(entries_frame, textvariable=department, font=("Calibri", 16), width=30)
txtDepartment.grid(row=3, column=3, padx=10, pady=10, sticky="w")

# Address Field
Label(entries_frame, text="Address", font=("Calibri", 16), bg="#535c68", fg="white").grid(row=4, column=0, padx=10, pady=10, sticky="w")
txtAddress = Text(entries_frame, width=85, height=5, font=("Calibri", 16))
txtAddress.grid(row=4, column=1, columnspan=3, padx=10, pady=10, sticky="w")


# Functions
def getData(event):
    selected_row = tv.focus()
    if not selected_row:  
        return
    data = tv.item(selected_row)
    global row
    row = data["values"]
    name.set(row[1])
    code.set(row[2])
    email.set(row[3])
    contact.set(row[4])
    gender.set(row[5])
    department.set(row[6])
    txtAddress.delete(1.0, END)
    txtAddress.insert(END, row[7])

def delete_employee():
    if not row:
        messagebox.showerror("Error", "No record selected")
        return
    db.remove(row[0])
    clearAll()
    displayAll()

def displayAll():
    tv.delete(*tv.get_children())
    for row in db.fetch():
        tv.insert("", END, values=(
            row[0],     # ID
            row[1],     # Name
            row[2],     # Code
            row[3],     # Email
            row[4],     # Contact
            row[5],     # Gender
            row[6],     # Department
            row[7]      # Address
        ))

def update_employee():
    if not row:
        messagebox.showerror("Error", "No record selected")
        return
    if txtName.get() == "" or txtCode.get() == "" or txtEmail.get() == "" or txtContact.get() == "" or txtAddress.get(1.0, END).strip() == "" or gender.get() == "" or txtDepartment.get() == "":
        messagebox.showerror("Error in Input", "Please fill all the details")
        return
    db.update(row[0], txtName.get(), txtCode.get(), txtEmail.get(), txtContact.get(), gender.get(), txtDepartment.get(), txtAddress.get(1.0, END))
    messagebox.showinfo("Success", "Record Updated")
    clearAll()
    displayAll()

def add_employee():
    if txtName.get() == "" or txtCode.get() == "" or txtEmail.get() == "" or txtContact.get() == "" or txtAddress.get(1.0, END).strip() == "" or gender.get() == "" or txtDepartment.get() == "":
        messagebox.showerror("Error in Input", "Please fill all the details")
        return
    db.insert(txtName.get(), txtCode.get(), txtEmail.get(), txtContact.get(), gender.get(), txtDepartment.get(), txtAddress.get(1.0, END))
    messagebox.showinfo("Success", "Record Inserted")
    clearAll()
    displayAll()

def clearAll():
    name.set("")
    code.set("")
    email.set("")
    contact.set("")
    gender.set("")
    department.set("")
    txtAddress.delete(1.0, END)
    global row
    row = None  

def export_to_csv():
    try:
        file_path = filedialog.asksaveasfilename(
            defaultextension='.csv',
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="Save CSV File"
        )
        
        if file_path:  
            # Get all data from database
            data = db.fetch()
            
            # Write to CSV file
            with open(file_path, 'w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file, quoting=csv.QUOTE_ALL)  # Add QUOTE_ALL to properly handle multiline text
                
                # Write headers
                writer.writerow(['ID', 'Name', 'Code', 'Email', 'Contact', 'Gender', 'Department', 'Address'])
                
                # Process and write data rows
                for row in data:
                    # Replace any newlines in the address with space
                    processed_row = list(row)
                    if processed_row[7]:  # Address is at index 7
                        processed_row[7] = processed_row[7].replace('\n', ' ').strip()
                    writer.writerow(processed_row)
                
            messagebox.showinfo("Success", "Data exported successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")


def generate_id_card():
    if not row:
        messagebox.showerror("Error", "Please select an employee first")
        return
    
    # Card dimensions and colors
    card_width = 600
    card_height = 400
    background_color = "#f5f5f5"  # Light gray
    header_color = "#b0bec5"      # Soft blue-gray
    accent_color = "#78909c"      # Muted teal
    text_color = "#37474f"        # Dark gray
    highlight_color = "#cfd8dc"
   
    img = Image.new('RGB', (card_width, card_height), color=background_color)
    d = ImageDraw.Draw(img)
    
    
    # Top header background
    d.rectangle([(0, 0), (card_width, 80)], fill=header_color)
    
    # Side accent bar
    d.rectangle([(0, 0), (20, card_height)], fill=accent_color)
    
    # Add decorative elements
    d.rectangle([(20, 75), (card_width-20, 77)], fill=highlight_color)  # Horizontal line below title
    d.rectangle([(20, card_height-60), (card_width-20, card_height-58)], fill=highlight_color)  # Bottom line
    
    try:
       
        try:
            title_font = ImageFont.truetype("arial.ttf", 40)
            header_font = ImageFont.truetype("arial.ttf", 24)
            detail_font = ImageFont.truetype("arial.ttf", 18)
            label_font = ImageFont.truetype("arial.ttf", 16)
        except:
            # If custom font fails, use default font
            title_font = ImageFont.load_default()
            header_font = ImageFont.load_default()
            detail_font = ImageFont.load_default()
            label_font = ImageFont.load_default()
        
        # Add title with shadow effect
        d.text((53, 23), "EMPLOYEE ID CARD", font=title_font, fill=accent_color)  # Shadow
        d.text((50, 20), "EMPLOYEE ID CARD", font=title_font, fill='white')
        
        # Add employee details with improved layout
        details = [
            ("Employee Name", row[1]),
            ("Employee ID", row[2]),
            ("Department", row[6]),  
            ("Email Address", row[3]),
            ("Contact No.", row[4]),
            ("Gender", row[5]),
            ("Address", row[7])     
        ]
        
        y_position = 100
        for label, value in details:
            # Draw label background
            d.rectangle([(40, y_position-5), (180, y_position+25)], fill=highlight_color, width=0)
            
            d.text((45, y_position), f"{label}:", font=label_font, fill=text_color)
            d.text((190, y_position), str(value), font=detail_font, fill=text_color)
            y_position += 40
        
        footer_y = card_height - 40
        d.rectangle([(0, footer_y), (card_width, card_height)], fill=header_color)
        
        current_year = datetime.now().year
        validity_text = f"Valid Until: December {current_year + 2}"
        d.text((card_width//2 - 100, footer_y + 10), 
               validity_text, 
               font=detail_font, fill='white')
        
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")
        return
    
    try:
        if not os.path.exists('id_cards'):
            os.makedirs('id_cards')
        file_path = f'id_cards/{row[2]}_id_card.png'
        img.save(file_path)
        img.show()
        messagebox.showinfo("Success", f"ID Card generated successfully!\nSaved at: {file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save ID card: {str(e)}")

# Buttons
btn_frame = Frame(entries_frame, bg="#535c68")
btn_frame.grid(row=5, column=0, columnspan=4, padx=10, pady=10, sticky="w")

Button(btn_frame, command=update_employee, text="Update Details", width=15, font=("Calibri", 16, "bold"), fg="white", bg="#2980b9", bd=0).grid(row=0, column=1, padx=10)
Button(btn_frame, command=add_employee, text="Add Details", width=15, font=("Calibri", 16, "bold"), fg="white", bg="#16a085", bd=0).grid(row=0, column=0)
Button(btn_frame, command=clearAll, text="Clear Details", width=15, font=("Calibri", 16, "bold"), fg="white", bg="#f39c12", bd=0).grid(row=0, column=3, padx=10)
Button(btn_frame, command=delete_employee, text="Delete Details", width=15, font=("Calibri", 16, "bold"), fg="white", bg="#c0392b", bd=0).grid(row=0, column=2, padx=10)
Button(btn_frame, command=export_to_csv, text="Export CSV", width=15, font=("Calibri", 16, "bold"), fg="white", bg="#27ae60", bd=0).grid(row=0, column=4, padx=10)
Button(btn_frame, command=generate_id_card, text="Generate ID", width=15, 
       font=("Calibri", 16, "bold"), fg="white", bg="#e74c3c", bd=0).grid(row=0, column=6, padx=10)
# Table frame with Scrollbars
tree_frame = Frame(root, bg="#ecf0f1")
tree_frame.place(x=0, y=480, width=1920, height=520)

# Add scrollbars
scroll_y = Scrollbar(tree_frame, orient=VERTICAL)
scroll_x = Scrollbar(tree_frame, orient=HORIZONTAL)
scroll_y.pack(side=RIGHT, fill=Y)
scroll_x.pack(side=BOTTOM, fill=X)

style = ttk.Style()
style.configure("mystyle.Treeview", font=("Calibri", 14), rowheight=40)
style.configure("mystyle.Treeview.Heading", font=("Calibri", 14, "bold"))

# Treeview configuration
tv = ttk.Treeview(tree_frame, columns=("ID", "Name", "Code", "Email", "Contact", "Gender", "Department", "Address"), 
                  style="mystyle.Treeview", 
                  yscrollcommand=scroll_y.set, 
                  xscrollcommand=scroll_x.set)

# Configure scrollbars
scroll_y.config(command=tv.yview)
scroll_x.config(command=tv.xview)

# Define column headings with different alignments
tv.heading("ID", text="ID", anchor=CENTER)
tv.heading("Name", text="Name", anchor=CENTER)
tv.heading("Code", text="Code", anchor=CENTER)
tv.heading("Email", text="Email", anchor=CENTER)
tv.heading("Contact", text="Contact", anchor=CENTER)
tv.heading("Gender", text="Gender", anchor=CENTER)
tv.heading("Department", text="Department", anchor=CENTER)
tv.heading("Address", text="Address", anchor=W)  

# Set column widths and alignments
tv.column("ID", width=20, minwidth=20, anchor=W)
tv.column("Name", width=100, minwidth=100, anchor=W)
tv.column("Code", width=40, minwidth=40, anchor=W)
tv.column("Email", width=160, minwidth=160, anchor=W)
tv.column("Contact", width=90, minwidth=90, anchor=W)
tv.column("Gender", width=50, minwidth=50, anchor=W)
tv.column("Department", width=200, minwidth=200, anchor=W)
tv.column("Address", width=450, minwidth=450, anchor=W)

tv["show"] = "headings"
tv.bind("<ButtonRelease-1>", getData)
tv.pack(fill=BOTH, expand=1)

# Make sure the tree frame uses all available space
tree_frame.pack_propagate(False)

# Display all records initially
displayAll()

root.mainloop()
