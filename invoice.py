import threading
import pdfkit
import smtplib
import datetime
import os
from tkinter import *
from tkinter import messagebox, ttk
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

items = []

def add_item():
    try:
        name = item_name.get().title()
        quantity = int(item_quantity.get())
        price = float(item_price.get())
        total = quantity * price
        items.append({'name': name, 'quantity': quantity, 'price': price, 'total': total})
        update_item_list()
        update_total()
        clear_item_fields()
        status_label.config(text="Item added successfully!", foreground="green")
    except ValueError:
        messagebox.showerror("Error", "Quantity and Price must be numbers")
        status_label.config(text="Error adding item!", foreground="red")

def update_item_list():
    item_list.delete(*item_list.get_children())
    for index, item in enumerate(items):
        item_list.insert("", "end", iid=index, values=(item['name'], item['quantity'], item['price'], item['total']))

def clear_item_fields():
    item_name.delete(0, 'end')
    item_quantity.delete(0, 'end')
    item_price.delete(0, 'end')

def clear_items():
    if messagebox.askyesno("Confirmation", "Are you sure you want to clear all items?"):
        global items
        items = []
        update_item_list()
        update_total()
        status_label.config(text="All items cleared!", foreground="green")

def delete_item():
    selected_item = item_list.selection()
    if selected_item:
        if messagebox.askyesno("Confirmation", "Are you sure you want to delete the selected item?"):
            index = int(selected_item[0])
            del items[index]
            update_item_list()
            update_total()
            status_label.config(text="Item deleted successfully!", foreground="green")
    else:
        messagebox.showerror("Error", "No item selected")
        status_label.config(text="Error deleting item!", foreground="red")

def generate_pdf(html_content, pdf_output_path, customer_name):
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    daily_folder = os.path.join(pdf_output_path, today_str)
    if not os.path.exists(daily_folder):
        os.makedirs(daily_folder)

    invoice_number = 1
    while os.path.exists(os.path.join(daily_folder, f"{customer_name}_{invoice_number}.pdf")):
        invoice_number += 1

    with open('invoice.html', 'w') as f:
        f.write(html_content)

    pdf_path = os.path.join(daily_folder, f"{customer_name}_{invoice_number}.pdf")
    wkhtmltopdf_path = r'C:\Users\91990\Downloads\k\wkhtmltopdf\bin\wkhtmltopdf.exe'

    config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)
    try:
        pdfkit.from_file('invoice.html', pdf_path, configuration=config)
        return pdf_path
    except Exception as e:
        print(f"Failed to generate PDF: {e}")
        return None

def send_email_with_attachment(pdf_path, customer_email):
    if not os.path.exists(pdf_path):
        print(f"PDF file does not exist: {pdf_path}")
        return

    sender_email = 'krupangdholakia143@gmail.com'
    sender_password = 'frjz dcqp jorz rczd'

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = customer_email
    msg['Subject'] = 'Invoice'

    try:
        with open(pdf_path, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename=invoice.pdf')
        msg.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, customer_email, msg.as_string())

        print(f"PDF sent to {customer_email} via email!")
        messagebox.showinfo("Success", "Invoice sent successfully via email!")
        status_label.config(text="Invoice sent successfully via email!", foreground="green")
    except Exception as e:
        print(f"Failed to send email: {e}")
        status_label.config(text="Failed to send email!", foreground="red")

def generate_invoice():
    customer_name = customer_name_entry.get().title()
    customer_email = customer_email_entry.get()
    customer_phone = customer_phone_entry.get()
    delivery_method = delivery_method_var.get()

    html_content = generate_html(customer_name, customer_email, customer_phone, items)
    pdf_output_path = r'C:\Users\91990\Downloads\k\invoice log'

    if not os.path.exists(pdf_output_path):
        os.makedirs(pdf_output_path)

    pdf_path = generate_pdf(html_content, pdf_output_path, customer_name)
    if pdf_path:
        if delivery_method == "Email":
            threading.Thread(target=send_email_with_attachment, args=(pdf_path, customer_email)).start()
        else:
            messagebox.showinfo("Success", f"Invoice saved locally at: {pdf_path}")
            status_label.config(text=f"Invoice saved locally at: {pdf_path}", foreground="green")
    else:
        messagebox.showerror("Error", "Failed to generate invoice PDF. Please try again.")
        status_label.config(text="Failed to generate invoice PDF!", foreground="red")

def generate_html(customer_name, customer_email, customer_phone, items):
    current_date = datetime.date.today().strftime("%B %d, %Y")
    item_rows = ""
    total_amount = 0
    for item in items:
        item_total = item['quantity'] * item['price']
        total_amount += item_total
        item_row = f"""
        <tr>
          <td>{item['name']}</td>
          <td>{item['quantity']}</td>
          <td>${item['price']:.2f}</td>
          <td>${item_total:.2f}</td>
        </tr>
        """
        item_rows += item_row

    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <title>Invoice</title>
    <style>
    body {{
      font-family: Arial, sans-serif;
      background-color: #f9f9f9;
      color: #333;
    }}
    .invoice {{
      width: 80%;
      margin: 0 auto;
      padding: 20px;
      border: 1px solid #ccc;
      background-color: #fff;
      box-shadow: 0 0 10px rgba(0,0,0,0.1);
    }}
    .header {{
      text-align: center;
    }}
    .customer, .items, .total {{
      margin-top: 30px;
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
    }}
    table, th, td {{
      border: 1px solid #ddd;
    }}
    th , td {{
      padding: 8px;
      text-align: left;
    }}
    th {{
      background-color: #f2f2f2;
    }}
    </style>
    </head>
    <body>
    <div class="invoice">
      <div class="header">
        <h1 style="color: #336699;">Invoice</h1>
        <p>Date: {current_date}</p>
      </div>
      <div class="customer">
        <h2>Customer Information</h2>
        <p>Name: {customer_name}</p>
        <p>Email: {customer_email}</p>
        <p>Phone: {customer_phone}</p>
      </div>
      <div class="items">
        <h2>Items</h2>
        <table>
          <thead>
            <tr>
              <th style="background-color: #336699; color: #fff;">Item</th>
              <th style="background-color: #336699; color: #fff;">Quantity</th>
              <th style="background-color: #336699; color: #fff;">Price</th>
              <th style="background-color: #336699; color: #fff;">Total</th>
            </tr>
          </thead>
          <tbody>
            {item_rows}
          </tbody>
        </table>
      </div>
      <div class="total">
        <h2>Total</h2>
        <p>${total_amount:.2f}</p>
      </div>
    </div>
    </body>
    </html>
    """
    return html_content

def update_total():
    total_amount = sum(item['total'] for item in items)
    total_label.config(text=f"Total: ${total_amount:.2f}")

def edit_item():
    selected_item = item_list.selection()
    if selected_item:
        index = int(selected_item[0])
        item = items[index]
        item_name.delete(0, 'end')
        item_name.insert(0, item['name'])
        item_quantity.delete(0, 'end')
        item_quantity.insert(0, item['quantity'])
        item_price.delete(0, 'end')
        item_price.insert(0, item['price'])
        delete_item()
        status_label.config(text="Edit the item and click 'Add Item'", foreground="blue")
    else:
        messagebox.showerror("Error", "No item selected")
        status_label.config(text="Error editing item!", foreground="red")

root = Tk()
root.title("Invoice Generator")
root.geometry("800x600")
root.config(bg='#f0f0f0')

style = ttk.Style()
style.configure('TLabel', font=('Arial', 12))
style.configure('TEntry', font=('Arial', 12))
style.configure('TButton', font=('Arial', 12))
style.configure('TFrame', background='#f0f0f0')

# Customer Details
customer_frame = ttk.LabelFrame(root, text="Customer Details", padding=(10, 10, 10, 10), style='TFrame')
customer_frame.grid(row=0, column=0, columnspan=3, padx=20, pady=10, sticky=W+E)

ttk.Label(customer_frame, text="Customer Name:", font=('Arial', 12, 'bold')).grid(row=0, column=0, padx=10, pady=5, sticky=W)
customer_name_entry = ttk.Entry(customer_frame)
customer_name_entry.grid(row=0, column=1, padx=10, pady=5, sticky=W)

ttk.Label(customer_frame, text="Customer Email:", font=('Arial', 12, 'bold')).grid(row=1, column=0, padx=10, pady=5, sticky=W)
customer_email_entry = ttk.Entry(customer_frame)
customer_email_entry.grid(row=1, column=1, padx=10, pady=5, sticky=W)

ttk.Label(customer_frame, text="Customer Phone:", font=('Arial', 12, 'bold')).grid(row=2, column=0, padx=10, pady=5, sticky=W)
customer_phone_entry = ttk.Entry(customer_frame)
customer_phone_entry.grid(row=2, column=1, padx=10, pady=5, sticky=W)

# Item Details
item_frame = ttk.LabelFrame(root, text="Item Details", padding=(10, 10, 10, 10), style='TFrame')
item_frame.grid(row=1, column=0, columnspan=3, padx=20, pady=10, sticky=W+E)

ttk.Label(item_frame, text="Item Name:", font=('Arial', 12, 'bold')).grid(row=0, column=0, padx=10, pady=5, sticky=W)
item_name = ttk.Entry(item_frame)
item_name.grid(row=0, column=1, padx=10, pady=5, sticky=W)

ttk.Label(item_frame, text="Quantity:", font=('Arial', 12, 'bold')).grid(row=1, column=0, padx=10, pady=5, sticky=W)
item_quantity = ttk.Entry(item_frame)
item_quantity.grid(row=1, column=1, padx=10, pady=5, sticky=W)

ttk.Label(item_frame, text="Price:", font=('Arial', 12, 'bold')).grid(row=2, column=0, padx=10, pady=5, sticky=W)
item_price = ttk.Entry(item_frame)
item_price.grid(row=2, column=1, padx=10, pady=5, sticky=W)

# Buttons
add_item_button = ttk.Button(item_frame, text="Add Item", command=add_item, style='TButton')
add_item_button.grid(row=0, column=2, padx=10, pady=5, sticky=W)

edit_item_button = ttk.Button(item_frame, text="Edit Item", command=edit_item, style='TButton')
edit_item_button.grid(row=0, column=3, padx=10, pady=5, sticky=W)

delete_item_button = ttk.Button(item_frame, text="Delete Item", command=delete_item, style='TButton')
delete_item_button.grid(row=1, column=2, padx=10, pady=5, sticky=W)

clear_items_button = ttk.Button(item_frame, text="Clear Items", command=clear_items, style='TButton')
clear_items_button.grid(row=2, column=2, padx=10, pady=5, sticky=W)

# Item List
item_list_frame = ttk.LabelFrame(root, text="Item List", padding=(10, 10, 10, 10), style='TFrame')
item_list_frame.grid(row=2, column=0, columnspan=3, padx=20, pady=10, sticky=W+E)

item_list_scrollbar = ttk.Scrollbar(item_list_frame, orient=VERTICAL)
item_list = ttk.Treeview(item_list_frame, columns=("Item", "Quantity", "Price", "Total"), show="headings", yscrollcommand=item_list_scrollbar.set)
item_list_scrollbar.config(command=item_list.yview)
item_list_scrollbar.pack(side=RIGHT, fill=Y)
item_list.heading("Item", text="Item", anchor=CENTER)
item_list.heading("Quantity", text="Quantity", anchor=CENTER)
item_list.heading("Price", text="Price", anchor=CENTER)
item_list.heading("Total", text="Total", anchor=CENTER)
item_list.pack(side=LEFT, fill=BOTH, expand=True)

# Total Label
total_label = ttk.Label(root, text="Total: $0.00", font=('Arial', 14, 'bold'), background='#f0f0f0')
total_label.grid(row=3, column=0, columnspan=3, pady=10)

# Status Label
status_label = ttk.Label(root, text="", font=('Arial', 12), background='#f0f0f0')
status_label.grid(row=4, column=0, columnspan=3, pady=10)

# Delivery Method
delivery_method_label = ttk.Label(root, text="Delivery Method:", background='#f0f0f0')
delivery_method_label.grid(row=5, column=0, padx=20, pady=(10, 0), sticky=W)

delivery_method_var = StringVar(root)
delivery_method_var.set("Email")
delivery_method_option = ttk.OptionMenu(root, delivery_method_var, "Email", "Email", "Local Save")
delivery_method_option.grid(row=5, column=1, padx=10, pady=(10, 0), sticky=W)

# Generate Invoice Button
generate_invoice_button = ttk.Button(root, text="Generate Invoice", command=generate_invoice, style='TButton')
generate_invoice_button.grid(row=5, column=2, padx=20, pady=(10, 0), sticky=W)

root.mainloop()
