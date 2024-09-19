import tkinter
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
import customtkinter

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("dark-blue")

invoice_list=[]

def clear_item():
    product_name_input.delete(0,tkinter.END)
    product_price_input.delete(0,tkinter.END)
    product_quantity_input.delete(0,tkinter.END)
    product_quantity_input.insert(0,"1")

def add_item():


    product_name=product_name_input.get()
    product_price= float(product_price_input.get())
    product_quantity = int(product_quantity_input.get())
    total = product_quantity *product_price

    invoice_items = [product_quantity,product_name,product_price,total]
    tree.insert('',0,values=invoice_items)
    invoice_list.append(invoice_items)
    clear_item()





def remove_item():
    tree.delete(tree.get_children()[0])
    invoice_list.pop()

def new_invoice():
    first_name_input.delete(0,tkinter.END)
    sur_name_input.delete(0,tkinter.END)
    phone_num_input.delete(0,tkinter.END)
    clear_item()
    tree.delete(*tree.get_children())
    invoice_list.clear()

def generate_invoice():

    doc  = DocxTemplate('invoice_template.docx')

    name = first_name_input.get()+" "+sur_name_input.get()
    phone = phone_num_input.get()
    subtotal=0

    for item in invoice_list:
        subtotal+=int(item[3])
        salesTax = 0.05
        salestax = str(salesTax*100)+" %"

        total = (1-salesTax)*subtotal


    doc.render({"name":name,"invoice_list":invoice_list,"subtotal":subtotal,"salestax":salestax,"phone":phone,"total":total})

    doc_name = "newInvoice_"+name+"_"+datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S")+".docx"
    doc.save(doc_name)


window = customtkinter.CTk()
window.title("Invoice Generator")

frame = tkinter.Frame(window,padx=20,pady=10,bg="#343434")

frame.pack()

first_name_label = customtkinter.CTkLabel(frame,text="First Name")
first_name_label.grid(row=0,column=0)

sur_name_label = customtkinter.CTkLabel(frame,text="Surname")
sur_name_label.grid(row=0,column=1)

phone_num_label = customtkinter.CTkLabel(frame,text="Phone Number")
phone_num_label.grid(row=0,column=2)


first_name_input = customtkinter.CTkEntry(frame)
first_name_input.grid(row=1,column=0)

sur_name_input = customtkinter.CTkEntry(frame)
sur_name_input.grid(row=1,column=1)

phone_num_input = customtkinter.CTkEntry(frame)
phone_num_input.grid(row=1,column=2)

product_name_label = customtkinter.CTkLabel(frame,text="Product Name")
product_name_label.grid(row=2,column=0)

product_quantity_label = customtkinter.CTkLabel(frame,text="Quantity")
product_quantity_label.grid(row=2,column=1)

product_price_label = customtkinter.CTkLabel(frame,text="Unit Price")
product_price_label.grid(row=2,column=2)

product_name_input = customtkinter.CTkEntry(frame)
product_name_input.grid(row=3,column=0)



product_quantity_input = customtkinter.CTkEntry(frame)
product_quantity_input.grid(row=3,column=1)

product_price_input = customtkinter.CTkEntry(frame)
product_price_input.grid(row=3,column=2)


remove_prev_item_btn = customtkinter.CTkButton(frame,text="Remove Previous Item",command=remove_item,corner_radius=32)
remove_prev_item_btn.grid(row=4,column=1,pady=5)

add_item_btn= customtkinter.CTkButton(frame,text="Add Item",command =add_item,corner_radius=32)
add_item_btn.grid(row=4,column=2,pady=5)

coltup=('qty','desc','price','total')

tree =ttk.Treeview(frame,columns=coltup,show="headings")
tree.grid(row=5,column=0,columnspan=3,padx=20,pady=10)

tree.heading('qty',text="Quantity")
tree.heading('desc',text="Description")
tree.heading('price',text="Price")
tree.heading('total',text="Total")

style = tkinter.ttk.Style()
style.theme_use("clam")
style.configure("Treeview",
                background="red",
                foreground="white",
                fieldbackground="#2e2e2e",
                font=('Helvetica', 12))

style.configure("Treeview.Heading",
                background="#444444",
                foreground="white",
                font=('Helvetica', 12, 'bold'))

style.map("Treeview.Heading",
          background=[('active', '#555555')],
          foreground=[('active', 'white')])
tree.tag_configure('colored', background='red')


generate_invoice_btn =customtkinter.CTkButton(frame,text="Generate Invoice",command =generate_invoice)
generate_invoice_btn.grid(row=6,column=0,columnspan=3,sticky="news",pady=5)

new_invoice_btn = customtkinter.CTkButton(frame,text="New Invoice",command=new_invoice)
new_invoice_btn.grid(row=7,column=0,columnspan=3,sticky="news",pady=5)





window.mainloop()