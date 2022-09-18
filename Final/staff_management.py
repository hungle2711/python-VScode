import tkinter as tk
import re
from tkinter import ttk
from tkinter import messagebox as msb
from openpyxl import load_workbook

def check_blank_data():
    errors = []
    message=''
    if entry_value_id.get()=='':
        errors.append('Id')
    if entry_value_age.get()=='':
        errors.append('Age')
    if entry_value_full_name.get()=='':
        errors.append('Full_name')
    if entry_value_email.get()=='':
        errors.append('Email')
    if entry_value_phone.get()=='':
        errors.append('Phone')
    if cbx_selected_department.get()=='--- Lựa chọn phòng ban ---':
        errors.append('Department')
    for error_index in range(0,len(errors)):
        if error_index == len(errors)-1:
            message += errors[error_index] + ' '
        else:    
            message += errors[error_index] + ', '
    if message =='':
        return True
    else:
        msb.showwarning('Warnning',message+'is blank!')
        return False

def validate_phone(value):
    # Bắt lỗi nhập số điện thoại
    # '^0': ký tự đầu phải là 0; 'd{9}': yêu cầu phải là số 
    pattern = '^0\d{9}'
    if entry_value_phone.get() !='':
        if re.fullmatch(pattern, entry_value_phone.get()) is None:
            return False
        else:
            return True
    
def validate_age(value):
    pattern = '\d+'
    if entry_value_age.get() !='':
        if re.fullmatch(pattern, entry_value_age.get()) is None:
            return False
        else:
            return True

def on_invalid_phone():
    msb.showwarning('Warnning','Invalid phone number!\nIt should be like 0xxxxxxxxx') 
        
def on_invalid_age():
    msb.showwarning('Warnning','Invalid Age!\nIt should be number which is greater than zero')   
        
def save_data():
    wb.save('staff_list.xlsx')

def clear_data():
    entry_value_id.set('')
    entry_value_age.set('')
    entry_value_full_name.set('')
    entry_value_email.set('')
    entry_value_phone.set('')
    cbx_department.current(0)

def item_selected():
    selected_item = tvw_info.selection()
    item = tvw_info.item(selected_item,option='values')
#     item = items['values']
    return item

def view_staff():
    try:
        item = item_selected()
        entry_value_id.set(item[0])
        entry_value_full_name.set(item[1])
        entry_value_email.set(item[2])
        entry_value_age.set(item[3])
        entry_value_phone.set(item[4])
        cbx_department.current(cbx_department['value'].index(item[5]))
    except:
        clear_data()

def view_all_staff(): 
    global wb
    global ws
    
    wb = load_workbook('staff_list.xlsx')
    ws = wb['Sheet1']
    
    for item in tvw_info.get_children():
        tvw_info.delete(item)
    for row in range(2,ws.max_row+1):
        tvw_value = []
        for col in range(1,ws.max_column+1):
            tvw_value.append(ws.cell(row=row,column=col).value)
        tvw_info.insert('',tk.END,value=tvw_value)

def add_staff():
    new_id=entry_value_id.get()    
    new_age = ''
    if validate_age == False:
        entry_value_age.set('')
    new_age=entry_value_age.get()    
    new_full_name=entry_value_full_name.get()    
    new_email=entry_value_email.get()    
    new_phone = ''
    if validate_phone == False:        
        entry_value_phone.set('')
    new_phone=entry_value_phone.get()    
    new_deparment=cbx_selected_department.get()    
    new_staff=[new_id,new_full_name,new_email,new_age,new_phone,new_deparment]

    if check_blank_data() == True:        
        # Kiểm tra xem ID có bị trùng không
        for row in ws:
            if row[0].row > 1:
                if row[0].value == new_id:
                    print(row[0].value)
                    msb.showwarning('Warnning','ID is existed!')
                    return 
        # Thực hiện thêm vào file excel nếu không trùng
        last_row = ws.max_row
        last_column = 6
        for col in range(1,last_column+1):
            ws.cell(row=last_row+1,column=col).value=new_staff[col-1]
        msb.showinfo('Anoucement','Add successful!')
        save_data()
        clear_data()
        view_all_staff()
    
def update_staff():
    global wb
    global ws 
    
    update_id = entry_value_id.get()
    update_age = ''
    if validate_age == False:
        entry_value_age.set('')
    update_age=entry_value_age.get()    
    update_full_name=entry_value_full_name.get()    
    update_email=entry_value_email.get()    
    update_phone = ''
    if validate_phone == False:        
        entry_value_phone.set('')
    update_phone=entry_value_phone.get()    
    update_department=cbx_selected_department.get()    
    update_staff=[update_id,update_full_name,update_email,update_age,update_phone,update_department]
    
    for row in ws:
        if row[0].value == update_id:
            for col in range(0,ws.max_column):
                row[col].value=update_staff[col]
                
    msb.showinfo('Anoucement','Update successful!')
    save_data()
    clear_data()        
    wb = load_workbook('staff_list.xlsx')
    ws = wb['Sheet1']
    view_all_staff()
    
def delete_staff():
    item = item_selected()
    answer = msb.askyesno('Confirm','Do you want to delete?')
    if answer == 1:
        for row in ws:
            if row[0].value == item[0]:
                ws.delete_rows(row[0].row) 
                save_data()
                msb.showinfo('Anoucement','Delete successful!')        
    view_all_staff()
        
def delete_all_staff():
    answer = msb.askyesno('Confirm','Do you want to delete?')
    if answer == 1:
        max_row = ws.max_row+1
        for row in range(2,max_row):
            ws.delete_rows(2)
        save_data()
        msb.showinfo('Anoucement','Delete successful!')  
    
    view_all_staff()

def search():
    tvw_value = []
    index = -1
    for item in tvw_info.get_children():
                tvw_info.delete(item)
    if cbx_selected_search.get() == 'ID':
        index = 0
    if cbx_selected_search.get() == 'Full Name':
        index = 1
    if cbx_selected_search.get() == 'Email':
        index = 2
    if cbx_selected_search.get() == 'Age':
        index = 3
    if cbx_selected_search.get() == 'Phone':
        index = 4
    if cbx_selected_search.get() == 'Department':
        index = 5
    for row in ws:
        if row[index].value == entry_value_search.get():
            for col in range(1,ws.max_column+1):
                tvw_value.append(ws.cell(row[index].row,column=col).value)
            tvw_info.insert('',tk.END,value=tvw_value)

def sort():
    sort_list = []    
    index = -1
    for item in tvw_info.get_children():
        tvw_info.delete(item)
    if cbx_selected_sort.get() == 'ID':
        index = 0
    if cbx_selected_sort.get() == 'Full Name':
        index = 1
    if cbx_selected_sort.get() == 'Email':
        index = 2
    if cbx_selected_sort.get() == 'Age':
        index = 3
    if cbx_selected_sort.get() == 'Phone':
        index = 4
    if cbx_selected_sort.get() == 'Department':
        index = 5
    for row in ws:
        sort_list.append(row[index].value)
    sort_list.pop(0)
    sort_list.sort()
    for item in sort_list:
        for r in range(2,ws.max_row+1):
            tvw_value = []     
            cell_value = ws.cell(row=r,column=index+1).value
            if cell_value == item:
                for c in range(1,ws.max_column+1):
                    tvw_value.append(ws.cell(row=r,column=c).value)
                tvw_info.insert('',tk.END,value=tvw_value)

def exit():
    window.destroy()

# Khởi tạo window chính
window = tk.Tk()
window.title('MSD STAFF MANAGEMENT')
window.geometry('465x600')
window.resizable(False,False)

# Khởi tạo đối tượng đọc file excel
wb = load_workbook('staff_list.xlsx')
ws = wb['Sheet1']

# Khởi tạo biến variable của các widget
entry_value_id = tk.StringVar()
entry_value_full_name = tk.StringVar()
entry_value_email = tk.StringVar()
entry_value_age = tk.StringVar()
entry_value_phone = tk.StringVar()
entry_value_search = tk.StringVar()
cbx_selected_search = tk.StringVar()
cbx_selected_department = tk.StringVar()
cbx_selected_sort = tk.StringVar()
info_column = ('id','full_name','email','age','phone','department')

# Khởi tạo các trường để nhập thông tin
frame_input_info = ttk.Labelframe(window,text='Input information')

label_id = ttk.Label(frame_input_info,text='ID')
label_id.grid(row=0,column=0,sticky='e',padx='5',pady='5')

label_age = ttk.Label(frame_input_info,text='Age')
label_age.grid(row=0,column=2,sticky='e',padx='5',pady='5')

label_full_name = ttk.Label(frame_input_info,text='Full Name')
label_full_name.grid(row=1,column=0,sticky='e',padx='5',pady='5')

label_email = ttk.Label(frame_input_info,text='Email')
label_email.grid(row=1,column=2,sticky='e',padx='5',pady='5')

label_phone = ttk.Label(frame_input_info,text='Phone.No')
label_phone.grid(row=2,column=0,sticky='e',padx='5',pady='5')

label_Department = ttk.Label(frame_input_info,text='Department')
label_Department.grid(row=2,column=2,sticky='e',padx='5',pady='5')

entry_id = ttk.Entry(frame_input_info,textvariable=entry_value_id)
entry_id.grid(row=0,column=1,sticky='w',pady='5')

entry_age = ttk.Entry(frame_input_info,textvariable=entry_value_age)
vcmd_age = (entry_age.register(validate_age),'%P')
ivcmd_age = (entry_age.register(on_invalid_age))
entry_age.config(validate='focusout', validatecommand=vcmd_age, invalidcommand=ivcmd_age)
entry_age.grid(row=0,column=3,sticky='w',pady='5')

entry_full_name = ttk.Entry(frame_input_info,textvariable=entry_value_full_name)
entry_full_name.grid(row=1,column=1,sticky='w',pady='5')

entry_email = ttk.Entry(frame_input_info,textvariable=entry_value_email)
entry_email.grid(row=1,column=3,sticky='w',pady='5')

entry_phone = ttk.Entry(frame_input_info,textvariable=entry_value_phone)
vcmd_phone = (entry_phone.register(validate_phone),'%P')
ivcmd_phone = (entry_phone.register(on_invalid_phone))
entry_phone.config(validate='focusout', validatecommand=vcmd_phone, invalidcommand=ivcmd_phone)
entry_phone.grid(row=2,column=1,sticky='w',pady='5')

cbx_department = ttk.Combobox(frame_input_info,state= "readonly",textvariable=cbx_selected_department)
cbx_department['value']=('--- Select a Department ---','Kế toán','Giám đốc','Bán hàng')
cbx_department.current(0)
cbx_department.grid(row=2,column=3,sticky='w',pady='5')

# Khởi tạo các Button để thực hiện các chức năng
frame_function = ttk.Labelframe(window,text='Function')

btn_view = ttk.Button(frame_function, text='View All', command=view_all_staff)
btn_view.grid(row=0,column=0,sticky='w',padx='5',pady='5')

btn_add = ttk.Button(frame_function, text='Add', command=add_staff)
btn_add.grid(row=0,column=1,sticky='w',padx='5',pady='5')

btn_update = ttk.Button(frame_function, text='Update', command=update_staff)
btn_update.grid(row=0,column=2,sticky='w',padx='5',pady='5')

btn_delete = ttk.Button(frame_function, text='Delete', command=delete_staff)
btn_delete.grid(row=0,column=3,sticky='w',padx='5',pady='5')

btn_delete_all_btn = ttk.Button(frame_function, text='Delete All', command=delete_all_staff)
btn_delete_all_btn.grid(row=0,column=4,sticky='w',padx='5',pady='5')

btn_get_info = ttk.Button(frame_function, text='Get Info', command=view_staff)
btn_get_info.grid(row=1,column=0,sticky='w',padx='5',pady='5')

btn_clear = ttk.Button(frame_function, text='Clear', command=clear_data)
btn_clear.grid(row=1,column=1,sticky='w',padx='5',pady='5')

btn_exit = ttk.Button(frame_function, text='Exit', command=exit)
btn_exit.grid(row=1,column=2,sticky='w',padx='5',pady='5')

# Khởi tạo các trường để search
frame_search = ttk.Frame(window)

entry_search = ttk.Entry(frame_search, textvariable=entry_value_search)
entry_search.grid(row=0,column=0,sticky='w',padx='5',pady='5')

cbx_search = ttk.Combobox(frame_search,state= "readonly",textvariable=cbx_selected_search, width=13)
cbx_search['value']=('-- Search by --','ID','Full Name','Email','Age','Phone','Department')
cbx_search.current(0)
cbx_search.grid(row=0,column=1,sticky='e',pady='5')

btn_search = ttk.Button(frame_search,text='Search',command=search)
btn_search.grid(row=0,column=2,sticky='w',padx='5',pady='5')

cbx_sort = ttk.Combobox(frame_search,state= "readonly",textvariable=cbx_selected_sort, width=10)
cbx_sort['value']=('-- Sort by --','ID','Full Name','Email','Age','Phone','Department')
cbx_sort.current(0)
cbx_sort.grid(row=1,column=1,sticky='e',pady='5')

btn_sort = ttk.Button(frame_search,text='Sort',command=sort)
btn_sort.grid(row=1,column=2,sticky='e',padx='5',pady='5')

# Khởi tạo các trường để hiển thị thông tin
frame_display_info = ttk.Labelframe(window, text='Display Information')

tvw_info = ttk.Treeview(frame_display_info,column=info_column,show='headings')
tvw_info.column('id',width=30)
tvw_info.column('full_name',width=100)
tvw_info.column('email',width=100)
tvw_info.column('age',width=30)
tvw_info.column('phone',width=80)
tvw_info.column('department',width=50)

tvw_info.heading('id',text='ID')
tvw_info.heading('full_name',text='Full Name')
tvw_info.heading('email',text='Email')
tvw_info.heading('age',text='Age')
tvw_info.heading('phone',text='Phone.No')
tvw_info.heading('department',text='Department')

tvw_info.grid(row=1,column=0,sticky='w',padx='5',pady='5')
tvw_info.bind('<<TreeviewSelect>>',item_selected())

scrollbar = ttk.Scrollbar(frame_display_info,orient='vertical',command=tvw_info.yview)
tvw_info['yscrollcommand']=scrollbar.set
scrollbar.grid(row=1,column=1,sticky='w',padx='5',pady='5')

# Sắp xếp các frame
frame_input_info.pack(padx='10',pady='20',anchor='w')
frame_function.pack(padx='10',pady='20',anchor='w')
frame_search.pack(padx='10',pady='0',anchor='w')
frame_display_info.pack(padx='10',pady='0',expand=True,fill='y',anchor='w')

window.mainloop()