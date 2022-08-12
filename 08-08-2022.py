def add_item(myTempList, item):    
    new_myTemplist = myTempList.append(item)

def find_index_item(myTempList, item_name):
    result = -1
    length = len(myTempList)
    for i in range(length):
        if myTempList[i]['name'] == item_name:
            result = i
    return result

def remove_item(myTempList, item_name):
    if find_index_item(myTempList, item_name) > -1:
        del myTempList[find_index_item(myTempList, item_name)]
    else:
        print(item_name + " not in list")

expenses = input('Tell me your expenses?  ' )
print("Your expenses:", expenses)
        
print("What do you want to do? -\n"\
        "1. Add\n" \
        "2. Remove")
option = int(input("Select option 1 or 2: "))
name_input = input("Item name: ")
if option == 1:
    cost_input = int(input("Item cost: "))
    date_input = input("Date: ")
    item = {'name': name_input, 'cost':cost_input, 'date':date_input}
    add_item(expenses, item)
    print("Your expenses: ", expenses)
elif option == 2:
    remove_item(expenses, name_input)
    print("Your expenses: ", expenses)
else:
    print("Invalid input")