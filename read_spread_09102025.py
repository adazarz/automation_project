import pandas as pd
import win32com.client
from datetime import date
from pprint import pprint

class bcolors:
    FAIL = '\033[91m'
    ENDC = '\033[0m'

#Cleaning the details of letters, separating the numbers
def clean(detail):
    #Auxiliary variables
    cleaned_details = []
    a = str()
    #Directly adding stand-alone numbers
    if isinstance(detail,int):
        cleaned_details.append(detail)
    else:
        l = len(detail) - 1
        for i,sign in enumerate(detail):
            if sign.isdigit():
                a += sign
                if i == l:
                    cleaned_details.append(a)
                    a = str()
            elif a:
                cleaned_details.append(a)
                a = str()
    #Elaborating shortened details, ex. 122345/46/47/48 into 122345, 122346, 122347, 122348
    #Elaboration happens if subsequent invoice numbers are shorter by at least 5 digits or are shorter by at least three digits and the next 
    #number is greater by one than the first, as in the example above.
    count = 0
    while count < len(cleaned_details) - 1:
        first_number = cleaned_details[count]
        next_number = cleaned_details[count+1]
        dif = len(str(first_number)) - len(str(next_number))
        if dif > 4:
            cleaned_details[count + 1] = first_number[:dif] + next_number
        elif dif > 2:
            c = first_number[dif:]
            if int(next_number) - int(c) == 1:
                cleaned_details[count + 1] = first_number[:dif] + next_number
        count += 1
    return cleaned_details

def text_only(detail):
    if isinstance(detail, int):
        return False
    for sign in detail:
        if sign.isdigit():
            return False
    return True

#Checking if the BOE references can be processed in a specifically integer-like way, that is, if they include letters, ex. AXN12345
def is_int(number):
    count_letters = 0
    for i in str(number):
        if not i.isdigit():
            count_letters += 1
    if count_letters == 0:
        return True
    else:
        return False
    
def differs_by_up_to_fifty(s1, s2):
    return is_int(s1) and is_int(s2) and abs(int(s1) - int(s2)) < 100

#Checking if we can merge bills of exchange's references
#The program compares "cleaned" details, it means sheer numbers, so as to recognise the same details written in a slightly different way
#ex. "AVANCE 253" and "SOLDE 253" will be recognised as identical
def can_merge(i, j):
    return (
        clients[i] == clients[j] and
        clean(details[i]) == clean(details[j]) and
        banq_domi[i].upper() == banq_domi[j].upper()
    ) 

def remove_line(h):
    clients.pop(h)
    nlcn.pop(h)
    banq_domi.pop(h)
    echeance.pop(h)
    details.pop(h)
    montants.pop(h)
    
#Checking the format of date d'echeance, formatting accordingly
def date_format(echeance):
    if isinstance(echeance,date):
        return echeance.strftime("%d/%m/%Y")
    else:
        return echeance

def connecting(numbers):
    #Checking if we can connect shortened nlcn with a dash
    if not all(is_int(z) for z in numbers):
        return
    y = 1
    while y < len(numbers):
        if len(numbers) > 3 and all(len(x) == 2 for x in numbers[y:]) and numbers[y-1].isdigit():
            numbers_parallel = [int(x) for x in numbers if x.isdigit()]
            numbers_parallel[y-1] = int(str(numbers[y-1])[-2:])
            from_the_beginning = y - 1
            while from_the_beginning < len(numbers) + 1 - y:
                from_the_end = 1
                while from_the_end <= len(numbers) + 1 - y:
                    if numbers_parallel[-from_the_end] != "-" and numbers_parallel[from_the_beginning] != "-" and numbers_parallel[-from_the_end] - numbers_parallel[from_the_beginning] == len(numbers) - from_the_beginning - from_the_end > 2:
                        difference = numbers_parallel[-from_the_end] - numbers_parallel[from_the_beginning]
                        co = 0
                        while co < difference -1:
                            numbers.pop(from_the_beginning + 1)
                            co += 1
                        numbers.insert(from_the_beginning + 1, "-")
                        #numbers[y:] = [str(x) for x in numbers_parallel[y:]]
                    from_the_end += 1
                from_the_beginning += 1
        y += 1

#Reading the data from excel macro
excel = pd.read_excel(PATH_TO_EXCEL_FILE)
clients = excel["Unnamed: 1"][3:].to_list()
nlcn = excel["Unnamed: 3"][3:].to_list()
names = excel["Unnamed: 2"][3:].to_list()
banq_domi = excel["Unnamed: 5"][3:].to_list()
echeance = excel["Unnamed: 7"][3:].to_list()
details = excel["Unnamed: 8"][3:].to_list()
montants = excel["Unnamed: 4"][3:].to_list()

#Ensuring customer numbers are of type int
clients = [int(i) for i in clients]

#Detecting incosistencies in customers' names and numbers
df2 = pd.DataFrame({"clients":clients, "names":names})

name_counts = df2.groupby('names')['clients'].nunique()
inconsistent_rows = df2[df2['names'].isin(name_counts[name_counts > 1].index)]

if len(inconsistent_rows) > 0:
    print(f"{bcolors.FAIL}There are inconsistencies in customers' numbers and names, please check.{bcolors.ENDC}\n")
    print(inconsistent_rows)
    print()

#Getting rid of the leading zeros from nlcn
y = -1
while y < len(nlcn) - 1:
    y += 1
    try:
        nlcn[y] = int(nlcn[y])
    except:
        continue

#Program can use automatically the current date or ask the user, it can be adjusted, just (un)comment this:
saving_date = input("Enter saving date or y to use the current one.\n")
if saving_date == "y":
    saving_date = date.today().strftime("%d/%m/%Y")

#And (un)comment this:
# saving_date = date.today().strftime("%d/%m/%Y")

date_length = len(saving_date) + 2

#Removing lines with ci-joint and avance and merging references from lines with no nlcn and checking for "SOLDE" in the details
text_only_line_removed = False
to_remove = []
for h, de in enumerate(details):
    try:
        if str(de).upper().strip() in ["SOLDE DU COMPTE","SOLDE","LE SOLDE","SOLDE COMPTE"]:
            details[h] = 0
        elif pd.isna(nlcn[h]):
            if clients[h] and de:
                details[h+1] = str(details[h+1]) + "/" + str(details[h])
                to_remove.append(h)
        elif str(de).upper() in ["AVANCE", "CI-JOINT","CI JOINT", "CI ATTACHE", "CI-ATTACHE"] or pd.isna(de):
            to_remove.append(h)
        elif text_only(de):
            text_only_line_removed = True
            to_remove.append(h)
    except AttributeError:
        print("AttributeError")

for h in reversed(to_remove):
    remove_line(h)
if not clients:
    input(f"{bcolors.FAIL}The file has no lines with details. You probably forgot to save it.{bcolors.ENDC}\n")

#Function for shortening consecutive bills of exchange
def shortening():
    gathered_nlcn = []
    for n in box:
        gathered_nlcn.append(nlcn[n])
    sorted_nlcn = sorted(gathered_nlcn)
    for n, m in  zip(box, range(len(box))):
        nlcn[n] = sorted_nlcn[m]
    #nlcn[box[0]:box[-1]+1] = sorted_nlcn
    previous = int(nlcn[box[0]])
    for z in box[1:]:
        current = int(nlcn[z])
        for power in range(2, len(str(nlcn[z]))):
            if current // 10**power == previous // 10**power:
                nlcn[z] = str(nlcn[z])[-power:]
                break
        previous = current

#Detecting consecutive bills of exchange numbers in order to shorten them
unique_details = {(client, detail, bq) : list() for detail, client, bq in zip(details, clients, banq_domi)}

for i, (client, detail, bq) in enumerate(zip(clients, details, banq_domi)):
    unique_details[(client,detail, bq)].append(i)

for detail in unique_details:
    if len(unique_details[detail]) > 1:
        box = [unique_details[detail][0]]
        for y in unique_details[detail][1:]:
            if can_merge(box[-1],y) and differs_by_up_to_fifty(nlcn[box[-1]], nlcn[y]):
                box.append(y)
            else:
                if len(box) > 1:
                    shortening()
                box = [y]

        #Handle the last box
        if len(box) > 1:
            shortening()

#Merging bills of exchange numers and bank abbreviations
boes = [pd.NA] * len(nlcn)
for detail in unique_details:
    if len(unique_details[detail]) == 1:
        boes[unique_details[detail][0]] = str(nlcn[unique_details[detail][0]]) + "/" + banq_domi[unique_details[detail][0]]
    else:
        group = [unique_details[detail][0]]  # Start with the first index

        for i,x in enumerate(unique_details[detail][1:]):
            if can_merge(group[-1], x):
                group.append(x)
            else:
                # Merge current group
                numbers = [str(nlcn[i]) for i in group]
                #Checking if we can connect shortened nlcn with a dash
                connecting(numbers)
                suffix = "/" + banq_domi[group[0]]
                merged = "+".join(numbers) + suffix
                if "+-+" in merged:
                    merged = merged.replace("+-+", "-")
                for i in group:
                    boes[i] = merged
                group = [x]
                
            # Handle the last group
            if group and i == len(unique_details[detail]) - 2:
                numbers = [str(nlcn[i]) for i in group]
                connecting(numbers)
                suffix = "/" + banq_domi[group[0]]
                merged = "+".join(numbers) + suffix
                if "+-+" in merged:
                    merged = merged.replace("+-+", "-")
                for i in group:
                    boes[i] = merged

#Creating final data
final_data = dict()
accounts = list(dict.fromkeys(clients))

details = list(tuple(clean(x)) for x in details)

repeated_accounts = []
repeated_details = []
for i, account in enumerate(accounts):
    if clients.count(account) == 1:
        ind = clients.index(account)
        final_data[account] = {details[ind] : (boes[ind], saving_date + " due " + date_format(echeance[ind]), str(round(montants[ind], 2)) + " MAD")}
    else:
        for j, account_in_clients in enumerate(clients):
            if account == account_in_clients:
                repeated_accounts.append(j)
        final_data[account] = {details[k] : (boes[k], saving_date + \
        " due " + date_format(echeance[k]), str(round(montants[k], 2)) + " MAD") for k in repeated_accounts}
        repeated_accounts = []

#Checking for duplicate details and customer numbers
invoices = list(dict.fromkeys(details))
for detail in invoices:
    if details.count(detail) > 1:
        new_amount = 0
        for a, det in enumerate(details):
            if detail == det and (not repeated_accounts or clients[repeated_accounts[-1]] == clients[a]):
                repeated_details.append(a)
                if not repeated_accounts:
                    repeated_accounts.append(a)
            else:
                continue
        for g in repeated_details:
            new_amount += montants[g]
        for e in repeated_details:
            montants[e] = new_amount
        
        #Removing duplicate merged boe references
        if repeated_details:
          final_data[clients[repeated_details[0]]][detail] = tuple(set(boes[b] for b in repeated_details)) + \
          (saving_date,) + tuple(set(" " + str(round(montants[b], 2)) + " MAD" for b in repeated_details))
        repeated_details = []
        repeated_accounts = []

pprint(final_data,sort_dicts=False)

if text_only_line_removed:
    print(f"\n{bcolors.FAIL}A line with no number in detail has been removed, it should be examined.{bcolors.ENDC}\n")

input("Write anything to continue.\n")
clipboard = list()

# Connecting to SAP
sap = win32com.client.GetObject("SAPGUI").GetScriptingEngine
session = sap.Children(0).Children(0)  # Access first open session

#Main loop
for key in final_data:
    skip = False
    # Open fbl5n
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nfbl5n"
    session.findById("wnd[0]").sendVKey(0)
    # Filling in the customer number and company code
    session.findById("wnd[0]/usr/ctxtDD_KUNNR-LOW").Text = key
    session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").Text = COMPANY_CODE
    session.findById("wnd[0]/usr/ctxtPA_VARI").Text = LAYOUT_NAME
    # Ensuring that normal items are checked
    session.findById("wnd[0]/usr/chkX_NORM").selected = True
    # Executing
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    # Filling in the references
    for ref in final_data[key].keys():
        if len(final_data[key].keys()) > 1:
            #Refreshing if the customer has more than one detail
            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/mbar/menu[0]/menu[2]").select()
        # Filtering by references if there was no "le solde" or sth similar
        if ref[0] != 0:
            try:
                session.findById("wnd[0]/usr/lbl[24,5]").caretPosition = 6
                session.findById("wnd[0]").sendVKey(2)
            except:
                probl = ""
                while probl != "y":
                    probl = input("Wrong customer number, write the correct one in SAP and write y to continue/ctrl+c to terminate\n").lower()
                #Two reapeted lines
                session.findById("wnd[0]/tbar[1]/btn[8]").press()
                session.findById("wnd[0]/usr/lbl[24,5]").caretPosition = 6
                session.findById("wnd[0]").sendVKey(2)

            session.findById("wnd[0]/tbar[1]/btn[38]").press()
            session.findById(r"wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press()
            #Resetting the clipboard
            clipboard = []
            for i, bare_ref in enumerate(ref):
                clipboard.append("*" + str(bare_ref))
            df = pd.DataFrame(clipboard[1:], columns=[clipboard[0]])
            # Copy the DataFrame to the clipboard
            df.to_clipboard(index=False)
            #Copying from clipboard
            session.findById("wnd[2]/tbar[0]/btn[24]").press()
            #Executing filtr
            session.findById("wnd[2]/tbar[0]/btn[8]").press()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        else:
            items_on_account = int(session.findById("wnd[0]/sbar").text[:-16])
        #Checking the sum and reason code's of the filtered invoices in SAP
        items_sum = 0
        empty_rc = False
        with_rc = 0
        #Necessary for handling solde
        if ref[0] != 0:
            len_ref = len(ref)
        else:
            len_ref = items_on_account
        try:
            if len_ref < 35:
                items_sum = float(session.findById(f"wnd[0]/usr/lbl[106,{len_ref+8}]").text.strip().replace("," ,""))
                line_number = 7
                while line_number < len_ref+7 and not with_rc:
                    #session.findById(f"wnd[0]/usr/lbl[123,{line_number}]").setFocus()
                    if session.findById(f"wnd[0]/usr/lbl[123,{line_number}]").text != "":
                        with_rc += 1
                    line_number += 1
                if not with_rc:
                    empty_rc = True
                else:
                    print("There is an item with reason code.")

                #Getting the document number below the last item. If empty, indicates that there are not too many items
                if ref[0] != 0:
                    blank = session.findById(f"wnd[0]/usr/lbl[95,{len(ref)+8}]").text
                    if blank == "":
                        invoice_amount_match = True
                    else:
                        invoice_amount_match = False
                        print("The amount of invoices doesn't match")
            else:
                left_lines = len_ref
                line_number = 7
                while line_number < left_lines + 7:
                    #session.findById(f"wnd[0]/usr/lbl[123,{line_number}]").setFocus()
                    try:
                        if session.findById(f"wnd[0]/usr/lbl[123,{line_number}]").text != "":
                            with_rc += 1
                        line_number += 1
                    except:
                        left_lines = left_lines - line_number + 7
                        if left_lines > 0:
                            session.findById("wnd[0]").maximize()
                            session.findById("wnd[0]").sendVKey(82)
                            line_number = 7
                        else:
                            break
                if not with_rc:
                    empty_rc = True
                else:
                    print("There is an item with reason code.")
                #session.findById(f"wnd[0]/usr/lbl[106,{line_number+2}]").setFocus()
                try:
                    items_sum = float(session.findById(f"wnd[0]/usr/lbl[106,{line_number+1}]").text.strip().replace("," ,""))
                    blank = session.findById(f"wnd[0]/usr/lbl[95,{line_number+1}]").text
                except:
                    #Handling cases where only the sum of the invoices is on the next page
                    session.findById("wnd[0]").maximize()
                    session.findById("wnd[0]").sendVKey(82)
                    items_sum = float(session.findById(f"wnd[0]/usr/lbl[106,8]").text.strip().replace("," ,""))
                    blank = session.findById(f"wnd[0]/usr/lbl[95,8]").text
                    
                #Ensuring the number of items in SAP match the number items from excel
                if blank == "":
                    invoice_amount_match = True
                else:
                    invoice_amount_match = False
                    print("The amount of invoices doesn't match")
        except:
            print("\nNot all of the invoices have been found.")
        print(items_sum)
        with_rc = 0
        #Selecting all
        #session.findById("wnd[0]/usr/lbl[1,5]").setFocus()
        session.findById("wnd[0]/usr/lbl[1,5]").caretPosition = 0
        session.findById("wnd[0]").sendVKey(5)
        #Opening mass change
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        #Filling in the text and reason code
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        #Creating description
        description = str()
        for piece in final_data[key][ref][:-1]:
            description += piece
            description += " "
        #Shortening too long descriptions by removing a date
        description_too_long = False
        if len(description) > 50:
            description = description[:-date_length]
            if len(description) > 50:
                description_too_long = True
                print("The description was far too long, you will have to fill it in manually.")
            else:
                print("The description was too long, the date has been removed")
        #If there are no such invoices in a customer account, this one goes wrong
        try:
            if description_too_long:
                session.findById("wnd[1]/usr/ctxt*BSEG-RSTGR").text = "ABD"
            else:
                session.findById("wnd[1]/usr/txt*BSEG-SGTXT").text = description
                session.findById("wnd[1]/usr/ctxt*BSEG-RSTGR").text = "ABD"
        except:
            print(ref, final_data[key][ref])
            problem = input("There are no such invoices in this account."
                                " You can open correct customer account and write y to continue, write n to skip the current detail or s to skip this customer altogether. If the customer is correct but there was a typo in the detail, save manually and write n.\n")
            while problem not in ["y","n","s"]:
                problem = input()
            if problem == "y":
                #Repeated lines from above till the next try
                session.findById("wnd[0]/usr/lbl[24,5]").caretPosition = 6
                session.findById("wnd[0]").sendVKey(2)
                session.findById("wnd[0]/tbar[1]/btn[38]").press()
                session.findById(r"wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press()
                for i, bare_ref in enumerate(ref):
                    session.findById(f"wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,{i}]")\
                    .text = "*" + str(bare_ref)
                    session.findById("wnd[2]/tbar[0]/btn[8]").press()
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
                    session.findById("wnd[0]/usr/lbl[1,5]").setFocus()
                    session.findById("wnd[0]/usr/lbl[1,5]").caretPosition = 0
                    #Opening mass change
                    session.findById("wnd[0]").sendVKey(5)
                    session.findById("wnd[0]/tbar[1]/btn[45]").press()
                    #Filling in the text and reason code
                    session.findById("wnd[0]").maximize()
                    session.findById("wnd[0]/tbar[1]/btn[45]").press()
                    #Creating description
                    description = str()
                    for piece in final_data[key][ref][:-1]:
                        description += piece
                        description += " "
                    session.findById("wnd[1]/usr/txt*BSEG-SGTXT").text = description
                    session.findById("wnd[1]/usr/ctxt*BSEG-RSTGR").text = "ABD"
            elif problem == "n":
                print(pd.DataFrame(final_data[key][ref]))
                continue
            elif problem == "s":
                skip = True
                break
        try:
            correct_amount = items_sum == float(final_data[key][ref][-1].replace("MAD", "").strip())
            if not correct_amount:
                print(f"Incorrect amount, diff {round(items_sum - float(final_data[key][ref][-1].replace("MAD", "").strip()), 2)}")
                
            if description_too_long:
                answer = input(f"{ref} {final_data[key][ref]} Fill in the description manually and write y/n.\n").lower()
            elif correct_amount and empty_rc:
                #Saving automatically
                print(f"Saved. {ref} {final_data[key][ref]}\n")
                answer = "y"
            else:
                answer = input(f"Execute mass change? y/n {ref} {final_data[key][ref]}\n").lower()
        except Exception as e:
            print(e)
        while answer not in ["y","n"]:
            answer = input("Write y or n.\n")
        if answer == "y":
            #Executing mass change
            try:
                session.findById("wnd[1]/usr/ctxt*BSEG-RSTGR").setFocus()
                session.findById("wnd[1]/usr/ctxt*BSEG-RSTGR").caretPosition = 3
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except:
                print("I haven't described")
                print(final_data[key][ref])
        elif answer == "n":
            try:
                print(pd.DataFrame(final_data[key][ref]))
                session.findById("wnd[1]").close()
            except:
                pass
    if skip:
        continue

ex = input("Write anything to exit the program\n").lower()

while ex not in list("{:c}".format(x) for x in range(97, 123)) + list(str(x) for x in range(10)):

    ex = input("Write anything to exit the program\n").lower()
