
from input_voice_functions import *
from messages import *
from config import *
from docxtpl import DocxTemplate
import subprocess
import datetime 
import time
import os
import win32com.client


def select_language():
    while True:
        present_time = datetime.datetime.now()
        text_gr, text_en = Select_Language()

        if text_gr in LANGUAGE_COMMANDS["el"] or text_en in LANGUAGE_COMMANDS["el"]:
            language = 'el'
        elif text_en in LANGUAGE_COMMANDS["en"] or text_gr in LANGUAGE_COMMANDS["en"]:
            language = 'en'
        else:
            print(f'Πείτε: Ελληνικά\nSay: English')
            continue
        print(MESSAGES[language]['language_selected'])
        if present_time.hour < 12 :
            print(MESSAGES[language]['morning'])
        else:
            print(MESSAGES[language]['evening'])
        time.sleep(1)
        break

    return language

# έλεγχος για None απάντηση
def ValueCheck(newinput, lang):
    while True:
        if newinput is not None:
            return newinput
        
        newinput = Speech(lang)

# έλεγχος για valid choice    
def GetValidChoice(language, lang, valid_choice):
    while True:
        answer = Speech(lang)

        if answer in valid_choice:
            return answer

        print(MESSAGES[language]["valid_choice"])
    
# Να κλείνει το πρόγραμμα

# Να μην επιτρέπει λεξιλόγιο


# επιλογή γλώσσας
print('Ελληνικά - English')
language = select_language()
lang = SPEECH_LANGUAGES[language]

# Δημηουργία menu για 1) Δημιουργία εγγράφου 2) Επιλογή Έτοιμου template 3) Έξοδος από την εφαρμογή.

# επιλογή template
print(MESSAGES[language]['choose_template'])
time.sleep(0.5)
valid_templates = ['Πτυχιακή', 'Thesis', 'Αίτηση', 'Application']
choice = GetValidChoice(language, lang, valid_templates)

# Δημιουργία λιστών
if choice == 'Πτυχιακή' or choice == 'Thesis':
    choice_list = ['name', 'last_name', 'fathers_name', 'registration_number', 'address', 'postal_code', 'city', 'phone', 'email', 'professors_name', 'thesis_topic']
elif choice == 'Αίτηση' or choice == 'Application':
    choice_list = ['name', 'last_name', 'fathers_name', 'mothers_name', 'registration_number', 'semester', 'street_name', 'street_number', 'postal_code', 'city', 'phone', 'mobile_phone']

newinputlist = []

# προσπέλαση της λίστας με ονόματα των templates
for item in choice_list:
    print(MESSAGES[language][item])
    if item == 'email':
        newinput = input(MESSAGES[language]['email_type'])
        # επιλογή κατάλληξης email
        newinput = str(newinput) + '@gmail.com'
        newinputlist.append(newinput)
        continue
    newinput = Speech(lang)
    newinput = ValueCheck(newinput, lang)
    newinputlist.append(newinput)

# Πρόσθετη επεξεργασία για το application
yes_list = ['ναι', 'yes']

# Τι χρειάζεται ο χρήστης
what_need = [
    'certificate_of_studies',
    'detailed_score',
    'other'
]
selected_what_need = []

for item in what_need:
    print(MESSAGES[language][item])
    newinput = Speech(lang)
    newinput = ValueCheck(newinput, lang)
    
    if newinput in yes_list:
        if item == 'other':
            print(DOCUMENT_VALUES[language][item])
            newinput = Speech(lang)
            first_messsage = f'{DOCUMENT_VALUES[language]['other_label']}'
            second_message = ValueCheck(newinput, lang)
            full_message_other = first_messsage + str(second_message)
            selected_what_need.append(full_message_other)
            continue
        selected_what_need.append(DOCUMENT_VALUES[language][item])
    else:
        pass

what_need_value = '\n'.join(selected_what_need)
choice_list.append('what_need')
newinputlist.append(what_need_value)

# Για πιο λόγο το θέλει ο χρήστης
what_for = [
    'every_use',
    'tax_office',
    'recruitment'
]
selected_what_for = []

for item in what_for:
    print(MESSAGES[language][item])
    newinput = Speech(lang)
    newinput = ValueCheck(newinput, lang)

    if newinput in yes_list:
        selected_what_for.append(DOCUMENT_VALUES[language][item])

what_for_value = '\n'.join(selected_what_for)
choice_list.append('what_for')
newinputlist.append(what_for_value)

# Ημερομηνία
present_date = datetime.datetime.now()
choice_list.append('day')
newinputlist.append(present_date.day)
choice_list.append('month')
newinputlist.append(present_date.month)
choice_list.append('year')
newinputlist.append(present_date.year)

# δημιουργία λεξικού για να το κάνω στην πορεία αντικατάσταση.
data = dict(zip(choice_list, newinputlist))
print()
for key, value in data.items():
    print(f'{key}: {value}')


# Ονομασία, επιλογή αρχείου και αποθήκευση του.
timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

filename = f"{choice}_{data['name']}_{data['last_name']}_{timestamp}.docx"

if choice == 'Πτυχιακή':  
    doc = DocxTemplate("templates/ΠΤΥΧΙΑΚΗ_doc.docx")

elif choice == 'Thesis':
    doc = DocxTemplate("templates/THESIS_doc.docx")

elif choice == 'Αίτηση':
    doc = DocxTemplate("templates/ΑΙΤΗΣΗ_doc.docx")

elif choice == 'Application':
    doc = DocxTemplate("templates/APPLICATION_doc.docx")

doc.render(data)
os.makedirs("saved_files", exist_ok=True)
docx_path = f"saved_files/{filename}"
doc.save(docx_path)

print(f'Το αρχείο αποθηκεύτηκε στη συσκευή σας επιτυχώς με όνομα {filename}.')
time.sleep(1.5)

# Ανίχνευση για microsoft word ή libreoffice καθώς και έλεγχος ορθής λειτουργίας τους.
filename = f"{choice}_{data['name']}_{data['last_name']}_{timestamp}.pdf"
try:
    word = win32com.client.Dispatch("Word.Application")
    document = word.Documents.Open(os.path.abspath(docx_path))
    pdf_path = os.path.abspath(f"saved_files/{filename}")
    document.SaveAs(pdf_path, FileFormat=17)
    document.Close()
    word.Quit()
    os.startfile(pdf_path)

except Exception as e:
    print(f"Microsoft Word: {type(e).__name__}: {e}")
    try:
        libreoffice_path = r"C:\Program Files\LibreOffice\program\soffice.exe"
        print(os.path.exists(libreoffice_path))

        subprocess.run([
        libreoffice_path,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", "saved_files",
        docx_path
        ], check=True)
        pdf_path = os.path.abspath(f"saved_files/{filename}")
        os.startfile(pdf_path)

    except Exception as e:
        print(f"LibreOffice: {type(e).__name__}: {e}")
        print("You don't have the right tool to convert to PDF!\nI recomment LibreOffice!")
print("The process is completed.")