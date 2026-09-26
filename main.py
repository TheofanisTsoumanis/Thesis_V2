
from input_voice_functions import *
from messages import *
from config import *
from docxtpl import DocxTemplate
import subprocess
import datetime 
import time
import os




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

# επιλογή template
print(MESSAGES[language]['choose_template'])
time.sleep(0.5)
valid_templates = ['Πτυχιακή', 'Thesis', 'Αίτηση', 'Application']
choice = GetValidChoice(language, lang, valid_templates)

# Δημιουργία λιστών
if choice == 'Πτυχιακή' or choice == 'Thesis':
    choice_list = ['name', 'last_name', 'fathers_name', 'registration_number', 'address', 'postal_code', 'city', 'phone', 'email', 'professors_name', 'thesis_topic']
elif choice == 'Αίτηση' or choice == 'Application':
    choice_list = ['name', 'last_name', 'require']

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

# Μετατροπή Word σε PDF
libreoffice_path = r"C:\Program Files\LibreOffice\program\soffice.exe"
subprocess.run([
    libreoffice_path,
    "--headless",
    "--convert-to", "pdf",
    "--outdir", "saved_files",
    docx_path
])


print(f'Το αρχείο αποθηκεύτηκε στη συσκευή σας επιτυχώς με όνομα {filename}.')
time.sleep(1.5)

