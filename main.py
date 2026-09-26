
from input_voice_functions import *
from messages import *
from config import *
from docxtpl import DocxTemplate
import subprocess
import datetime 
import time
import os


libreoffice_path = r"C:\Program Files\LibreOffice\program\soffice.exe"

def select_language():
    while True:
        present_time = datetime.datetime.now()
        text_gr, text_en = Select_Language()

        if text_gr is None and text_en is None:
            continue

        if text_gr == 'Ελληνικά' or text_en == 'Greek':
            print('Επιλέξατε Ελληνικά')
            if present_time.hour < 12 :
                print('Καλημέρα')
            else:
                print('Καλησπέρα')
            time.sleep(1)
            language = 'el'
            break

        elif text_gr == 'Αγγλικά' or text_en == 'English':
            print('You chose English')
            if present_time.hour < 12 :
                print('Good morning')
            else:
                print('Good evening')
            time.sleep(1)
            language = 'en'
            break
        else:
            print('Ελληνικά or English')

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
    

# επιλογή γλώσσας
print('Ελληνικά - English')
language = select_language()
lang = SPEECH_LANGUAGES[language]

# επιλογή template
print(MESSAGES[language]['choose_template'])
time.sleep(1)
valid_templates = ['Πτυχιακή', 'Thesis', 'Αίτηση', 'Application']
choice = GetValidChoice(language, lang, valid_templates)

if choice == 'Πτυχιακή' or choice == 'Thesis':
    choice_list = ['name', 'last_name', 'fathers_name', 'registration_number', 'address', 'postal_code', 'city', 'phone', 'email', 'professors_name', 'thesis_topic']
elif choice == 'Αίτηση' or choice == 'Application':
    choice_list = ['name', 'last_name', 'require']

newinputlist = []

# προσπέλαση της λίστας με ονόματα των templates
for item in choice_list:
    print(MESSAGES[language][item])

    newinput = Speech(lang)
    newinput = ValueCheck(newinput, lang)
    if item == 'email':
        newinput = str(newinput) + '@gmail.com'
    newinputlist.append(newinput)


# δημιουργία λεξικού για να το κάνω στην πορεία αντικατάσταση.
data = dict(zip(choice_list, newinputlist))
print()
for key, value in data.items():
    print(f'{key}: {value}')



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
subprocess.run([
    libreoffice_path,
    "--headless",
    "--convert-to", "pdf",
    "--outdir", "saved_files",
    docx_path
])


print(f'Το αρχείο αποθηκεύτηκε στη συσκευή σας επιτυχώς με όνομα {filename}.')
time.sleep(1.5)

# αντικατάσταση και αποθήκευση.