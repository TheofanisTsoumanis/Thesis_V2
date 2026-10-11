
from input_voice_functions import *
from messages import *
from config import *
from docxtpl import DocxTemplate
import subprocess
import datetime 
import time
import os
import win32com.client
import sys


# Να κλείνει το πρόγραμμα
def ProgramTermination(newinput, lang):
    if newinput in PROGRAM_TERMINATION_COMMANDS[lang]:
        print(PROGRAM_TERMINATION_SALUTE[language])
        sys.exit(0)

# Να μην επιτρέπει λεξιλόγιο
def InapropriateWords(newinput, lang):
    if newinput in INAPROPRIATE_WORDS[lang]:
        return True

# έλεγχος για None απάντηση
def NoneCheck(newinput):
        return newinput is None

# για την περίπτωση που δεν απαντήσω χάνεται το μήνυμα του χρήστη.
def InputValidation(lang, question_text):
    while True:
        #print('Calling the function for Speech')
        newinput = Speech(lang)
        #print('Check for None words')
        newinput_none = NoneCheck(newinput)
        if newinput_none:
            print(question_text)
            continue
        #print('Check for program termination')
        ProgramTermination(newinput, language)
        #print('check for inapropriate words')
        newinput_inapropriate = InapropriateWords(newinput, language)
        if newinput_inapropriate:
            continue
        #print('field type\ncomming soon!')
        return newinput

# έλεγχος για valid choice    
def GetValidChoice(language, lang, valid_choice):
    while True:
        answer = Speech(lang)
        if answer in valid_choice:
            return answer
        print(MESSAGES[language]["valid_choice"])

def select_language():
    while True:
        present_time = datetime.datetime.now()
        text_gr, text_en = Select_Language()
        if text_gr in LANGUAGE_COMMANDS["el"] or text_en in LANGUAGE_COMMANDS["el"]:
            language = 'el'
        elif text_en in LANGUAGE_COMMANDS["en"] or text_gr in LANGUAGE_COMMANDS["en"]:
            language = 'en'
        elif (text_gr in PROGRAM_TERMINATION_COMMANDS['el']) or (text_gr in PROGRAM_TERMINATION_COMMANDS['en']):
            sys.exit(0)
        elif (text_en in PROGRAM_TERMINATION_COMMANDS['el']) or (text_en in PROGRAM_TERMINATION_COMMANDS['en']):
            sys.exit(0)
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

# create new template
def Template_Creation():
    pass

# επιλογή template
def Template_Selection():
    pass
    print(MESSAGES[language]['choose_template'])
    time.sleep(0.5)
    valid_templates = ['πτυχιακή εργασία', 'thesis', 'αίτηση', 'application']
    choice = GetValidChoice(language, lang, valid_templates)

    # Δημιουργία λιστών
    if choice == 'πτυχιακή εργασία' or choice == 'thesis':
        choice_list = ['name', 'last_name', 'fathers_name', 'university_registration_number', 'address', 'postal_code', 'city', 'phone', 'email', 'professors_name', 'thesis_topic']
    elif choice == 'αίτηση' or choice == 'application':
        choice_list = ['name', 'last_name', 'fathers_name', 'mothers_name', 'university_registration_number', 'semester', 'street_name', 'street_number', 'postal_code', 'city', 'phone', 'mobile_phone']
    newinputlist = []

    # προσπέλαση της λίστας με ονόματα των templates
    for item in choice_list:
        print(MESSAGES[language][item])
        question_text = MESSAGES[language][item]

        if item == 'email':
            newinput = input(MESSAGES[language]['email_type'])
            newinputlist.append(newinput)
            continue

        newinput = InputValidation(lang, question_text)
        print(newinput)
        newinputlist.append(newinput)

    # Πρόσθετη επεξεργασία για το application
    yes_list = ['ναι', 'yes']

    if choice == 'αίτηση' or choice == 'application':
        # Τουλάχιστον 1
        # Τι χρειάζεται ο χρήστης
        what_need = [
            'certificate_of_studies',
            'detailed_score',
            'other'
        ]
        selected_what_need = []

        loop_again = True
        while loop_again:
            for item in what_need:
                while True:
                    print(MESSAGES[language][item])
                    question_text = MESSAGES[language][item]
                    newinput = InputValidation(lang, question_text)
                    
                    if newinput in yes_list:
                        if item == 'other':
                            print(DOCUMENT_VALUES[language][item])
                            first_message = f'{DOCUMENT_VALUES[language]['other_label']}'
                            second_message = InputValidation(lang, question_text)
                            full_message_other = first_message + str(second_message)
                            selected_what_need.append(full_message_other)
                            break
                        selected_what_need.append(DOCUMENT_VALUES[language][item])
                        break
                    else:
                        break
            if not selected_what_need:
                print()
                print(MESSAGES[language]['application_loop_select_at_least_one'])
                time.sleep(1)
                continue
            else:
                loop_again = False

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

        loop_again = True
        while loop_again:
            for item in what_for:
                while True:
                    print(MESSAGES[language][item])
                    question_text = MESSAGES[language][item]
                    newinput = InputValidation(lang, question_text)

                    if newinput in yes_list:
                        selected_what_for.append(DOCUMENT_VALUES[language][item])
                        break
                    else:
                        break

            if not selected_what_for:
                print()
                print(MESSAGES[language]['application_loop_select_at_least_one'])
                time.sleep(1)
                continue
            else:
                loop_again = False

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
    time.sleep(1.5)

    # Ονομασία, επιλογή αρχείου και αποθήκευση του.
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    filename = f"{choice}_{data['name']}_{data['last_name']}_{timestamp}.docx"

    if choice == 'πτυχιακή εργασία':  
        doc = DocxTemplate("templates/ΠΤΥΧΙΑΚΗ_doc.docx")

    elif choice == 'thesis':
        doc = DocxTemplate("templates/THESIS_doc.docx")

    elif choice == 'αίτηση':
        doc = DocxTemplate("templates/ΑΙΤΗΣΗ_doc.docx")

    elif choice == 'application':
        doc = DocxTemplate("templates/APPLICATION_doc.docx")

    doc.render(data)
    os.makedirs("saved_files", exist_ok=True)
    docx_path = f"saved_files/{filename}"
    doc.save(docx_path)

    print(f'Το αρχείο αποθηκεύτηκε στη συσκευή σας επιτυχώς με όνομα {filename}.')

    print(MESSAGES[language]['doc_convert'])
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

# επιλογή γλώσσας
print('Ελληνικά - English')
language = select_language()
lang = SPEECH_LANGUAGES[language]

# Δημιουργία menu για 1) Δημιουργία εγγράφου 2) Επιλογή Έτοιμου template 3) Έξοδος από την εφαρμογή.
menu_options = f'''{MESSAGES[language]['template_creation']},
{MESSAGES[language]['template_choice']},
{MESSAGES[language]['template_exit']},
{MESSAGES[language]['template_info']}'''


while True:
    print(menu_options)
    question_text = menu_options
    newinput = InputValidation(lang, question_text)

    match newinput:
        case _ if newinput in( "1", MENU_ANSWERS[language]['template_creation_answer']):
            print('entered template creation')
            Template_Creation()
        case _ if newinput in( "2", MENU_ANSWERS[language]['template_choice_answer']):
            print('entered template choice')
            Template_Selection()
        case _ if newinput in( "3", MENU_ANSWERS[language]['template_exit_answer']):
            break
        case _ if newinput in( "4", MENU_ANSWERS[language]['template_info_answer']):
            # Μπορεί να είναι και έτοιμο pdf template που το κάνει generate για να το έχουν σε ξεχωριστό παράθυρο
            print('info for navigation.')
        case _:
            print(MESSAGES[language]['template_invalid_choice'])





