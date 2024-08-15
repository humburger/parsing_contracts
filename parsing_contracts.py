from docx import Document
import xlwt
import docx2txt
import os, re, sys
import win32com.client as win32
from win32com.client import constants

CONST_DIR_SLASH = "\\"
CONST_FILE_TYPE_DOCX = ".docx"
regEx_old_contract = r"^sadarbibas ligums\.doc$"
regEx_old_attachment = r"^.*pielikums.*.doc$"
regEx_tempfile = r"^~.*"
regEx_new_contract = r"^.*\.doc$"
regEx_doc = r"\.doc$"

"""
    Svarīgi ir norādīt, ka regEx darbojas ar dokumentiem tikai tad, kad tiem ir izslēgtas/apstiprinātas visas IZSEKOTĀS IZMAIŅAS/labojumi (sarkanās malu līnijas malu tekstiem/paragrāfiem, kur ir redzams, ka kāds ir svītrojis vai arī kaut ko labojis)
    
    Informācijas iztrūkums dažās excel šūnās norāda uz to, ka iespējams ir jāpapildina kādu regEx izteksmi VAI arī vienkārši tāda veida infromācijas patterns/šablons vienkārši nav atrodams failā kā tādā
    
    N.B. Ja ir kļūda, ka nevar saglabāt jau esošu word failu, tad TaskMngr logā ir jāislēdz MS Word procesu, kas ir palicis ieslēgts kaut kur fonā.
    N.B. Ja ir kļūda, ka nevar peikļūt failam, nedŗikst/nevar rediģēt, iespējams uz to brīdi ar roku ir atvērts atteicīgais dokumnets (piem., testē kodu un pārbauda atvērtu excel rezultātu), to tad ir jāaizver, jo python nevar izmainīt jau atvērtu failu (izņemot, ja tas ir vienkāršs txt fails un to darbina ar notepad vai ko citu vienkāršu, kā notepad++)
"""

regEx_SIA = re.compile( r"SIA.+,|AS.+,|Sabiedrība ar ierobežotu atbildību.+,|Saimnieciskās darbības veicēj.+,|Akciju sabiedrība.+,|APP.+,|Z/S.+,|ZS.+,|IK.+,|Biedrība.+,|MPKS.+,|Mežsaimniecības pakalpojumu kooperatīvā sabiedrība.+," )
# regEx_SIA = re.compile( r"SIA.+," )
regEx_CONTACT_INFO = re.compile( r"SIA.+tālrunis \d+.+adrese.+@.+\.\w{2,}|Sabiedrības.+tālrunis \d+.+adrese.+@.+\.\w{2,}|Pārvaldes.+tālrunis \d+.+adrese.+@.+\.\w{2,}|DAP.+tālrunis \d+.+adrese.+@.+\.\w{2,}|Lietotāja.+tālrunis \d+.+adrese.+@.+\.\w{2,}|APP.+tālrunis \d+.+adrese.+@.+\.\w{2,}|Puses.+tālrunis \d+.+adrese.+@.+\.\w{2,}|Z/S.+tālrunis \d+.+adrese.+@.+\.\w{2,}|Sabiedrības.+tālrunis \+372\s?\d+.+adrese.+@.+\.\w{2,}|ZS.+tālrunis \d+.+adrese.+@.+\.\w{2,}|IK.+tālrunis \d+.+adrese.+@.+\.\w{2,}|Biedrība.+tālrunis \d+.+adrese.+@.+\.\w{2,}|MPKS.+tālrunis \d+.+adrese.+@.+\.\w{2,}|Mežsaimniecības pakalpojumu kooperatīvā sabiedrība.+tālrunis \d+.+adrese.+@.+\.\w{2,}")
regEx_CONTRACT_EXPIRATION_DATE = re.compile( r"Līgums stājas spēkā ar brīdi, kad to ir parakstījušas abas Puses un ir noslēgts līdz \d{4}.gada \d{1,}.\w+|Līgums stājas spēkā ar brīdi, kad to ir parakstījuši abi līdzēji un ir noslēgts bez termiņa ierobežojuma|Sadarbības līgums stājas spēkā tā abpusējas parakstīšanas dienā un tiek noslēgts uz nenoteiktu laiku" )
regEx_EXTRA_PHONE = re.compile(r"Tālrunis:* (?!67509545)\d+|Tālrunis:* \+371\s?\d+|Tālrunis:* \+372\s?\d+")
regEx_EXTRA_EMAIL = re.compile(r"e-pasts:* (?!pasts@daba.gov.lv)\w+@\w+\.\w{2,}|e-pasts:* (?!pasts@daba.gov.lv)\w+.\w+@\w+\.\w{2,}")
regEx_REG_NUMBER = re.compile(r"Reģ\.\s*Nr\.*\s*(?!90009099027)\d+|Reģ\.\s*Nr\.*\s*\w{2}\d{11}|Personas kods.*\d+-\d+|Reģ\.\s*Nr\.*\sIgaunijas uzņēmumu reģistrā: \d+")

# 0, 1, 2, 3, 4, 5
excel_columns = ["Nosaukums", "Reg.Nr. vai personas kods", "Sabiedrības kontakti (vārds uzvārds, telefons, epasts)", "Pārvaldes kontakti", "Extra sabiedrības kontakti (telefons, epasts)", "Līguma termiņš"]
excel_range = range(len(excel_columns))
CONST_EXCEL_NAME = "Ozols lietotāju izveides līgumu kopsavilkums"

directory = os.getcwd()
# if directory dont have any subdir, then by logic of code we still can work in working main dir
working_dir_list = [directory + CONST_DIR_SLASH]
doc_extension = 0
doc_SIA = []
doc_contacts = []
doc_reg_number = []
doc_expiration_date = []
doc_extra_phone = []
doc_extra_email = []
contract_info = []

class Contract_args:
  def __init__(self, name, reg_number, contacts, expiration_date, extra_phone, extra_email, working_dir):
    self.name = name
    self.reg_number = reg_number
    self.contacts = contacts
    self.expiration_date = expiration_date
    self.extra_phone = extra_phone
    self.extra_email = extra_email
    self.working_dir = working_dir  # lai vieglāk saprastu, kur varbūt nav pilnvērtīgi nolasīti dati un var atrast dokumentu

# ideja no stackoverflow.com/a/48832989
# def doc_2_docx(dir, doc_extension):
def doc_2_docx(dir):
    # if (doc_extension == 1):
    # opens MS word app
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc_file = word.Documents.Open(dir)
    doc_file.Activate ()
    
    # changes full path name for file with 'doc' to 'docx'
    docx_file = dir
    docx_file = re.sub(regEx_doc, CONST_FILE_TYPE_DOCX, docx_file)
    
    # saves with new file format and closes MS word app
    word.ActiveDocument.SaveAs(docx_file, FileFormat=constants.wdFormatXMLDocument)
    doc_file.Close(False) # ja nav uzstādīts, tad vēlāk kaut kā skripts nevar rediģēt šos pašus dokumentus VAI arī MS Drive dublicē (maisās pa vidu) un tajā brīdī arī rodas pieejamības kļūda (dont know for sure)
    return
    # else:
        # print("There was no 'doc' file!")
        
def extract_data(dir):
    docx_file_path = ""
    
    # doc_2_docx(dir, doc_extension);
    doc_2_docx(dir);
    # ### for now its only for testing
    # if doc_extension == 0:
        # sys.exit()
    # ###    
    docx_file_path = dir + "x"

    docx_file = docx2txt.process(docx_file_path)

    doc_SIA = regEx_SIA.findall(docx_file)
    doc_reg_number = regEx_REG_NUMBER.findall(docx_file)
    doc_contacts = regEx_CONTACT_INFO.findall(docx_file)
    doc_expiration_date = regEx_CONTRACT_EXPIRATION_DATE.findall(docx_file)
    doc_extra_phone = regEx_EXTRA_PHONE.findall(docx_file)
    doc_extra_email = regEx_EXTRA_EMAIL.findall(docx_file)

    # uzskatāmībai dati ir sagalbāti objektā, lai ir nedalāmība
    c1 = Contract_args(doc_SIA, doc_reg_number, doc_contacts, doc_expiration_date, doc_extra_phone, doc_extra_email, dir)
    # tiks veidots objektu datu saraksts
    contract_info.append(c1)
    return

# kaut cik pārveido datus, lai vēlāk ekselī izskatītos ok
def edit_data(index):
    # regEx_sia_name = re.compile(r"SIA.+,|Sabiedrība ar ierobežotu atbildību.+,")
    regEx_fullname = re.compile(r"[A-ZĀ-Ž]{1}[a-zā-ž]+ [A-ZĀ-Ž]{1}[a-zā-ž]+-[A-ZĀ-Ž]{1}[a-zā-ž]+,|[A-ZĀ-Ž]{1}[a-zā-ž]+ [A-ZĀ-Ž]{1}[a-zā-ž]+,") # komats beigās, jo citādi dokumentā atrod citas vēr'tibas, kas nav vārds un uzvārds // dubultā uzvārda versija
        # vārds uzvārds-uzvārds
        # vārds uzvārds
        
    regEx_phone = re.compile(r"\s\d{8}|\+371\s?\d+|\+372\s?\d+") # N.B. \s simbols pirms telefona numura atvieglo to, ka tālāk datu apstrādē telefona numuriem excel failā būs atstarpe pirms tiem, jo ir gadījumi, kurus uz momentu nevaru izskatīt, lai visiem numuriem būtu atstarpes pirms tiem, piem., vairāki kā viens utml.
    regEx_digit = re.compile(r"\d+|Igaunijas uzņēmumu reģistrā: \d+")
    
    regEx_email = re.compile(r"\s\w+\.\w+@\w+\.\w+\.\w{2,}|\w+\.\w+@\w+\.\w+\.\w{2,}|\s\w+\.\w+-\w+@\w+\.\w{2,}|\w+\.\w+-\w+@\w+\.\w{2,}|\s\w+\.\w+@\w+\.\w{2,}|\w+\.\w+@\w+\.\w{2,}|\w+@\w+\.\w{2,}") # citi epasti un vards.uzvards@daba.gov.lv; regexim laikam ir jāievēro secību, ka vispirms atrod garās versijas un tad meklē īsās, jo ja ir otrādāk, tad rezultātos atdod no garajiem epasteim to 'īsās' versijas
        # \svārds.uzvārds@daba.gov.lv
        # vārds.uzvārds@daba.gov.lv
        # \svārds.uzvārds-uzvārds@sia.lv
        # vārds.uzvārds-uzvārds@sia.lv
        # \svārds.uzvārds@sia.lv
        # vārds.uzvārds@sia.lv
        # \svārds@sia.lv
        # vārds@sia.lv
    
    regEx_exp_date = re.compile(r"bez termiņa ierobežojuma|uz nenoteiktu laiku|\d{4}.gada \d{1,}.\w+")
    temp_str = ""
    temp_arr = []
    
    # print(contract_info[index].name)
    # print(contract_info[index].reg_number)
    # print(contract_info[index].contacts)
    # print(contract_info[index].extra_phone)
    # print(contract_info[index].extra_email)
    # print(contract_info[index].expiration_date)
    
    temp_str = "".join(contract_info[index].name) # makes list as one whole string
    temp_str = temp_str.split(",")[0]
    temp_str = temp_str.replace("Sabiedrība ar ierobežotu atbildību", "SIA") # string replace(), ja neko neatrod, tad arī neko nemaina
    temp_str = temp_str.replace("Akciju sabiedrība", "AS")
    temp_str = temp_str.replace("Mežsaimniecības pakalpojumu kooperatīvā sabiedrība", "MPKS") 
    temp_str = temp_str.replace(" (turpmāk – Sabiedrība)", "")
    temp_str = temp_str.replace(" (turpmāk – Lietotājs)", "")
    temp_arr.append(temp_str)
    contract_info[index].name = temp_arr
    
    contract_info[index].reg_number = regEx_digit.findall(str(contract_info[index].reg_number))
    contract_info[index].extra_phone = regEx_phone.findall(str(contract_info[index].extra_phone))
    contract_info[index].extra_email = regEx_email.findall(str(contract_info[index].extra_email))
    contract_info[index].contacts[0] = regEx_fullname.findall(str(contract_info[index].contacts[0])) + regEx_phone.findall(str(contract_info[index].contacts[0])) + [", "] + regEx_email.findall(str(contract_info[index].contacts[0])) # konkatenācija noteik starp sarakstiem/array nevis stringiem, tāpēc ir +[" "]+ nevis +" "+, kā tas ir pierasts ar stingu konkatenāciju
    if len(contract_info[index].contacts) > 1:
        contract_info[index].contacts[1] = regEx_fullname.findall(str(contract_info[index].contacts[1])) + regEx_phone.findall(str(contract_info[index].contacts[1])) + [", "] + regEx_email.findall(str(contract_info[index].contacts[1]))
    
    contract_info[index].expiration_date = regEx_exp_date.findall(str(contract_info[index].expiration_date))
    
    # print(contract_info[index].name)
    # print(contract_info[index].reg_number)
    # print(contract_info[index].contacts)
    # print(contract_info[index].extra_phone)
    # print(contract_info[index].extra_email)
    # print(contract_info[index].expiration_date)
    
    return
    
# will get array with listdir and does it recursively
def read_folder_list(dir, dir_slash):

    dir_list = os.listdir(dir)
    for dir_incr in dir_list:
        next_dir = dir + dir_slash + dir_incr
        
        if os.path.isdir(next_dir):
        
            working_dir_list.append(next_dir)
            read_folder_list(next_dir, dir_slash)
        
    # print("script end")
    return
    
def export_data_to_excel(excel_new_name):
    # izveido svaigu excel failu, bet vēl nav saglabāts
    book = xlwt.Workbook()
    sheet1 = book.add_sheet("Sheet 1", cell_overwrite_ok=True) # šis atļaus pārraskt'ti pāri šunām vairākas reizes

    # izveido kolonu nosaukumus
    for y in range(len(excel_columns)):
        sheet1.row(0).write(y, excel_columns[y])
        
    for x, value in enumerate(contract_info):
        row = sheet1.row(x + 1) # pirmā rinda jau ir aizņemta ar kolonas nosaukumiem, tāpēc viss ir par vienu rindu uz priekšu
        for y in excel_range:
            # checks if list is not empty
            if value.name:
                row.write(0, value.name[0])
            row.write(1, value.reg_number)
            if len(value.contacts) > 1:
                row.write(2, value.contacts[1])  # sabiedrības kontakti
            row.write(3, value.contacts[0])  # pārvaldes kontakti
            if value.extra_email:
                row.write(4, value.extra_phone + [", "] + value.extra_email)
            else:
                row.write(4, value.extra_phone)
            row.write(5, value.expiration_date)
            row.write(6, value.working_dir)
            
    book.save(excel_new_name + ".xls") # works on old format excels, dont know why, bur for now it is enought
        
### this will be main script foundation of while loop

# directory ecomes 'item' from workig_dir_list
read_folder_list(directory, CONST_DIR_SLASH)
for item in working_dir_list:
    print(item)
    # directory mainīsies atkarībā no direktoriju saraksta, ka ir pirms tam iegūts
    file_names = os.listdir(item)
    # print(file_names)
    doc_file = ""

    # file_names ir atseviķi katrai direktrijai savs!!!!!!!!!!!!!!

    # print("working dir: \n \t" + item)
    for d in file_names:
        x = re.search(regEx_doc, d)
        y = re.search(regEx_old_attachment, d)
        z = re.search(regEx_tempfile, d)
        
        # atrod tikai vienu galveno līguma failu
        if x is not None and y is None and z is None:
            # print("file name: \n \t" + d)
            doc_file = d
            
            extract_data(item + CONST_DIR_SLASH + doc_file)
            edit_data(doc_extension)
            # šo vēl izdomāšu, kur nomainīt uz 1, lai var lietot ciklu
            doc_extension = doc_extension + 1
            # break
        else:
            pass
    
        
print('\nProcessed doc file count: ' + str(doc_extension))

export_data_to_excel(CONST_EXCEL_NAME)

### TAGA TIKAI IR JĀPIELITO IZDRUKU EKSELĪ UN TAD VAR MĒĢINĀT REKURSIJAS PA VISĀM MAPĒM (bet to atsevišķi ir jānotestē ar tukšo programmu)




