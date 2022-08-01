from time import sleep
import pywhatkit
from datetime import datetime, date, timedelta
import openpyxl as excel
##################################################


def readContacts(fileName):
    lst = []
    file = excel.load_workbook(fileName)
    sheet = file.active
    firstCol = sheet['A']
    for cell in range(len(firstCol)):

        contact = str(firstCol[cell].value)
        contact = "+"+contact
        lst.append(contact)
    return lst


phones = readContacts("contacts.xlsx")


#####################################################

########## hour = ساعت ارسال minute = دقیقه ارسال ###############
start = datetime.today().replace(hour=13, minute=55)

##### phones[:30] اینو با توجه به تعداد شماره های باید عوض کرد ############
for i, phone in enumerate(phones[:30]):

    ####timedelta = فاصله بین هر ارسال ####################
    scheduled_time = start + i * timedelta(minutes=0.3)


###########################################################
    pywhatkit.sendwhatmsg(phone, 'hello, how are you?',
                          scheduled_time.hour, scheduled_time.minute)


print('finish')
