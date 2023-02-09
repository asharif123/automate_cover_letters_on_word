#### create multiple cover letters ####################


##library to read word documents
from docxtpl import DocxTemplate
from datetime import datetime

##used to read data off excel files
import pandas as pd

##open document you wish to use
doc = DocxTemplate("my_word_template.docx")

##content to render
# context = {'company_name': "World company"}
my_name = "Awad Sharif"
my_address = "9761 Johannah Avenue"
city = "Garden Grove"
phone = 7145527522
email = "awadsharif9@gmail.com"
#convert string to time
date = datetime.now().strftime("%m/%d/%y")
# {{ recipient_name }}
# {{ recipient_title }}
# {{ company_name }}
# {{ company_address }}
# {{ company_city_and_zip_code }}
context = {'my_name': my_name, 'phone': phone, 'city': city, 'email': email, "my_address": my_address, "date": date}

##read info from csv file
df = pd.read_csv('fake_data.csv')

##read data off each row (prints row and corresponding index)
for index, row in df.iterrows():
##ex: parameters from word document: row['column name']
    my_context = {
         'recipient_name': row['name'],
         'email': row["email"],
         'my_address': row["address"],
         'phone': row["phone_number"],
         'company_name': row["company"]
    }
    ##update context with my_context having your info
    context.update(my_context)

    ##add data to the document
    doc.render(context)

    ##save updated info under new document
    doc.save(f"generated_doc_{index}.docx")

