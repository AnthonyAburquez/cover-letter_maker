from docxtpl import DocxTemplate #module for working with docx
from datetime import datetime
import time
import pandas as pd #dataframe to work with csv file
import openai #ChatGpt
import os
import shutil #switching directory

def myGPT(prompt):

    openai.api_key="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"
    # Set up the model and prompt
    model_engine = "text-davinci-003"

    # Generate a response
    completion = openai.Completion.create(
        engine=model_engine,
        prompt=prompt,
        max_tokens=2000,
        n=1,
        stop=None,
        temperature=0.5,
    )
    response = str(completion.choices[0].text)
    return response


def main(template,company_list):
    doc=DocxTemplate(template)
    df = pd.read_csv(company_list)

    #Setting up constant key value pair irrelevant to the csv file constant.
    month_day_year = datetime.today().strftime("%d %b, %Y")
    my_context={"month_day_year":month_day_year}


    for index, row in df.iterrows():

        #key value pairs in docx and the GPT command in csv file.
        context = {"company_name": row['company_name'],
                   "company_address": myGPT(row['company_address']),
                   'company_position': row['position'],
                   'recruiter_name': row['recruiter_name'],
                   'field': row['field'],
                   'attribute_3': myGPT(row['attribute_3']),
                   'lang_1': row['lang_1'],
                   'lang_2': row['lang_2'],
                   'lang_3': row['lang_3'],
                   'why_company': myGPT(row['why_company']),
                   }
        #updating the the dictionary to incorporate unrelated to contents of cover letter
        context.update(my_context)

        doc.render(context)

        #Setting up the directories
        filename=f"{row['company_name']}_Cover_Letter.docx"
        path = os.getcwd()
        new_path=path+"/Cover_Letters"

        #Saving the docx in an another folder
        doc.save(filename)
        shutil.move(f"{path}\{filename}",new_path)


if __name__=="__main__":
    template="cover_letter.docx"
    company_list="company_list.csv"

    start=time.time()
    main(template,company_list)
    end=time.time()
    print(f"{round((end-start)/60,2)} minutes")

