from docxtpl import DocxTemplate #module for working with docx
from datetime import datetime
import time
import pandas as pd #dataframe to work with csv file
import openai #ChatGpt
import os
import shutil #switching directory

def myGPT(prompt):
    #Insert your api, search for openai api key, if you don't have an account, create one. 
    api=""
    openai.api_key=api
    # Set up the model and prompt
    model_engine = "text-davinci-003"

    # Generate a response
    completion = openai.Completion.create(
        engine=model_engine,
        prompt=prompt,
        max_tokens=2000,
        n=1,
        stop=None,
        temperature=1,
    )
    response = str(completion.choices[0].text)
    return response


def main(template,company_list):
    doc=DocxTemplate(template)
    df = pd.read_excel(company_list)

    #Setting up constant key value pair irrelevant to the csv file constant.
    month_day_year = datetime.today().strftime("%d %b, %Y")
    my_context={"month_day_year":month_day_year}


    for index, row in df.iterrows():

        #key value pairs in docx and the GPT command in csv file.
        context = {"first_last_name": row['first_last_name'],
                   "phone": row['phone'],
                   "email": row['email'],
                   "linkedin": row['linkedin'],
                   "company_name": row['company_name'],
                   "company_address": myGPT(row['company_address']),
                   'company_position': row['position'],
                   'recruiter_name': row['recruiter_name'],
                   'field': row['field'],
                   'attribute_3': myGPT(row['attribute_3']),
                   "skill_1": myGPT(row['skill_1']),
                   "skill_2": myGPT(row['skill_2']),
                   "skill_3": myGPT(row['skill_3']),
                   "comp_skill_1": myGPT(row['comp_skill_1']),
                   "comp_skill_2": myGPT(row['comp_skill_2']),
                   'lang_1': myGPT(row['lang_1']),
                   'lang_2': myGPT(row['lang_2']),
                   'lang_3': myGPT(row['lang_3']),
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
    company_list="company_list.xlsx"

    start=time.time()
    main(template,company_list)
    end=time.time()
    print(f"{round((end-start)/60,2)} minutes")

