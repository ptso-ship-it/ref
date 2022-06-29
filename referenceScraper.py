import docx

def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)

import pandas as pd
import re

def get_all_lines(infile):
    lines = []
    for result in re.findall('##Top(.*?)##Bottom', text, re.S):
        lines.append(result)
    return lines

file = 'Reference_Clean.docx'
text = getText(file)
lines = get_all_lines(text)

countries = []
for line in lines:
    countryinfo = re.findall('#Start(.*?)#End', line, re.S)
    countries += countryinfo


country_parameters = "#Country(.*?)Tax authority and relevant"

beps_parameters = "implementation overview(.*?)c\) Is"

mcaa_parameters = "Multilateral Competent Authority(.*?)4. Transfer pricing"

contemp_parameters = "submitted or prepared contemporaneously(.*?)Does a local branch"


def processInfo(searchCriteria, column_name, countries):
    def findInfo(searchCriteria, country_string):
        result = re.search(rf'{searchCriteria}', country_string, re.S)
        if result:
            return result.group(0)
        else:
            return "No match found"

    column_values = []
    for country in countries:
        column_values.append(findInfo(searchCriteria, country))
    df[column_name] = column_values

df = pd.DataFrame()

processInfo(country_parameters, "country_names", countries)
processInfo(beps_parameters, "beps_status", countries)

df.to_csv('reference.csv', index =False)


