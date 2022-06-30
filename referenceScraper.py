import docx
import pandas as pd
import re
import unicodedata

def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    unicode_text = '\n'.join(fullText)
    lines_unicode = unicodedata.normalize(u'NFKD', unicode_text).encode('ascii', 'ignore').decode('utf8')
    return lines_unicode



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
contemp_parameters = "submitted or prepared contemporaneously(.*?)> Does a local branch"
# Materiality Thresholds
tpdoc_materiality_parameters = "> Transfer pricing documentation(.*?)> Master File"
mf_materiality_parameters = "> Master File(.*?)> Local File"
lf_materiality_parameters = "> Local File(.*?)> CbCR"
cbcr_materiality_parameters = "> CbCR(.*?)> Economic analysis"
# Other requirements
language_parameters = ">Local language documentation requirement(.*?)> Safe harbor"
tp_return_parameters = "> Transfer pricing-specific returns(.*?)> CbCR"
tp_taxreturn_parameters = "> Related-party disclosures along with corporate income tax return(.*?)> Related-party disclosures in financial"
#Deadlines
tp_disclosure_deadline_parameters = "> Other transfer pricing disclosures and return(.*?)> Master File"
deadline_first_parameters = "a\) Filing deadline(.*?)per COVID"
cit_secondary_deadline_parameters = "> Corporate income tax return(.*?)> Other transfer pricing"
mf_secondary_deadline_parameters = "> Master File(.*?)> CbCR"
cbcr_secondary_deadline_parameters = "> CbCR preparation and submission(.*?)b\)"
tpdoc_prep_deadline_parameters = "Transfer pricing documentation/Local File preparation deadline(.*?)c\) Transfer pricing"
tpdoc_submit_deadline_parameters = "or Local File\?(.*?)d\)"



def processInfo(searchCriteria, column_name, countries):
    def findInfo(searchCriteria, country_string):
        result = re.findall(rf'{searchCriteria}', country_string, re.S)
        if result:
            return result[0]
        else:
            return "No match found"

    column_values = []
    for country in countries:
        column_values.append(findInfo(searchCriteria, country))
    df[column_name] = column_values

def processSecondaryInfo(first_search_criteria, second_search_criteria, secondary_column_name, countries):
    def find_first_info(first_search_criteria, country_string):
        result = re.findall(rf'{first_search_criteria}', country_string, re.S)
        if result:
            return result[0]
        else:
            return "No match found"

    column_values = []
    for country in countries:
        column_values.append(find_first_info(first_search_criteria, country))

    def find_second_info(second_search_criteria, column_values):
        second_result = re.findall(rf'{second_search_criteria}', column_values, re.S)
        if second_result:
            return second_result[0]
        else:
            return "No match found"

    secondary_column_values = []    
    for value in column_values:
        secondary_column_values.append(find_second_info(second_search_criteria, value))
    df[secondary_column_name] = secondary_column_values


df = pd.DataFrame()

processInfo(country_parameters, "country_names", countries)
processInfo(beps_parameters, "beps_status", countries)
processInfo(mcaa_parameters, "mcaa_status", countries)
processInfo(contemp_parameters, "Contemporaneous Doc Requirement", countries)
processInfo(tpdoc_materiality_parameters, "TP Doc Thresholds", countries)
processInfo(mf_materiality_parameters, "MF Threshold", countries)
processInfo(lf_materiality_parameters, "LF Thresholds", countries)
processInfo(cbcr_materiality_parameters, "CbCR Thresholds", countries)
processSecondaryInfo(deadline_first_parameters, cit_secondary_deadline_parameters, "CIT Deadline", countries)
processSecondaryInfo(deadline_first_parameters, mf_secondary_deadline_parameters, "MF Deadline", countries)
processSecondaryInfo(deadline_first_parameters, cbcr_secondary_deadline_parameters,"CbCR Deadlines", countries)
processInfo(tpdoc_prep_deadline_parameters, "TP Doc Preparation Deadline", countries)
processInfo(tpdoc_submit_deadline_parameters, "TP Doc Submit Deadline", countries)


df.to_csv('reference.csv', index =False)


