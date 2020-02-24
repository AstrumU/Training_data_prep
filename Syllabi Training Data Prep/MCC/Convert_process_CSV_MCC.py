"""
   Get Useful Content
"""
import csv
import re
import os
import string
import xlwt
import datetime
import pandas as pd

# File Path {str}
read_excel_file = '/Users/cheng/Desktop/astrumu_parsers-master/tests/test_files/Courses/output/batch_process_result.csv'
write_excel_file = '/Users/cheng/Desktop/Testing'

# All Possible Markers
Certified_Markers = {'DATE SUBMITTED','CATALOG NO.','DATE DICC APPROVED','DATE LAST REVIEWED','COURSE INFORMATION FORM','DISCIPLINE', 
                     'COURSE TITLE','CR.HR','LECT HR.','LAB HR.','CLIN/INTERN HR.','CLOCK HR.','CATALOG DESCRIPTION','PREREQUISITES',
                     'EXPECTED STUDENT OUTCOMES IN THE COURSE','GENERAL EDUCATION OUTCOMES','PROGRAM-LEVEL OUTCOMES',
                     'CAREER AND TECHNICAL EDUCATION PROGRAM OUTCOMES','CLASS-LEVEL ASSESSMENT MEASURES','COURSE OUTLINE FORM',
                     'PROGRAM-LEVEL OUTCOMES ADDRESSED','COURSE INFORMATION','COURSE OUTLINE','DATE DICC','DICC APPROVAL','NO.',
                     'PROGRAM OUTCOMES','ASSESSMENT MEASURES'}

# Read in input csv document and activate output xlsx document
df = pd.read_csv(read_excel_file)   

# Regular Expression
p = r'((^|\t|\n){0,}[A-Z.]+[ \-\/\n\t]{0,1}){1,}(([A-Z()]|\.){2,}(\t|\s{4}|\n\$){0,})'
pattern = re.compile(p)

process_cnt = 0
results = pd.DataFrame(columns = ['File Name', 'DATE SUBMITTED', 'CATALOG NO.', 'DATE DICC APPROVED', 'DATE LAST REVIEWED', 'DISCIPLINE', 
                                  'COURSE TITLE', 'CR.HR.', 'LECT.HR.', 'LAB.HR.', 'CLIN/INTERN.HR.', 'CLOCK.HR.', 'CATALOG DESCRIPTION', 
                                  'PREREQUISITES', 'EXPECTED STUDENT OUTCOMES IN THE COURSE', 'GENERAL EDUCATION OUTCOMES', 
                                  'CAREER AND TECHNICAL EDUCATION PROGRAM OUTCOMES', 'CLASS-LEVEL ASSESSMENT MEASURES', 
                                  'PROGRAM-LEVEL OUTCOMES ADDRESSED', 'COURSE OUTLINE FORM', 'UPDATE AT'])

# Iterate Over the CSV Document
for i in range(len(df['file_fullname'])):

    File_Name = ""
    DATE_SUBMITTED = ""
    CATALOG_NO = ""
    DATE_DICC_APPROVED = ""
    DATE_LAST_REVIEWED = ""
    DISCIPLINE = ""
    COURSE_TITLE = ""
    CR_HR = ""
    LECT_HR = ""
    LAB_HR = ""
    CLIN_or_INTERN_HR = ""
    CLOCK_HR = ""
    CATALOG_DESCRIPTION = ""
    PREREQUISITES = ""
    EXPECTED_STUDENT_OUTCOMES_IN_THE_COURSE = ""
    GENERAL_EDUCATION_OUTCOMES = ""
    CAREER_AND_TECHNICAL_EDUCATION_PROGRAM_OUTCOMES = ""
    CLASS_LEVEL_ASSESSMENT_MEASURES = ""
    PROGRAM_LEVEL_OUTCOMES_ADDRESSED = ""
    COURSE_OUTLINE_FORM = ""

    File_Name = df['file_fullname'][i]

    # Read in and process file text
    buff = df['file_text'][i]
    if isinstance(buff, str):
        buff = buff.replace("|","")      # Some parsed results have unwanted symbols
        words = re.split('[\n\t]', buff)      # Break down file text
        markers = []

        for word in words:
            for uppercase in re.finditer(pattern, word):

                marker = uppercase[0]
                marker = marker.strip()

                if marker in Certified_Markers:
                    markers.append(marker)

                """
                    Special Cases 1: Two markers are put together
                """

                if marker == 'COURSE INFORMATION FORM DISCIPLINE':
                    markers.append('COURSE INFORMATION FORM')                    
                    markers.append('DISCIPLINE')
                if marker == 'COURSE OUTLINE FORM DISCIPLINE':
                    markers.append('COURSE OUTLINE FORM')                    
                    markers.append('DISCIPLINE') 
                if marker == 'CLOCK HR. CATALOG DESCRIPTION':
                    markers.append('CLOCK HR.')
                    markers.append('CATALOG DESCRIPTION')

                """
                    Special Cases 2: Remove redundant '(ESO)'
                """
                
                if marker == 'EXPECTED STUDENT OUTCOMES IN THE COURSE (ESO)':
                    markers.append('EXPECTED STUDENT OUTCOMES IN THE COURSE')
                if marker == 'GENERAL EDUCATION OUTCOMES (ESO)':
                    markers.append('GENERAL EDUCATION OUTCOMES')
                if marker == 'IN THE COURSE (ESO)':
                    markers.append('IN THE COURSE')
                if marker == 'OUTCOMES (ESO)':
                    markers.append('OUTCOMES')
                    
        markers.append('QWERTYUIOP')      # Add an extra unique marker to find the end of the text
        
        for j in range(len(markers) -1):
            buff = buff + "QWERTYUIOP"      # Append string corresponding to the unique marker

            """
                This is the last part. Some duplicate markers in this part need not to be treated as markers
            """
            if markers[j] == 'COURSE OUTLINE FORM' or markers[j] == 'COURSE OUTLINE':
                pattern0 = re.compile(markers[j] + '(.*?)' + markers[-1],re.S)
                result = pattern0.findall(buff)
                if result:
                    COURSE_OUTLINE_FORM = result[0].strip()
                    break
            else:
                pattern0 = re.compile(markers[j] + '(.*?)' + markers[j+1],re.S)
                result = pattern0.findall(buff)

            """
                Put the content under the corresponding marker
            """

            if result:
                if markers[j] == 'DATE SUBMITTED':
                    DATE_SUBMITTED = result[0].strip()
                elif markers[j] == 'CATALOG NO.' or markers[j] == 'NO.':
                    CATALOG_NO = result[0].strip()
                elif markers[j] == 'DATE DICC APPROVED' or markers[j] == 'DATE DICC' or markers[j] == 'DICC APPROVAL':
                    DATE_DICC_APPROVED = result[0].strip()
                elif markers[j] == 'DATE LAST REVIEWED':
                    DATE_LAST_REVIEWED = result[0].strip()
                elif markers[j] == 'DISCIPLINE':
                    DISCIPLINE = result[0].strip()
                elif markers[j] == 'COURSE TITLE':
                    COURSE_TITLE = result[0].strip()               
                elif markers[j] == 'CR.HR':
                    CR_HR = result[0].strip()
                elif markers[j] == 'LECT HR.':
                    LECT_HR = result[0].strip()          
                elif markers[j] == 'LAB HR.':
                    LAB_HR = result[0].strip()           
                elif markers[j] == 'CLIN/INTERN HR.':
                    CLIN_or_INTERN_HR = result[0].strip()
                elif markers[j] == 'CLOCK HR.':
                    CLOCK_HR = result[0].strip()
                elif markers[j] == 'CATALOG DESCRIPTION':
                    CATALOG_DESCRIPTION = result[0].strip()
                elif markers[j] == 'PREREQUISITES':
                    PREREQUISITES = result[0].strip()

                # Get rid of '(ESO)' if there is any
                elif markers[j] == 'EXPECTED STUDENT OUTCOMES IN THE COURSE' or markers[j] == 'IN THE COURSE':
                    temporary = result[0].strip()
                    if temporary[:5] == '(ESO)':
                        EXPECTED_STUDENT_OUTCOMES_IN_THE_COURSE = temporary[6:]
                    else:
                        EXPECTED_STUDENT_OUTCOMES_IN_THE_COURSE = temporary
                elif markers[j] == 'GENERAL EDUCATION OUTCOMES' or markers[j] == 'OUTCOMES':
                    temporary = result[0].strip()
                    if temporary[:5] == '(ESO)':
                        GENERAL_EDUCATION_OUTCOMES = temporary[6:]
                    else:
                        GENERAL_EDUCATION_OUTCOMES = temporary

                elif markers[j] == 'CAREER AND TECHNICAL EDUCATION PROGRAM OUTCOMES' or markers[j] == 'PROGRAM OUTCOMES':
                    CAREER_AND_TECHNICAL_EDUCATION_PROGRAM_OUTCOMES = result[0].strip()
                elif markers[j] == 'CLASS-LEVEL ASSESSMENT MEASURES' or markers[j] == 'ASSESSMENT MEASURES':
                    CLASS_LEVEL_ASSESSMENT_MEASURES = result[0].strip()
                elif markers[j] == 'PROGRAM-LEVEL OUTCOMES ADDRESSED':
                    PROGRAM_LEVEL_OUTCOMES_ADDRESSED = result[0].strip()
                else:
                    continue
            
            """
                Extract the remaing text
            """
            UPDATE_AT = datetime.datetime.now()
            pattern1 = re.compile(markers[j]+'(.*?)'+markers[-1],re.S)
            temp = pattern1.findall(buff)
            buff = temp[0]

        results.loc[process_cnt] = [File_Name, DATE_SUBMITTED, CATALOG_NO, DATE_DICC_APPROVED, DATE_LAST_REVIEWED, DISCIPLINE, 
                                    COURSE_TITLE, CR_HR, LECT_HR, LAB_HR, CLIN_or_INTERN_HR, CLOCK_HR, CATALOG_DESCRIPTION, 
                                    PREREQUISITES, EXPECTED_STUDENT_OUTCOMES_IN_THE_COURSE, GENERAL_EDUCATION_OUTCOMES, 
                                    CAREER_AND_TECHNICAL_EDUCATION_PROGRAM_OUTCOMES, CLASS_LEVEL_ASSESSMENT_MEASURES, 
                                    PROGRAM_LEVEL_OUTCOMES_ADDRESSED, COURSE_OUTLINE_FORM, UPDATE_AT]

        process_cnt += 1

    else:
        continue

"""
    Save the document
"""
if results.shape[0] > 0:
    results.to_csv(os.path.join(write_excel_file, 'Results.csv'), index=False)