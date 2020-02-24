"""
   Get Useful Content
"""
import csv
import os
import re
import string
import datetime
import pathlib
import pandas as pd

# File Path {str}
read_excel_file = '/Users/cheng/Desktop/astrumu_parsers-master/tests/test_files/output_Fall_2018/batch_process_result.csv'
write_excel_file = '/Users/cheng/Desktop/Testing'

# All Possible Markers
key_words = ['start', 'course content', 'structure', 'course info', 'title', 'number', 'prerequisite', 'description', 'objective', 'goal', 'outcome', 'overview', 
             'syllabus', 'calendar', 'administrative dates', 'course dates', 'text', 'book', 'reading', 'material', 'skill', 'requirement', 'following', 'topics', 
             'instruction', 'professional development', 'schedule', 'outline', 'topic', 'course format', 'course format']

useless_words =['Credits', 'Assessment', 'Paradigm', 'Grade', 'Grading', 'Assignment', 'department', 'college', 'Evaluation', 'Honors', 'Online', 'Clinical', 
                'Course Meets', 'Evacuation Plan', 'Policy', 'policies', 'Safety', 'Academic Integrity', 'Inclement Weather', 'Attendance', 'Plagiarism', 
                'Center', 'absence', 'Response', 'Rounding up', 'Code of Conduct', 'HONEST', 'Accessibility', 'Disability', 'Special Requests', 'accommodation', 
                'Late Work', 'Participation', 'discrimination', 'harassment', 'Academic Continuity', 'EXPECTATION', 'Class', 'Term', 'Location', 'Place', 'Time', 
                'Laptop', 'Faculty', 'Office', 'Instructor', 'Email', 'E-mail', 'mail', 'Phone', 'Professor', 'TA ', 'Contact Information', 'Taping of Lecture', 
                'Recording of Lecture', 'Course Webpage', 'Calculator', 'Excel projects', 'E-Learning Day', 'Section Information', 'Studying', 
                'General Education Competency', 'Contribution', 'Learning Strategies', 'Teaching Strategies', 'GENERAL EDUCATION CROSS‐CULTURAL AWARENESS',
                'NAAB STUDENT PERFORMANCE CRITERIA', 'Course Feedback', 'Plan', 'Comments', 'Collaborative Process', 'Course Approach', 'Learning Environment', 
                'Standards of Professional Practice', 'Emergency', 'Web Assign', 'Due Dates', 'quizzes', 'Tests', 'Exam', 'Student Responsibility', 'in This Course', 
                'Resource', 'Method', 'Lab Access', 'Fees', 'Music', 'Homework', 'General Questions', 'Electronic', 'Civility Statement', 'Communication', 
                'Activities Explained', 'Deliverables', 'Current Opportunities', 'DESIGN STUDIO PROCESS CHAPTERS', 'Odds and Ends', 'Specific University Reminder', 
                'Submission of Work from Other Courses', 'Copyright', 'Academic Grievances', 'Cooper Library', 'Support', 'Academic Advising', 'Registrar', 
                'Website', 'Canvas', 'Others', 'Additional makeup products you need to order', 'Project Checklist', 'When', 'Semester', 'Where', 'Remarks']

# Read in input csv document and activate output xlsx document
df=pd.read_csv(read_excel_file)

# Regular Expression
patt = re.compile(r'^[a-z,A-Z,\s, \\, \/, \-’]{3,50}$', re.I|re.U|re.X)

"""
   Initialize DataFrame
"""
process_cnt = 0

# Get course ID form File Fullname, then may get Course Name from File Beginning, Course Title, Course Description, and Course Syllabus based on course ID
results = pd.DataFrame(columns=['File Index', 'File Fullname', 'File Beginning', 'Course Title', 'Course Description', 'Course Syllabus', 'File Text'])

# Iterate Over the CSV Document
for i in range(len(df['file_fullname'])):
    buff = df['file_text'][i]
    buff = "START" + "\t" + str(buff)      # Add “START” to the front to track the beginning part of the file
    
    try:
        if isinstance(buff, str):
            buff=buff.replace("|","")
            Phrases = re.split('[\n\t]', buff)
            Candidates = []      # The set of all candidates of Markers

            for phrase in Phrases:
                if phrase:
                    if phrase[0].isupper():
                        if patt.match(phrase):
                            marker_Cand = phrase.strip()
                            Candidates.append(marker_Cand)
                        else:
                            if ':' in phrase:
                                marker_Cand = re.split(':', phrase)[0].strip()
                                if patt.match(marker_Cand):
                                    Candidates.append(marker_Cand)

            # find all markers and useful markers
            markers = []
            useful_markers = []
            
            for j in range(len(Candidates)):
                if Candidates[j] == 'Social Structure and Interaction' or Candidates[j] == 'Social Class':
                    continue
                elif 'reading assignments' in Candidates[j].lower():
                    continue
                elif 'exam schedule' in Candidates[j].lower():
                    continue
                elif Candidates[j].lower() == 'ta':
                    markers.append(Candidates[j])
                elif 'title ix' in Candidates[j].lower():
                    markers.append(Candidates[j])
                elif Candidates[j].lower() == 'read the syllabus':
                    markers.append(Candidates[j])
                elif Candidates[j].lower() == 'requirements and grades':
                    markers.append(Candidates[j])
                elif Candidates[j].lower() == 'course': 
                    markers.append(Candidates[j])
                    useful_markers.append(Candidates[j])
                elif Candidates[j].lower() == 'Advanced topics and quizzes':
                    markers.append(Candidates[j])
                else:
                    for key_word in key_words:
                        if key_word.lower() in Candidates[j].lower():
                            markers.append(Candidates[j])
                            useful_markers.append(Candidates[j])
                            break
                    else:
                        for useless_word in useless_words:
                            if useless_word.lower() in Candidates[j].lower():
                                markers.append(Candidates[j])
                                break
            
            markers.append('QWERTYUIOP')      # Add an extra unique marker to find the end of the text

            """
                Find markers we want
            """
            Final_markers = []
            text = ''
            
            for marker in markers:
                if '(' in marker:
                    final_marker = re.match(r'(.*?)\((.*?)\)(.*?)',marker).group(1).strip()
                else:
                    final_marker = marker
                Final_markers.append(final_marker)

            # Results
            file_index = process_cnt
            file_fullname = ""
            file_beginning = ""
            Course_Title = ""
            Course_Description = ""
            Course_Syllabus = ""
            file_text = ""

            file_fullname = df['file_fullname'][i]

            for j in range(len(Final_markers) -1):
                buff = buff + " QWERTYUIOP"      # Append string corresponding to the unique marker
                pattern=re.compile(Final_markers[j]+'(.*?)'+Final_markers[j+1],re.S)
                result=pattern.findall(buff)

                """
                    Put the content under the corresponding marker
                """
                if result:
                    if j == 0:
                        file_beginning = result[0].strip()
                        continue
                    if 'course' in Final_markers[j].lower() and 'title' in Final_markers[j].lower():
                        Course_Title = Final_markers[j] + '\t' + result[0].strip()
                        continue
                    if 'description' in Final_markers[j].lower():
                        Course_Description = Final_markers[j] + '\t' + result[0].strip()
                    if 'syllabus' in Final_markers[j].lower():
                        Course_Syllabus = Final_markers[j] + '\t' + result[0].strip()                
                    if Final_markers[j] in useful_markers:
                        text = text + '\n' + '\n' + Final_markers[j] + '\t' + result[0].strip()

                """
                    Extract the remaing text
                """
                pattern0=re.compile(Final_markers[j]+'(.*?)'+ Final_markers[-1],re.S)
                temp=pattern0.findall(buff)

                if temp:
                    buff=temp[0]
                else:
                    buff = ''
        
            file_text = text

            results.loc[process_cnt] = [file_index, file_fullname, file_beginning, Course_Title, Course_Description, Course_Syllabus, file_text]
            process_cnt += 1

    except Exception:
        print('The file number is:')
        print(i)

"""
    Save the document
"""
if results.shape[0] > 0:
    results.to_csv(os.path.join(write_excel_file, 'Results.csv'), index=False)