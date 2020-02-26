"""
    Get File Name from parsed text for each course
"""
import os
import re
import pandas as pd

READ_EXCEL_FILE = '/Users/astrumu_parsers-master/tests/test_files/batch_process_result.csv'
WRITE_EXCEL_FILE = '/Users/cheng/Desktop/Testing'
TITLEFILE = '/Users/cheng/Desktop/Testing/Course_Title.xlsm'

# All Possible Markers
KEY_WORDS = ['start', 'course content', 'structure', 'course info', 'title', 'number',
             'prerequisite', 'description', 'objective', 'goal', 'outcome', 'overview', 'syllabus',
             'calendar', 'administrative dates', 'course dates', 'text', 'book', 'reading',
             'material', 'skill', 'requirement', 'following', 'topics', 'instruction',
             'professional development', 'schedule', 'outline', 'topic', 'course format',
             'course format']

USELESS_WORDS = ['Credits', 'Assessment', 'Paradigm', 'Grade', 'Grading', 'Assignment', 'Online',
                 'college', 'Evaluation', 'Honors', 'Clinical', 'Course Meets', 'department',
                 'Evacuation Plan', 'Policy', 'policies', 'Safety', 'Academic Integrity',
                 'Inclement Weather', 'Attendance', 'Plagiarism', 'Center', 'absence', 'Response',
                 'Rounding up', 'Code of Conduct', 'HONEST', 'Accessibility', 'Disability',
                 'Special Requests', 'accommodation', 'Late Work', 'Participation', 'Studying',
                 'harassment', 'Academic Continuity', 'EXPECTATION', 'Class', 'Term', 'Location',
                 'Place', 'Time', 'Laptop', 'Faculty', 'Office', 'Instructor', 'Email', 'E-mail',
                 'mail', 'Phone', 'Professor', 'TA ', 'Contact Information', 'Taping of Lecture',
                 'Recording of Lecture', 'Course Webpage', 'Calculator', 'Excel projects',
                 'E-Learning Day', 'Section Information', 'General Education Competency',
                 'Contribution', 'Learning Strategies', 'Teaching Strategies', 'discrimination',
                 'GENERAL EDUCATION CROSS‐CULTURAL AWARENESS', 'NAAB STUDENT PERFORMANCE CRITERIA',
                 'Course Feedback', 'Plan', 'Comments', 'Collaborative Process', 'Course Approach',
                 'Learning Environment', 'Standards of Professional Practice', 'in This Course',
                 'Due Dates', 'quizzes', 'Tests', 'Exam', 'Student Responsibility', 'Emergency',
                 'Resource', 'Method', 'Lab Access', 'Fees', 'Music', 'Homework',
                 'General Questions', 'Electronic', 'Civility Statement', 'Communication',
                 'Activities Explained', 'Deliverables', 'Current Opportunities',
                 'DESIGN STUDIO PROCESS CHAPTERS', 'Odds and Ends', 'Specific University Reminder',
                 'Submission of Work from Other Courses', 'Copyright', 'Academic Grievances',
                 'Cooper Library', 'Support', 'Academic Advising', 'Registrar', 'Website', 'Canvas',
                 'Others', 'When', 'Semester', 'Where', 'Remarks',
                 'Additional makeup products you need to order', 'Project Checklist', 'Web Assign']

# Read in input csv document and activate output xlsx document

ALLTITLES = pd.DataFrame(pd.read_excel(TITLEFILE, keep_default_na=False))
DF = pd.read_csv(READ_EXCEL_FILE)
TITLES = []

for i in range(ALLTITLES.iloc[:, 0].size):
    buff_title = ALLTITLES.loc[i][0]
    buff_title = str(buff_title)
    TITLES.append(buff_title)

# Regular Expression
PATT = re.compile(r'^[a-z,A-Z,\s, \\, \/, \-’]{3,50}$', re.I|re.U|re.X)

"""
   Initialize DataFrame
"""
PROCESS_CNT = 0

RESULTS = pd.DataFrame(columns=['Course Title', 'File Text'])

for i in range(len(DF['file_fullname'])):
    buff = DF['file_text'][i]
    # Add “START” to the front to track the beginning part of the file
    buff = "START" + "\t" + str(buff)

    try:
        if isinstance(buff, str):
            buff = buff.replace("|", "")
            Phrases = re.split('[\n\t]', buff)
            Candidates = []      # The set of all candidates of Markers

            for phrase in Phrases:
                if phrase:
                    if phrase[0].isupper():
                        if PATT.match(phrase):
                            marker_Cand = phrase.strip()
                            Candidates.append(marker_Cand)
                        else:
                            if ':' in phrase:
                                marker_Cand = re.split(':', phrase)[0].strip()
                                if PATT.match(marker_Cand):
                                    Candidates.append(marker_Cand)

            # find all markers and useful markers
            markers = []
            useful_markers = []

            for candidate in Candidates:
                if candidate in ('Social Structure and Interaction', 'Social Class'):
                    continue
                elif 'reading assignments' in candidate.lower():
                    continue
                elif 'exam schedule' in candidate.lower():
                    continue
                elif candidate.lower() == 'ta':
                    markers.append(candidate)
                elif 'title ix' in candidate.lower():
                    markers.append(candidate)
                elif candidate.lower() == 'read the syllabus':
                    markers.append(candidate)
                elif candidate.lower() == 'requirements and grades':
                    markers.append(candidate)
                elif candidate.lower() == 'course':
                    markers.append(candidate)
                    useful_markers.append(candidate)
                elif candidate.lower() == 'Advanced topics and quizzes':
                    markers.append(candidate)
                else:
                    for key_word in KEY_WORDS:
                        if key_word.lower() in candidate.lower():
                            markers.append(candidate)
                            useful_markers.append(candidate)
                            break
                    else:
                        for useless_word in USELESS_WORDS:
                            if useless_word.lower() in candidate.lower():
                                markers.append(candidate)
                                break
            # Add an extra unique marker to find the end of the text
            markers.append('QWERTYUIOP')

            # Find markers we want
            Final_markers = []
            text = ''

            for marker in markers:
                if '(' in marker:
                    final_marker = re.match(r'(.*?)\((.*?)\)(.*?)', marker).group(1).strip()
                else:
                    final_marker = marker
                Final_markers.append(final_marker)

            # Results
            file_index = PROCESS_CNT
            file_fullname = ""
            file_beginning = ""
            Course_Title = ""
            Course_Description = ""
            Course_Syllabus = ""
            file_text = ""
            Course_Name = ""

            file_fullname = DF['file_fullname'][i]

            for j in range(len(Final_markers) -1):
                buff = buff + " QWERTYUIOP"      # Append string corresponding to the unique marker
                pattern = re.compile(Final_markers[j]+'(.*?)'+Final_markers[j+1], re.S)
                result = pattern.findall(buff)

                """
                    Put the content under the corresponding marker
                """
                if result:
                    if j == 0:
                        file_beginning = result[0].strip()

                        length = 0
                        info = str(file_beginning)

                        for title in TITLES:
                            if title.lower() in info.lower():
                                # useually the longest match string is the title for that course
                                if len(title) > length:
                                    Course_Name = title
                                    length = len(Course_Name)

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


                # Extract the remaing text
                pattern0 = re.compile(Final_markers[j]+'(.*?)'+ Final_markers[-1], re.S)
                temp = pattern0.findall(buff)

                if temp:
                    buff = temp[0]
                else:
                    buff = ''

            if Course_Name != '':
                file_text = text
                RESULTS.loc[PROCESS_CNT] = [Course_Name, file_text]
                PROCESS_CNT += 1

    except SyntaxWarning:
        print('The file number is:')
        print(i)

# Save the document
if RESULTS.shape[0] > 0:
    RESULTS.to_pickle(os.path.join(WRITE_EXCEL_FILE, 'frame_pickle'))
