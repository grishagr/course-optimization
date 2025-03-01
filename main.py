import csv
import time
import sys
from gurobipy import *
from collections import *
from datetime import datetime
from copy import deepcopy
from math import ceil
from pandas import DataFrame, ExcelWriter


# TODO: add other AP conversion to placement
# ------------------------------------------------------------------------------
# CONSTANTS
START = "25/SP"
YEAR = ""

LUNCH_PRIORITY = 15 / 5 # lunch constant -> how much does each lunch weight (per day)
LUNCH2PM_PRIORITY = 12.5 / 5 # put a lower weight on 2pm lunch

# lookup puts premium on priorities (1st priority is key 12 which maps to premium weight value 20)
LOOKUP = {12: 25, 11: 17, 10:13, 9:12, 8:9, 7:8, 6:7, 5:6, 4:4, 3:3, 2:2, 1:1, 0:0,
           LUNCH_PRIORITY: LUNCH_PRIORITY, LUNCH2PM_PRIORITY: LUNCH2PM_PRIORITY}

# placement levels (1st level, second level, etc.)
PLACEMENTS = {
    "MATH" : ["CALCULUS I", "CALCULUS II", "MULTIVARIABLE CALCULUS", "LINEAR ALGEBRA"],
    "FRNCH" : ["ELEMENTARY FRENCH", "INTERMEDIATE FRENCH I", "INTERMEDIATE FRENCH II", "WRITTEN & ORAL ARGUMENTATN"],
    "ECON" : ["INTRODUCTION TO ECONOMICS", "ECON THEORY & EVIDENCE"],
    "CHNSE" : ["FIRST TERM CHINESE", "THIRD TERM CHINESE", "THIRD YEAR CHINESE 1"],
    "CPSCI" : ["COMPUTER SCIENCE FOR ALL", "DESIGN PRINCIPLES"],
    "JAPN" : ["FIRST TERM JAPANESE", "THIRD TERM JAPANESE"],
    "LATIN" : ["ELEMENTARY LATIN I", "INTERMEDIATE LATIN"],
    "MUSIC" : ["INTRO MUSIC THEORY", "FUNDAMNTL & CHROM HARMONY", "CHROMATIC HARMONY"],
    "HSPST" : ["SPANISH IMMERSION I", "THIRD TERM SPANISH", "GRAMMAR AND COMPOSITION", "GRAMMAR FOR HERITAGE", "EXPLORING HISPANIC TEXTS"],
    "ITALN" : ["FIRST TERM ITALIAN", "THIRD TERM ITALIAN"]
}

# departments that require placement
REQUIRED_PLACEMENT = ["MATH", "FRNCH", "HSPST"]

# special placements (in the value of dict, put title first, then sections of it; or just title)
# numbers are weights that rank is multiplied with
SPECIAL_PLACEMENT_CASES = {"MATH 113 FYC" : {"CALCULUS I" : 0.9, "CALCULUS I-2" : 0.9, "CALCULUS I-3" : 0.9, "CALCULUS I-4" : 1},
                           "FRNCH 211 OR EQUIVALENT" : {"WHAT'S NEW? OR COMM FRNCH I": 1},
                           "MATH 216/224" : {"LINEAR ALGEBRA" : 1, "MULTIVARIABLE CALCULUS" : 1},
                           "HSPST 200+" : {"SPANISH IMMERSION I" : 0},
                           "ITALN 200" : {"FIRST TERM ITALIAN" : 0},
                           "113/116/216" : {"CALCULUS I" : 1, "CALCULUS II" : 1, "MULTIVARIABLE CALCULUS" : 1},
                           "LATIN 390" : {"ELEMENTARY LATIN I" : 0},
                           "JAPN 200F" : {"FIRST TERM JAPANESE" : 0},
                           "SPEAK TO KYOKO OMORI" : {"FIRST TERM JAPANESE" : 0},
                           "FRNCH 211, 212, 250 or 280" : {"WHAT'S NEW? OR COMM FRNCH I": 1}
}

# DIVISIONS -> [departments] dictionary 
DIVISIONS = { "arts" : ["ART", "DANCE", "DARTS", "MUSIC", "THETR"],
             "sciences" : ["ARCH", "PHYS", "BIO", "CHEM", "CPSCI", "GEOSC", "MATH", "PSYCH"],
             "humanities" : ["ARTH", "ASNST", "CLASC", "CLNG", "FRNCH", "HEBRW",
                             "HSPST", "LATIN", "LIT", "PHIL", "RUSSN", "HIST" ],
             "social sciences" : ["ANTHR", "ECON", "EDUC", "GOVT", "LING", "SOC", "PPOL" ],
             "interdisciplinary" : ["AFRST", "AMST", "ARABC", "CHNSE", "CNMS", "ENVST",
                                    "GERMN", "ITALN", "JAPN", "MDRST", "MEIWS", "RELST",
                                    "RSNST", "WMGST", "JLJS", "LTAM", "COLEG" ]
}

# titles that do not count towards writing intensive constraint
NOT_WRITING_INTENSIVE = ["WRITTEN & ORAL ARGUMENTATN", "EXPLORING HISPANIC TEXTS", "LINEAR ALGEBRA"] 

# titles that do not count towards FYC constraint
NOT_FYC = ["CALCULUS I"]

# Ignore these courseNames
IGNORE = ["PHYS 100L"]

# For printing messages
class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

# ------------------------------------------------------------------------------

# Utility functions

# Helper function to convert time string to datetime object
def time_to_datetime(time_str: str) -> datetime:
    return datetime.strptime(time_str, '%I:%M%p')


# Loading message
def load_log(percentage: int):
    message = f"Progress -- {percentage}%"
    sys.stdout.write(f"{bcolors.OKCYAN}\r{message}{bcolors.ENDC}")  # Use '\r' to return the cursor to the start of the line
    sys.stdout.flush() 
    if percentage == 100:
        sys.stdout.write(f"\r{bcolors.BOLD}{bcolors.OKCYAN}Done. Evaluating the model ...{bcolors.ENDC}")
        sys.stdout.flush()


# Helper function to check if two time intervals overlap
def times_overlap(start1: datetime, end1: datetime, start2: datetime, end2: datetime) -> bool:
    return start1 < end2 and start2 < end1


# Main function to find overlapping classes
def find_overlapping_classes(classDict: dict, target_class: str) -> list:
    target_meetings = classDict[target_class][5]
    overlapping_classes = []

    for class_title, class_info in classDict.items():
        if class_title == target_class:
            continue

        for target_days, target_start, target_end in target_meetings:
            target_start_dt = time_to_datetime(target_start)
            target_end_dt = time_to_datetime(target_end)

            for days, start, end in class_info[5]:
                start_dt = time_to_datetime(start)
                end_dt = time_to_datetime(end)

                # Check if there is any overlap in the days
                if any(day in days for day in target_days):
                    # Check if the times overlap
                    if times_overlap(target_start_dt, target_end_dt, start_dt, end_dt):
                        overlapping_classes.append(class_title)
                        break

    return overlapping_classes


# Looks up studentDict by courseName (Ex. PHYS 100)
# Returns tuple (True, title) if found, or (False, courseName) if not found
def findTitle(courseName: str) -> tuple:

    # if someone specified only department (comment out if don't want to assign priority to these students)
    if courseName in one_class_dept:
        return (True, one_class_dept[courseName])

    # check crosslisted coursenames
    for title in crossListed:
        if courseName in crossListed[title]:
            return (True, title)
        
    # check all others
    for index, info in enumerate(classDict.values()):
        if courseName == info[2]:
            return (True, list(classDict.keys())[index])
        
    return (False, courseName)


# Extracts days and times of class meetings
# Input: meeting info section as a string
# Returns list of days and times strings (ex. ["WF 02:30PM 03:45PM", "F 01:00PM 04:00PM"])
def extractTimes(timeInfo: str) -> list:
    if timeInfo == "":
        return None
    
    timeInfo = timeInfo.split('\n')
    previousTime = ["", "", ""]
    times = []
    try:
        for line in timeInfo:

            start = line.find("LEC")
            if start == -1:
                start = line.find("LAB")

            if start == -1:
                start = line.find("STU")
            
            if start == -1:
                continue

            time = line[start + 5:]
            time = time.split()

            if len(time) < 3:
                time.append(time[1])
                temp = time[0][:3]
                time[1] = time[0][3:]
                time[0] = temp

            if previousTime[0] != time[0] or previousTime[1] != time[1] or previousTime[2] != time[2]:
                times.append(time)
            previousTime = time
    except:
        times = []
    return times


# Extracts placement information from the string
# Input: placement info section as a string
# Returns list of placements by titles, for now ([True, <Title>] or [False, <Course Name>])
def extractPlacements(placementInfo: str) -> list:
    placements = []
    if "French: FRNCH 211, 212, 250 or 280" in placementInfo:
        start = ""
        end = ""
        index = placementInfo.find("French: FRNCH 211, 212, 250 or 280")
        if index - 2 > 0:
            if placementInfo[index - 2] == ",":
                start = ", "
        index += len("French: FRNCH 211, 212, 250 or 280")
        if index < len(placementInfo):
            if placementInfo[index] == "," and start == "":
                end = ", "
        placementInfo = placementInfo.replace(start + "French: FRNCH 211, 212, 250 or 280" + end, "")
        placements.append([False, "FRNCH 211, 212, 250 or 280"])

    placementInfo = placementInfo.split(",")
    for placement in placementInfo:
        try:
            courseName = placement.split(":")[1].strip().upper()
            found, placementTitle = findTitle(courseName)

            if not found:
                found, placementTitle = findTitle(courseName[:-1])
                if not found:
                    found, placementTitle = findTitle(courseName + "W")
                    if not found:
                        placementTitle = courseName

            placements.append([found, placementTitle]) 
        except:
            return []
    return placements

# Converts AP exams to respective placements according to 
# Placement rules
# Returns concatenated placement list
def convertAPtoPlacements(ap_exams: dict, placements: list = []) -> list:
    # Econ 166 placement
    if "MICRO" in ap_exams and "MACRO" in ap_exams:
        if ap_exams["MICRO"] + ap_exams["MACRO"] >= 9:
            placements.append([True, "ECON THEORY & EVIDENCE"])

   
    for p in placements:
        # Math 113/116 placement
        if "MATH 113/116" in p:
            placement = "CALCULUS I"
            if "CALCAB" in ap_exams:
                score = ap_exams["CALCAB"]
                if score >= 3:
                    placement = "CALCULUS II"
            placements.append([True, placement])
        # Math 116/216 placement
        elif "MATH 116/216" in p:
            placement = "CALCULUS II"
            if "CALCAB" in ap_exams:
                score = ap_exams["CALCAB"]
                if score >= 3:
                    placement = "MULTIVARIABLE CALCULUS"
            placements.append([True, placement])

    return placements

# Sorts data from csv file by columns specified in sortBy
# Returns sorted data list
def sortData(data: list, sortBy: list = ["Short Title", "Section"]) -> list:
    header = data[0]
    rows = data[1:]
    sort_indices = [header.index(col) for col in sortBy]
    rows.sort(key=lambda x: tuple(x[index] for index in sort_indices))
    return [header] + rows




print()
print(f"{bcolors.BOLD}{bcolors.OKCYAN}Proccesing CSV files...{bcolors.ENDC}\n")

#---------------------------------------------------------------------------------
# Process Classes Info (classes.csv)

# first-year courses
with open('classes.csv', encoding='utf-8-sig') as csv_file:
    sortBy = ['Short Title', 'Section']
    rawdata = csv.reader(csv_file, delimiter = ',')
    classInfo = [row for row in rawdata]
classInfo = sortData(classInfo)
indexList = classInfo[0]


'''
Class Dictionary template
[key] title: Title of class  -> list of class info under unique title or section if multisection
classInfo:
[0] - courseDept: department 
[1] - courseNum: course number
[2] - courseName: course name -- used to look up class in preferences
[3] - courseSection: section number
[4] - seats: available seats -- seats = capacity - total enrolled
        [!] for crosslisted courses, refer to crossListed dict for available seats

[5] - meetings: 2D list with meeting info including days, start time, end time strings
      (ex. [['MWF', '10:00AM', '10:50AM'], ['TR', '10:30AM', '11:45AM']])

[6] - credit: credit number
[7] - isWI: 1 if course is counted towards writing intensive constraint; 0, otherwise
[8] - isFYC: 1 if course is counted towards FYC constraint; 0, otherwise
'''
classDict = {}

departments = {} # holds titles for each department (dept -> [titles])
labs = {} # lab titles -> course name (Ex. "MECHANICAL UNIV LAB" -> "PHYS 190L")
multiSection = {} # class titles with multiple sections -> list of titles of all sections (title -> [title, title-2, ...])
crossListed = {} # class titles that are cross listed in other deparments -> list of course names under that title
                # i.e title -> [courseName1, courseName2, ...] 
one_class_dept = {} # departments that offer one class (dept -> title)
non_full = {} # classes without full credits
no_meetings = [] # classes with no meeting information available



# Record info about each class in classDict

for row in classInfo[1:]:
    title = row[indexList.index("Short Title")]
    courseDept = row[indexList.index("Dept")]
    courseNum = row[indexList.index("Course Number")]
    courseSection = row[indexList.index("Section")]
    totalCap = row[indexList.index("Sched Capacity")] if row[indexList.index("XList Capacity")] == "" else row[indexList.index("XList Capacity")] 
    totalEnrolled = row[indexList.index("Total Enr")]
    credit = float(row[indexList.index("Sched Min Cred")])
    courseType = row[indexList.index("Course Types")]

    isWI = 0
    if "W" in row[indexList.index("Course Types")]:
        if not title in NOT_WRITING_INTENSIVE:
            isWI = 1 # writing intensive course

    isFYC = 0
    if "FYC" in courseType and not title in NOT_FYC:
        isFYC = 1 # FYC course

    if not courseDept in one_class_dept:
        one_class_dept[courseDept] = title
    else:
        if one_class_dept[courseDept] != title:
            one_class_dept[courseDept] = "N/A"

    seats = int(totalCap) - int(totalEnrolled)
    if seats < 0:
        seats = 0

    courseName = " ".join([courseDept, courseNum])
    
    if courseName in IGNORE:
        continue

    meetingInfo = row[indexList.index("Start/End Date Bldg Room Meth Days Start/End time")]
    meetings = extractTimes(meetingInfo)

    # Class does not meet/could not find relative information
    if meetings is None:
        no_meetings.append(title + " : No information given")
        continue

    elif len(meetings) == 0:
        no_meetings.append(title)
    
    # record labs (title -> courseName) into labs dict
    if "L" in courseNum and not (title in labs):
        labs[title] = courseName
        credit = 0
    
    if title in classDict:

        # Record crossListed courses (title -> remaining seats) into crossListed dict
        if courseDept != classDict[title][0]:
            if title not in crossListed:
                crossListed[title] = [classDict[title][2], courseName]
            else:
                crossListed[title].append(courseName)
                classDict[title][4] = min(seats, classDict[title][4])
            continue
        
        # Record multisection courses (title -> [title, title-2, title-3,...]) in multiSection dict
        elif courseDept == classDict[title][0] and not (title in multiSection):
            multiSection[title] = [title]
            multiSection[title].append(title + "-" + courseSection)
        elif courseDept == classDict[title][0] and title in multiSection:
            multiSection[title].append(title + "-" + courseSection)

        classDict[title + "-" + courseSection] = [courseDept, courseNum, courseName, courseSection, seats, meetings, meetingInfo, credit, isWI, courseType, isFYC]
    else:
        classDict[title] = [courseDept, courseNum, courseName, courseSection, seats, meetings, meetingInfo, credit, isWI, courseType, isFYC]
    
# record title into corresponding department within departments dict
for title in classDict:
    if not title in labs and not title[:-2] in labs and not title in crossListed: # labs are okay but crosslisted need to be changed in the future
        dept = classDict[title][0]
        if dept in departments:
            departments[dept].append(title)
        else:
            departments[dept] = [title]

lunch_times = ["11:00AM", "11:30AM", "12:00PM", "12:30PM", "1:00PM", "1:30PM", "2:00PM"]

# record lunches as classes in the classDict and add them to multiSection courses by day
for day in "MTWRF":
    multiSection["Lunch " + day] = []
    for index, start in enumerate(lunch_times):
        title = "LUNCH " + day + " " + start
        courseDept = "lunch"
        courseNum = "lunch"
        courseName = "lunch"
        courseSection = day
        seats = 1000
        if index + 1 < len(lunch_times):
            end = lunch_times[index + 1]
        else:
            end = "2:30PM" # end time for the last lunch (change if last lunch is not 2PM)
        meetings = [[day, start, end]]
        credit = 0.00001
        isWI = 0
        isFYC = 0
        courseType = "N/A"
        classDict[title] = [courseDept, courseNum, courseName, courseSection, seats, meetings, meetingInfo, credit, isWI, courseType, isFYC]
        multiSection["Lunch " + day].append(title)

classTitles, courseDept, courseNum, courseName, courseSection, seats, meetings, meetingInfo, credit, isWI, courseType, isFYC = multidict(classDict)
          

# deletes unnecesary titles from one_class_dept
one_class_dept1 = one_class_dept.copy()
for i in one_class_dept1:
    if one_class_dept1[i] == "N/A":
        del one_class_dept[i]
        
# Preliminary information recordings about class

# Information about each class sections and meetings
with open("classes.txt", 'w') as f:
    for i in classDict:
        if i in crossListed:
            f.write("XList ({}) ".format(crossListed[i]))
        if i in labs:
            f.write("LAB ({}) ".format(labs[i][:-1]))
        f.write(i + " meeting: \n")
        count = 1
        if len(classDict[i][5]) == 0:
            f.write("TBA sec: " +  classDict[i][3] + " with {} open seats\n".format(classDict[i][4]))
        else:
            for k in classDict[i][5]:
                f.write(" ".join(k) + " sec: " +  classDict[i][3] + " with {} open seats\n".format(classDict[i][4]) )
        count += 1
        f.write("\n")

# Classes that have no meeting information
with open("noMeetings.txt", "w") as f:
    f.write("Classes with no meeting information found \n\n\n")
    for i in no_meetings:
        f.write(i + "\n")

# Information about multisection courses
with open('multisection.txt', 'w') as f:
    f.write("Multisection Courses \n\n\n")
    for i in multiSection:
        f.write(i + " " + str(multiSection[i]) + "\n")

# Information about crosslisted courses
with open('crosslisted.txt', 'w') as f:
    f.write("Crosslisted courses \n\n\n")
    for i in crossListed:
        f.write(i + " " + str(crossListed[i]) + "\n")

# Information about classes grouped by departments
with open("departments.txt", "w") as f:
    f.write("Courses grouped by departments\n")
    f.write("[!] Note these do not include crosslisted courses (each student has a personal departments dict) and labs.\n\n")
    for d in departments:
        f.write(d + ":")
        f.write(str(departments[d]))
        f.write("\n\n")

#---------------------------------------------------------------------------------
# Student Info (priorities.csv)


# process csv file
studentInfo = []
with open('priorities.csv', encoding='utf-8-sig') as csv_file:
    rawdata = csv.reader(csv_file, delimiter = ',')
    studentInfo = [row for row in rawdata]
indexList = studentInfo[0]


'''
Student Dict template
    [key] id -> [priorities, placements, title_to_courseName, studentDepts, name, email]
    [0] - priorities = dictionary of all classes
        class -> ranking (0 if it is not in any of priorities)
    [1] - placements = list of exam placements
    [2] - title_to_courseName = dict of courseNames used in preferences of each student
        title -> courseName
    [3] - studentDepts = dictionary of courses grouped by departments where crosslisted courses
                    are in the department of what student specified
            dept -> [course titles]
    [4] - name = name of the student
    [5] - email = email of the student
    [6] - majorInterest = area1 + area2 + area3 of student excel sheet
    [7] - stuType = 01 (Full time) or 02 (Full Time HEOP)
    [8] - grad = Graduate Education info
    
'''
studentDict = {}

notfound = [] # for courses that were not found at all 
unique_placements = [] # unique placements
no_priorities = [] # students with no priorities specified
studentsWrongPlacement = {} # id -> [[placements that are higher/not found, original choice, original courseName]]

# Record information about each student in studentDict

for row in studentInfo[1:]:
    id = row[indexList.index("id")] # student's id
    name = row[indexList.index("name")]
    email = row[indexList.index("ham email")]
    adjustment = 0 # used for skipping a class in priority list
    priorities = {} # classes ranking
    title_to_courseName = {} # title -> courseName dict used to extract department from student preferences
    studentDepts = deepcopy(departments) # titles of classes under specified departments by student
    ap_exams = {} # ap exam -> score
    exploredTitles = [] # used to check if a person put the same class under two different departments/course names
                        # which will ultimately ignore already explored titles
    majorInterest = "1) " + row[indexList.index("area 1")] + " 2) " + row[indexList.index("area 2")] + " 3) " + row[indexList.index("area 3")]
    stuType = "02 (Full Time HEOP)" if "y" in row[indexList.index("HEOP")].lower() else "01 (Full Time)"
    grad = row[indexList.index("graduate education")]
    # skip students with no priorities and record them later
    if row[indexList.index('priority 1')] == "":
        no_priorities.append([name, id])
        continue

    placementInfo = row[indexList.index('placements')]
    placements = extractPlacements(placementInfo) # tuples of [found (true/false), title/coursename] of extracted placements
    
    for p in placements:
        if not p[1] in unique_placements:
            unique_placements.append(p[1])

    for i in range(15):
        ap = row[indexList.index('AP 1')+i]
        if ap == "":
            break
        ap = ap.split("*")
        ap_exams[ap[0]] = int(ap[1])
    
    placements = convertAPtoPlacements(ap_exams, placements)

    # Initialize all classes to be 0
    for title in classDict:
        priorities[title] = 0

    # assign lunch priorities
    for day in "MTWRF":
        for index, lunch in enumerate(multiSection["Lunch " + day]):
            priorities[lunch] = LUNCH_PRIORITY
            if index == len(multiSection["Lunch " + day]) - 1:
                priorities[lunch] = LUNCH2PM_PRIORITY
            
    visited = {} # placements already counted for, so adjusment is correct
    for dept in PLACEMENTS:
        visited[dept] = False

    for i in range(12):
        shortTitle = row[indexList.index('priority 1')+i].upper()
        found, title = findTitle(shortTitle)
        try:
            dept = shortTitle.split()[0]
        except:
            dept = shortTitle

        if title in exploredTitles or title in notfound or title in labs:
            adjustment += 1
            continue
        
        rank = 12 - i + adjustment

        if not found:
            found, title = findTitle(shortTitle + "W") # adjust for writing intensive courses
            if title in exploredTitles or title in notfound:
                adjustment += 1
                continue

            if found:
                print("Adjusting priority for student {}: from {} to {} at {}".format(id, shortTitle, shortTitle + "W", rank))

            else:
                found, title = findTitle(shortTitle[:-1]) # adjust for non-writing intensive courses
                if title in exploredTitles or title in notfound:
                    adjustment += 1
                    continue

                if found:
                    print("Adjusting priority for student {}: from {} to {} at {}".format(id, shortTitle, shortTitle[:-1], rank))

                else:
                    notfound.append(shortTitle)
                    notfound.append(" no alternatives found...")
                    adjustment += 1
                    continue

        exploredTitles.append(title)

        # Placement precedence over preferences
        if dept in PLACEMENTS:
            placementFound = False

            if title in PLACEMENTS[dept]:

                if visited[dept]:
                    adjustment += 1
                    continue
                else:
                    visited[dept] = True

                for found, placement in placements:
                    
                    if found and placement in PLACEMENTS[dept]:
                        placementFound = True
                        if placement != title:
                            title = placement
                            exploredTitles.append(title)
                            break

                    elif not found and placement in SPECIAL_PLACEMENT_CASES and list(SPECIAL_PLACEMENT_CASES[placement].keys())[0] in PLACEMENTS[dept]:  
                        for j in SPECIAL_PLACEMENT_CASES[placement]:
                            if SPECIAL_PLACEMENT_CASES[placement][j] == 0:
                                adjustment += 1
                                if id in studentsWrongPlacement:
                                    studentsWrongPlacement[id].append([placement, 13 - rank, shortTitle])
                                else:
                                    studentsWrongPlacement[id] = [[placement, 13 - rank, shortTitle]]
                                
                            priorities[j] = round(rank * SPECIAL_PLACEMENT_CASES[placement][j])
                            title_to_courseName[j] = shortTitle
                            if j in multiSection:
                                for sec in multiSection[j]:
                                    priorities[sec] = round(rank * SPECIAL_PLACEMENT_CASES[placement][j])
                                    title_to_courseName[sec] = shortTitle
                            exploredTitles.append(j)
                        placementFound = True
                        rank = 0
                        break
                        

                # if no placement found, default should be level 1 class
                if not placementFound:
                    if not dept in REQUIRED_PLACEMENT:
                        title = PLACEMENTS[dept][0]
                        exploredTitles.append(title)
                    else: # except for the departments that require placements
                        adjustment += 1
                        continue

        # add title to the student picked department if crosslisted
        if title in crossListed:
            if dept in studentDepts:
                studentDepts[dept].append(title)
            else:
                studentDepts[dept] = [title]

        if rank == 0:
            continue
        
        priorities[title] = rank
        title_to_courseName[title] = shortTitle

        # add preferences for all sections
        if title in multiSection:
            for sec in multiSection[title]:
                priorities[sec] = rank
                title_to_courseName[sec] = shortTitle

    studentDict[id] = [priorities, placements, title_to_courseName, studentDepts, name, email, majorInterest, stuType, grad]

id, priorities, placements, title_to_courseName, studentDepts, name, email, majorInterest, stuType, grad = multidict(studentDict)

# Preliminary information recordings about students

# Unique placements found
with open("unique_placements.txt", "w") as f:
    f.write("All unique placements found on students preferences\n\n\n")
    for p in unique_placements:
        f.write(p + "\n")

# Classes on preference sheet that were not found in tour guide
with open('notfound.txt', 'w') as f:
    f.write("Courses chosen by students that were not found on Tour Guide\n\n\n")
    for i in notfound:
        try:
            if not i[0] == " ":
                f.write(i)
            else:
                f.write(i + "\n")
        except:
            f.write(i + "\n")

# Information about student's preferences and placements
with open("students.txt", 'w') as f:
    f.write("Students' preferences/choices and placements \n\n")
    f.write("[!] Note that the highest score is the top priority, and the highest for each student should be 12.\n\n")
    for i in studentDict:
        f.write("{} {} ({}) priorities: \n".format(studentDict[i][4], i, studentDict[i][5]))
        for j in studentDict[i][0]:
            if studentDict[i][0][j] != 0 and not "LUNCH" in j:
                f.write(" " + j + " " + str(studentDict[i][0][j]) + "\n")
        f.write("\n")
        f.write("Placements: \n")
        for j in studentDict[i][1]:
            f.write(str(j[0]) + " " + j[1] + "\n")
        f.write("\n")


print()
print(f"\r{bcolors.BOLD}{bcolors.OKCYAN}Done.{bcolors.ENDC}")
print()

print(bcolors.OKGREEN + "[!] Check 'students.txt' for assigned preferences." + bcolors.ENDC)
print(bcolors.OKGREEN + "[!] Check 'classes.txt' for meeting info of each class." + bcolors.ENDC)
print(bcolors.OKGREEN + "[!] Check 'multisection.txt' for multisection courses." + bcolors.ENDC)
print(bcolors.OKGREEN + "[!] Check 'crosslisted.txt' for cross listed dictionary information." + bcolors.ENDC)
print(bcolors.OKGREEN + "[!] Check 'departments.txt' for department info." + bcolors.ENDC)
print(bcolors.OKGREEN + "[!] Check 'notfound.txt' for course names in students' preferences but not found on tour guide." + bcolors.ENDC)
print(bcolors.OKGREEN + "[!] Check 'unique_placements.txt' for unique placements found in students' preferences." + bcolors.ENDC)
print(bcolors.OKGREEN + "[!] Check 'noMeetings.txt' for classes with no meeting information." + bcolors.ENDC)



print()
print(f"\r{bcolors.BOLD}{bcolors.OKCYAN}Building the model.{bcolors.ENDC}")

#-------------------------------------------------------------------------------------------------
print()


# MODEL
M = len(classTitles)

m = Model("Student Registration")

# PARAMS 


m.Params.IntegralityFocus = 1 # focuses on integer solutions (do not change)
m.setParam('MIPGap', 0) # finds a feasible solution within specified MIPGap of the optimal

print()
print(f"{bcolors.BOLD}{bcolors.OKCYAN}Setting up the model...{bcolors.ENDC}\n")

# # Variables
# # x_ij = 1 if student i is placed to class j.

x = {}
for i in id:
    for j in classTitles:
        x[i,j] = m.addVar(vtype = GRB.BINARY)
m.update()

# CONSTRAINTS

# 4 credits max
for i in id:
    m.addConstr(quicksum(x[i,j] * credit[j] for j in classTitles) <= 4.1, name = "credits")
m.update()


# No more than two labs per student
lab_constraint = []

for l in labs:
    if l in multiSection:
        for j in multiSection[l]:
            lab_constraint.append(j)
    else:
        lab_constraint.append(l)

for i in id:
    m.addConstr(quicksum(x[i,l] for l in lab_constraint)  <= 2, name = "upper bound on labs")
m.update()


# No two sections from the same course
for i in id:
    for k in multiSection:
        m.addConstr(quicksum(x[i,j] for j in multiSection[k]) <= 1, name = "multisections")
m.update()


# Don't put student in something they did not choose
for i in id:
    for j in classTitles:
        if (not j in labs) and (not j[:-2] in labs):
            m.addConstr(x[i,j] <= priorities[i][j], name = "no_bad_choices")
m.update()


# enrollment cap
for j in classTitles:
    m.addConstr(quicksum(x[i,j] for i in id) <= seats[j], name = "cap")
m.update()


# if a lab course, student in a lab section
# look at all bio classes prior and change the letters
bio_letters = ["D", "E", "G", "H", "I"]
for i in id:
    for l in labs:
        found, title = findTitle(labs[l][:-1])
        if not found: # bio -> have to add "a", "b", "e", "j", "n"
            m.addConstr(quicksum(x[i,l1] for l1 in multiSection[l]) == quicksum(x[i, findTitle(title + j)[1]] 
                        for j in bio_letters), name = "labs1")
        elif l in multiSection:
            if title in multiSection:
                m.addConstr(quicksum(x[i,l1] for l1 in multiSection[l]) == quicksum(x[i, j] 
                            for j in multiSection[title]), name = "labs2")
            else:
                m.addConstr(quicksum(x[i,l1] for l1 in multiSection[l]) == x[i, title], name = "labs3")
        else:
            if title in multiSection:
                m.addConstr(x[i,l] == quicksum(x[i, j] 
                            for j in multiSection[title]), name = "labs4")
            else:
                m.addConstr(x[i,l] == x[i, title], name = "labs5")
m.update()


# No more than one class per department: crosslisted -> look at students' preference and get department from there
for i in id:
    for d in studentDepts[i]:
        m.addConstr(quicksum(x[i,j] for j in studentDepts[i][d]) <= 1, name = "departments")
m.update()        


# No more than 3 classes in a division
for i in id:
    for d in DIVISIONS:
        divSections = []
        for dept in DIVISIONS[d]:
            if dept in studentDepts[i]:
                for j in studentDepts[i][dept]:
                    divSections.append(j)
        m.addConstr(quicksum(x[i,j] for j in divSections) <= 3, name = "DIVISIONS")
m.update()        


# Writing intensive: one WI unless they have a language 
for i in id:
    m.addConstr(quicksum(x[i,j] * isWI[j] for j in classTitles) <= 1, name = "WI")
m.update()

# No more than 1 FYC course
for i in id:
    m.addConstr(quicksum(isFYC[j] * x[i,j] for j in classTitles) <= 1, name = "FYC")

# Can only take one of math 152 + econ 100 + econ 166
restrictedTitles = ["STAT ANALYSIS OF DATA", "INTRODUCTION TO ECONOMICS", "ECON THEORY & EVIDENCE"]
restrictedSections = []

for j in restrictedTitles:
    if j in multiSection:
        for section in multiSection[j]:
            restrictedSections.append(section)
    else:
        restrictedSections.append(j)

for i in id:
    m.addConstr(quicksum(x[i,j] for j in restrictedSections) <= 1, name = "econ_fuss")
m.update()


# Overlapping times
count = 0
previous_count = count
interval = 1  # interval in seconds
start = time.time()
next_time = start + interval
for i in id:
    count += 1
    for j in classTitles:
        if time.time() >= next_time:
            load_log(round(count / len(id) * 100))
            next_time += interval
        m.addConstr(quicksum(x[i,k] for k in find_overlapping_classes(classDict, j)) <= M - M * x[i,j], name = "overlap")
m.update()


print()

# OBJECTIVE FUNCTION

m.setObjective(quicksum(LOOKUP[priorities[i][j]] * x[i,j] * ceil(credit[j]) for i in id for j in classTitles), GRB.MAXIMIZE)

m.update()
m.optimize()


# ------------------------------------------------------------------------------

# RESULTS
grand_total = 0
no_first_choice = []
not_four = {}
top_priorities = [1,2,3,4]
no_top_choices = []
no_top2 = []
no_lunch = {}


# Record each student's schedule in 'schedules.txt'
with ExcelWriter('result.xlsx', engine='xlsxwriter') as writer:
    with open("schedules.txt", "w") as f:
        f.write("""[!] Note that original course names may be different from titles due to placement overtaking precedence
                Original course name is whatever the student put down on the preference sheet\n\n""")
        studentID = [] #
        start = [] #
        year = [] #
        studentName = [] #
        studentEmail = [] #
        stuType_out = [] #
        major = [] #
        graduateEd = [] #
        placements_out = [] #
        sectionName = [] #
        courseTitle = [] #
        courseType_out = [] #
        meetingInfo_out = [] #
        choice = [] #
        
        for i in id:
            first_choice = False
            second_choice = True
            credits = 0
            count_top_choice = 0
            total = 0
            total_lunches = 0

            f.write("PLACEMENTS info for {}:\n".format(name[i]))

            placement_string = ""
            for p in placements[i]:
                placement_string += p[1]
                if p[0] == True:
                    found = ""
                    placement_string += " (" + courseName[p[1]] + ") |"
                else:
                    placement_string += " (Not found) | "
                    found = " (Not found)"
                    if i in studentsWrongPlacement:
                        for wrongplacement in studentsWrongPlacement[i]:
                            if wrongplacement[0] == placement:
                                found = " (Not found | ignored {} as choice {}) ".format(wrongplacement[2], wrongplacement[1]) 

                f.write(found + " " + p[1] + "\n")
            
            f.write("{} {} ({}) got:\n".format(name[i], i, email[i]))

            for j in classTitles:
                if x[i,j].x > .7: # x[i,j] = 1
                    priority = 0
                    record = True

                    if "LUNCH" in j:
                        total_lunches += 1
                        f.write("{} \n".format(j))
                        record = False

                    elif not j in lab_constraint:
                        if j in multiSection:
                            priority = min([13 - priorities[i][sec] for sec in multiSection[j]])
                        elif j[:-2] in multiSection:
                            priority = min([13 - priorities[i][sec] for sec in multiSection[j[:-2]]])
                        else:
                            priority = (13 - priorities[i][j])

                        if priority in top_priorities:
                            count_top_choice += 1

                        if priority == 1:
                            first_choice = True

                        if priority == 2:
                            second_choice = True

                        total += priority
                        f.write("{} | choice {} | original course name {}\n".format(j, priority, title_to_courseName[i][j]))

                    else:
                        f.write("{} \n".format(j))

                    if not "LUNCH" in j:
                        credits += credit[j]

                    seats[j] -= 1 # updates seats for each class
                        
                    f.write("@\t")
                    if len(meetings[j]) == 0:
                        f.write("TBA")
                    else:
                        for k in meetings[j]:
                            f.write(" ".join(k) + "; ")
                    f.write('\n\n')
                    
                    if record:
                        studentID.append(i)
                        start.append(START)
                        year.append(YEAR)
                        studentName.append(name[i])
                        studentEmail.append(email[i])
                        stuType_out.append(stuType[i])
                        major.append(majorInterest[i])
                        graduateEd.append(grad[i]) 
                        placements_out.append(placement_string)
                        sectionName.append("-".join(title_to_courseName[i][j].split()) + "-" + courseSection[j] if not j in lab_constraint
                                            else courseDept[j] + "-" + courseNum[j] + "-" + courseSection[j])
                        courseTitle.append(j)
                        courseType_out.append(" ".join(courseType[j].split()) if courseType != "" else "")
                        meetingInfo_out.append(meetingInfo[j])
                        choice.append(priority)

            if total_lunches < 5:
                no_lunch[i] = total_lunches

            if not first_choice:
                no_first_choice.append(i)

            if not first_choice and not second_choice:
                no_top2.append(i)

            if count_top_choice == 0:
                no_top_choices.append(i)

            if credits != 4:
                not_four[i] = credits
            grand_total += total   

            f.write("-----------------------------------------")
            f.write("\n\n")

        df1 = DataFrame({"ID": studentID, "Start": start, "Year": year, "Student Name": studentName, "email": studentEmail, 
                    "Stu Type": stuType_out, "Major Interests": major, "Graduate Education": graduateEd, "Placements": placements_out,
                    "Section Name": sectionName, "Title": courseTitle, "Course Type": courseType_out,  "Meeting Info": meetingInfo_out, "Priority": choice})  
        

    df1.to_excel(writer, sheet_name = "Student Info", index = False)
    workbook = writer.book
    worksheet = writer.sheets['Student Info']
    
    # Adjusts the width of columns in excel
    for i, col in enumerate(df1.columns):
        max_length = max(df1[col].astype(str).map(len).max(), len(col))
        worksheet.set_column(i, i, max_length + 2) 


    # Record statistics of the run in results.txt
    with open("results.txt", "w") as f:
        f.write("Average ranking: {}\n\n".format(grand_total / len(id)))
        f.write("------------------------------------------------------------------------------\n\n")
        f.write("Students that did not complete preferences (total = {}):\n\n".format(len(no_priorities)))
        for i in no_priorities:
            f.write("{} {}\n\n".format(i[0], i[1]))
        f.write("------------------------------------------------------------------------------\n\n\n\n")
        f.write("Students that did not get all four credits (total = {}):\n\n".format(len(not_four)))
        for i in not_four:
            f.write("{} {} with only {} credits\n\n".format(name[i], i, not_four[i]))
        f.write("------------------------------------------------------------------------------\n\n\n\n")

        f.write("Students that did not get any of the 4 top choices (total = {}):\n\n".format(len(no_top_choices)))
        for i in no_top_choices:
            f.write("{} {}\n".format(name[i], i))
        f.write("------------------------------------------------------------------------------\n\n\n\n")

        f.write("Students that did not get first choice (total = {}):\n\n".format(len(no_first_choice)))
        for i in no_first_choice:
            f.write("{} {}\n".format(name[i], i))
        f.write("------------------------------------------------------------------------------\n\n\n\n")

        f.write("Students that did not get first or second choice (total = {}):\n\n".format(len(no_top2)))
        for i in no_top2:
            f.write("{} {}\n".format(name[i], i))
        f.write("------------------------------------------------------------------------------\n\n\n\n")

        f.write("Students that did not get all 5 lunches (total = {}):\n\n".format(len(no_lunch)))
        for i in no_lunch:
            f.write("{} {}\n".format(name[i], i))
        f.write("------------------------------------------------------------------------------\n\n\n\n")

        empty_seats = 0

        sectionName = []
        courseTitle = []
        courseSeats = []
        meetingInfo_out = []
        crossListed_out = []
        for j in seats:
            if seats[j] > 0 and not "LUNCH" in j:
                f.write("{} has {} empty seats\n".format(j, seats[j]))
                empty_seats += seats[j]

            if not "LUNCH" in j:
                sectionName.append(courseDept[j] + "-" + courseNum[j] + "-" + courseSection[j])
                courseTitle.append(j)
                courseSeats.append(seats[j])
                meetingInfo_out.append(meetingInfo[j])
                crossListed_out.append("Y" if j in crossListed else "N")


        f.write("\nTotal of {} empty seats.\n\n".format(empty_seats))

        f.write("------------------------------------------------------------------------------\n\n\n\n")
        df2 = DataFrame({"Section Name": sectionName, "Short Title": courseTitle, "Seats Available": courseSeats, "Meeting Info": meetingInfo_out, "Crosslisted?": crossListed_out})
    df2.to_excel(writer, sheet_name = "Courses Info", index = False)
    workbook = writer.book
    worksheet = writer.sheets['Courses Info']
    
    # Adjusts the width of columns in excel
    for i, col in enumerate(df2.columns):
        max_length = max(df2[col].astype(str).map(len).max(), len(col))
        worksheet.set_column(i, i, max_length + 2) 

print()
print(bcolors.OKGREEN + "[!] Check 'results.txt' for statistics of the latest run." + bcolors.ENDC)
print(bcolors.OKGREEN + "[!] Check 'result.xlsx' for schedules of each student." + bcolors.ENDC)
