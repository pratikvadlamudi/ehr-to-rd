#!pip install pandas
#!pip install datetime
#!pip install openpyxl

import pandas as pd
import numpy as np
import os
import re
import glob

def removesuffix(input: str, suffix: str) -> str:
      if input.endswith(suffix):
        return input[:-1*len(suffix)]
      else :
        return input[ :]

## DEFINITIONS

path = './files/' #need to set
newlines = True #NEED TO MAKE SURE THIS IS CORRECT

os.chdir(path)
files = glob.glob("*.txt")
file_name_list = []

for f in files:
#   print(f)
  file_name_list.append(f)
  file_name_list = sorted(file_name_list, reverse=False)

strings = ['Date of Exam:' , 'Exam Start Time:' , 'Was the patient a walk in or were they scheduled?', 'Does participant have flashes or floaters which have changed recently', 'Has participant had any surgery or laser surgery to their eyes?', 'Has participant had any history of trauma to your eyes such as boxing,', 'Is participant taking any eye drops other than artificial tears or over the counter medicatio', 'If "Yes", participant is taking eye drops other than artificial tears or over the counter medications, please list:', 'Was the visual acuity taken with correction', 'Presenting visual acuity (Right/OD) DISTANCE', 'If "Other", presenting visual acuity score (Right/OD) DISTANCE, specify: ', 'Presenting visual acuity (Left/OS) DISTANCE', 'If "Other", Presenting visual acuity score (Left/OS) DISTANCE, specify:', 'Post-refraction Best corrected visual acuity (Right/OD) DISTANCE', 'If "Other", Best corrected visual acuity score (Right/OD) DISTANCE, specify:', 'Post-refraction Best corrected visual acuity (Right/OD) NEAR', 'If "Other", Best corrected visual acuity score (Right/OD) NEAR, specify:', 'Post-refraction Best corrected visual acuity (Left/OS) DISTANCE:', 'If "Other" Best corrected visual acuity score (Left/OS) DISTANCE, specify:', 'Post-refraction Best corrected visual acuity (Left/OS) NEAR:', 'If "Other", Best corrected visual acuity score (Left/OS) NEAR, specify:', 'Contrast sensitivity (Right/OD):', 'If "Other", Contrast sensitivity (Right/OD), specify:', 'Contrast sensitivity (Left/OS):', 'If "Other", Contrast sensitivity (Left/OS), specify:', 'Autorefraction (Right/OD) Sphere:', 'Autorefraction (Right/OD) Cylinder:', 'Autorefraction (Right/OD) Axis:', 'Autorefraction (Left/OS) Sphere:', 'Autorefraction (Left/OS) Cylinder:', 'Autorefraction (Left/OS) Axis:', 'Manifest refraction (Right/OD) Sphere:', 'Manifest refraction (Right/OD) Cylinder:', 'Manifest refraction (Right/OD) Axis:', 'Manifest refraction (Right/OD) Add', 'Manifest refraction (Left/OS) Sphere:', 'Manifest refraction (Left/OS) Cylinder:', 'Manifest refraction (Left/OS) Axis:', 'Manifest refraction (Left/OS) Add:', 'Was trial frame performed?', 'If "No", trial frame, specify why:', 'Afferent pupillary defect (Right/OD):', 'Afferent pupillary defect (Left/OS):', 'Intraocular Pressure Measurement Tim', 'Intraocular pressure measurement(s) (Right/OD)', 'Intraocular pressure measurement(s) (Left/OS)', 'Is third intraocular pressure measurement required?', 'Pachymetry (Right/OD) NOT DONE:', 'Pachymetry (Right/OD):', 'Pachymetry (Left/OS) NOT DONE:', 'Pachymetry (Left/OS):',  'Narrow anterior chamber angle by penlight (Right/OD):', 'Narrow anterior chamber angle by penlight (Left/OS):',  'Dilation performed', 'Which eye did you dilate',  'If "No", specify why:', 'Extraocular motility (Right/OD):',  'Extraocular motility (Left/OS):', 'Extraocular alignment (Right/OD):',  'Extraocular alignment (Left/OS):', 'Inter-pupillary distance NOT DONE:',  'Inter-pupillary distance:', 'Was a glasses order started?',  'If "Yes", where were the glasses ordered from?',  'If "Other", specify:', 'If "Yes", how much did glasses cost including shipping?',  'If "Yes", how many pairs of glasses were ordered?',  'If "Yes", how did the participant feel about their glasses?',  'Comments:', 'Exam End Time']

mislinedStrings = ['Presenting visual acuity (Right/OD) DISTANCE', 'If "Other", presenting visual acuity score (Right/OD) DISTANCE, specify: ', 'Presenting visual acuity (Left/OS) DISTANCE', 'If "Other", Presenting visual acuity score (Left/OS) DISTANCE, specify:', 'Post-refraction Best corrected visual acuity (Right/OD) DISTANCE', 'If "Other", Best corrected visual acuity score (Right/OD) DISTANCE, specify:', 'Post-refraction Best corrected visual acuity (Right/OD) NEAR', 'If "Other", Best corrected visual acuity score (Right/OD) NEAR, specify:', 'Post-refraction Best corrected visual acuity (Left/OS) DISTANCE:', 'If "Other" Best corrected visual acuity score (Left/OS) DISTANCE, specify:', 'Post-refraction Best corrected visual acuity (Left/OS) NEAR:', 'If "Other", Best corrected visual acuity score (Left/OS) NEAR, specify:', 'Contrast sensitivity (Right/OD):', 'If "Other", Contrast sensitivity (Right/OD), specify:', 'Contrast sensitivity (Left/OS):', 'If "Other", Contrast sensitivity (Left/OS), specify:', 'Autorefraction (Right/OD) Sphere:', 'Autorefraction (Right/OD) Cylinder:', 'Autorefraction (Right/OD) Axis:', 'Autorefraction (Left/OS) Sphere:', 'Autorefraction (Left/OS) Cylinder:', 'Autorefraction (Left/OS) Axis:', 'Manifest refraction (Right/OD) Sphere:', 'Manifest refraction (Right/OD) Cylinder:', 'Manifest refraction (Right/OD) Axis:', 'Manifest refraction (Right/OD) Add', 'Manifest refraction (Left/OS) Sphere:', 'Manifest refraction (Left/OS) Cylinder:', 'Manifest refraction (Left/OS) Axis:', 'Manifest refraction (Left/OS) Add:', 'Afferent pupillary defect (Right/OD):', 'Afferent pupillary defect (Left/OS):', 'Pachymetry (Right/OD) NOT DONE:', 'Pachymetry (Right/OD):', 'Pachymetry (Left/OS) NOT DONE:', 'Pachymetry (Left/OS):',  'Which eye did you dilate',  'Inter-pupillary distance:']

data = []
cols = ['record_id', 'examdat', 'examstrttm', 'examvsttype', 'flshfloatr',  'eyesurgery', 'eyetrauma', 'eyedrops', 'eyedropstx', 'pvawthc',  'pvasdisrod', 'pvasdisrodtx', 'pvasdislos', 'pvasdislostx', 'bcvasdisrod', 'bcvasdisrodtx', 'bcvasnearrod', 'bcvasnearrodtx', 'bcvasdislos', 'bcvasdislostx', 'bcvasnearlos', 'bcvasnearlostx', 'csrod', 'csrodtx', 'cslos', 'cslostx', 'arodsphere', 'arodcylin', 'arodaxis', 'arlossphere', 'arloscylin', 'arlosaxis', 'mrrodsphere', 'mrrodcylin', 'mrrodaxis', 'mrrodadd', 'mrlossphere', 'mrloscylin', 'mrlosaxis', 'mrlosadd', 'trialframe', 'trialframetx', 'apdrod', 'apdlos', 'ipmtm', 'frstipmrod', 'frstipmlos', 'secipmrod', 'secipmlos', 'thrdipmreq', 'thrdipmrod', 'thrdipmlos', 'phmeryrodnd___87', 'phmeryrod', 'phmerylosnd___87', 'phmerylos', 'nacabprod', 'nacabplos', 'dilation', 'dilationeye', 'nodilation', 'emrod', 'emlos', 'earod', 'ealos', 'ipdnd___87', 'ipd', 'glassesord', 'glassesfrom', 'glassesfromtx', 'glassescost', 'glassesnum', 'glassesfeel', 'examcnotes', 'examendtm']

df = pd.DataFrame(data, columns = cols)

## MAIN

# setting flag and index to 0
flag = 0
index = 0
count = 0
filec = 0

while (filec < len(file_name_list)):
  count = 0
  name = file_name_list[filec]
  name = name[0:6]
  variable_list = [name] #record_id
  while (count < len(strings)):
    # Loop through the file line by line
    string = file_name_list[filec]
    file = open(string, encoding = "ISO-8859-1")
    flag = 0
    mislinedFlag = 0
    index = 0
    for line in file:  
        index += 1 
        # checking string is present in line or not
        if (strings[count] in line):
          flag = 1
          if (strings[count] in mislinedStrings) and newlines == True:
            mislinedFlag = 1 #see, typically when you paste from michart there are newlines 
             #between some of the q's (in the list mislinedStrings) and a's
          break
      # checking condition for string found or not
    if flag == 0:
      variable_list.append('')
    else: #string is in the line
      if mislinedFlag == 0: #question and variable are in same line
        variable = line
        if (':' in variable):
          variable = variable.split(':', 1)
          variable = variable[1]
          variable = variable.split('\n', 1)
          variable_list.append(variable[0])
        elif ('?' in variable):
          variable = variable.split('?', 1)
          variable = variable[1]
          variable = variable.split('\n', 1)
          variable_list.append(variable[0])
        else:
          variable_list.append(variable)
      else: #if the variable is mislined from question (question and variable are in different line)
        if next(file).strip() == "": #in-between line is blank
          variable = next(file) #following line is pulled and assigned to 'variable'
        else:
          variable = 'MANUAL: In-between line was not blank'
        variable_list.append(variable)

    count = count + 1

  if ',' in variable_list[44]: #intraocular pressure measurement time #MODIFIED
    time = variable_list[44].split(',')
    time = time[0] #redcap doesn't care about the times for intraocular preasure measurement #2 or #3
    del variable_list[44]
    variable_list.insert(44, time)

  if ',' in variable_list[45]: #intraocular pressure right/OD #MODIFIED
    right = variable_list[45].split(',')
    first_right = right[0]
    second_right = right[1]
    third_right = right[2]
  else:
    first_right = "MANUAL: see original file"
    second_right = "MANUAL: see original file"
    third_right = "MANUAL: see original file"

  if ',' in variable_list[46]: #intraocular pressure left/OS #MODIFIED
    left = variable_list[46].split(',')
    first_left = left[0]
    second_left = left[1]
    try:
      third_left = left[2]
    except IndexError:
      third_left = "MANUAL: IndexError"
  else:
    first_left = "MANUAL: see original file"
    second_left = "MANUAL: see original file"
    third_left = "MANUAL: see original file"


  del variable_list[45]
  del variable_list[45]
  variable_list.insert(45, first_right)
  variable_list.insert(46, first_left)
  variable_list.insert(47, second_right)
  variable_list.insert(48, second_left)
  variable_list.insert(50, third_right)
  variable_list.insert(51, third_left)

  list_count = 0

  while (list_count < len(variable_list)):
    temp = variable_list[list_count].lstrip()
    if (temp.startswith('yes') or temp.startswith('Yes') or temp.startswith('Yes')):
      variable_list[list_count] = 'Yes'
    if (temp.startswith('no') or temp.startswith('No')):
      variable_list[list_count] = 'No'
    if ('plano' in temp):
      variable_list[list_count] = '0'
    if ('sphere' in temp):
      variable_list[list_count] = '0' 
    if ('Plano' in temp):
      variable_list[list_count] = '0'
    if ('Sphere' in temp):
      variable_list[list_count] = '0' 
    list_count = list_count + 1

# ['record_id', 'examdat', 'examstrttm', 'examvsttype', 'flshfloatr', 'eyesurgery', 'eyetrauma', 'eyedrops', 'eyedropstx', 'pvawthc', 'pvasdisrod', 'pvasdisrodtx', 'pvasdislos', 'pvasdislostx', 'bcvasdisrod', 'bcvasdisrodtx', 'bcvasnearrod', 'bcvasnearrodtx', 'bcvasdislos', 'bcvasdislostx', 'bcvasnearlos', 'bcvasnearlostx', 'csrod', 'csrodtx', 'cslos', 'cslostx', 'arodsphere', 'arodcylin', 'arodaxis', 'arlossphere', 'arloscylin', 'arlosaxis', 'mrrodsphere', 'mrrodcylin', 'mrrodaxis', 'mrrodadd', 'mrlossphere', 'mrloscylin', 'mrlosaxis', 'mrlosadd', 'trialframe', 'trialframetx', 'apdrod', 'apdlos', 'ipmtm', 'frstipmrod', 'frstipmlos', 'secipmrod', 'secipmlos', 'thrdipmreq', 'thrdipmrod', 'thrdipmlos', 'phmeryrodnd', 'phmeryrod', 'phmerylosnd', 'phmerylos', 'nacabprod', 'nacabplos', 'dilation', 'dilationeye', 'nodilation', 'emrod', 'emlos', 'earod', 'ealos', 'ipdnd', 'ipd', 'glassesord', 'glassesfrom', 'glassesfromtx', 'glassescost', 'glassesnum', 'glassesfeel', 'examcnotes', 'examendtm']
  df2 = pd.DataFrame([variable_list], columns = cols)
  df = pd.concat([df2, df])
  filec = filec + 1

df = df.reset_index(level=0, drop=True)

## FORMATTING

# Reformat the exam start time and the exam end time
def reformat_time(time): #TODO this doesn't work for 8:__ or 9:__ times
    s = re.match(r"(.*)(\d{2})(.*)(\d{2})(.*)", time)
    if s is not None and s.group(1).rstrip()=="": #need to check that there is nothing before the time parsed
        return s.group(2) + ":" + s.group(4) #s.group(2) is the hour and s.group(4) is the minute
    else:
        return "MANUAL: "+time

df['examstrttm'] = df['examstrttm'].apply(lambda x: reformat_time(x))
df['examendtm'] = df['examendtm'].apply(lambda x: reformat_time(x))

# Reformat ipmtm
def reformat_ipmtm(time):
    if time[-2:] == 'PM':
        s = re.match(r"([1-9]):(\d{2})", time)
        if s is not None:
            return str(int(s.group(1))+12) + ":" + s.group(2)
        else:
            s1 = re.match(r"(1?[0-9]):(\d{2})", time)
            if s1 is not None:
                return s1.group(1) + ":" + s1.group(2)
            else:
                return "N/A"
    else:
        s = re.match(r"(1?[0-9]):(\d{2})", time)
        if s is not None:
            return s.group(1) + ":" + s.group(2)
        else:
            return "MANUAL: "+time

df['ipmtm'] = df['ipmtm'].apply(lambda x: x.replace(' ', ''))
df['ipmtm'] = df['ipmtm'].apply(lambda x: reformat_ipmtm(x))

# Remove " from all columns
df = df.applymap(lambda x: x.replace('"', '') if (isinstance(x, str)) else x)
df = df.applymap(lambda x: x.strip() if (isinstance(x, str)) else x)

# Replace N/A with empty value
df[cols] = df[cols].replace("N/A", np.nan)

def format_flshfloatr(entry):
  if entry == "No" or entry == "None":
    return 0
  elif entry == "Yes" or entry == "Trace":
    return 1
  else:
    return 'MANUAL: '+str(entry)

def format_trialframe(entry):
  if entry=="Yes":
    return 1
  else:
    return 'MANUAL: ALSO edit trialframetx: '+str(entry)

def format_pvasdisrod(entry):
  if str(entry).startswith("20/15"):
    return 1
  elif str(entry).startswith("20/20") and not str(entry).startswith("20/200"):
    return 2
  elif str(entry).startswith("20/25"):
    return 3
  elif str(entry).startswith("20/30"):
    return 4
  elif str(entry).startswith("20/40") and not str(entry).startswith("20/400"):
    return 5
  elif str(entry).startswith("20/50"):
    return 6
  elif str(entry).startswith("20/60"):
    return 7
  elif str(entry).startswith("20/70"):
    return 8
  elif str(entry).startswith("20/80"):
    return 9
  elif str(entry).startswith("20/100"):
    return 10
  elif str(entry).startswith("20/200"):
    return 11
  elif str(entry).startswith("20/400"):
    return 12
  elif str(entry).startswith("< 20/400") or str(entry).startswith("<20/400"):
    return 13
  elif entry == "CF" or entry=="CF @ 3'":
    return 14
  elif entry == "HM":
    return 15
  elif entry == "LP":
    return 16
  elif entry == "NLP":
    return 17
  elif str(entry).startswith("20/13"):
    return 18
  elif str(entry).startswith("20/10"):
    return 19
  elif entry == "Other":
    return 88
  else:
    return 'MANUAL: '+entry

def format_csrod(entry):
  if entry == "0.00":
    return 1
  elif entry == "0.15":
    return 2
  elif entry == "0.30":
    return 3
  elif entry == "0.45":
    return 4
  elif entry == "0.60":
    return 5
  elif entry == "0.75":
    return 6
  elif entry == "0.90":
    return 7
  elif entry == "1.05":
    return 8
  elif entry == "1.20":
    return 9
  elif entry == "1.35":
    return 10
  elif entry == "1.50":
    return 11
  elif entry == "1.65":
    return 12
  elif entry == "1.80":
    return 13
  elif entry == "1.95":
    return 14
  elif entry == "2.10":
    return 15
  elif entry == "2.25":
    return 16
  else:
    return 'MANUAL: '+str(entry)

def format_arodsphere(entry):
  if entry=="none" or entry=="None" or entry == "no" or entry=="No":
    return 0
  elif entry.replace('-','').replace('+','').replace('.','').isnumeric():
    return entry
  else:
    return 'MANUAL: '+str(entry)

def format_frstipmrod(entry):
  if str(entry).isnumeric():
    return entry
  else:
    return 'MANUAL: '+str(entry)

def format_phmeryrod(entry):
  if entry.isnumeric():
    return entry
  else:
    return 'MANUAL: make this field BLANK and enter the integer \'1\' into phmeryod__87 or phmeryos__87'

def format_dilation(entry):
  if entry == "No" or entry == "None":
    return 'MANUAL: Dilation was not done. Check dilationeye and nodilation.'
  elif entry == "Yes" or entry == "Trace":
    return 1
  else:
    return 'MANUAL: '+str(entry)

def format_dilationeye(entry):
  if entry == "Yes" or entry == "Both eyes":
    return 3
  elif entry == "Left eye":
    return 2
  elif entry == "Right eye":
    return 1
  elif entry == "None":
    return "MANUAL: change this to 0, and check nodilation"
  else:
    return 'MANUAL: '+str(entry)+'. And check nodilation'

def format_emrod(entry):
  if entry == "Full":
    return 1
  elif entry == "Limited abduction":
    return 2
  elif entry == "Limited adduction":
    return 3
  elif entry == "Limited supraduction":
    return 4
  elif entry == "Limited infraduction":
    return 5
  else:
    try:
      return 'MANUAL: '+str(entry)
    except TypeError:
      return 'MANUAL'

def format_ipd(entry):
  entry = removesuffix(entry, ' mm')
  if entry.replace('.','',1).isdigit(): #checks if the string without first decimal point is made of digits
    return entry
  else:
    return 'MANUAL: '+str(entry)

def format_examvsttype(entry):
  if entry == "Walk-in":
    return 1
  elif entry == "Walk in":
    return 1
  elif entry == "Scheduled":
    return 2
  else: 
    return 'MANUAL: '+str(entry)

df['flshfloatr'] = df['flshfloatr'].apply(lambda x: format_flshfloatr(x))
df['eyesurgery'] = df['eyesurgery'].apply(lambda x: format_flshfloatr(x))
df['eyetrauma'] = df['eyetrauma'].apply(lambda x: format_flshfloatr(x))
df['eyedrops'] = df['eyedrops'].apply(lambda x: format_flshfloatr(x))
df['pvawthc'] = df['pvawthc'].apply(lambda x: format_flshfloatr(x))

df['pvasdisrod'] = df['pvasdisrod'].apply(lambda x: format_pvasdisrod(x))
df['pvasdislos'] = df['pvasdislos'].apply(lambda x: format_pvasdisrod(x))
df['bcvasdisrod'] = df['bcvasdisrod'].apply(lambda x: format_pvasdisrod(x))
df['bcvasnearrod'] = df['bcvasnearrod'].apply(lambda x: format_pvasdisrod(x))
df['bcvasdislos'] = df['bcvasdislos'].apply(lambda x: format_pvasdisrod(x))
df['bcvasnearlos'] = df['bcvasnearlos'].apply(lambda x: format_pvasdisrod(x))

df['csrod'] = df['csrod'].apply(lambda x: format_csrod(x))
df['cslos'] = df['cslos'].apply(lambda x: format_csrod(x))

df['arodsphere'] = df['arodsphere'].apply(lambda x: format_arodsphere(x))
df['arodcylin'] = df['arodcylin'].apply(lambda x: format_arodsphere(x))
df['arodaxis'] = df['arodaxis'].apply(lambda x: format_arodsphere(x))
df['arlossphere'] = df['arlossphere'].apply(lambda x: format_arodsphere(x))
df['arloscylin'] = df['arloscylin'].apply(lambda x: format_arodsphere(x))
df['arlosaxis'] = df['arlosaxis'].apply(lambda x: format_arodsphere(x))
df['mrrodsphere'] = df['mrrodsphere'].apply(lambda x: format_arodsphere(x))
df['mrrodcylin'] = df['mrrodcylin'].apply(lambda x: format_arodsphere(x))
df['mrrodaxis'] = df['mrrodaxis'].apply(lambda x: format_arodsphere(x))
df['mrrodadd'] = df['mrrodadd'].apply(lambda x: format_arodsphere(x))
df['mrlossphere'] = df['mrlossphere'].apply(lambda x: format_arodsphere(x))
df['mrloscylin'] = df['mrloscylin'].apply(lambda x: format_arodsphere(x))
df['mrlosaxis'] = df['mrlosaxis'].apply(lambda x: format_arodsphere(x))
df['mrlosadd'] = df['mrlosadd'].apply(lambda x: format_arodsphere(x))

df['trialframe'] = df['trialframe'].apply(lambda x: format_flshfloatr(x))
df['apdrod'] = df['apdrod'].apply(lambda x: format_flshfloatr(x))
df['apdlos'] = df['apdlos'].apply(lambda x: format_flshfloatr(x))

df['frstipmrod'] = df['frstipmrod'].apply(lambda x: format_frstipmrod(x))
df['frstipmlos'] = df['frstipmlos'].apply(lambda x: format_frstipmrod(x))
df['secipmrod'] = df['secipmrod'].apply(lambda x: format_frstipmrod(x))
df['secipmlos'] = df['secipmlos'].apply(lambda x: format_frstipmrod(x))
df['thrdipmrod'] = df['thrdipmrod'].apply(lambda x: format_frstipmrod(x))
df['thrdipmlos'] = df['thrdipmlos'].apply(lambda x: format_frstipmrod(x))

df['phmeryrod'] = df['phmeryrod'].apply(lambda x: format_phmeryrod(x))
df['phmerylos'] = df['phmerylos'].apply(lambda x: format_phmeryrod(x))

df['nacabprod'] = df['nacabprod'].apply(lambda x: format_flshfloatr(x))
df['nacabplos'] = df['nacabplos'].apply(lambda x: format_flshfloatr(x))

df['dilation'] = df['dilation'].apply(lambda x: format_dilation(x))

df['dilationeye'] = df['dilationeye'].apply(lambda x: format_dilationeye(x))

df['emrod'] = df['emrod'].apply(lambda x: format_emrod(x))
df['emlos'] = df['emlos'].apply(lambda x: format_emrod(x))
df['earod'] = df['earod'].apply(lambda x: format_emrod(x))
df['ealos'] = df['ealos'].apply(lambda x: format_emrod(x))

df['glassesord'] = df['glassesord'].apply(lambda x: format_flshfloatr(x))

df['ipd']=df['ipd'].apply(lambda x: format_ipd(x))

df['examvsttype'] = df['examvsttype'].apply(lambda x: format_examvsttype(x))

df.to_excel('a'+str(newlines)+'TechExam.xlsx', index=False)
