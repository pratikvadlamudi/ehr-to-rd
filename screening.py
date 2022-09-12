#!pip install python-dateutil
#!pip install pandas
#!pip install datetime
#!pip install openpyxl
# When you install openpyxl, if you are on MacOS, you can decline the prompt to install Command Line Tools

import pandas as pd
import numpy as np
import os
from datetime import datetime
from datetime import timedelta
from dateutil import parser
import re
import glob

def removeprefix(input: str, prefix: str) -> str:
      if input.startswith(prefix):
       return input[len(prefix): ]
      else :
       return input[: ]

def removesuffix(input: str, suffix: str) -> str:
      if input.endswith(suffix):
        return input[:-1*len(suffix)]
      else :
        return input[ :]

## DEFINITIONS

path = './files/' # need to add path

os.chdir(path)
file_name_list = sorted(glob.glob("*.txt")) # import sorted filenames into list files

strings = ['Start time' , 'Was presenting visual acuity worse than 20/40 in either eye (presented with visual impairment)?' , 'Did the patient have refractive error?', 'I agree with the measured manifest refraction and recommend prescribing it.', 'Were external photos taken?', 'External photos missing?', 'Were external photos of sufficient quality for interpretation?', 'Were external photos normal?', 'Were all fundus photos taken?', 'Fundus photos missing?', 'Were fundus photos of sufficient quality for interpretation?', 'Vertical cup to disc ratio OD:', 'Vertical cup to disc ratio OS:', 'Were all Optical Coherence Tomography (OCT) photos taken?', 'OCT photos missing?', 'Were OCT images of sufficient quality for interpretation?', 'OD: overall thickness', 'OD checklist', 'OS: overall thickness', 'OS checklist', 'Any signs of glaucoma?', 'Were there any abnormalities on the fundus exam?', 'Follow-up with an ophthalmologist?', 'Assessment and Plan', 'Stop time', 'Letter sent to', 'Progress Notes']

cols = ['record_ID', 'extnlpdat', 'extnlpsrttm', 'extnlpname', 'extnlpnametx', 'visimpairmnt', 'visimpairmnteye', 'refctverror', 'srdescrip', 'visimpairmnteye_2', 'refractyperod___1', 'refractyperod___2', 'refractyperod___3', 'refractyperod___4', 'refractyperod___5', 'refractyperod___6', 'refractyperod___7', 'refractyperod___99', 'refractfurod___1', 'refractfurod___2', 'refractfurod___3', 'refractfurod___99', 'refractypelos___1', 'refractypelos___2', 'refractypelos___3', 'refractypelos___4', 'refractypelos___5', 'refractypelos___6', 'refractypelos___7', 'refractypelos___99', 'refractfulos___1', 'refractfulos___2', 'refractfulos___3', 'refractfulos___99', 'refractrecpresc', 'extnlp', 'extnlpmiss', 'rodextnlpqlty', 'rodextnlpqltytx', 'losextnlpqlty', 'losextnlpqltytx', 'extnlpnorm', 'extnlpabnrmloc___1', 'extnlpabnrmloc___2',  'extnlpabnrmloc___3',  'extnlpabnrmloc___4',  'extnlpabnrmloc___88',  'extnlpabnrmloctx',  'lensabnrm___1',  'lensabnrm___2',  'lensabnrm___88',  'intraocirlneye',  'intralentcom',  'cataracteye',  'cataract',  'cataractcom',  'lensabnrmtx',  'lensabnrmotheye',  'lensabnrmothtx',  'cornabnrm___1',  'cornabnrm___2',  'cornabnrm___3',  'cornabnrm___4',  'cornabnrm___88',  'dryeye',  'dryeyeartears',  'dryeyetx',  'pterygiumeye',  'pterygiumrod',  'pterygiumlos',  'pterygiumtx',  'cornopactyeye',  'cornopactytx',  'cornulcereye',  'cornulcertx',  'cornabnrmtx',  'cornabnrmotheye',  'irisabnrm___1',  'irisabnrm___88',  'irislesioneye',  'irislesiontx',  'irisabnrmtx',  'irisabnrmotheye',  'irisabnrmothtx',  'extnabnrm___1',  'extnabnrm___2',  'extnabnrm___88',  'lidlesioneye',  'lidlesionrod',  'lidlesionlos',  'lidlesiontx',  'ptosisobviseye',  'ptosisobvistx',  'extnabnrmtx',  'extnabnrmotheye',  'extnabnrmothtx',  'fundusp', 'funduspmiss',  'rodfunduspqlty', 'rodfunduspqltytx',  'losfunduspqlty',  'losfunduspqltytx',  'ctdrrod', 'ctdrlos', 'octp', 'octpmiss',  'rodoctpqlty', 'rodoctpqltytx',  'losoctpqlty',  'losoctpqltytx',  'rodthickns', 'rodsupr', 'rodinferior', 'rodnasal', 'rodtemprl', 'losthickns', 'lossupr', 'losinferior', 'losnasal', 'lostemprl', 'glaucoma', 'glaucstrod',  'glaucstlos',  'srdescrip2',  'glaucsignrod___1',  'glaucsignrod___2',  'glaucsignrod___3',  'glaucsignrod___4',  'glaucsignrod___5',  'glaucsignrod___6',  'glaucsignrod___7',  'glaucsignrod___8',  'glaucsignrod___88',  'glaucsignrodtx',  'glaucsignlos___1',  'glaucsignlos___2',  'glaucsignlos___3',  'glaucsignlos___4',  'glaucsignlos___5',  'glaucsignlos___6',  'glaucsignlos___7',  'glaucsignlos___8',  'glaucsignlos___88',  'glaucsignlostx',  'fundusabnrm', 'fundusreslts___1',  'fundusreslts___2',  'fundusreslts___88',  'dretingraderod',  'rodmacedema',  'dretingradelos',  'losmacedema',  'amd',  'amdtyperod',  'amdtypelos',  'addretinadiag___1',  'addretinadiag___2',  'addretinadiag___3',  'addretinadiag___4',  'addretinadiag___5',  'addretinadiag___6',  'addretinadiag___88',  'addretinadiag___99',  'eprenlmemeye',  'hyprtnsvretineye',  'hollnhrsplaqeye',  'vasclroccleye',  'chrodlnevseye',  'chrodlnevsrod',  'chrodlnevslos',  'melanomaeye',  'retinatx',  'retinatxeye',  'nerveabnrm',  'nervereslts___1',  'nervereslts___2',  'nervereslts___3',  'nervereslts___4',  'nervereslts___88',  'optcswelleye',  'optcpalloreye',  'optcdrusneye',  'optcapdeye',  'nerveabnrmtx',  'nerveabnrmtxeye',  'fuophthalmgst',  'fucomopthal',  'fucomopthalhr',  'recmndcom', 'spcserviceyn',  'spcservice___1',  'spcservice___2',  'spcservice___3',  'spcservice___4',  'spcservice___5',  'spcservice___6',  'spcservice___7',  'spcservice___8',  'spcservice___88',  'spcservicetx',  'extnlpstptm']

## FORMATTING FUNCTIONS
def get_ODOSChecklistResults(file, count):
  temp_list = []
  for i in ['Super','Infer','Nasal','Tempo']: #Superior, Inferior, Nasal, Temporal -- first letter of each
    try:
      variable = next(file) #grabs next line
      variable = variable.replace(u'\xa0',u'')
      if(variable[2:7] == i or variable[10:15] == i): #participant YP0396 will make you 
        #understand all this goddamn error-checking
        # and also the reason why you're also checking variable[3:8] is because sometimes 
        # there's a \xc2 in front of the bulletpoint
        variable = variable.split(strings[count])
        variable = variable[len(variable)-1]
        variable = variable.replace(u'\n', u'').replace(u'\xa0',u'')
        variable = variable.split(':')[1]
        variable = variable.strip()
        temp_list.append(variable)
      else:
        temp_list.append("MANUAL: "+variable)
    except IndexError:
      temp_list.append("MANUAL: MISSING DATA")
  return temp_list

def format_time(timeToParse):
  for format in ('%I:%M %p', '%I:%M%p', '%H:%M', '%h:%M'):
    try:
      time = datetime.strptime(timeToParse, format)
    except ValueError:
      continue
    if (time.hour < 7):
      time = time + timedelta(hours=12)
    return datetime.strftime(time, "%H:%M")
  return ("MANUAL: "+timeToParse)

def format_extnlpname_and_tx(rawName):
  if 'Bicket, Amanda K, MD' in rawName:
    name = 1
    nametx = ''
  elif 'Elam, Angela Renee, MD' in rawName:
    name = 2
    nametx = ''
  elif 'John, Denise Ann-Marie, MD' in rawName:
    name = 3
    nametx = ''
  elif 'Newman-Casey, Paula Anne, MD' in rawName:
    name = 4
    nametx = ''
  elif 'Woodward, Maria Anneke, MD' in rawName:
    name = 5
    nametx = ''
  elif 'Zhang, Jason, MD' in rawName:
    name = 6
    nametx = ''
  else:
    name = 'MANUAL'
    nametx = 'MANUAL'
  
  return [name, nametx]

def formatOCTColor(color):
  color = color.split(': ')
  color = color[len(color)-1].strip()
  if (color == 'green' or color == 'Green'):
    return '1'
  elif (color == 'yellow' or color == 'Yellow'):
    return '2'
  elif (color == 'red' or color == 'Red'):
    return '3'
  elif (color == 'white' or color == 'White'):
    return '4'
  else:
    return ('MANUAL: NOT FOUND: ' + str(color))

def formatCTDRatio(ctdr):
  try:
    formatted_ctdr = "{:.2f}".format(float(ctdr))
  except ValueError:
    return ("MANUAL: "+str(ctdr))
  return formatted_ctdr

def format_refctverror(error):
  values = []

  if (error == "no" or error == "No"): #no error present
    values.append('0') #refctverror
    for i in range(26):
      values.append('')
  else: #error is present
    error = removeprefix(error,'Yes; ')
    errorOD = 'NONE'
    errorOS = 'NONE'

    values.append('1') #error is present
    values.append(error) #srdescrip
    if ' and OS; ' in error: #error is in both eyes
      error = error.split(' and OS; ') 
      if error[0].startswith('OD'): #split removes ' and OS; ' so we have to remove 'OD; '
        errorOD = removeprefix(error[0], 'OD; ')
        errorOS = error[1]
      elif error[0].startswith('OS'):
        errorOS = error[0]
        errorOD = removeprefix(error[1], 'OD; ')
    else: #error is in just one eye
      if error.startswith('OD'):
        errorOD = removeprefix(error, 'OD; ') #remove 'OD; '
        errorOS = 'NONE'
      elif error.startswith('OS'):
        errorOS = removeprefix(error, 'OS; ') #remove 'OS; '
        errorOD = 'NONE'
    if (errorOD != 'NONE' and errorOS == 'NONE'): #if error is in just right eye
      values.append('1')
      errorOD = removeprefix(errorOD, 'Refractive error type: ')
      values.extend(processRefctError(errorOD))
      values.extend(['','','','','','','','','','','','']) #all left eye stuff is unneeded since left eye has no refractive errors
    elif (errorOS != 'NONE' and errorOD == 'NONE'): #if error is in just left eye
      values.append('2')
      errorOS = removeprefix(errorOS, 'Refractive error type: ')
      values.extend(['','','','','','','','','','','','']) #all right eye stuff is uneeded since right eye has no refractive errors
      values.extend(processRefctError(errorOS))
    elif (errorOS != 'NONE' and errorOD != 'NONE'):
      values.append('3')
      errorOD = removeprefix(errorOD, 'Refractive error type: ')
      values.extend(processRefctError(errorOD))
      errorOS = removeprefix(errorOS, 'Refractive error type: ')
      values.extend(processRefctError(errorOS))

  return values

def processRefctError(error):
  values = []
  if '; ' in error: #this indicates follow-up was recommended
    error = error.split('; ')
    errorType = error[0]
    errorFU = error[1]
    values.extend(processRefctErrorType(errorType))
    values.extend(processRefctErrorFU(errorFU))
  else: #no follow-up was recommended
    values.extend(processRefctErrorType(error))
    values.extend(['0','0','0','1'])
  return values

def processRefctErrorType(error):
  values = []
#  if ' and ' in error:
#    errors = error.split(' and ')
#    if ', ' in errors[0]:
#      errorsFinal = errors[1] #final error, after ' and '
#      errors0 = errors[0].split(', ') #all errors before ' and ', split by ', '
#      errors = errors0.append(errorsFinal) #combine the errors before ' and ' with the errors after ' and '

  if 'Hyperopia' in error:
    values.append('1')
  else:
    values.append('0')
  if 'High hyperopia (>+5.00)' in error:
    values.append('1')
  else:
    values.append('0')
  if 'High myopia (<-5.00)' in error:
    values.append('1')
  else:
    values.append('0')
  if 'Myopia' in error:
    values.append('1')
  else:
    values.append('0')
  if 'Astigmatism' in error:
    values.append('1')
  else:
    values.append('0')
  if 'High astigmatism (>+3.00)' in error:
    values.append('1')
  else:
    values.append('0')
  if 'Presbyopia' in error:
    values.append('1')
  else:
    values.append('0')
  if not ('Hyperopia' in error or 'High hyperopia (>+5.00)' in error 
      or 'High myopia (<-5.00)' in error or 'Myopia' in error or
      'Astigmatism' in error or 'High astigmatism (>+3.00)' in error or 
      'Presbyopia' in error):
    values.append('1')
  else:
    values.append('0')
  return values

def processRefctErrorFU(errorFU):
  values = []
  if 'Refer for in-person evaluation in 1-6 months for evaluation of peripheral retina' == errorFU:
    values = ['0','1','0','0']
  else:
    values = ['MANUAL: ' + errorFU, 'MANUAL: SEE LEFT', 'MANUAL: SEE LEFT', 'MANUAL: SEE LEFT']
  return values

def format_extnlpAbnormalities(abnrml):
  values = []

  for i in range(55):
    values.append('')

  if (abnrml == "yes" or abnrml == "Yes"): #extnlpnorm thru extnabnrmothtx
    values[0] = '1'
  else:
    if abnrml == 'No; Lens; Cataract; OD; cataract is not visually significant':
      values[0]='0' #extnlpnorm
      values[1]='1' #extnlpabnrmloc__1
      values[8]='1' #lensabnrm___2 (cataract)
      values[12]='1' #cataracteye
      values[13]='1' #cataract
    elif abnrml == 'No; Lens; Cataract; OD; cataract is visually significant and needs referral for further evaluation':
      values[0]='0'
      values[1]='1'
      values[8]='1'
      values[12]='1'
      values[13]='2'
    elif abnrml == 'No; Lens; Cataract; OS; cataract is not visually significant':
      values[0]='0'
      values[1]='1'
      values[8]='1'
      values[12]='2'
      values[13]='1'
    elif abnrml == 'No; Lens; Cataract; OS; cataract is visually significant and needs referral for further evaluation':
      values[0]='0'
      values[1]='1'
      values[8]='1'
      values[12]='2'
      values[13]='2'
    elif abnrml == 'No; Lens; Cataract; OU; cataract is not visually significant':
      values[0]='0'
      values[1]='1'
      values[8]='1'
      values[12]='3'
      values[13]='1'
    elif abnrml == 'No; Lens; Cataract; OU; cataract is visually significant and needs referral for further evaluation':
      values[0]='0'
      values[1]='1'
      values[8]='1'
      values[12]='3'
      values[13]='2'
    elif abnrml == 'No; Lens; Intraocular lens; OU':
      values[0]='0'
      values[1]='1'
      values[7]='1' #lensabnrm___1 (intraocular)
      values[10]='3' #intraocirlneye
    elif abnrml == "No; Lens; Intraocular lens; OD":
      values[0]='0'
      values[1]='1'
      values[7]='1' #lensabnrm___1 (intraocular)
      values[10]='1' #intraocirlneye
    elif abnrml == "No; Lens; Intraocular lens; OS":
      values[0]='0'
      values[1]='1'
      values[7]='1' #lensabnrm___1 (intraocular)
      values[10]='2' #intraocirlneye
    else:
      values[0]='MANUAL: '+ abnrml
      for i in range(54): #extnlpabnrmloc thru extnabnrmothtx needs to be marked with MANUAL
        values[i+1] = 'MANUAL: see left'
  
  return values

def format_glaucoma(glauc): 
  values = []
  for i in range(24):
      values.append('')
  if (glauc == 'no' or glauc == 'No'):
    values[0] = '0' #glaucoma absent
  else:
    values[0] = '1' #glaucoma present
    values[3] = glauc #srdescrip2
    glaucOS = 'NONE'
    glaucOD = 'NONE'
    glauc = removeprefix(glauc, 'Yes; ')
    if ' and OS; ' in glauc: #glaucoma is in both eyes
      glauc = glauc.split(' and OS; ') 
      if glauc[0].startswith('OD'): #split removes ' and OS; ' so we have to remove 'OD; '
        glaucOD = removeprefix(glauc[0], 'OD; ')
        glaucOS = glauc[1]
      elif glauc[0].startswith('OS'):
        glaucOS = glauc[0]
        glaucOD = removeprefix(glauc[1], 'OD; ')
    else: #error is in just one eye
      if glauc.startswith('OD'):
        glaucOD = removeprefix(glauc, 'OD; ') #remove 'OD; '
        glaucOS = 'NONE'
      elif glauc.startswith('OS'):
        glaucOS = removeprefix(glauc, 'OS; ') #remove 'OS; '
        glaucOD = 'NONE'
    if (glaucOD != 'NONE' and glaucOS == 'NONE'): #if glaucoma is in just right eye
      values[2] = 99
      if '; Suspect' in glaucOD:
        values[1] = 1
        glaucOD = removesuffix(glaucOD,'; Suspect')
      elif '; Definite' in glaucOD:
        values[1] = 2
        glaucOD = removesuffix(glaucOD,'; Definite')
      else:
        values[1] = 'MANUAL: UNKNOWN VALUE'
      glaucOD = glaucOD.rstrip()
      values = glaucSigns(glaucOD, values, 'OD')
    elif (glaucOS != 'NONE' and glaucOD == 'NONE'): #if glaucoma is in just left eye
      values[1] = 99
      if '; Suspect' in glaucOS:
        values[2] = 1
        glaucOS = removesuffix(glaucOS,'; Suspect')
      elif '; Definite' in glaucOS:
        values[2] = 2
        glaucOS = removesuffix(glaucOS,'; Definite')
      else:
        values[2] = 'MANUAL: UNKNOWN VALUE'
      glaucOS = glaucOS.rstrip()
      values = glaucSigns(glaucOS, values, 'OS')
    elif (glaucOS != 'NONE' and glaucOD != 'NONE'): #if glaucoma is in both eyes
      if '; Suspect' in glaucOD:
        values[1] = 1
        glaucOD = removesuffix(glaucOD,'; Suspect')
      elif '; Definite' in glaucOD:
        values[1] = 2
        glaucOD = removesuffix(glaucOD,'; Definite')
      else:
        values[1] = 'MANUAL: UNKNOWN VALUE'
      if '; Suspect' in glaucOS:
        values[2] = 1
        glaucOS = removesuffix(glaucOS,'; Suspect')
      elif '; Definite' in glaucOS:
        values[2] = 2
        glaucOS = removesuffix(glaucOS,'; Definite')
      else:
        values[2] = 'MANUAL: UNKNOWN VALUE'
      glaucOD = glaucOD.rstrip()
      glaucOS = glaucOS.rstrip()
      values = glaucSigns(glaucOD, values, 'OD')
      values = glaucSigns(glaucOS, values, 'OS')
  
  return values

def glaucSigns(glauc, values, eye):
  offset = 'NONE'
  if eye == 'OS':
    offset = 10
  elif eye == 'OD':
    offset = 0
  
  glaucs = glauc

#  if ' and ' in glauc:
#    glaucs = glauc.split(' and ')
#    if ', ' in glaucs[0]:
#      glaucsFinal = glaucs[1] #final glaucoma sign, after ' and '
#      glaucs0 = glaucs[0].split(', ') #all glaucoma signs before ' and ', split by ', '
#      glaucs = glaucs0.append(glaucsFinal) #combine the glauc signs before ' and ' with the glauc signs after ' and '

  if 'Narrow angle on penlight exam' in glauc:
    values[offset+4] = 1
  else:
    values[offset+4] = 0
  if 'Patient previously treated for glaucoma (e.g. already taking glaucoma medications or previous glaucoma surgery)' in glauc:
    values[offset+5] = 1
  else:
    values[offset+5] = 0
  if 'Cup to disc ratio greater than or equal to 0.7' in glauc:
    values[offset+6] = 1
  else:
    values[offset+6] = 0
  if 'Intraocular pressure greater than or equal to 21 mmHg' in glauc:
    values[offset+7] = 1
  else:
    values[offset+7] = 0
  if 'Disc hemorrhage' in glauc:
    values[offset+8] = 1
  else:
    values[offset+8] = 0
  if 'Asymmetry of the cup-to-disc by ?0.2 where the larger cup is ?0.6' in glauc:
    values[offset+9] = 1
  else:
    values[offset+9] = 0
  if 'Rim thinning or focal notch' in glauc:
    values[offset+10] = 1
  else:
    values[offset+10] = 0
  if 'Abnormal OCT (overall RNFL thickness <80 or thinning at <1% certainty (RED) in the inferior or superior quadrants)' in glauc:
    values[offset+11] = 1
  else:
    values[offset+11] = 0
  if not ('Narrow angle on penlight exam' in glauc or 
          'Patient previously treated for glaucoma (e.g. already taking glaucoma medications or previous glaucoma surgery)' in glauc or
          'Cup to disc ratio greater than or equal to 0.7' in glauc or
          'Intraocular pressure greater than or equal to 21 mmHg' in glauc or
          'Disc hemorrhage' in glauc or
          'Asymmetry of the cup-to-disc by ?0.2 where the larger cup is ?0.6' in glauc or
          'Rim thinning or focal notch' in glauc or
          'Abnormal OCT (overall RNFL thickness <80 or thinning at <1% certainty (RED) in the inferior or superior quadrants)' in glauc):
    values[offset+12] = 1
    values[offset+13] = glauc
  else:
    values[offset+12] = 0
    values[offset+13] = ''

  return values

def format_fundusAbnormalities(fundAbn):
  values = []
  for i in range(41):
    values.append('')
  
  if (fundAbn == "none" or fundAbn == "None"): #fundusabnrm == 0
    values[0] = '0' #fundusabnrm
  elif ("none" in fundAbn or "None" in fundAbn): #this case is if fundAbn says "None" but also says other stuff
    values[0] = 'MANUAL: ' + fundAbn
    for i in range(1,30):
      values[i] = 'MANUAL: see fundAbn'
  else: #fundusabnrm == 1
    values[0] = 'MANUAL: ' + fundAbn #in case nothing works, we'll still know what the original text said 
    values[1] = '0' #fundusreslts___1
    values[2] = '0' #fundusreslts___2
    values[3] = '0' #fundusreslts___88
    values[8] = '0' #amd
    values[11] = '0' #addretinadiag___1
    values[12] = '0' #addretinadiag___2
    values[13] = '0' #addretinadiag___3
    values[14] = '0' #addretinadiag___4
    values[15] = '0' #addretinadiag___5
    values[16] = '0' #addretinadiag___6
    values[17] = '0' #addretinadiag___88
    values[18] = '0' #addretinadiag___99
    values[29] = '0' #nerveabnrm
    if 'Diabetic retinopathy' in fundAbn and not 'Diabetic retinopathy; Grade OD No diabetic retinopathy; Diabetic macular edema OD no; Grade OS No diabetic retinopathy; Diabetic macular edema OS no' in fundAbn:
      values[0] = '1'
      values[1] = '1'
      if 'Grade OD No diabetic retinopathy' in fundAbn and not 'Diabetic retinopathy; Grade OD No diabetic retinopathy; Diabetic macular edema OD no; Grade OS No diabetic retinopathy; Diabetic macular edema OS no' in fundAbn:
        values[4] = '0' #dretingraderod
      elif 'Grade OD Mild NPDR' in fundAbn:
        values[4] = '1'
      elif 'Grade OD Moderate NPDR' in fundAbn:
        values[4] = '2'
      elif 'Grade OD Severe NPDR' in fundAbn:
        values[4] = '3'
      elif 'Grade OD Proliferative diabetic retinopathy' in fundAbn:
        values[4] = '4'
      else:
        values[4] = 'MANUAL: ERROR: ' + fundAbn

      if 'Diabetic macular edema OD no' in fundAbn and not 'Diabetic retinopathy; Grade OD No diabetic retinopathy; Diabetic macular edema OD no; Grade OS No diabetic retinopathy; Diabetic macular edema OS no' in fundAbn:
        values[0] = '1'
        values[5] = '0' #rodmacedema
      elif 'Diabetic macular edema OD yes' in fundAbn:
        values[0] = '1'
        values[5] = '1'
      else:
        values[5] = 'MANUAL: ERROR: ' + fundAbn
      
      if 'Grade OS No diabetic retinopathy' in fundAbn and not 'Diabetic retinopathy; Grade OD No diabetic retinopathy; Diabetic macular edema OD no; Grade OS No diabetic retinopathy; Diabetic macular edema OS no' in fundAbn:
        values[0] = '1'
        values[6] = '0' #dretingradelos
      elif 'Grade OS Mild NPDR' in fundAbn:
        values[0] = '1'
        values[6] = '1'
      elif 'Grade OS Moderate NPDR' in fundAbn:
        values[0] = '1'
        values[6] = '2'
      elif 'Grade OS Severe NPDR' in fundAbn:
        values[0] = '1'
        values[6] = '3'
      elif 'Grade OS Proliferative diabetic retinopathy' in fundAbn:
        values[0] = '1'
        values[6] = '4'
      else:
        values[6] = 'MANUAL: ERROR: ' + fundAbn

      if 'Diabetic macular edema OS no' in fundAbn and not 'Diabetic retinopathy; Grade OD No diabetic retinopathy; Diabetic macular edema OD no; Grade OS No diabetic retinopathy; Diabetic macular edema OS no' in fundAbn:
        values[0] = '1'
        values[7] = '0' #losmacedema
      elif 'Diabetic macular edema OS yes' in fundAbn:
        values[0] = '1'
        values[7] = '1'
      else:
        values[7] = 'MANUAL: ERROR: ' + fundAbn

      if ('Diabetic retinopathy; Grade OD No diabetic retinopathy; '+
      'Diabetic macular edema OD no; Grade OS No diabetic retinopathy; '+
      'Diabetic macular edema OS no') in fundAbn: #SPECIAL CASE: this is how the docs note that pt has diabetes but no diabetic retinopathy
      # In this case the pt DOES NOT ACTUALLY HAVE diabetic retinopathy
        values[0] = '0' #fundusabnrm
        values[1] = '' #fundusreslts__1
        values[5] = ''
        values[7] = ''
    else:
      values[1] = '0' #fundusreslts__1
      
    if 'Macular degeneration' in fundAbn:
      values[0] = '1'
      values[2] = '1' #fundusreslts___2
    else:
      values[2] = '0'

    if 'Other retina' in fundAbn:
      values[0] = '1'
      values[3] = 'MANUAL: ' + fundAbn #fundusreslts___88
      values[11] = 'MANUAL: see fundusreslts___88' #addretinadiag___1
      values[12] = 'MANUAL: see fundusreslts___88' #addretinadiag___2
      values[13] = 'MANUAL: see fundusreslts___88' #addretinadiag___3
      values[14] = 'MANUAL: see fundusreslts___88' #addretinadiag___4
      values[15] = 'MANUAL: see fundusreslts___88' #addretinadiag___5
      values[16] = 'MANUAL: see fundusreslts___88' #addretinadiag___6
      values[17] = 'MANUAL: see fundusreslts___88' #addretinadiag___88
      values[18] = 'MANUAL: see fundusreslts___88' #addretinadiag___99
      values[19] = 'MANUAL: see addretinadiag__1' #eprenlmemeye
      values[20] = 'MANUAL: see addretinadiag__2' #hyprtnsvretineye
      values[21] = 'MANUAL: see addretinadiag__3' #hollnhrsplaqeye
      values[22] = 'MANUAL: see addretinadiag__4' #vasclroccleye
      values[23] = 'MANUAL: see addretinadiag__5' #chrodlnevseye
      values[24] = 'MANUAL: see chrodlnevseye' #chrodlnevsrod
      values[25] = 'MANUAL: see chrodlnevseye' #chrodlnevslos
      values[26] = 'MANUAL: see addretinadiag__6' #melanomaeye
      values[27] = 'MANUAL: see addretinadiag__88' #retinatx
      values[28] = 'MANUAL: see addretinadiag__88' #retinatxeye
    else:
      values[3] = '0'
      values[18] = '1' #addretinadiag_99 (None)

    if not ('Diabetic retinopathy' in fundAbn or 'Macular degeneration' not in fundAbn or 'Other retina' not in fundAbn): #yet, fundusabnrm == 1 still
      #this would be an unexpected case so it warrants manual intervention
      for i in range(28):
        values[i+1] = 'MANUAL: see fundusabnrm'

    if 'AMD type OD Early dry AMD' in fundAbn: #for some reason, this isn't something
      # that only shows up in redcap when fundusreslts == 2. This question's 
      #always available in redcap if fundusabnrm=1, so i've treated it as such here
      values[0] = '1'
      values[8] = '1' #amd
      values[9] = '1'
    elif 'AMD type OD Moderate dry AMD' in fundAbn:
      values[0] = '1'
      values[8] = '1'
      values[9] = '2'
    elif 'AMD type OD Severe dry AMD' in fundAbn:
      values[0] = '1'
      values[8] = '1'
      values[9] = '3'
    elif 'AMD type OD Wet AMD' in fundAbn:
      values[0] = '1'
      values[8] = '1'
      values[9] = '4'
    else:
      values[0] = '1'
      values[8] = '0' #amd
      values[9] = '' #amdtyperod

    if 'AMD type OS Early dry AMD' in fundAbn:
      values[0] = '1'
      values[8] = '1'
      values[10] = '1'
    elif 'AMD type OS Moderate dry AMD' in fundAbn:
      values[0] = '1'
      values[8] = '1'
      values[10] = '2'
    elif 'AMD type OS Severe dry AMD' in fundAbn:
      values[0] = '1'
      values[8] = '1'
      values[10] = '3'
    elif 'AMD type OS Wet AMD' in fundAbn:
      values[0] = '1'
      values[8] = '1'
      values[10] = '4' #amdtypelos
    else:
      values[0] = '1'
      values[8] = '0' #amd
      values[10] = '' #amdtypelos
    
    if 'Nerve' in fundAbn:
      values[3] = 1 #fundusreslts__88
      values[29] = 'MANUAL: ' + fundAbn #nerveabnrm
    else:
      values[29] = '0'
      
  return values

def format_fuophthalmgst(fu, AP):
  values = []
  for i in range(15):
    values.append('')

  values[3] = AP

  if fu == 'No':
    values[0] = '0' #fuopththalmgst
  elif fu == 'Yes; Diabetic with mild non-proliferative retinopathy, follow up in 1 year':
    values[0] = '1' #fuopththalmgst
    values[1] = '6' #fucomopthal
    values[4] = '0' #spcserviceyn
  elif fu == 'Yes; Diabetic without retinopathy, follow up in 2 years':
    values[0] = '1' #fuopththalmgsta
    values[1] = '5' #fucomopthal
    values[4] = '0' #spcserviceyn
  elif fu == 'Yes; Higher risk individual 1-3 Months':
    values[0] = '1' #fuophthalmgst
    values[1] = '7' #fucomopthal
    values[2] = '3' #fucomopthalhr
    values[4] = '0' #spcserviceyn
  elif fu == 'Yes; Higher risk individual 10-12 Months':
    values[0] = '1' #fuophthalmgst
    values[1] = '7' #fucomopthal
    values[2] = '6' #fucomopthalhr
    values[4] = '0' #spcserviceyn
  elif fu == 'Yes; Higher risk individual 4-6 Months':
    values[0] = '1'
    values[1] = '7'
    values[2] = '4'
    values[4] = '0'
  elif fu == 'Yes; Higher risk individual 7-9 Months':
    values[0] = '1'
    values[1] = '7'
    values[2] = '5'
    values[4] = '0'
  elif fu == 'Yes; Higher risk individual Within 2 Weeks':
    values[0] = '1'
    values[1] = '7'
    values[2] = '1'
    values[4] = '0'
  elif fu == 'Yes; Low risk individual 40-49 years, follow up in 3-4 years':
    values[0] = '1'
    values[1] = '2'
    values[4] = '0'
  elif fu == 'Yes; Low risk individual 50-64 years, follow up in 2-3 years':
    values[0] = '1'
    values[1] = '3'
    values[4] = '0'
  elif fu == 'Yes; Low risk individual 65 years and older, follow up in 1-2 years':
    values[0] = '1'
    values[1] = '4'
    values[4] = '0'
  elif fu == 'Yes; Low risk individual under 40 years, follow up in 5-10 years':
    values[0] = '1'
    values[1] = '1'
    values[4] = '0'
  else:
    values[0] = 'MANUAL: '+fu #fuophthalmgst needs to be done manually
    for i in range(14): #need to note that fucomopthal thru spcservicetx need to be done manually
      values[i+1] = 'MANUAL: see fuophthalmgst'
    values[3] = AP #recmndcom
  
  if ('specialist' in AP or 'specialty' in AP or 'special' in AP or 'follow' in AP or 'follow-up' in AP or 'evaluation' in 
      AP or 'refer' in AP or 'FU' in AP or 'baseline' in AP or 'testing' in AP or 'repeat' in AP):
      #generally, physicians tend to put the need for specialty services 
      #in the assessment and plan (recmndcom), usually with some phase (a possible phrase 
      #is "recommend follow-up for assessment with a retina specialist")
    for i in range(11): #mark spcserviceyn thru spcservicetx as MANUAL based on recmndcom
      values[i+4] = 'MANUAL: see recmndcom' #values[4] is spcserviceyn

  return values

## PARTICIPANT FUNCTIONS (called column functions, but each participant is actually a row)

def formatParticipantColumn(column):
  formattedColumn = []
  formattedColumn.append(column[0])

  if (column[33] == "MANUAL"):
    formattedColumn.append("MANUAL")
  else:
    formattedColumn.append(datetime.isoformat(column[33])) #extnlpdat is all the way at the end

  formattedColumn.append(format_time(column[1])) #extnlpsrttm

  formattedColumn.extend(format_extnlpname_and_tx(column[32])) #extnlpname and extnlpnametx

  if (column[2] == "No" or column[2] == "no"): #visimpairmnt and visimpairmnteye
    formattedColumn.append('0')
    formattedColumn.append('')
  else:
    eye = column[2][4:]
    formattedColumn.append('1')
    if (eye.strip() == "OD"):
      formattedColumn.append('1')
    elif (eye.strip() == "OS"):
      formattedColumn.append('2')
    elif (eye.strip() == "OU"):
      formattedColumn.append('3')
    else:
      formattedColumn.append('MANUAL: unknown eye')

  formattedColumn.extend(format_refctverror(column[3])) #refctverror thru refractfulos

  if (column[4] == "yes" or column[4] == "Yes"): #refractrecpresc
    formattedColumn.append('1')
  elif (column[4] == "no" or column[4] == "No"):
    formattedColumn.append('0')
  else:
    formattedColumn.append("MANUAL: " + str(column[4]))
  
  if ((column[5] == "yes" or column[5] == "Yes") and (column[6] == "no" or column[6] == "No")): #extnlp and extnlpmiss
    formattedColumn.append('1')
    formattedColumn.append('')
  else:
    formattedColumn.append('0')
    formattedColumn.append('MANUAL: external photos taken = \'' + str(column[5]) + '\' and external photos missing = \'' + str(column[6]) + '\'')

  if (column[7] == "yes" or column[7] == "Yes"): #rodextnlpqlty, rodextnlpqltytx, losextnlpqlty, and losextnlpqltytx
    formattedColumn.append('1')
    formattedColumn.append('')
    formattedColumn.append('1')
    formattedColumn.append('')
  else:
    formattedColumn.append('MANUAL: ' + column[7])
    formattedColumn.append('MANUAL: SEE LEFT')
    formattedColumn.append('MANUAL: SEE LEFT')
    formattedColumn.append('MANUAL: SEE LEFT')

  formattedColumn.extend(format_extnlpAbnormalities(column[8])) #extnlpnorm thru extnabnrmothtx
  
  if ((column[9] == "yes" or column[9] == "Yes") and (column[10] == "no" or column[10] == "No")): #fundusp and funduspmiss
    formattedColumn.append('1')
    formattedColumn.append('')
  else:
    formattedColumn.append('0')
    formattedColumn.append('MANUAL: photos taken = \'' + str(column[9]) + '\' and photos missing = \'' + str(column[10]) + '\'')

  if (column[11] == "yes" or column[11] == "Yes"): #rodfunduspqlty, rodfunduspqltytx, losfunduspqlty, and losfunduspqltytx
    formattedColumn.append('1')
    formattedColumn.append('')
    formattedColumn.append('1')
    formattedColumn.append('')
  else:
    formattedColumn.append('MANUAL: ' + column[11])
    formattedColumn.append('MANUAL: SEE LEFT')
    formattedColumn.append('MANUAL: SEE LEFT')
    formattedColumn.append('MANUAL: SEE LEFT')
  
  formattedColumn.append(formatCTDRatio(column[12])) #ctdrrod
  formattedColumn.append(formatCTDRatio(column[13])) #ctdrlos

  if ((column[14] == "yes" or column[14] == "Yes") and (column[15] == "no" or column[15] == "No")): #octp and octpmiss
    formattedColumn.append('1') #octp
    formattedColumn.append('') #octpmiss
  else:
    formattedColumn.append('MANUAL: photos taken = \'' + str(column[14]) + '\' and photos missing = \'' + str(column[15]) + '\'') #octp
    formattedColumn.append('MANUAL: photos taken = \'' + str(column[14]) + '\' and photos missing = \'' + str(column[15]) + '\'') #octpmiss

  if (column[16] == "yes" or column[16] == "Yes"): #rodoctpqlty, rodoctpqltytx, losoctpqlty, and losoctpqltytx
    formattedColumn.append('1')
    formattedColumn.append('')
    formattedColumn.append('1')
    formattedColumn.append('')
  else:
    formattedColumn.append('MANUAL: ' + column[16])
    formattedColumn.append('MANUAL: SEE LEFT')
    formattedColumn.append('MANUAL: SEE LEFT')
    formattedColumn.append('MANUAL: SEE LEFT')
  
  retinarodthickness = column[17]
  retinarodthickness = removesuffix(removesuffix(retinarodthickness, 'm'),' u')
  formattedColumn.append(retinarodthickness) #rodthickns

  formattedColumn.append(formatOCTColor(column[18])) #rodsupr
  formattedColumn.append(formatOCTColor(column[19])) #rodinferior
  formattedColumn.append(formatOCTColor(column[20])) #rodnasal
  formattedColumn.append(formatOCTColor(column[21])) #rodtemprl

  retinalosthickness = column[22]
  retinalosthickness = removesuffix(removesuffix(retinalosthickness, 'm'),' u')
  formattedColumn.append(retinalosthickness) #losthickns

  formattedColumn.append(formatOCTColor(column[23])) #lossupr
  formattedColumn.append(formatOCTColor(column[24])) #losinferior
  formattedColumn.append(formatOCTColor(column[25])) #losnasal
  formattedColumn.append(formatOCTColor(column[26])) #lostemprl


  formattedColumn.extend(format_glaucoma(column[27])) #glaucoma thru glaucsignlostx

  formattedColumn.extend(format_fundusAbnormalities(column[28])) #fundusabnrm thru nerveabnrmtxeye
  
  formattedColumn.extend(format_fuophthalmgst(column[29], column[31])) #column[29] is fuophtalmgst thru fucomopthalhr; column[31] is recmndcom;
  #together these yield spcserviceyn thru spcservicetx
  
  formattedColumn.append(format_time(column[30])) #extnlpstptm

  return formattedColumn

def makeParticipantColumn(part):
  count = 0
  variable_list=[]
  variable_list.append(part[0:6]) #first variable in list is participant ID

  docName = 'MANUAL'
  AP = ''
  date = 'MANUAL'

  while (count < len(strings)): #splits questions off from answers
    # Loop through the file line by line
    afterAP = 0 #if line is after "Assessment and Plan", then this becomes 1
    afterStop = 0
    string = part
    file = open(string, encoding = "ISO-8859-1")
    for line in file:
      line = line.replace(u'\xa0', u' ')
      #line = line.replace('\xc2', '') #I think this doesn't work anyway to remove 0xc2 (nbsp in UTF, Â in ISO-8859-1) character
      # checking string is present in line or not
      if strings[count] == 'Assessment and Plan' and 'Assessment and Plan' in line:
        afterAP = 1
      if afterAP == 0:
        if (strings[count] == 'Progress Notes') and (strings[count] in line): #to grab doc name
          docName = next(file)
        elif (strings[count] == 'OD checklist') and (strings[count] in line): #need to grab four lines after for OD OS checklist results
          variable_list.extend(get_ODOSChecklistResults(file, count))
        elif (strings[count] == 'OS checklist') and (strings[count] in line): #need to grab four lines after for OD OS checklist results
          variable_list.extend(get_ODOSChecklistResults(file, count))
        elif (strings[count] == 'Letter sent to') and (strings[count] in line): #need to grab next line for date letter was sent
          variable = next(file)
          try:
            date = parser.parse(variable, fuzzy=True, dayfirst=False, yearfirst=False)
          except parser.ParserError:
            date = 'MANUAL'
        elif strings[count] in line:
          variable = line
          # separate questions from answers
          variable = variable.split(strings[count])
          variable = variable[len(variable)-1]
          variable = variable.replace(u'\n', u'')
          variable = variable.strip()
          if (variable == ''):
            variable_list.append("MANUAL: MISSING DATA")
          else:
            variable_list.append(variable)
      else: #afterAP == 1
        if (strings[count] == 'Stop time' and strings[count] in line): #just stop time needs to go into different variable, not into AP
          variable = line
          # separate questions from answers
          variable = variable.split(strings[count])
          variable = variable[len(variable)-1]
          variable = variable.replace(u'\n', u'')
          variable = variable.strip()
          if (variable == ''):
            variable_list.append("MANUAL: MISSING DATA")
          else:
            variable_list.append(variable)
        else: #not stop time, after "Assessment and Plan"
          if 'Stop time ' in line:
            afterStop = 1
          if afterStop == 0: #before "Stop time "
            AP = AP + line + '\n'
    count = count + 1

  AP = AP.replace('\n\n','\n')

  variable_list.append(AP) #recmndcom, in column[31]
  variable_list.append(docName) #in column[32]
  variable_list.append(date) #examdat, in column[33]

  return variable_list #final space is AP = Assessment&Plan, but it should be moved

data = []

for file in file_name_list:
  try:
    data.append(formatParticipantColumn(makeParticipantColumn(file)))
  except IndexError: #This index error is because fields are missing – go fill them in with BLANKS. Usually OCT results are missing because OCT photos are missing.
    print('INDEX ERROR for: ' + file[0:6])

df = pd.DataFrame(data, columns = cols)

df.to_excel('aScreeningResults.xlsx', index=False)
