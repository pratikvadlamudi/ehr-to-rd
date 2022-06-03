# What is this?
ehr-to-rd is an algorithm of two files written in Python 3, `technician.py` and `screening.py`. If you paste ophthalmology notes from MiChart, for the MI-SIGHT program, into `.txt` files, then you can use this algorithm to extract the data and output as an Excel file that needs some manual editing before being converted to a `.csv` and uploaded to REDCap. 

# How do I use this?
If you do not already have Python 3 installed on your device, you will need to use that first. Additionally, you will need `.txt` files, one for each participant's full note, copy-pasted from MiChart. Additionally, you will need some time, as much manual editing still needs to be done even after the output `.xlsx` files are produced.

Read the following instructions, then place the `.txt` files into a directory named `files`, then run these two `.py` files from the directory that `files` is in.

### Instructions for Technician Eye Exam

*For person who produces text files:*

Ensure that newlines are in the correct places after the following questions. Generally, these newlines are present after the first time copy/pasting from MiChart. Each of the following questions should have two newlines between the question and the answer:
* Presenting visual acuity (Right/OD) DISTANCE
* Presenting visual acuity (Left/OS) DISTANCE
* Post-refraction Best corrected visual acuity (Right/OD) DISTANCE
* Post-refraction Best corrected visual acuity (Right/OD) NEAR
* Post-refraction Best corrected visual acuity (Left/OS) DISTANCE:
* Post-refraction Best corrected visual acuity (Left/OS) NEAR:
* Contrast sensitivity (Right/OD):
* Contrast sensitivity (Left/OS):
* Autorefraction (Right/OD) Sphere:
* Autorefraction (Right/OD) Cylinder:
* Autorefraction (Right/OD) Axis:
* Autorefraction (Left/OS) Sphere:
* Autorefraction (Left/OS) Cylinder:
* Autorefraction (Left/OS) Axis:
* Manifest refraction (Right/OD) Sphere:
* Manifest refraction (Right/OD) Cylinder:
* Manifest refraction (Right/OD) Axis:
* Manifest refraction (Right/OD) Add
* Manifest refraction (Left/OS) Sphere:
* Manifest refraction (Left/OS) Cylinder:
* Manifest refraction (Left/OS) Axis:
* Manifest refraction (Left/OS) Add:
* Afferent pupillary defect (Right/OD):
* Afferent pupillary defect (Left/OS):
* Pachymetry (Right/OD) NOT DONE: Pachymetry (Right/OD):
* Pachymetry (Left/OS) NOT DONE: Pachymetry (Left/OS): 
* Which eye did you dilate 
* Inter-pupillary distance:

---

*For person who runs the code:*
Make sure the variable `newlines` below is set to the right variable

---

*For person who produces output file:*

Produce the following new columns:

* Produce a column named `redcap_event_name`
  * The value is `screening_arm_1` for first-time visits or `m12m24_arm_1` for follow-up visits

* Produce a column named `technician_eye_exam_complete`
  * The values in this column should all be `1` which corresponds to a form status of `Unverified` in REDCap

* Fill in the column titled `thrdipmreq` based on whether `thrdipmrod` or `thrdipmlos` has a value

Visually check the following columns:
* All columns for “unable” – ensure this is coded correctly per column
* `bcvasdislos`. If this is in text rather than number, put an 88 here, and in the next column (bcvasdislostx), put that text
* `bcvasdisrod`. Same as above
* `bcvasnearlos`. Same as above
* `bcvasnearrod`. Same as above
* `pvasdislos`. Same as above
* `pvasdisros`. Same as above
* `cslos`. Same as above
* `csrod`. Same as above
* `eyedrops`. If the patient is taking eyedrops, it will say which one here – put this information in `eyedropstx`. And in `eyedrops`, put a `1`
* `examstrttime`. Look for missing times. Go to the original text file and grab that information
* `phmerylos_hand`. If there is text here, code it correctly and make sure `phmerylos__87` is coded correctly.
* `phmeryrod_hand`. Same as above
* `thrdipmreq`. If there is a number in `thrdipmrod` and `thrdipmlos`, then `thrdipmreq` should be `1`. Else, it should be `0`
* `trialframe`. If there is a `0` here, go to the original text file, figure out why the patient did not receive trial frames, and then put that in `trialframetx`\n"

---
---
---

### Instructions for Screening Results

*For person who produces text files:*
Just copy the whole note, including the physician's name at the top, and the section at the bottom that says "Letter sent to" at the bottom. Just do control-A (Windows) or command-A (Mac) from the MiChart entry, and then copy and paste it straight into a plaintext editor (NOT Microsoft Word) such as Notepad++ or Sublime Text 3. Then save the file, with the filename being the participant ID. Thank you!

---

*For person who produces output file:*

Manually enter the following columns for each participant:

* Produce a column named `redcap_event_name`
  * The value is `screening_arm_1` for first-time visits or `m12m24_arm_1` for follow-up visits
* Produce a column named `screening_results_complete`
  * The values in this column should all be `1` which corresponds to a form status of `Unverified` in REDCap
* In the columns for `refractfurod`, only keep the data in a cell if any of `refractperyod__2` OR `__3` OR `__6` is equal to `1`. Same for these variables but `*los`

For any participants who have \"Index Errors\" at the bottom of this code file:

* Go to their text files and edit them
* The issue is that some lines are missing. Typically, this is because the OCT photos were missing, so the provider removed the relevant lines. Add these lines back after the question `Were OCT images of sufficient quality for interpretation?`:
 * 
    `OCT Results`
    `OD: overall thickness`
    `OD checklist`
    `·`
    `OS: overall thickness`
    `OS checklist`
    `·`

Visually check the following columns for 'MANUAL': 

* `extnlpdat`
  * Here, you have to remove "T00:00:00" from every entry. Use the Excel SUBSTITUTE function
* `extnlpsrttm`
* `refractrecpresc`
* `extnlpnorm` thru `extnabnrmothtx`
* `rodoctpqlty`, `rodoctpqltytx`
* `losoctpqlty`, `losoctpqltytx`
* `rodsupr`, `rodinferior`, `rodnasal`, `rodtemprl`
* `lossupr`, `losinferior`, `losnasal`, `lostemprl`
* `fundusreslts__88`; `addretinadiag__1` thru `retinatxeye`
* `spcserviceyn` thru `spcservicetx`
* `extnlpstptm`
* `ctdrod` and `ctdlos` for **blank cells** (not 'MANUAL')

Apply this logic in Excel to edit fields of csv after exporting:
* If `fundusreslts == “”`, `amd == “” OR “0”`, `addretinadiag == “99”`, and `nerveabnrm == “” OR “0”`, then `fundusabnrm = “0”` (see YP0680)
* If there is a diagnosis in `addretinadiag`, then `fundusreslts__88 = “1”`
* If `refractyperod__2 == 1` OR `refractyperod__3 == 1` OR `refractyperod__6 == 1`, then `refractfurod__1`,`__2`,`__3`,... should have values. Else, `refractfurod__* == “”`. Same for `refractfulos__*`
* If `thrdipmrod` and `thrdipmlos` are present, then `thrdipmreq = “1”`"
