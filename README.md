# peereval

**Script for tabulating peer evaluation responses for group work**

The concept and template detailed below are from the [Welcome to My Classroom workshop](https://carleton.ca/tls/2019/welcome-to-my-classroom-using-peer-to-peer-approaches-to-solve-common-assessment-challenges/) by Audrey Girouard (School of Information Technology) and Morgan Rooney (English), Carleton University.  Code was written in Python by Derek Mueller (Geography), Carleton University. 


**The code works as follows:** 
1. students complete a survey form (Excel xlsx file and and it in)
2. script reads from a directory of completed Excel forms 
3. data are concatentated and validated by group and respondent
3. a "Peer Evaluation Multiplier" (PEM) is computed
4. all the records are output to csv files


## Peer evaluation: 
*Sample explanation (to insert into course syllabus for example)*

All group project grades will be adjusted by a multiplier according to a peer evaluation performed at the end of the class. The Peer Evaluation Mulitplier (PEM) is calculated by dividing the average peer evaluation score for each person by the average peer evaluation score of the group. It is then limited within the 0.75-1.25 range before being applied to group marks. For example:

PEM Calculation Example: For member X in group 1
Average peer evaluation score for member X = 32
Average peer evaluation score for entire group = 30
PEM = 32/30 or 1.07
Group Project grade (weighted mean of 6 group grades as above) = 80%
Member X Mark for Group Project = 80 * 1.07 = 85.6%

The instructor reserves the right to adjust the ratio from the peer evaluation in exceptional circumstances. 


## Instructions / Tutorial: 
1. Modify the template to suit your course (See detailled instructions below)
2. Distribute the form to students 
3. Collect all the completed forms in a directory (they can be in sub-folders - e.g., after unzipping an assignment bulk download)

Note step 4-6 assumes you have some familiarity with Python.  If you don't ask a friend to help with these steps

4. Install Python
  * from here: https://www.python.org/downloads/
  * or through conda:  https://www.anaconda.com/ or https://docs.conda.io/en/latest/miniconda.html
5. Install the [pandas library](https://pandas.pydata.org/), note that you need the optional package 'openpxyl' which should come with pandas
  * option 1, use [pip](https://pypi.org/project/pandas/) to install to your computer.  Type `pip install pandas` at the command line
  * option 2, use [conda](https://anaconda.org/anaconda/pandas) to install to your environment. Type  `conda install pandas` command line
6. Open a terminal (command-line), go to the directory where the peereval.py file is and type the following:
  * `python peereval.py -d <data directory>` , where `<data directory>` is where the completed forms are found.
  * If the completed forms are in the same directory as peereval.py then you can just type: `python peereval.py` 
7. If you see any warning messages, investigate and rerun the script - there may be a student who didn't complete. 
8. Open the csv output files (found in the `<data directory>`) and review.  There are 2 files: 
  * peereval.csv -- this is the raw data collated in one place (number of rows = groups x respondents x members), except the answer to the overarching question
  * pem.csv -- this is a summary table with the PEM, mean of score and all questions, plus the course/group level feedback


##Code Assumptions:     
* Any template modifications instructors make are captured correctly in the code (row,col of data cells are hard coded)
* Code works for groups of up to 8 students (but the example form only allows entry into 4 rows - this can easily be modified)
* Works for seven criteria (Q1 to Q7) - all weighted equally
* Retrieves a single response to a text-based overarching comment from each team member
* Assumes all files are in the same folder (including within subfolders), saved as *.xlsx with unique name
* Overwriting output files is ok
* That all forms are completed properly (there is some error checking, but not thorough)
* The PEM is not limited in any way. The instructor needs to do this post-hoc according to the course policy (see example above)


##Template Preparation: 
Original template was modified in Libre Office (Excel would work too)
1. Lay out the form in a suitable way
  * Highlight fields to be completed with grey
2. Add validation under Data/Validity 
  * Use list for group name and team member names (creates a dropdown selection). Here's a [youtube tutorial](https://www.youtube.com/watch?v=9i_-ErFVffs)
  * Put a valid range in for the answers to Q1 to Q7
  * Note that not allowing empty cells doesn't seem to work
3. Allow edits to only grey cells
  * Protect all cells, then unprotect the grey ones
  * Protect the entire sheet -- Note the password is "Muskoka" (use a different one for yours) Here's a [youtube tutorial](https://help.libreoffice.org/6.1/en-US/text/scalc/guide/cell_protect.html)


