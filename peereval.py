#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
peer_eval.py

Script to:
    1) take data from an directory of completed Excel forms 
    2) concatentate and validate the data
    3) compute participation multiplier
    4) output all the records in a csv file

Assumptions:     
* Template modifications are captured correctly in the code below (row,col of data cells)
* Works for groups of up to 8 students (must allow entry into more rows)
* Works for seven criteria (Q1 to Q7) and a comment for each team member - all weighted equally
* Can handle a single response to a text-based course/group level question
* Assumes all files are in the same folder, saved as *.xlsx with unique name
* Overwriting output files is ok
* That all forms are completed properly (there is some error checking, not much)

Template Preparation: 
 Original template was modified in Libre Office (Excel would work too)
*Lay out the form in a suitable way
 -Highlight fields to be completed with grey
*Add validation under Data/Validity 
 -Use list for group name and team member names (creates a dropdown selection) 
     --https://www.youtube.com/watch?v=9i_-ErFVffs
 -Put a valid range in for the answers to Q1 to Q7
 -Note that not allowing empty cells doesn't seem to work
*Allow edits to only grey cells
 -Protect all cells, then unprotect the grey ones
 -Protect the entire sheet -- Note password is "Muskoka"
 --https://help.libreoffice.org/6.1/en-US/text/scalc/guide/cell_protect.html


Acknowledgements: 
*Used concept and template from Morgan Rooney and Audrey Girouard, Carleton University
*Python code below under GNU GPLv3 Licence
*Author Derek Mueller, Carleton University
*Created April 2022

"""

# import libraries
import os
import sys
import glob
import pandas as pd  # need the optional package 'openpxyl'
import numpy as np
import argparse

def readform(f):
    """
    reads an excel form in the current directory
    positions of data hard coded in 
    
    returns 2 data frames - one of the data entry for each student
    another for the overarching comment

    Parameters
    ----------
    f : str
        Excel file name

    Returns
    -------
    team_eval: pandas dataframe
        contains the group, respondent, answers to q17 and comments on all team members
    grp_feedback: pandas dataframe
        contains the group, respondent, and text answer for group level question
    """
        
    # open file as pandas df 
    df = pd.read_excel(f)
    # knowing the format will stay the same, hard code the cell addresses
    respondent = df.iloc[2,2]  # name of respondent
    group = df.iloc[4,2]  # group
    feedback = df.iloc[9,11]  # overall feedback question

    names=df.iloc[17:25,1]  # team member names
    arr = df.iloc[17:25,2:9]  # question array - q1 to q7
    #avg = df.iloc[17:25,10]  # the average (which we will calculate in python)
    comments = df.iloc[17:25,11]  # comments on team members 

    #combine into group dataframe
    grp_feedback = pd.DataFrame({'group':  [group], 'respondent': [respondent], 'feedback': [feedback]})
    
    #combine into team member dataframe
    team_eval = pd.concat([names,arr,comments],axis=1)
    team_eval = team_eval.dropna(how='all')  # remove all blank rows
    team_eval.insert(0,'respondent',value=respondent)  # add the respondent as a column
    team_eval.insert(0,'group',value=group) # add the group as a column
    #name the columns
    team_eval.columns = ['group','respondent','name','q1','q2','q3','q4','q5','q6','q7','comments']
    
    return team_eval, grp_feedback


def dataValid(dfeval):
    """
    Checks data to make sure it is valid.  Run and look at output.  
    
    Pretty rudimentary, many other things could be checked, but let's not get 
    too carried away guessing what might go wrong for now... 

    Parameters
    ----------
    dfeval : pandas dataframe 
        Dataframe from readform function
        
    Returns
    -------
    None.

    """
    
    print('Checking data validity... ')
    groups = dfeval.group.unique()
    for grp in groups:
        print("\n Checking Group {}:".format(grp))
        gmembers = np.sort(dfeval.loc[dfeval['group'] == grp]['respondent'].unique())
        for member in gmembers:
            print('  Checking form from {}'.format(member))
            ### Every group should have the same team members across all respondents
            m = dfeval.loc[(dfeval['group'] == grp) & 
                           (dfeval['respondent'] == member)]['name'].sort_values().to_numpy()
            if not np.array_equal(gmembers,m):
                print('  ...Problem with form from {}, check that all group members are listed'.format(member))
            ### all data should be there    
            d = dfeval.loc[(dfeval['group'] == grp) & 
                           (dfeval['respondent'] == member),['q1','q2','q3','q4','q5','q6','q7']]
            if d.shape[0] != len(gmembers):
                print('  ...Problem with form from {}, check that all group members are evaluated'.format(member))
            if d.isnull().sum().sum() >0:
                print('  ...Problem with form from {}, check that all criteria are evaluated'.format(member))

    print('\n Data checking complete.  If there were issues, be sure to address them before finalizing the analysis \n')

def calcPEM(dfeval):    
    """
    Calculates the Peer Evaluation Multiplier (PEM) for each group member

    Parameters
    ----------
    dfeval : dataframe
        dataframe from the readeval function

    Returns
    -------
    score_avg : dataframe
        a dataframe with group, name, score and pem

    """

   
    # create column 'score' which gives the sum of all the scores from q1 to q7
    dfeval.insert(3,'score', dfeval[['q1','q2','q3','q4','q5','q6','q7']].sum(axis=1))
    
    # get the average score per group member
    score_avg = dfeval.groupby(['group', 'name'])[['score']].mean().round(2)
    
    # make new column for the Peer Evaluation Multiplier (PEM)
    score_avg = score_avg.assign(pem=0)
    
    # get the average score per group
    grp_avg = dfeval.groupby(['group'])[['score']].mean()
    
    
    # calculate the pem and add it to the dataframe
    for grp in dfeval.groupby(['group']).groups.keys():
        pem = score_avg.loc[[grp],'score']/grp_avg.loc[[grp],'score']
        score_avg.loc[grp,['pem']] = pem.round(3)


    return score_avg


if __name__ == '__main__':
    
    # argparse here
    
    description = "Script to calculate Peer Evaluation Modifier from student evaluations"    
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument("-d", "--dir", help="the directory where the xlsx files are found and the output csv files will be placed (defaults to current directory")
    args = parser.parse_args()

    print('\n Peer Evaluation Modifier calculation in progress...\n')

    # switch to directory with xls files
    if args.dir:
        try:
            os.chdir(args.dir)
        except:
            print('Cannot find that directory, exiting')
            sys.exit(1)
    
    # look for xlsx files                
    xlist = glob.glob("**/*.xlsx", recursive=True)
    
    if len(xlist) == 0:
        print('Could not find any xlsx files, exiting')
        sys.exit(1)

    # create empty dataframes        
    dfeval = pd.DataFrame(columns=['group','respondent','name','q1','q2','q3','q4','q5','q6','q7','comments'])
    dfgroup = pd.DataFrame(columns=['group', 'respondent', 'feedback'])

    for f in xlist:
        #read each file and add them to a larger dataframe
        t, g = readform(f)
        dfgroup = pd.concat([dfgroup,g])
        dfeval = pd.concat([dfeval,t])
        del [t,g]  # remove temporary data frames
    
    dfeval = dfeval.sort_values(['group','respondent', 'name'])
    dfgroup = dfgroup.sort_values(['group','respondent'])    
    
    #check to see if the data are valid
    dataValid(dfeval)
    score_pem = calcPEM(dfeval)
    
    ##export the dataframes
    dfgroup.to_csv('dfgroup.csv', index=False)
    dfeval.to_csv('dfeval.csv', index=False)
    score_pem.to_csv('pem.csv')
    
    print('Completed calculations and data export... \n')
