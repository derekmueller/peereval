#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
peer_eval.py

Script for tabulating peer evaluation responses for group work

The concept and template detailed below are from the [Welcome to My Classroom workshop](https://carleton.ca/tls/2019/welcome-to-my-classroom-using-peer-to-peer-approaches-to-solve-common-assessment-challenges/) by Audrey Girouard (School of Information Technology) and Morgan Rooney (English), Carleton University.  Code was written in Python by Derek Mueller (Geography), Carleton University. 

**The code works as follows:** 
1. students complete a survey form (Excel xlsx file and and it in)
2. script reads from a directory of completed Excel forms 
3. data are concatentated and validated by group and respondent
3. a "Peer Evaluation Multiplier" (PEM) is computed
4. all the records are output to csv files

See readme.md for more info

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
    respondent = df.iloc[2,2]  # member of respondent
    group = df.iloc[4,2]  # group
    feedback = df.iloc[9,11]  # overall feedback question

    members=df.iloc[17:25,1]  # team member members
    arr = df.iloc[17:25,2:9]  # question array - q1 to q7
    #avg = df.iloc[17:25,10]  # the average (which we will calculate in python)
    comments = df.iloc[17:25,11]  # comments on team members 

    #combine into group dataframe
    grp_feedback = pd.DataFrame({'group':  [group], 'respondent': [respondent], 'feedback': [feedback]})
    
    #combine into team member dataframe
    team_eval = pd.concat([members,arr,comments],axis=1)
    team_eval = team_eval.dropna(how='all')  # remove all blank rows
    team_eval.insert(0,'respondent',value=respondent)  # add the respondent as a column
    team_eval.insert(0,'group',value=group) # add the group as a column
    #member the columns
    team_eval.columns = ['group','respondent','member','q1','q2','q3','q4','q5','q6','q7','comments']
        
    return team_eval.reset_index(drop=True), grp_feedback.reset_index(drop=True)


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
                           (dfeval['respondent'] == member)]['member'].sort_values().to_numpy()
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

def calcPEM(dfeval, dfmember):    
    """
    Calculates the Peer Evaluation Multiplier (PEM) for each group member

    Parameters
    ----------
    dfeval : pandas dataframe
        dataframe from the readeval function (raw eval data)

    dfmember : pandas dataframe
        dataframe with one row per each member

    Returns
    -------
    score_avg : dataframe
        a dataframe with group, member, score, pem, avg score per q1-7, and feedback response

    """

   
    # create column 'score' which gives the sum of all the scores from q1 to q7
    dfeval.insert(3,'score', dfeval[['q1','q2','q3','q4','q5','q6','q7']].sum(axis=1))
    # these fields are not needed... so remove
    dfeval = dfeval.drop ( ['respondent', 'comments'], axis = 1) 
    # get the average score per group member
    score_avg = dfeval.groupby(['group', 'member']).agg(['mean']).round(4)
    # reset column index - they are all means
    score_avg.columns = score_avg.columns.droplevel(1)
    # make new column for the Peer Evaluation Multiplier (PEM)
    score_avg.insert(0,'pem',0)
    
    # get the average score per group
    grp_avg = dfeval.groupby(['group'])[['score']].mean()
    
    # calculate the pem and add it to the dataframe
    for grp in dfeval.groupby(['group']).groups.keys():
        pem = score_avg.loc[[grp],'score']/grp_avg.loc[[grp],'score']
        score_avg.loc[grp,['pem']] = pem.round(4)
    
    #merge the respondent feedback with score_avg dataframe
    dfmember.columns = ['group','member','feedback']  # fix col names
    dfmember = dfmember.set_index(['group','member'])  # create matching index
    score_avg = pd.concat([score_avg, dfmember], axis=1, join ="inner") #merge
    
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
    dfeval = pd.DataFrame(columns=['group','respondent','member','q1','q2','q3','q4','q5','q6','q7','comments'])
    dfmember = pd.DataFrame(columns=['group', 'respondent', 'feedback'])

    for f in xlist:
        #read each file and add them to a larger dataframe
        t, g = readform(f)
        dfmember = pd.concat([dfmember,g])
        dfeval = pd.concat([dfeval,t])
        del [t,g]  # remove temporary data frames
    
    dfeval = dfeval.sort_values(['group','respondent', 'member'])
    dfmember = dfmember.sort_values(['group','respondent'])    
    
    #check to see if the data are valid
    dataValid(dfeval)
    #calculate everything
    score_pem = calcPEM(dfeval, dfmember)
    
    ##export the dataframes
    dfeval.to_csv('peereval.csv', index=False)
    score_pem.to_csv('pem.csv')
    
    print('Completed calculations and data export... \n')
