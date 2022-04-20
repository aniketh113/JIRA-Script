from jira import JIRA
from structure import teamName,jiraOptions,auths,issueTypes
from datetime import date
import xlsxwriter
import time
today = date.today()
d2 = today.strftime("%B %d, %Y")
workbook = xlsxwriter.Workbook(f'Daily report-{d2}.xlsx')
 
# By default worksheet names in the spreadsheet will be
# Sheet1, Sheet2 etc., but we can also specify a name.
worksheet = workbook.add_worksheet("My sheet")
flag=0
today = date.today()
d2 = today.strftime("%B %d, %Y")

#URL and User Credentials for JIRA.
jira = JIRA(options=jiraOptions, basic_auth=auths)
m = []
j=[]
# Search all issues mentioned against a project name and specific issue types.
print("These are the issue with details:")
for singleIssue in jira.search_issues(jql_str=f'project = NIMBUS AND issuetype in {issueTypes}  AND status = "In Progress"AND team = {teamName} AND Sprint in openSprints()', fields=['key','summary', 'assignee', 'timetracking']):
    if singleIssue.fields.assignee != None or singleIssue.fields.timetracking.remainingEstimate != None:
        flag += 1
        j=[]
        j.append(singleIssue.key)
        j.append(singleIssue.fields.summary)
        j.append( singleIssue.fields.assignee)
        j.append(singleIssue.fields.timetracking.remainingEstimate)
        m.append(j)     
        print('{}:{}:{}:{}'.format(singleIssue.key, singleIssue.fields.summary,singleIssue.fields.assignee,singleIssue.fields.timetracking.remainingEstimate))        
if flag <=0:
    print(f"\nHI {auths[0]} ")
    print("----------------------------------------")
    print(" You are good to go for- ",d2)
    print("----------------------------------------")
    time.sleep(8)
else:
    print("\nYour report is now generated.")
    #for freezing the console for 10 secs
    time.sleep(4)
    row = 0
    col = 0
    for i in range(0, len(m)):
        worksheet.write(row, col, m[i][0])
        worksheet.write(row, col+1, m[i][1])
        worksheet.write(row, col+2, m[i][2].name)
        worksheet.write(row, col+3, m[i][3])
        row+= 1
    workbook.close()
#for freezing the console for 10 secs
#time.sleep(10)