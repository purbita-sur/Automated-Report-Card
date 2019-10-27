import gspread
from oauth2client.service_account import ServiceAccountCredentials
scope=['https://www.googleapis.com/auth/drive']
credentials=ServiceAccountCredentials.from_json_keyfile_name('milestone1.json',scope)
client=gspread.authorize(credentials)
import pandas as pd
import numpy as np
import seaborn as sns
import requests
import smtplib
import config
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import matplotlib.pyplot as plt
sheet=client.open_by_url("https://docs.google.com/spreadsheets/d/1HYjfEe3aCbufbqIXKs0Xz-gfoQNztGhCN1ivx0gZXnc/edit#gid=0")
idt=[]
task=[]
date=[]
module=[]
task_type=[]
student=[]
late_submission=[]
points=[]
total=[]
task_winner=[]
for i in sheet:
    data=i.get_all_records()
    for item in data:
        it=item['id']
        idt.append(it)
        t=item['Task']
        task.append(t)
        d=item['Date']
        date.append(d)
        m=item['Module']
        module.append(m)
        tp=item['Type']
        task_type.append(tp)
        s=item['Student']
        student.append(s)
        ls=item['Late Submission']
        late_submission.append(ls)
        pts=item['Points']
        points.append(pts)
        tot=item['Total']
        total.append(tot)
        tw=item['Task Winner']
        task_winner.append(tw)
df_list={}
df_list['Id']=idt
df_list['Task']=task
df_list['Date']=date
df_list['Module']=module
df_list['Type']=task_type
df_list['Student']=student
df_list['Late Submission']=late_submission
df_list['Points']=points
df_list['Total']=total
df_list['Task Winner']=task_winner
df= pd.DataFrame(data=df_list)
df['Date']=pd.to_datetime(df['Date'])
def analysis_date_range(start_date,end_date):
    mask=df[(df['Date']>start_date) & (df['Date']<end_date)]
    return mask
new_df=analysis_date_range('2019-01-01','2019-09-16')
new_df['Task']=new_df['Task'].str.lower()
new_df['Month']=new_df['Date'].dt.month
new_df['Month']=new_df['Month'].replace({3:8,4:8})
new_df.replace({'Swaastik':'Swaastick','Sushree':'Siddhishikha'},inplace=True)
new_df=new_df[new_df['Student']!=""]
new_df=new_df[new_df['Student']!='Wrick']
new_df=new_df[new_df['Student']!='Nitish']
new_df=new_df[new_df['Student']!='Sudhanshu']
def duplicates_removal(task):
    temp=new_df[new_df['Task']==task]
    months=temp['Date'].dt.month.unique().tolist()
    indx=np.array([])
    for date,month in zip(temp['Date'].values,temp['Month'].values):
        mask_date=temp['Date']==date
        mask_month=temp['Month']==month
        d=temp[mask_date & mask_month].sort_values('Student')
        indx=np.append(indx,d[d.duplicated(subset='Student',keep='last')].index)
    temp_1=temp.drop(indx)
    return temp_1
my_df=duplicates_removal('ajkyaukhada')
knowledge_df=duplicates_removal('knowledge sharing')
t=knowledge_df[knowledge_df['Date']>'2019-08-14']
print(t.head())
sunday=['Arya','Anjali','Surabhi']
monday=['Kunal','Bhavna','Apurwa']
tuesday=['Purbita','Chandrima','Prasoon']
wednesday=['Roumyak','Kaushal','Ujjainee']
thursday=['Shakib','Dipam','Siddhishikha']
friday=['Shantanu','Durga','Sonali']
saturday=['Swaastick','Sharika','Vishal']
t['weekdays']=t['Date'].dt.strftime("%A")
def knowledge_cleaner(knowledge):
    indx=np.array([])
    c=0
    for c in range (0,len(t['weekdays'])):
        i=t['weekdays'].iloc[c]
        student=knowledge['Student'].iloc[c]
        idx=knowledge['Student'].index[c]
        if i=='Monday':
            if(student not in monday):
                indx=np.append(indx,idx)
        elif i=='Tuesday':
            if(student not in tuesday):
                indx=np.append(indx,idx)
        elif i=='Wednesday':
            if(student not in wednesday):
                indx=np.append(indx,idx)
        elif i=='Thursday':
            if(student not in thursday):
                indx=np.append(indx,idx)
        elif i=='Friday':
            if(student not in friday):
                indx=np.append(indx,idx)
        elif i=='Saturday':
            if(student not in saturday):
                indx=np.append(indx,idx)
        elif i=='Sunday':
            if(student not in sunday):
                indx=np.append(indx,idx)
        c+=1
    return indx
idx=knowledge_cleaner(t)
knowledge_df=knowledge_df.drop(idx)
len(knowledge_df)
temp=new_df[new_df['Task']!='ajkyaukhada']
other_tasks=temp[temp['Task']!='knowledge sharing']
student_list=new_df['Student'].unique().tolist()
indx=np.array([])
for i in student_list:
    temp_df=other_tasks[other_tasks['Student']==i]
    indx=np.append(indx,temp_df[temp_df.duplicated(subset='Task',keep='last')].index)
other_tasks=other_tasks.drop(indx)
tasks=other_tasks.groupby('Student')['Points'].sum().reset_index()
myday_df=my_df.groupby('Student')['Points'].sum().reset_index()
knowledge_sharing=knowledge_df.groupby('Student')['Points'].sum().reset_index()
late_submission=other_tasks.groupby('Student')['Late Submission'].sum().reset_index()
wins=other_tasks.groupby('Student')['Task Winner'].sum().reset_index()
tot_task_score=other_tasks.groupby('Student')['Total'].sum().max()
final_df=myday_df.merge(knowledge_sharing,on='Student',how='outer')
final_df=final_df.merge(tasks,on='Student',how='outer')
final_df.insert(4,"Total Task Points",tot_task_score)
temp=final_df.drop(columns={'Total Task Points'})
final_df['Total Score'] = temp.sum(axis=1)
final_df=final_df.merge(late_submission,on='Student',how='outer')
final_df=final_df.merge(wins,on='Student',how='outer')
final_df.rename(columns={'Points_x':'MyDay Score','Points_y':'Knowledge Sharing Score','Late Submission':'Total Late Submissions','Task Winner':'Number of Wins','Points':'Task Score Achieved'},inplace=True)
final_df=final_df.sort_values(by='Total Score',ascending=False).reset_index(drop=True)
final_df
p1=plt.figure(figsize=(20,7))
sns.countplot(x='Student',data=my_df)
plt.xlabel('Names', fontsize=20)
plt.ylabel('Numbers',fontsize=20)
plt.title("Myday",fontsize=30)
plt.show()
p2=plt.figure(figsize=(20,7))
sns.countplot(x='Student',data=knowledge_df)
plt.xlabel('Names', fontsize=20)
plt.ylabel('Numbers',fontsize=20)
plt.title("Knowledge Sharing",fontsize=30)
plt.show()
len(knowledge_df)
p3=plt.figure(figsize=(20,7))
total=final_df.groupby('Student')['Task Score Achieved'].sum().sort_values(ascending=False)
total.plot.bar(x='Student',y='Task Score Achieved')
plt.xlabel('Student Names',fontsize=20)
plt.ylabel("Task Scores",fontsize=20)
plt.title("Task Score Graph",fontsize=30)
plt.show()
p4=plt.figure(figsize=(20,7))
total=final_df.groupby('Student')['Total Score'].sum().sort_values(ascending=False)
total.plot.bar(x='Student',y='Total Score')
plt.xlabel('Student Names',fontsize=20)
plt.ylabel("Total Scores",fontsize=20)
plt.title("Total Score Graph",fontsize=30)
plt.show()
len(t)
len(knowledge_df)
print(other_tasks)
x=my_df.groupby(['Student','Date'])['Points'].sum().reset_index()
p1.savefig("myday.jpg",bbox_inches='tight')
p2.savefig("knowledge.jpg",bbox_inches='tight')
p3.savefig("task",bbox_inches='tight')
p4.savefig("score",bbox_inches='tight')
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from datetime import date
date_label=date.today().strftime("%B %d %Y")
tot_days=len(pd.date_range(start="22-07-19",end='15-09-19'))
def report_creator(name):
    table=final_df[final_df['Student']==name]
    rank=table.index.values[0]+1
    total=table['Task Score Achieved'].iloc[0]
    late=table['Total Late Submissions'].iloc[0]
    win=table['Number of Wins'].iloc[0]
    myday=table['MyDay Score'].iloc[0]
    gyan=table['Knowledge Sharing Score'].iloc[0]
    my_const=my_df[my_df['Student']=='Purbita']['Student'].count()
    table1=final_df.head(10)
    student=table1['Student'].values
    tot=table1['Total Score'].values
    my_canvas = canvas.Canvas("Result.pdf",pagesize=letter)
    my_canvas.setLineWidth(.3)
    my_canvas.drawImage("campusx.jpeg", 143, 730,width=52,height=50)
    my_canvas.setFont('Times-Bold', 20)
    my_canvas.drawString(200,750,'CAMPUSX RESULTS')
    my_canvas.setFont('Times-Roman', 14)
    my_canvas.drawString(455,770,'Date: '+str(date_label))
    my_canvas.drawString(50,700,'Hi '+name+', here are your results!!')
    my_canvas.setFont('Times-Bold', 13)
    my_canvas.drawString(55,670,'Rank')
    my_canvas.drawString(165,670,'Total Score')
    my_canvas.drawString(305,670,'Task Wins')
    my_canvas.drawString(444,670,'Late Submission')
    my_canvas.setFont('Times-Roman', 13)
    my_canvas.drawString(55,650,str(rank))
    my_canvas.drawString(165,650,str(total))
    my_canvas.drawString(305,650,str(late))
    my_canvas.drawString(444,650,str(win))
    my_canvas.drawString(55,620,"Your Myday Score: "+str(myday))
    my_canvas.drawString(55,605,"Your Knowledge Sharing Score: "+str(gyan))
    my_canvas.drawString(330,620,"Your Myday Consistency: "+str(my_const)+" / "+str(tot_days))
    #my_canvas.drawString(380,605,"Your Knowledge Sharing Consistency: "+str(gyan))
    my_canvas.setFont('Times-Bold', 15)
    my_canvas.drawString(55,565,"Leaderboard")
    my_canvas.setFont('Times-Bold', 13)
    my_canvas.drawString(55,545,"Name")
    my_canvas.drawString(175,545,"Total Score")
    my_canvas.drawImage("image.jpg",300,380,width=200,height=200)
    my_canvas.setFont('Times-Roman', 12)
    c=525
    for i,j in zip(student,tot):
        my_canvas.drawString(55,c,i)
        my_canvas.drawString(175,c,str(j))
        c-=15
    my_canvas.setFont('Times-Bold', 13)
    my_canvas.drawString(55,350,"Myday Graph")
    my_canvas.drawImage("myday.jpg", 20, 30,width=580,height=300)
    my_canvas.showPage()
    my_canvas.setFont('Times-Bold', 13)
    my_canvas.drawString(55,750,"Knowledge Sharing Graph")
    my_canvas.drawImage("knowledge.jpg", 20, 420,width=580,height=300)
    my_canvas.save()
students = final_df['Student'].unique().tolist()
m = new_df['Date'].dt.strftime("%B")
m = m.replace({'March': 'August', 'April': 'August'})
mon = m.unique().tolist()
st = ""
for i in mon:
    st = st + i + ', '
print(st)
mails = client.open_by_url("https://docs.google.com/spreadsheets/d/1hTKuo8BCw3wuIt7ITiFjxEE9Xs-Gb9HIPo9GUSis-Vc/edit#gid=252560420")
emails = mails.get_worksheet(0).get_all_records()
names = []
email_ids = []
for i in emails:
    n=i['fname']
    if (n=='Sushree'):
        names.append("Siddhishikha")
    elif (n=='Kunal N.'):
        names.append('Kunal')
    elif (n=='Md Shakib'):
        names.append("Shakib")
    else:
        names.append(n)
    e=i['email']
    email_ids.append(e)
email_list={}
email_list['Names']=names
email_list['Email_id']=email_ids
email_df=pd.DataFrame(data=email_list)
def report_sender_via_email(student_names):
    for i in student_names:
        report_creator(i)
        receivers=email_df['Names'].tolist()
        for j in receivers:
            if (i==j):
                temp=email_df[email_df["Names"]==j]
                e_id=temp['Email_id'].iloc[0]
                email_user = config.email_user
                email_password = config.email_password
                email_send = e_id

                subject = 'CampusX Results'

                msg = MIMEMultipart()
                msg['From'] = email_user
                msg['To'] = email_send
                msg['Subject'] = subject

                body = 'Your Results for '+st
                msg.attach(MIMEText(body,'plain'))

                filename='Result.pdf'
                attachment  =open(filename,'rb')

                part = MIMEBase('application','octet-stream')
                part.set_payload((attachment).read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition',"attachment; filename= "+filename)

                msg.attach(part)
                text = msg.as_string()
                server = smtplib.SMTP('smtp.gmail.com',587)
                server.starttls()
                server.login(email_user,email_password)


                server.sendmail(email_user,email_send,text)
                server.quit()
report_sender_via_email(students)