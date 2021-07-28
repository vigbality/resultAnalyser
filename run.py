

#tot appeared, tot pas, tot fail, pass%, 33-44%, 45-59, 60-74, 75-89, 90 and above, total, perc stud 70% and above, no of stud 70% and above '
def cal(subName):
	l=df[subName].tolist()
	d={1:0,2:0,3:0,4:0,5:0,6:0,7:0,8:0,9:0,10:0,11:0,12:0}
	tot=l.pop(0)
	d[10]=int(tot)
	for _ in range(l.count(-1)):
		l.remove(-1)
	d[1]=len(l)
	for i in l:
		p=floor((i/tot)*100)
		if p<33:
			d[3]+=1
		elif 33<=p<=44:
			d[5]+=1
		elif 45<=p<=59:
			d[6]+=1
		elif 60<=p<=74:
			d[7]+=1
		elif 75<=p<=89:
			d[8]+=1
		else:
			d[9]+=1
		if p>=70:
			d[12]+=1
		else:
			pass
	d[2]=d[1]-d[3]
	d[4]=round((d[2]/d[1])*100,2)
	d[11]=round((d[12]/d[1])*100,2)
	return list(d.values())

def calSC(df):
	l=[]
	for i in df.values:
		l.append(sum(list(i)[3:]))
	d={1:0,2:0,3:0,4:0,5:0,6:0,7:0,8:0,9:0,10:0,11:0,12:0}
	l.pop(0)
	tot=500
	d[10]=int(tot)
	for _ in range(l.count(-1)):
		l.remove(-1)
	d[1]=len(l)
	for i in l:
		p=floor((i/tot)*100)
		if p<33:
			d[3]+=1
		elif 33<=p<=44:
			d[5]+=1
		elif 45<=p<=59:
			d[6]+=1
		elif 60<=p<=74:
			d[7]+=1
		elif 75<=p<=89:
			d[8]+=1
		else:
			d[9]+=1
		if p>=70:
			d[12]+=1
		else:
			pass
	d[2]=d[1]-d[3]
	d[4]=round((d[2]/d[1])*100,2)
	d[11]=round((d[12]/d[1])*100,2)
	return list(d.values())

def calPI(subName):
	l=df[subName].tolist()
	d={1:0,2:0,3:0,4:0,5:0,6:0,7:0,8:0,9:0,10:0,11:0,12:0,13:0,14:0,15:0}
	tot=l.pop(0)
	d[13]=int(tot)
	for _ in range(l.count(-1)):
		l.remove(-1)
	d[1]=len(l)
	for i in l:
		p=floor((i/tot)*100)
		if p<33:
			d[3]+=1
		elif 33<=p<=38:
			d[5]+=1
		elif 39<=p<=44:
			d[6]+=1
		elif 45<=p<=50:
			d[7]+=1
		elif 51<=p<=60:
			d[8]+=1
		elif 61<=p<=70:
			d[9]+=1
		elif 71<=p<=80:
			d[10]+=1
		elif 81<=p<=90:
			d[11]+=1	
		else:
			d[12]+=1
		if p>=70:
			d[15]+=1
		else:
			pass
	d[2]=d[1]-d[3]
	d[4]=round((d[2]/d[1])*100,2)
	d[14]=round((d[15]/d[1])*100,2)
	return list(d.values())


def calSCPI(df):
	l=[]
	for i in df.values:
		l.append(sum(list(i)[3:]))
	d={1:0,2:0,3:0,4:0,5:0,6:0,7:0,8:0,9:0,10:0,11:0,12:0,13:0,14:0,15:0}
	l.pop(0)
	tot=500
	d[13]=int(tot)
	for _ in range(l.count(-1)):
		l.remove(-1)
	d[1]=len(l)
	for i in l:
		p=floor((i/tot)*100)
		if p<33:
			d[3]+=1
		elif 33<=p<=38:
			d[5]+=1
		elif 39<=p<=44:
			d[6]+=1
		elif 45<=p<=50:
			d[7]+=1
		elif 51<=p<=60:
			d[8]+=1
		elif 61<=p<=70:
			d[9]+=1
		elif 71<=p<=80:
			d[10]+=1
		elif 81<=p<=90:
			d[11]+=1	
		else:
			d[12]+=1
		if p>=70:
			d[15]+=1
		else:
			pass
	d[2]=d[1]-d[3]
	d[4]=round((d[2]/d[1])*100,2)
	d[14]=round((d[15]/d[1])*100,2)
	return list(d.values())

#----------------------------------------------------------------------------------------------
from pandas import read_excel
from math import floor
print('PLEASE WAIT...')
df = read_excel('all.xls').fillna(-1)

sub=list(df.columns)[3:]
res={}
for s in sub:
	res[s]=[cal(s),calPI(s)]

from openpyxl import load_workbook
myworkbook=load_workbook('xyz.xlsx')
worksheet=myworkbook['RESULT ANALYSIS XII']
worksheetPI=myworkbook['PI CALCULATION XII']


start=7
for subName,subData in res.items():
	worksheet['A'+str(start)]=subName
	worksheetPI['A'+str(start)]=subName
	for r in range(12):
		worksheet[chr(67+r)+str(start)]=subData[0][r]
	for r in range(15):
		worksheetPI[chr(67+r)+str(start)]=subData[1][r]
	start+=1


dfSci = read_excel('sci.xls').fillna(0)
dfCom= read_excel('com.xls').fillna(0)


subDs=calSC(dfSci)
subDc=calSC(dfCom)

subDsPI=calSCPI(dfSci)
subDcPI=calSCPI(dfCom)


for r in range(12):
	worksheet[chr(67+r)+'30']=subDs[r]
	worksheet[chr(67+r)+'31']=subDc[r]

for r in range(15):
	worksheetPI[chr(67+r)+'30']=subDsPI[r]
	worksheetPI[chr(67+r)+'31']=subDcPI[r]



myworkbook.save("RESULT_ANALYSIS_XII.xlsx")
	
print('SUCCESSFULLY SAVED')