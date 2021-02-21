# -*- coding: UTF-8 -*-
import os
import xlsxwriter
import requests as req
import json,sys,random

if os.getenv('ACCOUNT')== '' or os.getenv('OTHER_CONFIG') == '':
    print("<<<<<<<<<<<<<配置初始化中>>>>>>>>>>>>>")
    sys.exit()
else:
    account=json.loads(os.getenv('ACCOUNT'))
    other_config=json.loads(os.getenv('OTHER_CONFIG'))
if os.getenv('ACCOUNT_ADD') != '' or os.getenv('ACCOUNT_DEL') != '' or os.getenv('EMAIL') != '' or os.getenv('TG_BOT') != '':
    print("<<<<<<<<<<<<<配置初始化中>>>>>>>>>>>>>")
    sys.exit()    
if account == {'client_id':[],'client_secret':[],'ms_token':[]}:
    print("尚未设置账号")
    sys.exit()  
redirect_uri=os.getenv('REDIRECT_URI')
if redirect_uri =='':
    redirect_uri = r'https://login.microsoftonline.com/common/oauth2/nativeclient'
app_count=len(account['client_id'])
access_token_list=['wangziyingwen']*app_count
log_list=['']*app_count
###########################
# config选项说明
# 0：关闭  ， 1：开启
# allstart：是否全api开启调用，关闭默认随机抽取调用。默认0关闭
# rounds: 轮数，即每次启动跑几轮。
# rounds_delay: 是否开启每轮之间的随机延时，后面两参数代表延时的区间。默认0关闭
# api_delay: 是否开启api之间的延时，默认0关闭
# app_delay: 是否开启账号之间的延时，默认0关闭
########################################
config = {
         'allstart': 0,
         'rounds': 1,
         'rounds_delay': [0,0,5],
         'api_delay': [0,0,5],
         'app_delay': [0,0,5],
         }

#微软refresh_token获取
def getmstoken(appnum):
    #try:except?
    headers={
            'Content-Type':'application/x-www-form-urlencoded'
            }
    data={
         'grant_type': 'refresh_token',
         'refresh_token': ms_token,
         'client_id':client_id,
         'client_secret':client_secret,
         'redirect_uri':redirect_uri,
         }
    for retry_ in range(4):
        html = req.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',data=data,headers=headers)
        if html.status_code < 300:
            print(r'账号/应用 '+str(appnum)+' 的微软密钥获取成功')
            break
        else:
            if retry_ == 3:
                print(r'账号/应用 '+str(appnum)+' 的微软密钥获取失败\n'+'请检查secret里 CLIENT_ID , CLIENT_SECRET , MS_TOKEN 格式与内容是否正确，然后重新设置')
                if other_config['tg_bot'] != ['','']:
                    sendTgBot('AutoApi简报：'+'\n'+r'账号 '+str(appnum+1)+' token获取失败，运行中断')    
    jsontxt = json.loads(html.text)       
    return jsontxt['access_token']

#延时
def timeDelay(xdelay):
    if config[xdelay][0] == 1:
        time.sleep(random.randint(config[xdelay][1],config[xdelay][2]))
        
def apiReq(method,a,url,data='QAQ'):
    timeDelay('api_delay')
    access_token=access_token_list[a]
    headers={
            'Authorization': 'bearer ' + access_token,
            'Content-Type': 'application/json'
            }
    for retry_ in range(4):        
        if method == 'post':
            posttext=req.post(url,headers=headers,data=data)
        elif method == 'put':
            posttext=req.put(url,headers=headers,data=data)
        elif method == 'delete':
            posttext=req.delete(url,headers=headers)
        else :
            posttext=req.get(url,headers=headers)
        if posttext.status_code < 300:
            print('        操作成功')
            break
            #操作成功跳出循环
        else:
            if retry_ == 3:
                print('        操作失败')
    return posttext     
          

#上传文件到onedrive(小于4M)
def uploadFile(a,filesname,f):
    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/App'+str(a)+r'/'+filesname+r':/content'
    if apiReq('put',a,url,f).status_code >= 300 :
        if sys._getframe().f_code.co_name not in log_list[a]:
            log_list[a]=log_list[a]+sys._getframe().f_code.co_name+','
    
        
# 发送邮件到自定义邮箱
def sendEmail(a,subject,content):
    url=r'https://graph.microsoft.com/v1.0/me/sendMail'
    mailmessage={
                'message':{
                          'subject': subject,
                          'body': {'contentType': 'Text', 'content': content},
                          'toRecipients': [{'emailAddress': {'address': other_config['email'][0]}}],
                          },
                'saveToSentItems': 'true',
                }            
    if apiReq('post',a,url,json.dumps(mailmessage)).status_code >= 300 :
        if sys._getframe().f_code.co_name not in log_list[a]:
            log_list[a]=log_list[a]+sys._getframe().f_code.co_name+','
    	
	
#修改excel(这函数分离好像意义不大)
#api-获取itemid: https://graph.microsoft.com/v1.0/me/drive/root/search(q='.xlsx')?select=name,id,webUrl
def excelWrite(a,filesname,sheet):
    try:
        print('    添加工作表')
        url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/App'+str(a)+r'/'+filesname+r':/workbook/worksheets/add'
        data={
             "name": sheet
             }
        apiReq('post',a,url,json.dumps(data))
        print('    添加表格')
        url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/App'+str(a)+r'/'+filesname+r':/workbook/worksheets/'+sheet+r'/tables/add'
        data={
             "address": "A1:D8",
             "hasHeaders": False
             }
        jsontxt=json.loads(apiReq('post',a,url,json.dumps(data)).text)
        print('    添加行')
        url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/App'+str(a)+r'/'+filesname+r':/workbook/tables/'+jsontxt['id']+r'/rows/add'
        rowsvalues=[[0]*4]*2
        for v1 in range(0,2):
            for v2 in range(0,4):
                rowsvalues[v1][v2]=random.randint(1,1200)
        data={
             "values": rowsvalues
             }
        apiReq('post',a,url,json.dumps(data))
    except:
        print("        操作中断")
        if sys._getframe().f_code.co_name not in log_list[a]:
            log_list[a]=log_list[a]+sys._getframe().f_code.co_name+','
        return 
    
def taskWrite(a,taskname):
    try:
        print("    创建任务列表")
        url=r'https://graph.microsoft.com/v1.0/me/todo/lists'
        data={
             "displayName": taskname
             }
        listjson=json.loads(apiReq('post',a,url,json.dumps(data)).text)
        print("    创建任务")
        url=r'https://graph.microsoft.com/v1.0/me/todo/lists/'+listjson['id']+r'/tasks'
        data={
             "title": taskname,
             }
        taskjson=json.loads(apiReq('post',a,url,json.dumps(data)).text)
        print("    删除任务")
        url=r'https://graph.microsoft.com/v1.0/me/todo/lists/'+listjson['id']+r'/tasks/'+taskjson['id']
        apiReq('delete',a,url)
        print("    删除任务列表")
        url=r'https://graph.microsoft.com/v1.0/me/todo/lists/'+listjson['id']
        apiReq('delete',a,url)
    except:
        print("        操作中断")
        if sys._getframe().f_code.co_name not in log_list[a]:
            log_list[a]=log_list[a]+sys._getframe().f_code.co_name+','
        return 
    
def teamWrite(a,channelname):
    #新建team
    try:
        print('    新建team')
        url=r'https://graph.microsoft.com/v1.0/teams'
        data={
             "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
             "displayName": channelname,
             "description": "My Sample Team’s Description"
             }
        apiReq('post',a,url,json.dumps(data))
        print("    获取team信息")
        url=r'https://graph.microsoft.com/v1.0/me/joinedTeams'
        teamlist = json.loads(apiReq('get',a,url).text)
        for teamcount in range(teamlist['@odata.count']):
            if teamlist['value'][teamcount]['displayName'] == channelname:
                #创建频道
                print("    创建team频道")
                data={
                     "displayName": channelname,
                     "description": "This channel is where we debate all future architecture plans",
                     "membershipType": "standard"
                     }
                url=r'https://graph.microsoft.com/v1.0/teams/'+teamlist['value'][teamcount]['id']+r'/channels'
                jsontxt = json.loads(apiReq('post',a,url,json.dumps(data)).text)
                url=r'https://graph.microsoft.com/v1.0/teams/'+teamlist['value'][teamcount]['id']+r'/channels/'+jsontxt['id']
                print("    删除team频道")
                apiReq('delete',a,url)
                #删除teams
                print("    删除team")
                url=r'https://graph.microsoft.com/v1.0/groups/'+teamlist['value'][teamcount]['id']
                apiReq('delete',a,url)  
    except:
        print("        操作中断")
        if sys._getframe().f_code.co_name not in log_list[a]:
            log_list[a]=log_list[a]+sys._getframe().f_code.co_name+','
        return 
        
def onenoteWrite(a,notename):
    try:
        print('    创建笔记本')
        url=r'https://graph.microsoft.com/v1.0/me/onenote/notebooks'
        data={
             "displayName": notename,
             }
        notetxt = json.loads(apiReq('post',a,url,json.dumps(data)).text)
        print('    创建笔记本分区')
        url=r'https://graph.microsoft.com/v1.0/me/onenote/notebooks/'+notetxt['id']+r'/sections'
        data={
             "displayName": notename,
             }
        apiReq('post',a,url,json.dumps(data))
        print('    删除笔记本')
        url=r'https://graph.microsoft.com/v1.0/me/drive/root:/Notebooks/'+notename
        apiReq('delete',a,url)
    except:
        print("        操作中断")
        if sys._getframe().f_code.co_name not in log_list[a]:
            log_list[a]=log_list[a]+sys._getframe().f_code.co_name+','
        return
        
def sendTgBot(content):
    headers={
            'Content-Type': 'application/json'
            }
    data={
         'chat_id':other_config['tg_bot'][1],
         'text':content,
         'parse_mode':'HTML'
         }  
    for retry_ in range(4):  
        posttext=req.post(r'https://api.telegram.org/bot'+other_config['tg_bot'][0]+r'/sendMessage',headers=headers,data=json.dumps(data))
        if posttext.status_code < 300:
             print('tg推送成功')
             break
        else:
            if retry_ == 3:
                print('tg推送失败')
    print('')

#一次性获取access_token，降低获取率
for a in range(0, app_count):
    client_id=account['client_id'][a]
    client_secret=account['client_secret'][a]
    ms_token=account['ms_token'][a]
    access_token_list[a]=getmstoken(a)
print('')    
#获取天气
if len(other_config['email']) <= 1:
    city='Beijing'
else:
    city=other_config['email'][1]
headers={'Accept-Language': 'zh-CN'}
weather=req.get(r'http://wttr.in/'+city+r'?format=4&?m',headers=headers).text

#实际运行
for a in range(0, app_count):
    print('账号 '+str(a+1))
    print('发送邮件 ( 邮箱单独运行，每次运行只发送一次，防止封号 )')
    if other_config['email'][0] != '':
        sendEmail(a,'weather',weather)
    else:
        print("尚未配置邮箱")
print('')
#其他api
for _ in range(1,config['rounds']+1):
    timeDelay('rounds_delay')  
    print('第 '+str(_)+' 轮\n')        
    for a in range(0, app_count):
        timeDelay('app_delay')    
        print('账号 '+str(a+1))    
        #生成随机名称
        filesname='QAQ'+str(random.randint(1,600))+r'.xlsx'
        #新建随机xlsx文件
        xls = xlsxwriter.Workbook(filesname)
        xlssheet = xls.add_worksheet()
        for s1 in range(0,4):
            for s2 in range(0,4):
                xlssheet.write(s1,s2,str(random.randint(1,600)))
        xls.close()
        xlspath=sys.path[0]+r'/'+filesname
        print('上传文件')
        with open(xlspath,'rb') as f:
            uploadFile(a,filesname,f)
        choosenum = random.sample(range(1, 5),2)
        if config['allstart'] == 1 or 1 in choosenum:
            print('excel文件操作')
            excelWrite(a,filesname,'QVQ'+str(random.randint(1,600)))
        if config['allstart'] == 1 or 2 in choosenum:
            print('team操作')
            teamWrite(a,'QVQ'+str(random.randint(1,600)))
        if config['allstart'] == 1 or 3 in choosenum:
            print('task操作')
            taskWrite(a,'QVQ'+str(random.randint(1,600)))
        if config['allstart'] == 1 or 4 in choosenum:
            print('onenote操作')
            onenoteWrite(a,'QVQ'+str(random.randint(1,600)))
        print('-')
if other_config['tg_bot'] != []:
    content='AutoApi.W简报: '+'\n'
    for i in range(app_count):
        if log_list[i] != '':
            content=content+'账号 '+str(i)+' 失败 api：'+log_list[i]+'\n'
        else:
            content=content+'账号 '+str(i)+' 所有 api : OK'+'\n'
    sendTgBot(content)