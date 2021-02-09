# -*- coding: UTF-8 -*-
import os
import requests as req
import json,sys,time,random

reload(sys)
sys.setdefaultencoding('utf-8')
emailaddress=os.getenv('EMAIL')
app_num=os.getenv('APP_NUM')
#是否全启（3选1）
config = {'allstart':'Y','延时':'Y'}
if app_num == '':
    app_num = '1'
city=os.getenv('CITY')
if city == '':
    city = 'Beijing'
access_token_list=['wangziyingwen']*int(app_num)

#微软refresh_token获取
def getmstoken(ms_token,appnum):
    headers={'Content-Type':'application/x-www-form-urlencoded'
            }
    data={'grant_type': 'refresh_token',
        'refresh_token': ms_token,
        'client_id':client_id,
        'client_secret':client_secret,
        'redirect_uri':'http://localhost:53682/'
        }
    html = req.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',data=data,headers=headers)
    jsontxt = json.loads(html.text)
    if 'refresh_token' in jsontxt:
        print(r'账号/应用 '+str(appnum)+' 的微软密钥获取成功')
    else:
        print(r'账号/应用 '+str(appnum)+' 的微软密钥获取失败\n'+'请检查secret里 CLIENT_ID , CLIENT_SECRET , MS_TOKEN 格式与内容是否正确，然后重新设置')
    refresh_token = jsontxt['refresh_token']
    access_token = jsontxt['access_token']
    return access_token

def apiReq(method,a,url,data='QAQ'):
    access_token=access_token_list[a-1]
    headers={
            'Authorization': 'bearer ' + access_token,
            'Content-Type': 'application/json'
            }
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
    else:
        print('        操作失败')
#    if posttext.status_code > 300:
#        print('        操作失败')
#        #成功不提示
    return posttext.text
          

#上传文件到onedrive(小于4M)
def UploadFile(a,filesname,f):
    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/App'+str(a)+r'/'+filesname+r':/content'
    apiReq('put',a,url,f)
    
        
# 发送邮件到自定义邮箱
def SendEmail(a,subject,content):
    url=r'https://graph.microsoft.com/v1.0/me/sendMail'
    mailmessage={'message': {'subject': subject,
                             'body': {'contentType': 'Text', 'content': content},
                             'toRecipients': [{'emailAddress': {'address': emailaddress}}],
                             },
                 'saveToSentItems': 'true'}            
    apiReq('post',a,url,json.dumps(mailmessage))	
	
#修改excel(这函数分离好像意义不大)
#api-获取itemid: https://graph.microsoft.com/v1.0/me/drive/root/search(q='.xlsx')?select=name,id,webUrl
def excelWrite(a,filesname,sheet):
    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/App'+str(a)+r'/'+filesname+r':/workbook/worksheets/add'
    data={
         "name": sheet
         }
    print('    添加工作表')
    apiReq('post',a,url,json.dumps(data))
    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/App'+str(a)+r'/'+filesname+r':/workbook/worksheets/'+sheet+r'/tables/add'
    data={
         "address": "A1:D8",
         "hasHeaders": False
         }
    print('    添加表格')
    jsontxt=json.loads(apiReq('post',a,url,json.dumps(data)))
    print('    获取表格')
    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/App'+str(a)+r'/'+filesname+r':/workbook/worksheets/'+sheet+r'/tables'
    jsontxt=json.loads(apiReq('get',a,url))
    print(jsontxt['value'][0]['id'])
    print(jsontxt['value'][0]['name'])
#   添加行失败，搞不懂。
#    print('    添加行')
#    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/App'+str(a)+r'/'+filesname+r':/workbook/tables/'+jsontxt['id']+r'/rows/add'
#    data={
#          "values": [
#          [random.randint(1,1200) , random.randint(1,1200), random.randint(1,1200)],
#          [random.randint(1,1200) , random.randint(1,1200), random.randint(1,1200)]
#          ]
#         }
#    print(apiReq('post',a,url,json.dumps(data)))
    
def taskWrite(a,taskname):
    url=r'https://graph.microsoft.com/v1.0/me/todo/lists'
    data={
         "displayName": taskname
         }
    print("    创建任务列表")
    listjson=json.loads(apiReq('post',a,url,json.dumps(data)))
    url=r'https://graph.microsoft.com/v1.0/me/todo/lists/'+listjson['id']+r'/tasks'
    data={
         "title": taskname,
         }
    print("    创建任务")
    taskjson=json.loads(apiReq('post',a,url,json.dumps(data)))
    url=r'https://graph.microsoft.com/v1.0/me/todo/lists/'+listjson['id']+r'/tasks/'+taskjson['id']
    print("    删除任务")
    apiReq('delete',a,url)
    url=r'https://graph.microsoft.com/v1.0/me/todo/lists/'+listjson['id']
    print("    删除任务列表")
    apiReq('delete',a,url)    
    
def teamWrite(a,channelname):
    url=r'https://graph.microsoft.com/v1.0/me/joinedTeams'
    print("    获取team")
    jsontxt = json.loads(apiReq('get',a,url))
    objectlist=jsontxt['value']
    #创建
    print("    创建team频道")
    data={
         "displayName": channelname,
         "description": "This channel is where we debate all future architecture plans",
         "membershipType": "standard"
         }
    url=r'https://graph.microsoft.com/v1.0/teams/'+objectlist[0]['id']+r'/channels'
    jsontxt = json.loads(apiReq('post',a,url,json.dumps(data)))
    url=r'https://graph.microsoft.com/v1.0/teams/'+objectlist[0]['id']+r'/channels/'+jsontxt['id']
    print("    删除team频道")
    apiReq('delete',a,url)

def onenoteWrite(a,notename):
    url=r'https://graph.microsoft.com/v1.0/me/onenote/notebooks'
    data={
         "displayName": notename
         }
    print('    创建笔记本')
    notetxt = json.loads(apiReq('post',a,url,json.dumps(data)))
    print('    删除笔记本')
    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/Notebooks/'+notename
    apiReq('delete',a,url)
    
#一次性获取access_token，降低获取率
for a in range(1, int(app_num)+1):
    client_id=os.getenv('CLIENT_ID_'+str(a))
    client_secret=os.getenv('CLIENT_SECRET_'+str(a))
    ms_token=os.getenv('MS_TOKEN_'+str(a))
    access_token_list[a-1]=getmstoken(ms_token,a)
    
#获取天气
headers={'Accept-Language': 'zh-CN'}
weather=req.get(r'http://wttr.in/'+city+r'?format=4&?m',headers=headers).text

#实际运行   
for a in range(1, int(app_num)+1):
    #生成随机名称
    filesname='QAQ'+str(random.randint(1,600))+r'.xlsx'
    os.rename('AutoApi.xlsx',filesname)
    xlspath=sys.path[0]+r'/'+filesname
    print('可能会偶尔出现创建上传失败的情况'+'\n'+'上传文件')
    with open(xlspath,'rb') as f:
        UploadFile(a,filesname,f)
    print('发送邮件')
    if emailaddress != '':
        SendEmail(a,'weather',weather)
    choosenum = random.randint(1,4) 
    if config['allstart'] == 'Y' or choosenum == 1:
        print('excel文件操作')
        excelWrite(a,filesname,'QVQ'+str(random.randint(1,600)))
    if config['allstart'] == 'Y' or choosenum == 2:
        print('team操作')
        teamWrite(a,'QVQ'+str(random.randint(1,600)))
    if config['allstart'] == 'Y' or choosenum == 3:
        print('task操作')
        taskWrite(a,'QVQ'+str(random.randint(1,600)))
    if config['allstart'] == 'Y' or choosenum == 4:
        print('onenote操作')
        onenoteWrite(a,'QVQ'+str(random.randint(1,600)))