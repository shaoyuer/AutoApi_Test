# -*- coding: UTF-8 -*-
import os
import requests as req
import xlsxwriter
import json,sys,time,random

reload(sys)
sys.setdefaultencoding('utf-8')
<<<<<<< HEAD
emailaddress=os.getenv('EMAIL')
=======
path=sys.path[0]+r'/AutoApi.xlsx'
email_address=os.getenv('EMAIL')
>>>>>>> 1136cea6c987a690104f5001bdff62a760e36800
app_num=os.getenv('APP_NUM')
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
    
#apiPost函数
def apiPost(a,data,url):
    access_token=access_token_list[a-1]
    headers={
            'Authorization': 'bearer ' + access_token,
            'Content-Type': 'application/json'
            }
<<<<<<< HEAD
    posttext=req.post(url,headers=headers,data=data)
    if posttext.status_code < 300:
        print('    操作成功')
    else:
        print('    操作失败')
    return posttext

#apiDelete函数
def apiDelete(a,url):
=======
    if req.put(r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/'+filesname+r':/content',headers=headers,data=f).status_code < 300:
        print('文件上传onedrive成功\n')
    else:
        print('文件上传onedrive失败\n')
        
# 发送天气邮件到自定义邮箱
def SendEmail(a,content,email_address):
>>>>>>> 1136cea6c987a690104f5001bdff62a760e36800
    access_token=access_token_list[a-1]
    headers={
            'Authorization': 'bearer ' + access_token,
            'Content-Type': 'application/json'
            }
    if req.delete(url,headers=headers).status_code < 300:
        print('    操作成功')
    else:
        print('    操作失败')

#上传文件到onedrive(小于4M)
def UploadFile(a,filesname,files):
    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/app'+str(a)+r'/'+filesname+r':/content'
    data = files
    apiPost(a,data,url)
        
# 发送邮件到自定义邮箱
def SendEmail(a,subject,content):
    url=r'https://graph.microsoft.com/v1.0/me/sendMail'
    mailmessage={'message': {'subject': subject,
                             'body': {'contentType': 'Text', 'content': content},
                             'toRecipients': [{'emailAddress': {'address': email_address}}],
                             },
                 'saveToSentItems': 'true'}            
    apiPost(a,json.dumps(mailmessage),url)
		
#修改excel(这函数分离好像意义不大)
#api-获取itemid: https://graph.microsoft.com/v1.0/me/drive/root/search(q='.xlsx')?select=name,id,webUrl
def excelWrite(a,filesname,sheet):
    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/app'+str(a)+r'/'+filesname+r':/workbook/worksheets/add'
    data={
         'name': sheet
         }
    print('  添加工作表')
    apiPost(a,data,url)
    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/app'+str(a)+r'/'+filesname+r':/workbook/'+sheet+r'/tables/add'
    data={
         "address": "A1:D8",
         "hasHeaders": false
         }
    print('  添加表格')
    apiPost(a,data,url)

def taskWrite(a,taskname):
    url=r'https://graph.microsoft.com/v1.0/me/todo/lists'
    data={
         "displayName": taskname
         }
    listjson=json.loads(apiPost(a,data,url))
    url=r'https://graph.microsoft.com/v1.0/me/todo/lists/'+listjson['id']+r'/tasks'
    data={
         "title": taskname,
         }
    taskjson=json.loads(apiPost(a,data,url))
    url=r'https://graph.microsoft.com/v1.0/me/todo/lists/'+listjson['id']+r'/tasks/'+taskjson['id']
    apiDelete(a,url)
    url=r'https://graph.microsoft.com/v1.0/me/todo/lists'+listjson['id']
    apiDelete(a,url)
    
    
def teamWrite(a,channelname):
    access_token=access_token_list[a-1]
    headers={
            'Authorization': 'bearer ' + access_token,
            'Content-Type': 'application/json'
            }
    html=req.get('https://graph.microsoft.com/v1.0/me/joinedTeams',headers=headers)
    jsontxt = json.loads(html.text)
    objectlist=jsontxt['value']
    #创建
    print("  创建team频道")
    data={
         "displayName": channelname,
         "description": "This channel is where we debate all future architecture plans",
         "membershipType": "standard"
         }
    url=r'https://graph.microsoft.com/v1.0/teams/'+objectlist[0]['id']+r'/channels'
    jsontxt = json.loads(apiPost(a,data,url))
    url=r'https://graph.microsoft.com/v1.0/teams/'+objectlist[0]['id']+r'/channels/'+jsontxt['id']
    print("  删除team频道")
    apiDelete(a,url) 
     
#一次性获取access_token，降低获取率
for a in range(1, int(app_num)+1):
    client_id=os.getenv('CLIENT_ID_'+str(a))
    client_secret=os.getenv('CLIENT_SECRET_'+str(a))
    ms_token=os.getenv('MS_TOKEN_'+str(a))
    access_token_list[a-1]=getmstoken(ms_token,a)

<<<<<<< HEAD
#获取天气充当邮件内容
=======
#获取天气(待优化)
>>>>>>> 1136cea6c987a690104f5001bdff62a760e36800
headers={'Accept-Language': 'zh-CN'}
weather=req.get(r'http://wttr.in/'+city+r'?format=4&?m',headers=headers).text

#实际运行   
for a in range(1, int(app_num)+1):
<<<<<<< HEAD
    #生成随机名称
    filesname='QAQ'+str(random.randint(1,600))+r'.xlsx'
    #新建随机xlsx文件
    xls = xlsxwriter.Workbook(filesname)
    xlssheet = workbook.add_worksheet()
    xlssheet.write(0,0,str(random.randint(1,600)))
    xlssheet.write(0,1,str(random.randint(1,600)))
    xlssheet.write(1,0,str(random.randint(1,600)))
    xlssheet.write(1,1,str(random.randint(1,600)))
    xls.close()
    xlspath=sys.path[0]+r'/'+filesname
    print('可能会偶尔出现创建上传失败的情况\n'+'上传随机文件到onedrive')
    with open(xlspath, 'rb') as f:
        UploadFile(a,filesname,f)
    print('发送邮件')
    if emailaddress != '':
        SendEmail(a,'weather',weather)
    #excel,team,task3选1执行
    choosenum = random.randint(1,3) 
#    if choosenum == 1 :
        print('excel文件操作')
        excelWrite(a,filesname,'QVQ'+str(random.randint(1,600))
#    elif choosenum == 2 :
        print('team频道操作')
        teamWrite(a,'QVQ'+str(random.randint(1,600))
#    elif choosenum == 3 :
        print('tasks操作')
        taskWrite(a,'QVQ'+str(random.randint(1,600))
=======
    print('上传xlsx文件')
    #上传excel文件是为了能运行excel的api
    with open(path, 'rb') as f:
        UploadFile(a,'AutoApi.xlsx',f)
    print('上传随机txt文件')
    UploadFile(a,'log.txt',weather)
    if email_address != '':
        SendEmail(a,weather,email_address)
>>>>>>> 1136cea6c987a690104f5001bdff62a760e36800
    
