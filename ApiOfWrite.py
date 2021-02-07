# -*- coding: UTF-8 -*-
import os
import requests as req
import json,sys,time,random

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
        print(r'账号/应用 '+str(appnum)+' 的微软密钥获取失败'+'\n'+'请检查secret里 CLIENT_ID , CLIENT_SECRET , MS_TOKEN 格式与内容是否正确，然后重新设置')
    refresh_token = jsontxt['refresh_token']
    access_token = jsontxt['access_token']
    return access_token

#上传天气文件到onedrive
def UploadFile(a):
    localtime = time.asctime( time.localtime(time.time()) )
    access_token=access_token_list[a-1]
    headers={
            'Authorization': 'bearer ' + access_token,
            'Content-Type': 'application/json'
            }
    if req.put(r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/log.txt:/content',headers=headers,data=weather.text).status_code < 300:
        print('文件上传onedrive成功')
    else:
        print('文件上传onedrive失败')
        
# 发送天气邮件到自定义邮箱
def SendEmail(a):
    access_token=access_token_list[a-1]
    headers={
            'Authorization': 'bearer ' + access_token,
            'Content-Type': 'application/json'
            }
    ToRecipients= [r"{'EmailAddress': {'Address': wz.lxh@outlook.com}}"]
    mailmessage={'Message': {'Subject': 'Weather',
                             'Body': {'ContentType': 'Text', 'Content': weather.text},
                             'ToRecipients': ToRecipients,
                             }
                 'SaveToSentItems': 'true'}
    if req.post(r'https://graph.microsoft.com/v1.0/me/sendMail',headers=headers,data=mailmessage).status_code < 300:
        print('邮件发送成功')
    else:
        print('邮件发送失败')

#一次性获取access_token，降低获取率
for a in range(1, int(app_num)+1):
    client_id=os.getenv('CLIENT_ID_'+str(a))
    client_secret=os.getenv('CLIENT_SECRET_'+str(a))
    ms_token=os.getenv('MS_TOKEN_'+str(a))
    access_token_list[a-1]=getmstoken(ms_token,a)

#获取天气
weather=req.get(r'http://wttr.in/'+city)
        
for a in range(1, int(app_num)+1):
    UploadFile(a)
    SendEmail(a)
    
