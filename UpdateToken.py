# -*- coding: UTF-8 -*-
import requests as req
import json
import os
from base64 import b64encode
from nacl import encoding, public

gh_token=os.getenv('GH_TOKEN')
gh_repo=os.getenv('GH_REPO')
#账号信息生成
accountkey=['client_id','client_secret','ms_token']
#account存在？
if os.getenv('ACCOUNT')!='':
    account=json.loads(os.getenv('ACCOUNT'))
else:
    account={'client_id':[],'client_secret':[],'ms_token':[]}
account_add=os.getenv('ACCOUNT_ADD').split(",")
account_del=os.getenv('ACCOUNT_DEL')
#更新？
if account_del != '' or account_add != [''] :
    print('<<<<<<<<<<<<<<<配置信息更新中>>>>>>>>>>>>>>>')
#删除？
if account_del != '':
    print('删除账号中')
    for i in range(3):
        del account[accountkey[i]][int(account_del)-1]
#增加？
if account_add != ['']:
    print('增加账号中')
    for i in range(3):
        account[accountkey[i]].append(account_add[i])
#自定义url?
redirect_uri=os.getenv('REDIRECT_URI')
if redirect_uri =='':
    redirect_uri = r'https://login.microsoftonline.com/common/oauth2/nativeclient'  
#其他配置生成   
other_config=os.getenv('OTHER_CONFIG')
if os.getenv('OTHER_CONFIG') =='':
    other_config={'email':[],'tg_bot':[]}
else:
    other_config=json.loads(os.getenv('OTHER_CONFIG'))
if os.getenv('EMAIL') != '':
    print('更新邮箱')
    other_config['email']=[]
    for i in range(2):    
        other_config['email'].append(os.getenv('EMAIL').split(',')[i])
if os.getenv('TG_BOT') != '' and os.getenv('TG_BOT') != ' ':
    print('更新TG推送')
    other_config['tg_bot']=[]
    for i in range(2):    
        other_config['tg_bot'].append(os.getenv('TG_BOT').split(',')[i])
#删除TG推送
if os.getenv('TG_BOT') == ' ':
    other_config['tg_bot']=[]
    print('删除TG推送')
gh_url=r'https://api.github.com/repos/'+gh_repo+r'/actions/secrets/'
key_id='wangziyingwen'
print('<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>')
#微软refresh_token获取
def getmstoken(appnum):
    #try:except?
    ms_headers={
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
        html = req.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',data=data,headers=ms_headers)
        #json.dumps失败
        if html.status_code < 300:
            print(r'账号/应用 '+str(appnum+1)+' 的微软密钥获取成功')
            break
        else:
            if retry_ == 3:
                print(r'账号/应用 '+str(appnum+1)+' 的微软密钥获取失败'+'\n'+'请检查secret里 account_ID , account_SECRET , MS_TOKEN 格式与内容是否正确，然后重新设置')
    jsontxt = json.loads(html.text)
    refresh_token = jsontxt['refresh_token']
    return refresh_token

gh_headers={
           'Accept': 'application/vnd.github.v3+json',
           'Authorization': r'token '+gh_token,
           }
#公钥获取
def getpublickey(url_name):
    for retry_ in range(4):
        html = req.get(gh_url+url_name,headers=gh_headers)
        if html.status_code < 300:
            print("公钥获取成功")
            break
        else:
            if retry_ == 3:
                print("公钥获取失败，请检查secret里 GH_TOKEN 格式与设置是否正确")
    jsontxt = json.loads(html.text)
    global key_id 
    key_id = jsontxt['key_id']
    return jsontxt['key']

#token加密
def createsecret(secret_value,public_key):
    public_key = public.PublicKey(public_key.encode("utf-8"), encoding.Base64Encoder())
    sealed_box = public.SealedBox(public_key)
    encrypted = sealed_box.encrypt(secret_value.encode("utf-8"))
    return b64encode(encrypted).decode("utf-8")

#token上传
def setsecret(url_name,encrypted_value):
    data={
         'encrypted_value': encrypted_value,
         'key_id': key_id
         }
    for retry_ in range(4):
        putstatus=req.put(gh_url+url_name,headers=gh_headers,data=json.dumps(data))
        if putstatus.status_code < 300:
            print(r'账号配置更新成功')
            break
        else:
            if retry_ == 3:
                print(r'账号配置更新失败，请检查secret里 GH_TOKEN 格式与设置是否正确')        

#secret删除
def deletesecret(url_name):
    for retry_ in range(4):
        putstatus=req.delete(gh_url+url_name,headers=gh_headers)
        if putstatus.status_code < 300:
            print('--')
            break
 
#调用 
gh_public_key=getpublickey('public-key')
for a in range(0,len(account['client_id'])):
    client_id=account['client_id'][a]
    client_secret=account['client_secret'][a]
    ms_token=account['ms_token'][a]
    account['ms_token'][a]=getmstoken(a)
setsecret('ACCOUNT',createsecret(json.dumps(account),gh_public_key))
if os.getenv('EMAIL') != '' or os.getenv('TG_BOT') != '':
    setsecret('OTHER_CONFIG',createsecret(json.dumps(other_config),gh_public_key))
 
deletesecret('EMAIL')
deletesecret('TG_BOT')
deletesecret('ACCOUNT_ADD')
deletesecret('ACCOUNT_DEL')
