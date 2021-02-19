# -*- coding: UTF-8 -*-
import requests as req
import json
import os
from base64 import b64encode
from nacl import encoding, public

configkey=['client_id','client_secret','ms_token']
#config存在？
if os.getenv('CONFIG')!='':
    config=json.loads(os.getenv('CONFIG'))
else:
    config={'client_id':[],'client_secret':[],'ms_token':[]}
config_add=os.getenv('CONFIG_ADD').split(",")
def is_int(s):
    try:
        int(s)
        return True
    except:
        return False
#更新？
if config_add != ['']:
    #替换or增加？
    if is_int(config_add[0]):
        if is_int(config_add[1]):
            config['ms_token'][int(config_add[0])-1]=config_add[3]
        else:
            for i in range(3):
                config[configkey[i]][int(config_add[0])-1]=config_add[i+1]
    else:
        for i in range(3):
            config[configkey[i]].append(config_add[i])
#自定义url?
redirect_uri=os.getenv('REDIRECT_URI')
if redirect_uri =='':
    redirect_uri = r'https://login.microsoftonline.com/common/oauth2/nativeclient'
gh_token=os.getenv('GH_TOKEN')
gh_repo=os.getenv('GH_REPO')
Auth=r'token '+gh_token
geturl=r'https://api.github.com/repos/'+gh_repo+r'/actions/secrets/public-key'
puturl=r'https://api.github.com/repos/'+gh_repo+r'/actions/secrets/CONFIG'
deleteurl=r'https://api.github.com/repos/'+gh_repo+r'/actions/secrets/CONFIG_ADD'
key_id='wangziyingwen'

#公钥获取
def getpublickey():
    headers={
            'Accept': 'application/vnd.github.v3+json','Authorization': Auth
            }
    for retry_ in range(4):
        html = req.get(geturl,headers=headers)
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
                print(r'账号/应用 '+str(appnum+1)+' 的微软密钥获取失败'+'\n'+'请检查secret里 CLIENT_ID , CLIENT_SECRET , MS_TOKEN 格式与内容是否正确，然后重新设置')
    jsontxt = json.loads(html.text)
    refresh_token = jsontxt['refresh_token']
    return refresh_token


#token加密
def createsecret(secret_value,public_key):
    public_key = public.PublicKey(public_key.encode("utf-8"), encoding.Base64Encoder())
    sealed_box = public.SealedBox(public_key)
    encrypted = sealed_box.encrypt(secret_value.encode("utf-8"))
    return b64encode(encrypted).decode("utf-8")

gs_headers={
           'Accept': 'application/vnd.github.v3+json',
           'Authorization': Auth
           }
#token上传
def setsecret(encrypted_value):
    data={
         'encrypted_value': encrypted_value,
         'key_id': key_id
         }
    #data_str=r'{"encrypted_value":"'+encrypted_value+r'",'+r'"key_id":"'+key_id+r'"}'
    for retry_ in range(4):
        putstatus=req.put(puturl,headers=gs_headers,data=json.dumps(data))
        if putstatus.status_code < 300:
            print(r'账号配置上传成功')
            break
        else:
            if retry_ == 3:
                print(r'账号配置上传失败，请检查secret里 GH_TOKEN 格式与设置是否正确')        

#config_add删除
def deletesecret():
    for retry_ in range(4):
        putstatus=req.delete(deleteurl,headers=gs_headers)
        if putstatus.status_code < 300:
            break
        else:
            if retry_ == 3:
                print(r'CONFIG_ADD删除失败') 
 
#调用 
for a in range(0,len(config['client_id'])):
    client_id=config['client_id'][a]
    client_secret=config['client_secret'][a]
    ms_token=config['ms_token'][a]
    config['ms_token'][a]=getmstoken(a)
setsecret(createsecret(json.dumps(config),getpublickey()))
if config_add != ['']:
    deletesecret()
