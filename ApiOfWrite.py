# -*- coding: UTF-8 -*-
import os
import xlsxwriter
import requests as req
import json,sys,time,random
from datetime import datetime
from base64 import b64encode
from nacl import encoding, public

gh_repo=os.getenv('GH_REPO')
gh_token=os.getenv('GH_TOKEN')
Auth=r'token '+gh_token
geturl=r'https://api.github.com/repos/'+gh_repo+r'/actions/secrets/public-key'
#puturl=r'https://api.github.com/repos/'+gh_repo+r'/actions/secrets/MS_TOKEN'
key_id='wangziyingwen'

#reload(sys)
#sys.setdefaultencoding('utf-8')
emailaddress=os.getenv('EMAIL')
email_day=os.getenv('EMAIL_DAY')

app_num=os.getenv('APP_NUM')
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
         'rounds': 2,
         'rounds_delay': [1,0,5],
         'api_delay': [1,0,5],
         'app_delay': [0,0,5],
         }        
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
        print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),r'账号/应用 '+str(appnum)+' 的微软密钥获取成功')
    else:
        print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),r'账号/应用 '+str(appnum)+' 的微软密钥获取失败\n'+'请检查secret里 CLIENT_ID , CLIENT_SECRET , MS_TOKEN 格式与内容是否正确，然后重新设置')
    refresh_token = jsontxt['refresh_token']
    access_token = jsontxt['access_token']
    return access_token


#api延时
def apiDelay():
    if config['api_delay'][0] == 1:
        time.sleep(random.randint(config['api_delay'][1],config['api_delay'][2]))
        
def apiReq(method,a,url,data='QAQ'):
    apiDelay()
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
        print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'        操作成功')
    else:
        print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'        操作失败')
#    if posttext.status_code > 300:
#        print('        操作失败')
#        #成功不提示
#    print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'##',posttext.text)	
    return posttext.text
          


# 上传email_day START
#公钥获取
def getpublickey(Auth,geturl):
    headers={'Accept': 'application/vnd.github.v3+json','Authorization': Auth}
    html = req.get(geturl,headers=headers)
    jsontxt = json.loads(html.text)
    if 'key' in jsontxt:
        print("公钥获取成功")
    else:
        print("公钥获取失败，请检查secret里 GH_TOKEN 格式与设置是否正确")
    public_key = jsontxt['key']
    global key_id 
    key_id = jsontxt['key_id']
    return public_key
#token加密
def createsecret(public_key,secret_value):
    public_key = public.PublicKey(public_key.encode("utf-8"), encoding.Base64Encoder())
    sealed_box = public.SealedBox(public_key)
    encrypted = sealed_box.encrypt(secret_value.encode("utf-8"))
    return b64encode(encrypted).decode("utf-8")

#token上传
def setsecret(encrypted_value,key_id,puturl):
    headers={'Accept': 'application/vnd.github.v3+json','Authorization': Auth}
    #data={'encrypted_value': encrypted_value,'key_id': key_id}  ->400error
    data_str=r'{"encrypted_value":"'+encrypted_value+r'",'+r'"key_id":"'+key_id+r'"}'
    putstatus=req.put(puturl,headers=headers,data=data_str)
    if putstatus.status_code >= 300:
        print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),' EMAIL_DAY上传失败，请检查secret里 EMAIL_DAY 格式与设置是否正确')
    else:
        print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),' EMAIL_DAY上传成功')
    return putstatus

#上传发送日期 
def UploadDay(day):
    puturl=r'https://api.github.com/repos/'+gh_repo+r'/actions/secrets/EMAIL_DAY'

    encrypted_value=createsecret(getpublickey(Auth,geturl),day)
    setsecret(encrypted_value,key_id,puturl)

# 上传email_day END

#上传文件到onedrive(小于4M)
def UploadFile(a,filesname,f):
    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/App'+str(a)+r'/'+filesname+r':/content'
    apiReq('put',a,url,f)
	
# fuck邮件
def FuckEmail(a):
    url=r'https://graph.microsoft.com/v1.0/me/messages?$select=sender,subject'         
    apiDelay()
    access_token=access_token_list[a-1]
    headers={
            'Authorization': 'bearer ' + access_token,
            'Content-Type': 'application/json'
            }
    posttext=req.get(url,headers=headers)
    if posttext.status_code < 300:
        #print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),posttext.text)
        print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'   fuck OK')
        try:
            mdata = posttext.json()["value"]  # 将响应的JSON数据转换为Python对象
            # 循环遍历JSON数据
            for item in mdata:
                value = item['subject']
                if value == "未送达: Undeliverable: weathertheapi" or value == "weathertheapi":
                    DeleteEmail(a,item['id'])
        except (KeyError, ValueError) as e:
            # 处理键或值的访问异常
            print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),"JSON数据解析异常:", e)
    else:
        print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'   fuck NO')

# 删除邮件
def DeleteEmail(a,mid):
    url=r'https://graph.microsoft.com/v1.0/me/messages/'+str(mid)      
    apiReq('delete',a,url)	
	
# 列出邮件
def ListEmail(a):
    url=r'https://graph.microsoft.com/v1.0/me/messages'         
    apiReq('get',a,url)	
	
# 列出邮箱的文件夹
def ListEmailFolders(a):
    url=r'https://graph.microsoft.com/v1.0/me/mailFolders'         
    apiReq('get',a,url)	
	
# 发送邮件到自定义邮箱
def SendEmail(a,subject,content):
    url=r'https://graph.microsoft.com/v1.0//me/sendMail'
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
    print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'    添加工作表')
    apiReq('post',a,url,json.dumps(data))
    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/App'+str(a)+r'/'+filesname+r':/workbook/worksheets/'+sheet+r'/tables/add'
    data={
         "address": "A1:D8",
         "hasHeaders": False
         }
    print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'    添加表格')
    jsontxt=json.loads(apiReq('post',a,url,json.dumps(data)))
    print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'    添加行')
    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/App'+str(a)+r'/'+filesname+r':/workbook/tables/'+jsontxt['id']+r'/rows/add'
    rowsvalues=[[0]*4]*2
    for v1 in range(0,2):
        for v2 in range(0,4):
            rowsvalues[v1][v2]=random.randint(1,1200)
    data={
         "values": rowsvalues
         }
    apiReq('post',a,url,json.dumps(data))
    
def taskWrite(a,taskname):
    url=r'https://graph.microsoft.com/v1.0/me/todo/lists'
    data={
         "displayName": taskname
         }
    print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),"    创建任务列表")
    listjson=json.loads(apiReq('post',a,url,json.dumps(data)))
    url=r'https://graph.microsoft.com/v1.0/me/todo/lists/'+listjson['id']+r'/tasks'
    data={
         "title": taskname,
         }
    print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),"    创建任务")
    taskjson=json.loads(apiReq('post',a,url,json.dumps(data)))
    url=r'https://graph.microsoft.com/v1.0/me/todo/lists/'+listjson['id']+r'/tasks/'+taskjson['id']
    print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),"    删除任务")
    apiReq('delete',a,url)
    url=r'https://graph.microsoft.com/v1.0/me/todo/lists/'+listjson['id']
    print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),"    删除任务列表")
    apiReq('delete',a,url)    
    
def teamWrite(a,channelname):
    url=r'https://graph.microsoft.com/v1.0/me/joinedTeams'
    print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),"    获取team")
    jsontxt = json.loads(apiReq('get',a,url))
    objectlist=jsontxt['value']
    tmid=objectlist[0]['id']
    #创建
    print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),"    创建team频道")
    data={
         "displayName": channelname,
         "description": "This channel is where we debate all future architecture plans",
         "membershipType": "standard"
         }
    url=r'https://graph.microsoft.com/v1.0/teams/'+tmid+r'/channels'
    jsontxt = json.loads(apiReq('post',a,url,json.dumps(data)))
    if 'id' in jsontxt:
        url=r'https://graph.microsoft.com/v1.0/teams/'+tmid+r'/channels/'+jsontxt['id']   
        print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),"    删除team频道")
        apiReq('delete',a,url)

def onenoteWrite(a,notename):
    url=r'https://graph.microsoft.com/v1.0/me/onenote/notebooks'
    data={
         "displayName": notename
         }
    print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'    创建笔记本')
    notetxt = json.loads(apiReq('post',a,url,json.dumps(data)))
    print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'    删除笔记本')
    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/Notebooks/'+notename
    apiReq('delete',a,url)
    
#一次性获取access_token，降低获取率
for a in range(1, int(app_num)+1):
    client_id=os.getenv('CLIENT_ID_'+str(a))
    client_secret=os.getenv('CLIENT_SECRET_'+str(a))
    ms_token=os.getenv('MS_TOKEN_'+str(a))
    access_token_list[a-1]=getmstoken(ms_token,a)
print('')    
#获取天气
headers={'Accept-Language': 'zh-CN'}
weather=datetime.now().strftime("%Y-%m-%d %H:%M:%S")+"beijing: ☀️   🌡️-1°C 🌬️↘26km/h"

#实际运行
if emailaddress != '':
    if str(email_day) != datetime.now().strftime("%Y%m%d"):
        print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'weather',weather)
        for a in range(1, int(app_num)+1):
            print('账号 '+str(a))
            print('发送邮件 ( 邮箱单独运行，每次运行只发送一次，防止封号 )')
            SendEmail(a,'weathertheapi',weather)
            print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),' 发送邮件')
        UploadDay(datetime.now().strftime("%Y%m%d"))
    else:
        print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),' 今天已发送过邮件')
        
#其他api
for _ in range(1,config['rounds']+1):
    if config['rounds_delay'][0] == 1:
        time.sleep(random.randint(config['rounds_delay'][1],config['rounds_delay'][2]))     
    print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'第 '+str(_)+' 轮\n')        
    for a in range(1, int(app_num)+1):
        if config['app_delay'][0] == 1:
            time.sleep(random.randint(config['app_delay'][1],config['app_delay'][2]))        
        print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'账号 '+str(a))    
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
        print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'上传文件 ( 可能会偶尔出现创建上传失败的情况 ) ')
        with open(xlspath,'rb') as f:
            UploadFile(a,filesname,f)
        choosenum = random.sample(range(1, 5),2)
        if config['allstart'] == 1 or 1 in choosenum:
            print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'excel文件操作')
            excelWrite(a,filesname,'QVQ'+str(random.randint(1,600)))
        if config['allstart'] == 1 or 2 in choosenum:
            print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'team操作')
            teamWrite(a,'QVQ'+str(random.randint(1,600)))
        if config['allstart'] == 1 or 3 in choosenum:
            print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'task操作')
            taskWrite(a,'QVQ'+str(random.randint(1,600)))
        if config['allstart'] == 1 or 4 in choosenum:
            print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'onenote操作')
            onenoteWrite(a,'QVQ'+str(random.randint(1,600)))
        print('-')

#实际运行
for a in range(1, int(app_num)+1):
    print('账号 '+str(a))
    print('删除发送失败的邮件')
    if emailaddress != '':
        FuckEmail(a)
	    
print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'执行完毕')
