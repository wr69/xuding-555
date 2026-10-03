

## 说明 ##
* 自动续期程序，但是**不保证续期**
* 设置了**周六日(UTC时间)不启动**自动调用，周1-5每6小时自动启动一次 （修改看教程）
* 调用api保活：
     * 查询系api：onedrive,outkook,notebook,site等
     * 创建系api: 自动发送邮件，上传文件，修改excel等

## 说明 ##
依次点击页面上栏右边的 Setting -> 左栏 Secrets -> 右上 New repository secret，新建6个secret： GH_TOKEN、MS_TOKEN、CLIENT_ID、CLIENT_SECRET、CITY、EMAIL
* GH_TOKEN
* github密钥 (第三步获得)，例如获得的密钥是abc...xyz，则在secret页面直接粘贴进去，不用做任何修改，只需保证前后没有空格空行
* 
* MS_TOKEN
* 微软密钥（第二步获得的refresh_token）
* 
* CLIENT_ID
* 应用程序ID (第一步获得)
* 
* CLIENT_SECRET
* 应用程序密码 (第一步获得)
* 
* CITY
* 城市 (例如Beijing,自动发送天气邮件要用到)
* 
* EMAIL
* 收件邮箱 (自动发送天气邮件要用到)
* 
## 说明 ##
       env: 
       
        #github的账号信息
        GH_TOKEN: ${{ secrets.GH_TOKEN }} 
        GH_REPO: ${{ github.repository }}
        
        #以下是微软的账号信息(修改以下，复制增加账号)
        APP_NUM: ${{ secrets.APP_NUM }} 
        #账号/应用1
        MS_TOKEN_1: ${{ secrets.MS_TOKEN }} 
        CLIENT_ID_1: ${{ secrets.CLIENT_ID }}
        CLIENT_SECRET_1: ${{ secrets.CLIENT_SECRET }}
        #账号/应用2
        MS_TOKEN_2: ${{ secrets.MS_TOKEN_2 }} 
        CLIENT_ID_2: ${{ secrets.CLIENT_ID_2 }}
        CLIENT_SECRET_2: ${{ secrets.CLIENT_SECRET_2 }}
        #账号/应用3
        MS_TOKEN_3: ${{ secrets.MS_TOKEN_3 }} 
        CLIENT_ID_3: ${{ secrets.CLIENT_ID_3 }}
        CLIENT_SECRET_3: ${{ secrets.CLIENT_SECRET_3 }}
        #账号/应用4
        MS_TOKEN_4: ${{ secrets.MS_TOKEN_4 }} 
        CLIENT_ID_4: ${{ secrets.CLIENT_ID_4 }}
        CLIENT_SECRET_4: ${{ secrets.CLIENT_SECRET_4 }}
        #账号/应用5
        MS_TOKEN_5: ${{ secrets.MS_TOKEN_5 }} 
        CLIENT_ID_5: ${{ secrets.CLIENT_ID_5 }}
        CLIENT_SECRET_5: ${{ secrets.CLIENT_SECRET_5 }}
        #如此类推，自己复制增加账号/应用
        MS_TOKEN_n: ${{ secrets.MS_TOKEN_n }} 
        CLIENT_ID_n: ${{ secrets.CLIENT_ID_n }}
        CLIENT_SECRET_n: ${{ secrets.CLIENT_SECRET_n }}
       
 * 教程地址: https://www.vvhan.com/officeE5-AutoApi.html

