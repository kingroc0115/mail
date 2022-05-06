# coding:utf-8
import smtplib
from email.header import Header
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import xml
from xml.dom import minidom


# 1、首先对robot进行规范的管理，这样方便对最终的xml进行解析，以下是一个robot的示例。
''
# 以上内容共划分了脚本名、测试线名、对于脚本的描述（或用例名）、用例集合名4项
# 可以根据实际情况调整，并对xml解析部分进行修改。

# 2、邮件的正文：
  # 其中dealwithXML()对整合后的xml文件的解析，参数为xml文件的绝对路径，返回值为处理结果。
    # groupList,脚本所属测试线列表
    # passList,脚本通过情况列表
    # tagPassList,脚本所属用例集合列表
    # idList,脚本id列表，此项由robot生成，所以暂时没用到
    # nameList,脚本名列表
    # docList,脚本名的一些描述，详细为[Documentation]
    # timeList,脚本的执行时间信息
    # errList，脚本的报错信息，由于各脚本的不同这个位置是最容易出错的，需要详细分析xml
    # xml的解析学习参考xml.dom.minidom的使用方法

  # AddRobotReportXML()调用了dealwithXML()，把处理结果生成为html语言并生成到邮件的正文。
    # 这一部分参考日报的结果，对需要调整的内容修改即可。
    # 丰富日报页面参考学习html的语法进行调试

# 3、邮件的发送
class mail:
    def __init__(self,smtpserver,username,password,htmlPath = ''):
        self.smtpserver = smtpserver
        self.username = username
        self.password = password
        if htmlPath != '':
            htmlText = ''
            with open(htmlPath,'r',encoding='utf-8') as f:
                htmlText = f.read()

            self.msg = MIMEText(htmlText,'html')
        else:
            self.msg = MIMEMultipart('mixed')

    def dealwithXML(self,fileName):
        tagList = []#用例所属测试集
        groupList = []#用例所属测试线
        passList = []#用例执行结果
        tagPassList = []#以用例集分的测试结果
        idList = []#id
        nameList = []#脚本名
        docList = []#用例名
        timeList = []
        errList = []

        dom = xml.dom.minidom.parse(fileName)
        rootdata = dom.documentElement
        itemlist = rootdata.getElementsByTagName('arg')
        for i in range(len(itemlist)):
            if i%2 == 0:
                if len(itemlist[i].childNodes)>0:
                    tagList.append(itemlist[i].childNodes[0].data)

        itemlist = rootdata.getElementsByTagName('test')
        for i in range(len(itemlist)):
            groupList.append(itemlist[i].getAttribute('name'))

        itemlist = rootdata.getElementsByTagName('statistics')[0].getElementsByTagName('suite')[0].getElementsByTagName('stat')
        for i in range(len(itemlist)):
            passList.append([itemlist[i].getAttribute('pass'),itemlist[i].getAttribute('fail')])
            idList.append(itemlist[i].getAttribute('id'))
            nameList.append(itemlist[i].getAttribute('name'))

        itemlist = rootdata.getElementsByTagName('statistics')[0].getElementsByTagName('tag')[0].getElementsByTagName('stat')
        for i in range(len(itemlist)):
            tagPassList.append([itemlist[i].getAttribute('pass'),itemlist[i].getAttribute('fail'),itemlist[i].childNodes[0].data])

        itemlist = rootdata.getElementsByTagName('doc')
        for i in range(len(itemlist)):
            if i%3 == 2:
                if len(itemlist[i].childNodes)>0:
                    docList.append(itemlist[i].childNodes[0].data)

        itemlist = rootdata.getElementsByTagName('suite')[0].getElementsByTagName('suite')
        for i in range(len(itemlist)):
            temp = itemlist[i].getElementsByTagName('test')[0].getElementsByTagName('status')[-1]
            timeList.append([temp.getAttribute('starttime'),temp.getAttribute('endtime')])
            if len(temp.childNodes)>0:
                if temp.childNodes[0].data.startswith('Evaluating expression'):
                    errList.append(temp.childNodes[0].data.split('Evaluating expression \'')[1].split('==True')[0])
                else:
                    errList.append(temp.childNodes[0].data.split('==True')[0][1:])
            else:
                errList.append('')

        return [groupList,passList,tagPassList,idList,nameList,docList,timeList,errList]

    def SetSenderRecever(self,sender,receiverList):
        self.sender = sender
        self.receiverList = receiverList
        self.msg['From'] = sender
        self.msg['To'] = ';'.join(receiverList)

    def setSubject(self,subjectMail):
        #邮件主题
        subject = Header(subjectMail,'utf-8').encode()
        self.msg['subject'] = subject

    def AddText(self,sendtext):
        #文字内容
        text_plain = MIMEText(sendtext,'plain','utf-8')
        self.msg.attach(text_plain)

    def AddImage(self,ImageFilePath):
        #图片链接
        sendImagineFile = open(ImageFilePath,'rb').read()
        image = MIMEImage(sendImagineFile)
        image.add_header('Content-ID','<image1>')
        image['Content-Disposition'] = 'attachment;filename="%s"'%ImageFilePath.split('\\')[-1]
        self.msg.attach(image)

    def AddFile(self,FilePath,fileNameShow):
        #附件
        senfFile = open(FilePath,'rb').read()
        text_att = MIMEText(senfFile,'base64','utf-8')
        text_att['Content-Type'] = 'application/octet-stream'
        text_att['Content-Disposition'] = 'attachment;filename="%s"'%fileNameShow
        self.msg.attach(text_att)

    def AddRobotReportXML(self,FilePath):
        result = self.dealwithXML(FilePath)
        mailMsg = '<table border="1"cellspacing="1"width="800"cellpadding="2"style="font-size:15px"><tr><th colspan="5">无线性能组自动化测试日报</th></tr>'
        mailMsg += '<tr><th>组</th><th>用例数</th><th>通过用例</th><th>失败用例</th><th>通过率</th></tr>'
        if int(result[1][0][0])+int(result[1][0][1])>0:
            totalNums = int(result[1][0][0])+int(result[1][0][1])
            passNums = result[1][0][0]
            failNums = result[1][0][1]
            resultNums= int(result[1][0][0])/(int(result[1][0][0])+int(result[1][0][1]))*100
            mailMsg += '<tr><td align=center>无线性能</td><td align=center>%s</td><td align=center>%s</td><td align=center>%s</td><td align=center>%.1f%%</td></tr>'%(totalNums,passNums,failNums,resultNums)
        else:
            mailMsg += '<tr><td align=center>无线性能</td><td align=center>0</td><td align=center>0</td><td align=center>0</td><td align=center>0%</td></tr>'

        mailMsg += '<tr><th align=center>用例集</th><th>用例数</th><th>通过用例</th><th>失败用例</th><th>通过率</th></tr>'
        for i in range(len(result[2])):#用例集
            if int(result[2][i][0])+int(result[2][i][1])>0:
                mailMsg += '<tr><td align=center>%s</td><td align=center>%s</td><td align=center>%s</td><td align=center>%s</td><td align=center>%.1f%%</td></tr>'%(result[2][i][2],int(result[2][i][0])+int(result[2][i][1]),result[2][i][0],result[2][i][1],(int(result[2][i][0])/(int(result[2][i][0])+int(result[2][i][1])))*100)

        mailMsg += '<tr><th align=center>测试线</th><th>用例数</th><th>通过用例</th><th>失败用例</th><th>通过率</th></tr>'
        collection = []
        for i in range(len(result[0])):
            if not result[0][i] in collection:
                passNum = 0
                collectNum = 0
                collection.append(result[0][i])
                for j in range(len(result[0])):
                    if result[0][j] == result[0][i]:
                        collectNum += 1
                        passNum += int(result[1][j][0])
                if collectNum>0:
                    mailMsg += '<tr><td align=center>%s</td><td align=center>%s</td><td align=center>%s</td><td align=center>%s</td><td align=center>%.1f%%</td></tr>'%(result[0][i],collectNum,passNum,collectNum-passNum,passNum/collectNum*100)

        mailMsg += '<tr><th align=center>用例描述</th><th>脚本名</th><th>是否通过</th><th>失败原因</th><th>测试线</th></tr>'
        for i in range(len(result[4])-1):
            mailMsg += '<tr><td align=center>%s</td><td>%s</td><td align=center>%s</td><td align=center>%s</td><td>%s</td></tr>'%(result[5][i],result[4][i+1],'是' if int(result[1][i+1][0]) == 1 else '否',result[7][i],result[0][i])

        mailMsg += '</table>'
        self.msg.attach(MIMEText(mailMsg,'html','utf-8'))

    def SendMailQuit(self):
        #发送邮件
        smtp = smtplib.SMTP()
        smtp.connect(self.smtpserver)
        smtp.login(self.username,self.password)
        # smtp.sendmail(self.sender,self.receiverList,self.msg.as_string())
        smtp.sendmail(self.sender, self.receiverList, self.msg.as_string())
        #退出
        smtp.quit()

if __name__ == "__main__":
    smtpserver = 'smtp.qq.com'
    username = '34356548'
    password = '1szhyhvhiodkqbggh'
    sender = '34356548@qq.com'
    receiverList = ['kingroc0115@hotmail.com', 'kingroc0115@163.com']
    subject = '增加pic和附件的测试邮件'
    imagePath = 'D:\\Py_script\\test\\a3.PNG'
    filePath = 'D:\\Py_script\\test\\3.5GNR\\20121220_3.5GNR.xlsx'
    fileNameShow = '20121220_3.5GNR.xlsx'
    addtext = 'TEST!!!'
    m = mail(smtpserver,username, password)
    m.SetSenderRecever(sender, receiverList)
    m.setSubject(subject)
    m.AddImage(imagePath)
    m.AddFile(filePath, fileNameShow)
    m.AddText(addtext)
    #m.AddRobotReportXML("E:\\AUTOTEST\\pycharm_project\\send_mail\\log.html")
    m.SendMailQuit()