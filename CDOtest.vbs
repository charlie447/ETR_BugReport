Const Email_From = "etrmailserver@gmail.com"'"etrtestsmtp@163.com" '发件人邮箱 
Const Password = "Welcome2Pwc" '发件人邮箱密码 
Const Email_To = "412706126@qq.com" '收件人邮箱 

Set CDO = CreateObject("CDO.Message") '创建CDO.Message对象 
CDO.Subject = "From Demon" '邮件主题 
CDO.From = Email_From '发件人地址 
CDO.To = Email_To '收件人地址 
CDO.TextBody = "Hello world!" '邮件正文 
'CDO.AddAttachment = "C:\hello.txt" '邮件附件文件路径 
Const schema = "http://schemas.microsoft.com/cdo/configuration/" '规定必须是这个，我也不知道为什么 

With CDO.Configuration.Fields '用with关键字减少代码输入 
.Item(schema & "sendusing") = 2 '使用网络上的SMTP服务器而不是本地的SMTP服务器 
.Item(schema & "smtpserver") = "smtp.gmail.com" 'SMTP服务器地址 
.Item(schema & "smtpauthenticate") = 1 '服务器认证方式 
.Item(schema & "sendusername") = Email_From '发件人邮箱 
.Item(schema & "sendpassword") = Password '发件人邮箱密码 
.Item(schema & "smtpserverport") = 25 'SMTP服务器端口 
.Item(schema & "smtpusessl") = 1 '是否使用SSL 
.Item(schema & "smtpconnectiontimeout") = 60 '连接服务器的超时时间 
.Update '更新设置 
End With 

CDO.Send '发送邮件 

set CDO = Nothing

MsgBox "Mail Sent!"