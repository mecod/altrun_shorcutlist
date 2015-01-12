Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
Set Flds = iConf.Fields
schema = "http://schemas.microsoft.com/cdo/configuration/"
Flds.Item(schema & "sendusing") = 2
Flds.Item(schema & "smtpserver") = "smtp.126.com"
Flds.Item(schema & "smtpserverport") = 25
Flds.Item(schema & "smtpauthenticate") = 1
Flds.Item(schema & "sendusername") = "mecod@126.com"
Flds.Item(schema & "sendpassword") = "5f754cf159"
Flds.Item(schema & "smtpusessl") = 1
Flds.Update
With iMsg
.To = "wanggaolei.e1ad998@m.evernote.com"
.From = "mecod@126.com"
.Subject = wscript.arguments.item(0)
Set .Configuration = iConf
SendEmailGmail = .Send
End With

' 供altrun来向Evernote收件箱发送信息
' 使用方法：
' 将此ss.vbs右键发送到ALTRun中制成一个快捷项目，参数类型选择第二项参数无编码便可告成功。
' 另外也可以直接添加快捷项目命令行语句@cscript 路径\ss.vbs “{%p}”。