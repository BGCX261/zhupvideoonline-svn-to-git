<%
'发送邮件类
class sentmail
  dim outSmtp,outUser,outPsd,recUser,recSubmit,bodyContent,AddAttachment,ifsend
  '邮件函数执行的方法和参数
  function init_mail(str1,str2,str3,str4,str5,str6,str7)
	outSmtp=str1
	outUser=str2
	outPsd=str3
	recUser=str4
	recSubmit=str5
	bodyContent=str6
	AddAttachment=str7
	if isnull(outSmtp)or isnull(outUser)or isnull(outPsd) or isnull(recUser) or isnull(recSubmit) or isnull(bodyContent)  then
	 ifsend=false
	end if
  end function
  function send()
	if not ifsend then
	Const cdoSendUsingMethod="http://schemas.microsoft.com/cdo/configuration/sendusing"
	Const cdoSendUsingPort=2
	Const cdoSMTPServer="http://schemas.microsoft.com/cdo/configuration/smtpserver"
	Const cdoSMTPServerPort="http://schemas.microsoft.com/cdo/configuration/smtpserverport"
	Const cdoSMTPConnectionTimeout="http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout"
	Const cdoSMTPAuthenticate="http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"
	Const cdoBasic=1
	Const cdoSendUserName="http://schemas.microsoft.com/cdo/configuration/sendusername"
	Const cdoSendPassword="http://schemas.microsoft.com/cdo/configuration/sendpassword"
	
	Dim objConfig ' As CDO.Configuration
	Dim objMessage ' As CDO.Message
	Dim Fields ' As ADODB.Fields
	
	Set objConfig = Server.CreateObject("CDO.Configuration")
	Set Fields = objConfig.Fields
	
	With Fields
	.Item(cdoSendUsingMethod) = cdoSendUsingPort
	.Item(cdoSMTPServer) =outSmtp
	.Item(cdoSMTPServerPort) = 25
	.Item(cdoSMTPConnectionTimeout) = 10
	.Item(cdoSMTPAuthenticate) = cdoBasic
	.Item(cdoSendUserName) = outUser
	.Item(cdoSendPassword) = outPsd
	.Update
	End With
	
	Set objMessage = Server.CreateObject("CDO.Message")
	Set objMessage.Configuration = objConfig
	On Error Resume Next
	 With objMessage
	 .To =recUser
	 .From =outUser
	 .Subject = recSubmit
	 '.TextBody = bodyContent
	 .HTMLBody=bodyContent
	 '.AddAttachment "C:\Scripts\Output.txt"'邮件附件
	 .Send
	 End With
	On Error GoTo 0
	
	if err.description<>""then
	response.Write "aaa"
	end if
	 
	Set Fields = Nothing
	Set objMessage = Nothing
	Set objConfig = Nothing
	end if
  end function
end class
%>