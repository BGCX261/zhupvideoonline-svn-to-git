<!-- #include file="inc/Conn.asp" -->
<%
if request.form("save")="ok" then 
prodid=trim(request("prodid"))
name=trim(request.form("name"))
mail=trim(request.form("mail"))
face=request.form("face")
nr=request.form("nr")
jibie=request.form("jibie")


	set rsP=Server.CreateObject("ADODB.RecordSet")
	sqlP="select * from buyok_pinglun"
	rsP.open sqlP,conn,1,3

			rsP.Addnew
			rsP("name")=name
			rsP("prodid")=prodid
			rsP("mail")=mail
			rsP("face")=face
			rsP("nr")=nr
			rsP("jibie")=jibie
			rsP("IP")=Request.serverVariables("REMOTE_ADDR")
			rsP.Update
		rsP.close
		set rsP=nothing
	set rsV=Server.CreateObject("ADODB.RecordSet")
	sqlV="select * from Product where ArticleID=" & prodid & ""
	rsV.open sqlV,conn,1,3
			rsV("Votes")=rsV("Votes")+1
			rsV.Update
		rsV.close
		set rsV=nothing
	response.write "<script language='javascript'>"
	if shenhe=False then
	response.write "alert('评论提交成功，评论须经管理员审核才能发布。');"
	else
	response.write "alert('评论提交成功，感谢您的参与！');"
	end if
	
set rsmail=server.createobject("adodb.recordset") 
sqlmail="select * from mailset" 
rsmail.open sqlmail,conn,1,1

  '参数说明
  'Subject     : 邮件标题
  'MailAddress : 发件服务器的地址,如smtp.163.com
  'Email       : 收件人邮件地址
  'Sender      : 发件人姓名
  'Content     : 邮件内容
  'Fromer      : 发件人的邮件地址

  Sub SendAction(subject, email, sender, content) 
	Set JMail = Server.CreateObject("JMail.Message") 
	JMail.Charset = "utf-8" ' 邮件字符集，默认为"US-ASCII"
	JMail.From = strMailUser ' 发送者地址
	JMail.FromName = sender' 发送者姓名
	JMail.Subject =subject
	JMail.MailServerUserName = strMailUser' 身份验证的用户名
	JMail.MailServerPassword = strMailPass ' 身份验证的密码
	JMail.Priority = 3
	JMail.AddRecipient(email)
	JMail.Body = content
	JMail.Send(strMailAddress)
  End Sub
  
  '调用此Sub的例子
  Dim strSubject,strEmail,strMailAdress,strSender,strContent,strFromer
  strSubject     = "来自城职院学生视频网站:评论审核" & Request("name")
  strContent     = "详细内容:" & Request("nr")
  strSender	 = Request("name")
  strEmail       = rsmail("mailincept")  '这是收信的地址,可以改为其它的邮箱
  strMailAddress = rsmail("mailsrv") '我司企业邮局地址，请使用 mail.您的域名
  strMailUser	 = rsmail("mailuser") '我司企业邮局用户名
  strMailPass	 = rsmail("mailpwd") '邮局用户密码

  Call SendAction (strSubject,strEmail,strSender,strContent)
  rsmail.close
	
	response.write "location.href='ArticleShow.asp?ArticleID="&prodid&"';"
	response.write "</script>"
	response.end
end if
%>