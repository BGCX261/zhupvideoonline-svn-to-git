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
	response.write "alert('�����ύ�ɹ��������뾭����Ա��˲��ܷ�����');"
	else
	response.write "alert('�����ύ�ɹ�����л���Ĳ��룡');"
	end if
	
set rsmail=server.createobject("adodb.recordset") 
sqlmail="select * from mailset" 
rsmail.open sqlmail,conn,1,1

  '����˵��
  'Subject     : �ʼ�����
  'MailAddress : �����������ĵ�ַ,��smtp.163.com
  'Email       : �ռ����ʼ���ַ
  'Sender      : ����������
  'Content     : �ʼ�����
  'Fromer      : �����˵��ʼ���ַ

  Sub SendAction(subject, email, sender, content) 
	Set JMail = Server.CreateObject("JMail.Message") 
	JMail.Charset = "utf-8" ' �ʼ��ַ�����Ĭ��Ϊ"US-ASCII"
	JMail.From = strMailUser ' �����ߵ�ַ
	JMail.FromName = sender' ����������
	JMail.Subject =subject
	JMail.MailServerUserName = strMailUser' �����֤���û���
	JMail.MailServerPassword = strMailPass ' �����֤������
	JMail.Priority = 3
	JMail.AddRecipient(email)
	JMail.Body = content
	JMail.Send(strMailAddress)
  End Sub
  
  '���ô�Sub������
  Dim strSubject,strEmail,strMailAdress,strSender,strContent,strFromer
  strSubject     = "���Գ�ְԺѧ����Ƶ��վ:�������" & Request("name")
  strContent     = "��ϸ����:" & Request("nr")
  strSender	 = Request("name")
  strEmail       = rsmail("mailincept")  '�������ŵĵ�ַ,���Ը�Ϊ����������
  strMailAddress = rsmail("mailsrv") '��˾��ҵ�ʾֵ�ַ����ʹ�� mail.��������
  strMailUser	 = rsmail("mailuser") '��˾��ҵ�ʾ��û���
  strMailPass	 = rsmail("mailpwd") '�ʾ��û�����

  Call SendAction (strSubject,strEmail,strSender,strContent)
  rsmail.close
	
	response.write "location.href='ArticleShow.asp?ArticleID="&prodid&"';"
	response.write "</script>"
	response.end
end if
%>