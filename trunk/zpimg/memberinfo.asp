<!--#include file="include/head.asp"-->
<!--#include file="include/ofunctions.asp"--><%
'** �ļ�˵������ʾ�û�ע����Ϣ
Dim Form_UserName,Form_Password,Form_Sex,Form_Email,Form_PostNum,Form_RegTime,Form_logoncount,Form_logontime,Form_Userweb,Form_Userblog,Form_VIP,Form_Group,Form_Bonus,Form_Sign,Form_ReplyNum
Dim Msg
Form_UserName=Request.QueryString("id")
If Form_UserName="" Then
	Form_UserName=User_ID
End If
If Left(Form_UserName,1)="~"  Then
Call ShowError("���û�Ϊ�����û�������ϸ���ϣ�",0)
End If
if user_id="" then
Call ShowError("ע���Ա��¼����ܷ��ʣ�",0)
end if

LoadUser
Sub LoadUser
	Sql="select * from Users where UserName='"&SqlShow(Form_UserName)&"'"
	OpenRs(Sql)
	If Rs.RecordCount=0 Then Call ShowError("�޴��û�",0)
	Form_Password="88888888"
	Form_Email=Rs("Email")
	Form_Userweb=Rs("Userweb")
	Form_Sign=Rs("Sign")
	Form_PostNum=Rs("PostNum")
	Form_RegTime=Rs("RegTime")
	Form_LogonCount=Rs("LogonCount")
	Form_LogonTime=Rs("LastLogonTime")
	Form_Bonus=Rs("bonus")
	Form_ReplyNum=Rs("ReplyNum")
	If Rs("vip")=1 Then Form_VIP="&nbsp;&nbsp;[<font color=red>VIP</font>��Ա]"
	If Rs("usergroup")="member" and Form_VIP="" Then Form_Group="&nbsp;[��ͨ��Ա]"
	If Rs("usergroup")="admin" Then Form_Group="&nbsp;[����Ա]"
	If Cnum(Rs("usergroup"))>0 Then Form_Group="&nbsp;["&Application(systemkey&rs("usergroup"))&"-����]"
	Rs.Close
	If Left(Form_Userweb,4)<>"http" Then Form_Userweb="http://"&Trim(Form_Userweb)
	If Left(Form_Sign,4)<>"http" Then Form_Sign="http://"&Trim(Form_Sign)
End Sub
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=HtmlEncode(Form_UserName) & "�ĸ�������"%></title>
<link rel="stylesheet" type="text/css" href="images/<%=css_style%>">
<script language="javascript" src="include/js/edit.js"></script>
<script language="javascript">
if(window.opener==null){
	openpost(window.location);
	window.location="<%=strUrl%>"
}
</script>
</head>
<body>
<form name="form1" action="register.asp?act=<%=Act%>" method="post" onSubmit="javascript:return checkform();">
	<div class="reg_">
		 <div class="reg_1">��<%=HtmlEncode(Form_UserName)%>���ĸ������� </div>
		 <div class="reg_2"><span>��ݣ�</span><%=Form_VIP%><%=Form_Group%>&nbsp;&nbsp;������:<%=form_bonus%>��</div>
	     <div class="reg_2"><span>����������</span>&nbsp;&nbsp;<%=Form_PostNum%></div>
	     <div class="reg_2"><span>�ظ����ۣ�</span>&nbsp;&nbsp;<%=Form_ReplyNum%></div>
		 <div class="reg_2"><span>��¼������</span>&nbsp;&nbsp;<%=Form_LogonCount%></div>
		 <div class="reg_2"><span>ע��ʱ�䣺</span>&nbsp;&nbsp;<%=HtmlEncode(Form_RegTime)%></div>
	     <div class="reg_2"><span>����¼��</span>&nbsp;&nbsp;<%=Form_LogonTime%></div>
         <div class="reg_2"><span>�����ʼ���</span>&nbsp;&nbsp;<%=HtmlEncode(Form_Email)%></div>
	</div>
	
	<div class="reg_4">
		<input onClick="javascript:window.close();" type="button" value="�ر�" name="B1" class="reg_button">
	</div>
</form>
<%call closeall%>
</body>
</html>