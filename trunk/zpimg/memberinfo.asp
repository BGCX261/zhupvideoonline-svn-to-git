<!--#include file="include/head.asp"-->
<!--#include file="include/ofunctions.asp"--><%
'** 文件说明：显示用户注册信息
Dim Form_UserName,Form_Password,Form_Sex,Form_Email,Form_PostNum,Form_RegTime,Form_logoncount,Form_logontime,Form_Userweb,Form_Userblog,Form_VIP,Form_Group,Form_Bonus,Form_Sign,Form_ReplyNum
Dim Msg
Form_UserName=Request.QueryString("id")
If Form_UserName="" Then
	Form_UserName=User_ID
End If
If Left(Form_UserName,1)="~"  Then
Call ShowError("该用户为匿名用户，无详细资料！",0)
End If
if user_id="" then
Call ShowError("注册会员登录后才能访问！",0)
end if

LoadUser
Sub LoadUser
	Sql="select * from Users where UserName='"&SqlShow(Form_UserName)&"'"
	OpenRs(Sql)
	If Rs.RecordCount=0 Then Call ShowError("无此用户",0)
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
	If Rs("vip")=1 Then Form_VIP="&nbsp;&nbsp;[<font color=red>VIP</font>会员]"
	If Rs("usergroup")="member" and Form_VIP="" Then Form_Group="&nbsp;[普通会员]"
	If Rs("usergroup")="admin" Then Form_Group="&nbsp;[管理员]"
	If Cnum(Rs("usergroup"))>0 Then Form_Group="&nbsp;["&Application(systemkey&rs("usergroup"))&"-版主]"
	Rs.Close
	If Left(Form_Userweb,4)<>"http" Then Form_Userweb="http://"&Trim(Form_Userweb)
	If Left(Form_Sign,4)<>"http" Then Form_Sign="http://"&Trim(Form_Sign)
End Sub
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=HtmlEncode(Form_UserName) & "的个人资料"%></title>
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
		 <div class="reg_1">【<%=HtmlEncode(Form_UserName)%>】的个人资料 </div>
		 <div class="reg_2"><span>身份：</span><%=Form_VIP%><%=Form_Group%>&nbsp;&nbsp;【积分:<%=form_bonus%>】</div>
	     <div class="reg_2"><span>发表数量：</span>&nbsp;&nbsp;<%=Form_PostNum%></div>
	     <div class="reg_2"><span>回复评论：</span>&nbsp;&nbsp;<%=Form_ReplyNum%></div>
		 <div class="reg_2"><span>登录次数：</span>&nbsp;&nbsp;<%=Form_LogonCount%></div>
		 <div class="reg_2"><span>注册时间：</span>&nbsp;&nbsp;<%=HtmlEncode(Form_RegTime)%></div>
	     <div class="reg_2"><span>最后登录：</span>&nbsp;&nbsp;<%=Form_LogonTime%></div>
         <div class="reg_2"><span>电子邮件：</span>&nbsp;&nbsp;<%=HtmlEncode(Form_Email)%></div>
	</div>
	
	<div class="reg_4">
		<input onClick="javascript:window.close();" type="button" value="关闭" name="B1" class="reg_button">
	</div>
</form>
<%call closeall%>
</body>
</html>