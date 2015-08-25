<!--#include file="include/head.asp"-->
<!--#include file="include/ofunctions.asp"-->
<!--#include file="include/md5.asp"-->
<%
'** 文件说明：会员注册
Dim Form_UserName,Form_Password,Form_Email,Form_Userweb,Form_Sign,Form_userPicUrl,vl_code
Dim Msg
Act=Request.QueryString("act")
If Act<>"new" And Act<>"edit" Then
	If User_ID="" Then
		Act="new"
		If Enable_Register = 0 Then Call ShowError("目前暂停注册！",0)
	Else
		Act="edit"
		Call LoadUser
	End If
Else
	If Act="new" Then
		If Enable_Register = 0 Then Call ShowError("目前暂停注册！",0)
		Call SaveUser
	Else
		Call EditUser
	End If
End If

Sub LoadUser
	Sql="select * from Users where UserName='" & SqlShow(User_ID) & "'"
	OpenRs(Sql)
	Form_UserName=User_ID
	Form_Password="v8KWs5Uj"
	Form_Email=Rs("Email")
	Form_Userweb=Rs("Userweb")
	Form_Sign=Rs("Sign")
	Form_userPicUrl=Rs("userPicUrl")
	Rs.Close
End Sub

Sub EditUser
	Form_Password=Request.Form("userpass")
	Form_Email=Request.Form("email")
	Form_Userweb=Request.Form("userweb")
	Form_Sign=Request.Form("Sign")
	vl_code=trim(request.form("vl_code"))
	Form_userPicUrl=trim(request.form("userPicUrl"))
	If Form_userPicUrl="" Then Form_userPicUrl="images/face/user.gif"
	If isValidEmail(Form_Email)=False  Then
	    Msg="电子邮件地址错误"
	   Exit Sub
	End If
	If cstr(vl_code)="" or cstr(Session("GetCode"))<>cstr(vl_code) then
	Call ShowError("请输入正确的验证码！",1)
	Exit Sub
	End If

	Sql="select * from Users where UserName='" & SqlShow(User_ID) & "'"
	OpenRs(Sql)
	If Form_Password<>"v8KWs5Uj" Then
	Rs("UserPass")=MD5(Form_Password)
	Response.cookies(systemkey)("userpassword")=MD5(Form_Password)
	End If
	Rs("Email")=Form_Email
	Rs("Userweb")=Form_Userweb
	Rs("Sign")=Form_Sign
	Rs("userPicUrl")=Form_userPicUrl
	Rs.Update
	Rs.Close
	Act="Save"
	Call ShowError("修改成功！",0)
End Sub

Sub SaveUser
	vl_code=trim(request.form("vl_code"))
	Form_UserName=Request.Form("username")
	Form_Password=Request.Form("userpass")
	Form_Email=Request.Form("email")
	Form_Userweb=Request.Form("userweb")
	Form_Sign=Request.Form("Sign")
	Form_userPicUrl=trim(request.form("userPicUrl"))
	If Form_userPicUrl="" Then Form_userPicUrl="images/face/user.gif"
	If IsSpecial(Form_UserName)=True or CCEmpty(Form_UserName)<>Form_UserName or Form_UserName="" Then
		Msg="用户名中不可以包含特殊字符!"
		Exit Sub
	End If
	Sql="select * from Users where UserName='" & SqlShow(Form_UserName) & "'"
	OpenRs(Sql)
	If Rs.RecordCount>0 Then
		Rs.Close
		Msg="该用户名已经存在，请另外选择用户名"
		Exit Sub
	End If
	If isValidEmail(Form_Email)=False  Then
	        Rs.Close
	    Msg="电子邮件地址错误"
	    Exit Sub
	End If
	If strlen(Form_UserName)<3 Then
	        Rs.Close
		Msg="用户名不能小于3个字符"
		Exit Sub
	End If
	If cstr(vl_code)="" or cstr(Session("GetCode"))<>cstr(vl_code) then
	Msg="请输入正确的验证码！"
	Exit Sub
	End If
	Rs.AddNew
	Rs("UserName")=Form_UserName
	Rs("UserPass")=MD5(Form_Password)
	Rs("UserGroup")="member"
	Rs("Email")=Form_Email
	Rs("Userweb")=Form_Userweb
	Rs("Sign")=Form_Sign
	Rs("userPicUrl")=Form_userPicUrl
	Rs("RegTime")=Date
	Rs("vip")=0

	Rs.Update
	Rs.Close
	If Enable_Cookies = 1 Then
			Response.cookies(systemkey)("userid")=Form_UserName
			Response.cookies(systemkey)("userpassword")=MD5(Form_Password)
			Response.cookies(systemkey)("usergroup")="member"
       Else
		Session(SystemKey & "User_ID")=Form_UserName
		Session(SystemKey & "User_Group")=Rs("UserGroup")
    End If

	Msg="已成功注册，现在已经以注册用户身份登录"
	Act="Save"
End Sub
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%If Act="new" Then%>注册用户<%Else%>修改资料<%End If%></title>
<link rel="stylesheet" type="text/css" href="images/<%=strStyle%>">
<script language="javascript" src="include/js/edit.js"></script>
<script language="javascript">
function checkform()
{
	if(form1.username.value==""){
		form1.username.focus();
		return(false);
	}
	if(form1.userpass.value==""){
		form1.userpass.focus();
		return(false);
	}
	if(form1.passcheck.value==""){
		form1.passcheck.focus();
		return(false);
	}
	if(form1.passcheck.value!=form1.userpass.value){
		alert("两次输入的密码不一致!");
		form1.userpass.focus();
		return(false);
	}
}
<%
If Msg<>"" Then
	Printl "alert(""" & Msg & """);"
End If
If Act="Save" Then
	Act="edit"
%>
window.opener.document.location.reload();
window.close();
<%
End If
%>
</script>
<body >
<form name="form1" action="register.asp?act=<%=Act%>" method="post" onSubmit="javascript:return checkform();">

<div class="reg_">
	<div class="reg_1">用户资料</div>

	<div class="reg_2">
		<span class="red_color">姓名：</span><input <%if act="edit" then%>readonly <%end if%>maxlength="9" type="text" name="username" size="20" class="input" value="<%=HtmlEncode(Form_UserName)%>">
	</div>

	<div class="reg_2">
         <span class="red_color">密码：</span><input maxlength="20" type="password" name="userpass" size="20" class="input" value="<%=HtmlEncode(Form_Password)%>">
	</div>

	<div class="reg_2">
         <span class="red_color">确认密码：</span><input maxlength="20" type="password" name="passcheck" size="20" class="input" value="<%=HtmlEncode(Form_Password)%>">
	</div>

	<div class="reg_2">
        <span class="red_color"> Email：</span><input maxlength="48" type="text" name="email" size="28" class="input" value="<%=HtmlEncode(Form_Email)%>">
	</div>
	
<%
dim Form_userpicurl2
If Form_userpicurl<>"" Then Form_userpicurl2=Form_userpicurl Else Form_userpicurl2="images/face/user.gif"
%>
    <div style="display:none;">
	     <span>用户头像：</span><INPUT TYPE="text" size=42 NAME="userPicUrl"  id="userPicUrl" value="<%=Form_userPicUrl2%>" readonly>
	</div>

    <div class="display:none;">
		 <table title="选择头像" cellspacing=0 cellpadding=0 border=0 style="display:none;" onclick="javascript:open('INClude/Face.htm','','width=500,height=500,resizable,scrollbars')">
			  <tr>
			       <td>
			            <img src="<%=Form_userpicurl2%>" width=48 height=48 id="tus" name="tus"  border="0">
			       </td>
			  </tr>
	     </table>

	</div>

	<div class="reg_2">
	     <span class="red_color">验证码：</span>
		 <input name="vl_code" class="input" type="text" id="vl_code" size="8" maxlength="4" onKeyPress="if ((event.keyCode<48 || event.keyCode>57)) event.returnValue=false"><img src="include/GetCode.asp" alt= "如看不清楚验证码请点击刷新" style="cursor : pointer;" onClick="this.src='include/GetCode.asp'">
	</div>

    <div class="reg_4">
	  <input type="submit" value=" 确 定 " name="B1" onMouseOut="javascript:this.style.backgroundColor='';" class="reg_button">&nbsp;&nbsp;
	  <input type="reset" value=" 清 空 " name="B2" onMouseOut="javascript:this.style.backgroundColor='';" class="reg_button">
	</div>

</div>
</form>

</body>
</html>
<%CloseAll%>