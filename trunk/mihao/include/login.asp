<!--#include file="head.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="ofunctions.asp"-->
<%
'>>文件说明：登录验证文件<<'
Dim Action,Form_UserName,Form_Password
Action=Request("act")
Select Case Action
	'"&strUrl&"处登录
	Case "login"
		Form_UserName=Request("username")
		Form_Password=Request("userpass")
		Sql="select * from Users where UserName='" & SqlShow(Form_UserName) & "' and UserPass='" & MD5(Form_Password) & "'"
		OpenRs(Sql)
		If Rs.RecordCount>0 Then
			If Rs("lockuser")=TRUE Then
			Call ShowError("用户已被管理员锁定，不能登录！",1)
			End If
			Response.cookies(systemkey)("userid")=rs("username")
			Response.cookies(systemkey)("userpassword")=rs("userpass")
			Response.cookies(systemkey)("usergroup")=rs("usergroup")
			Response.cookies(systemkey)("uservip")=Cnum(Rs("vip"))
			Response.Cookies(SystemKey).Expires=Date+31
			Rs("LogonCount")=Rs("LogonCount")+1
			Rs("LastLogonTime")=Now
			Rs("Ip")=RemoteIP()
			Rs.update
			Action="OK"
		else
		Call ShowError("用户名或密码错误！",1)
		End If
	Case "logout"
		If Enable_Cookies = 1 Then
		Response.cookies(systemkey)("userid")=""
		Response.cookies(systemkey)("usergroup")=""
		Response.cookies(systemkey)("userpassword")=""
		Response.cookies(systemkey)("uservip")=""
		Session(SystemKey & "User_ID")=""
		Session(SystemKey & "User_Group")=""
		Session(SystemKey&"User_VIP")=""
		Else
		Session(SystemKey & "User_ID")=""
		Session(SystemKey & "User_Group")=""
		Session(SystemKey&"User_VIP")=""
		End If
		User_ID=""
		User_Group=""
		User_VIP=""
	Case ""
		TurnTo Request.ServerVariables("HTTP_REFERER")
End Select
If Action="OK" Then
TurnTo Request.ServerVariables("HTTP_REFERER")
Else
TurnTo Request.ServerVariables("HTTP_REFERER")
End If
CloseAll
%>