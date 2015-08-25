<!--#include file="../include/head.asp"-->
<%
'** 文件说明：用户管理/仅在用户信息浏览界面的删除

If User_Group<>"admin" Then
	TurnTo "../"&strUrl&""
End If

Dim UserName
Act=Request.QueryString("act")
UserName=Request.QueryString("id")
Select Case Act
	Case "delete"

		If username="admin" then
		Call ShowError("不能删除管理员！",0)
		End If

		Sql="delete * from Content where poster='"&SqlShow(UserName)&"'"
		OpenRs Sql
		Sql="delete * from Users where username='" & SqlShow(UserName) & "'"
		OpenRs Sql
		Sql="update Content set lastupdateuser='已删除' where lastupdateuser='" & SqlShow(UserName) & "'"
		OpenRs Sql
		Response.write "<script language='javascript'>window.close();</script>"
	Case "deleteanonymous"
		Sql="delete * from Content where poster='"&SqlShow(UserName)&"'"
		OpenRs Sql
		Sql="update Content set lastupdateuser='已删除' where lastupdateuser='" & SqlShow(UserName) & "'"
		OpenRs Sql
		Response.write "<script language='javascript'>window.close();</script>"
End Select
CloseAll
%>