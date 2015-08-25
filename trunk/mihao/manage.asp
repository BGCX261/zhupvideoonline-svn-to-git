<!--#include file="include/head.asp"-->
<!--#include file="include/ofunctions.asp"-->
<%If User_Group<>"admin" Then TurnTo "/"%>
<%
Dim ID,ContextID
Act=Request.QueryString("act")
ID=Request.QueryString("id")
ContextID=Request.QueryString("ContextID")

Select Case Act

'删除文章评论
Case "DeleteContext"
sql="delete * from Content where id="&id&""
OpenRs(sql)
Response.write("<script language='javascript'>alert('删除成功！');window.location='context.asp?ID="&ConTextID&"';</script>")

Case Else
End Select
%>
<%Call CloseAll()%>