<!--#include file="security.asp"-->
<!--#include file="inc/articlechar.inc"-->
<!--#include file="dbConn.asp"-->
<%
if request.form("title")="" then 
response.write"<script>alert('表单未填完整，请返回重填！');history.back();</Script>"
response.end
end if
%>
<%
if request.form("html")="on" then
content=request.form("content")
else
content=htmlencode2(request.form("content"))
end if

 
articleid=request("id")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from albums where articleid="&articleid
rs.open sql,conn,3,3

rs("typeid")=Request.Form("typeid")
rs("orgTime")=request.form("QYear")&"年"&request.form("QMonth")&"月"
rs("title")=request.form("title")
rs("content")=content
rs("images")=request.form("images")
rs("images1")=request.form("images1")
rs("images2")=request.form("images2")
rs("images3")=request.form("images3")
rs("images4")=request.form("images4")
rs("images5")=request.form("images5")
rs("images6")=request.form("images6")
rs("images7")=request.form("images7")
rs("images8")=request.form("images8")
rs("images9")=request.form("images9")
rs("images10")=request.form("images10")
rs("images11")=request.form("images11")
rs("images12")=request.form("images12")
rs("images13")=request.form("images13")
rs("images14")=request.form("images14")
rs("images15")=request.form("images15")
rs("images16")=request.form("images16")
rs("images17")=request.form("images17")
rs("images18")=request.form("images18")
rs("images19")=request.form("images19")
rs("images20")=request.form("images20")
rs("best")=request("best")
rs("bestpic")=request.form("minPic")
rs("dateandtime")=now()
rs("author")=request.form("author")
rs("source")=request.form("source")
rs("lx")=request.form("lx")
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing

response.redirect "managePhoto.asp"
%>


