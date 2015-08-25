<!--#include file="bsconfig.asp"-->
<!--#include file="inc/md5.asp"-->

<%
set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_User where Id=" & request.querystring("uid")
rs.open sqltext,conn,3,3

'更新管理员密码到数据库
rs("adminName")=request.form("uid")
rs("PassWord")=md5(request.form("pwd1"))
rs.update
rs.close
conn.close
response.redirect "Bs_Admin.asp"
%>