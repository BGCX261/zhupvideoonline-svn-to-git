<!--#include file="bsconfig.asp"-->
<!--#include file="inc/md5.asp"-->

<%
password=replace(trim(Request("pwd1")),"'","")
password=md5(password)
set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_User where UserName='" & request.form("uid") & "' and RealName='" & request.form("realname") & "' and PassWord='" & password & "'"
rs.open sqltext,conn,1,1

'查找数据库，检查此管理员是否已经存在
if rs.recordcount >= 1 then
if rs("UserName")=request.form("uid") then
Response.Redirect "Loginsb.asp?msg=此管理员帐号已经存在，请选用其它名称!"
response.end
rs.close
end if
end if


set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_User"
rs.open sqltext,conn,3,3

'添加一个管理员帐号到数据库
rs.addnew
rs("UserName")=request.form("uid")
rs("RealName")=request.form("realname")
rs("purview")=request.form("purview")
rs("identity")=request.form("identity")
rs("PassWord")=password
rs.update
Response.Redirect "Bs_Admin.asp"

%>