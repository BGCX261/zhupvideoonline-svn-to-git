<!--#include file="bsconfig.asp"-->
<!--#include file="inc/md5.asp"-->
<%
password=replace(trim(Request("pwd1")),"'","")
password=md5(password)
set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_User where Id=" & request.querystring("uid")
rs.open sqltext,conn,3,3

'更新管理员密码到数据库
rs("PassWord")=password
rs.update
rs.close
conn.close
Response.Write("<script language=""JavaScript"">alert(""修改成功,请退出后重新进入!"");history.go(-1);</script>")

%>
