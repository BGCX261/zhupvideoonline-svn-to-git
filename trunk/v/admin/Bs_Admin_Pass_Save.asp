<!--#include file="bsconfig.asp"-->
<!--#include file="inc/md5.asp"-->
<%
password=replace(trim(Request("pwd1")),"'","")
password=md5(password)
set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_User where Id=" & request.querystring("uid")
rs.open sqltext,conn,3,3

'���¹���Ա���뵽���ݿ�
rs("PassWord")=password
rs.update
rs.close
conn.close
Response.Write("<script language=""JavaScript"">alert(""�޸ĳɹ�,���˳������½���!"");history.go(-1);</script>")

%>
