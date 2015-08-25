<!--#include file="bsconfig.asp"-->

<%
dim SmallClassName,sql

SmallClassName=trim(Request("SmallClassName"))
if SmallClassName<>"" then
	sql="delete from SmallClass where SmallClassName='"&SmallClassName&"'"
	conn.Execute sql
	sql="delete from Product where SmallClassName='"&SmallClassName&"'"
	conn.Execute sql
end if
call CloseConn()      
response.redirect "Bs_Class.asp"
%>