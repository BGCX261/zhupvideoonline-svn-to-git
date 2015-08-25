<!--#include file="bsconfig.asp"-->

<%
dim UserID,sql,rs

UserID=trim(Request("UserID"))
if UserID<>"" then
	sql="delete from [User] where UserID=" & Clng(UserID)
	conn.Execute sql
end if
call CloseConn()      
response.redirect "Bs_User.asp"
%>