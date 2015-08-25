<!--#include file="bsconfig.asp"-->

<%
dim UserID,Action,sql

UserID=trim(Request("UserID"))
Action=trim(request("Action"))
if UserID<>"" then
	if Action="Lock" then
		sql="Update [User] set LockUser=true where UserID=" & CLng(UserID)
	else
		sql="Update [User] set LockUser=false where UserID=" & CLng(UserID)
	end if
	conn.Execute sql
    call CloseConn()      
end if
response.redirect "Bs_User.asp"
%>