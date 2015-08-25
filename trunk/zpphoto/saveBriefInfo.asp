<!--#include file="dbConn.asp"-->
<%
briefOk=request("briefOk")
set rs=server.createobject("adodb.recordset")

if briefOk<>"" then
   briefInfo=request("briefInfo")
   rs.open "select * from baseSet",conn,1,3
   rs("briefInfo")=briefInfo
   rs.update
   rs.close
end if

   set rs=nothing
   conn.close
   set conn=nothing
response.redirect "briefInfo.asp"
%>


