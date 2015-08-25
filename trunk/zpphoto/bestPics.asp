<!--#include file="dbConn.asp"-->
<%
   dim sql
   dim rs
   dim pic
   set rs=server.createobject("adodb.recordset")
   rs.open "select bestpic from albums where articleid="&request("id"),conn,1,1
   response.redirect ""&rs("bestpic")&""
rs.close
set rs=nothing
conn.close
set conn=nothing
%>


