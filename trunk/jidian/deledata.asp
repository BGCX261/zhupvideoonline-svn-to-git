<!--#include file="inc/Conn.asp"-->
<%
   useridtxt = Request("id")
   delesql="DELETE FROM gstbook WHERE id=" & useridtxt
   conn.Execute delesql
   response.redirect "admin.asp"
%>