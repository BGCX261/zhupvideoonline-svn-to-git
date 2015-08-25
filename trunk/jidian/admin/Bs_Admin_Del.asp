<!--#include file="bsconfig.asp"-->
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("03")%>
<%
dim SQL, Rs, contentID,CurrentPage

CurrentPage = request("Page")
contentID=request("ID")

set rs=server.createobject("adodb.recordset")
sqltext="delete from Bs_User where Id="& contentID
rs.open sqltext,conn,3,3
set rs=nothing
response.redirect "Bs_Admin.asp"
%>