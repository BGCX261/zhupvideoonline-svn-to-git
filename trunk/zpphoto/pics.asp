<!--#include file="dbConn.asp"-->
<%
   dim sql
   dim rs
   dim pic
   set rs=server.createobject("adodb.recordset")
   rs.open "select * from albums where articleid="&request("id"),conn,1,1
   pic=cint(request("pics"))
   select case pic
          case 1
             response.redirect ""&rs("images1")&""
          case 2
             response.redirect ""&rs("images2")&""
          case 3
             response.redirect ""&rs("images3")&""
          case 4
             response.redirect ""&rs("images4")&""
          case 5
             response.redirect ""&rs("images5")&""
          case 6
             response.redirect ""&rs("images6")&""
          case 7
             response.redirect ""&rs("images7")&""
          case 8
             response.redirect ""&rs("images8")&""
          case 9
             response.redirect ""&rs("images9")&""
          case 10
             response.redirect ""&rs("images10")&""
          case 11
             response.redirect ""&rs("images11")&""
          case 12
             response.redirect ""&rs("images12")&""
          case 13
             response.redirect ""&rs("images13")&""
          case 14
             response.redirect ""&rs("images14")&""
          case 15
             response.redirect ""&rs("images15")&""
          case 16
             response.redirect ""&rs("images16")&""
          case 17
             response.redirect ""&rs("images17")&""
          case 18
             response.redirect ""&rs("images18")&""
          case 19
             response.redirect ""&rs("images19")&""
          case 20
             response.redirect ""&rs("images20")&""
   end select
rs.close
set rs=nothing
conn.close
set conn=nothing
%>


