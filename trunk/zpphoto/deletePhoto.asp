<!--#include file="security.asp"-->
<!--#include file="dbConn.asp"-->
<%    
   if request("page")="" then
   gopage=0
   else
   gopage=request("page")
   end if
   conn.execute"delete from albums where articleid="&request("articleid")
   conn.close
   set conn=nothing
   response.redirect "managePhoto.asp?page="&gopage&"&typeid="&request("typeid")

%>


