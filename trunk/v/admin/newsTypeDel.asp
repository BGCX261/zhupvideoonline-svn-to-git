<!--#include file="../conn.asp"-->
<%newsTypeId=replace(request("checkme")," ","")
newsTypeId=split(newsTypeId,",")
for i=0 to UBound(newsTypeId)
  sql="delete from Bs_User where ID=" & newsTypeId(i)
  conn.execute(sql)
next
response.redirect "newsTypeList.asp"
%>
<%
rsSysAdmin.close
%>
<%
call closeConnDB()
%>