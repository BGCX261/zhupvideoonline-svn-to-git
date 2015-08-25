<!--#include file="dbConn.asp"-->
<%
Function EnCode(Str,Key)
for i=1 to len(Str)
TempNum=Hex(asc(mid(Str,i,1)) XOR Key)	
'response.write TempNum&"|"
if len(TempNum)=3 then
encode=encode &  cstr(TempNum) 
else
encode=encode & "0" & cstr(TempNum)
end if
next
End Function
%>
<%
my=request("my")
change=request("change")
delet=request("delet")
addname=request("addname")
cmdok=request("cmdok")
set rs=server.createobject("adodb.recordset")
if my<>"" then
name=request("name")
pwd=request("pwd")
pwd=Encode(pwd,21)
sql="select * from siteUsers where id=1"
rs.open sql,conn,1,3

rs("name")=name
rs("pwd")=pwd
rs.update
rs.close
end if

if change<>"" then
typena=request("typename")
sql="select * from type where typeid="&request("typeid")
rs.open sql,conn,1,3

rs("type")=typena
rs.update
rs.close
end if

if delet<>"" then
   conn.execute"delete from type where typeid="&request("typeid")
end if

if addname<>"" then
typer=request("type")
rs.open "select * from type",conn,1,3

rs.addnew
rs("type")=typer
rs.update
rs.close
end if

if cmdok<>"" then
   homes=request("homes")
   url=request("url")
   announces=request("announces")
   rs.open "select * from baseSet",conn,1,3
   
   rs("homes")=homes
   rs("url")=url
   rs("announces")=announces
   rs.update
   rs.close
end if

   set rs=nothing
   conn.close
   set conn=nothing
response.redirect "baseInfo.asp"
%>


