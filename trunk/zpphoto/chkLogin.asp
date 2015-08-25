<!--#include file=dbConn.asp-->
<%
Function DeCode(Str,Key)
for i=1 to len(Str) step 3
decode=decode & chr(int("&H" & mid(Str,i,3)) XOR Key)	
next
End Function
%>
<%
dim sql
dim rs
dim seekerrs
dim founduser
dim name
dim companyid
dim pwd
dim errmsg
dim founderr
founderr=false
FoundUser=false
	name=trim(request("name"))
	pwd=cstr(Request("pwd"))

set rs=server.createobject("adodb.recordset")
sql="select * from siteUsers where id=1"
rs.open sql,conn,1,1
 if not(rs.bof and rs.eof) then
 if pwd=Decode(rs("pwd"),21) and name=rs("name") then
   response.cookies("name49s")=true
   session("UserName")=rs("name")
   response.redirect "adminMain.asp"
 else           
response.write"<script>alert('您输入的用户名或密码不对，请输入正确的用户和密码，方可进入管理页面！');history.back();</Script>"
   response.end
end if
end if
rs.close
conn.close
set rs=nothing
set conn=nothing
%>


