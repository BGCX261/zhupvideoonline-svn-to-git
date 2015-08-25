<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%> 
<% 
Session.CodePage = 936           '注意这个语句在windows2000下IIS5中运行无效。 
Response.Charset = "gb2312" 
%>
<% dim sql_injdata 
SQL_injdata = "'|and|exec|insert|select|delete|update|count|*|%|chr|mid|master|truncate|char|declare" 
SQL_inj = split(SQL_Injdata,"|") 

If Request.QueryString<>"" Then 
For Each SQL_Get In Request.QueryString 
For SQL_Data=0 To Ubound(SQL_inj) 
if instr(Request.QueryString(SQL_Get),Sql_Inj(Sql_DATA))>0 Then 
Response.Write "<Script>alert('系统提示:请不要非法操作!');history.back(-1)</Script>" 
Response.end 
end if 
next 
Next 
End If 

If Request.Form<>"" Then 
For Each Sql_Post In Request.Form 
For SQL_Data=0 To Ubound(SQL_inj) 
if instr(Request.Form(Sql_Post),Sql_Inj(Sql_DATA))>0 Then 
Response.Write "<Script>alert('系统提示:请不要非法操作!');history.back(-1)</Script>" 
Response.end 
end if 
next 
next 
end if 
 %>
<% Response.Buffer=True
dim conn
dim connstr
db="./NjrAbgKOTRjS386a/SFJJT5MjNI8w.asp" '数据库文件位置
on error resume next
connstr="DBQ="+server.mappath(""&db&"")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
set conn=server.createobject("ADODB.CONNECTION")
if err then
err.clear
else
conn.open connstr
end if

Dim sql_BsCo,rs_BsCo,BsCompanyName,BsEnCompanyName,BsAddress,BsLinkman,BsEnLinkman,BsPhone,qq,BsFax,BsWeb,BsEmail,BsZipcode,Declare,BsICPcode
sql_BsCo="select * from syswork where code='101'"
Set rs_BsCo= Server.CreateObject("ADODB.Recordset")
rs_BsCo.open sql_BsCo,conn,1,1

BsCompanyName=trim(rs_BsCo("BsCompanyName"))
BsEnCompanyName=rs_BsCo("BsEnCompanyName")
BsAddress=trim(rs_BsCo("BsAddress"))
BsEnAddress=trim(rs_BsCo("BsEnAddress"))
BsLinkman=trim(rs_BsCo("BsLinkman"))
BsEnLinkman=trim(rs_BsCo("BsEnLinkman"))
BsPhone=trim(rs_BsCo("BsPhone"))
qq=trim(rs_BsCo("qq"))
BsFax=trim(rs_BsCo("BsFax"))
BsWeb=trim(rs_BsCo("BsWeb"))
BsEmail=trim(rs_BsCo("BsEmail"))
BsZipcode=trim(rs_BsCo("BsZipcode"))
Declare=trim(rs_BsCo("Declare"))
BsICPcode=trim(rs_BsCo("BsICPcode"))

rs_BsCo.close


%>
