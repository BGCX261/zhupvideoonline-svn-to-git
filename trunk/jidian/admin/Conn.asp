<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%> 
<% 
Session.CodePage = 936           'ע����������windows2000��IIS5��������Ч�� 
Response.Charset = "gb2312" 
%> 
<%Response.Buffer=True%>
<%
dim conn
dim connstr
dim db
db="../NjrAbgKOTRjS386a/SFJJT5MjNI8w.asp" '���ݿ��ļ�λ��
on error resume next
connstr="DBQ="+server.mappath(""&db&"")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
set conn=server.createobject("ADODB.CONNECTION")
if err then
err.clear
else
conn.open connstr
end if

sub CloseConn()
	conn.close
	set conn=nothing
end sub

%>

