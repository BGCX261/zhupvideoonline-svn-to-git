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


'����������checkmanage
'����  �ã������̨����Ȩ��                     
'����  ����Ȩ��ֵ
'����  �ڣ�2011/8/5
sub checkmanage(str)
Set mrs = conn.Execute("select * from Bs_User where adminName='"&request.cookies("buyok")("admin")&"'")
if not (mrs.bof and mrs.eof) then
  manage=mrs("admin_Manage")
  if instr(manage,str)<=0 then
	response.write "<script language='javascript'>"
	response.write "alert('���棺��û�д��������Ȩ�ޣ�');"
	response.write "parent.location.href='Login.asp';"
	response.write "</script>"
	response.end
'  else
'    session("adminName")=0
   end if
else 
  response.write "<script language='javascript'>"
  response.write "alert('û�е�½������ִ�д˲�����');"
  response.write "parent.location.href='Login.asp';"
  response.write "</script>"
  response.end
end if
set mrs=nothing
end sub


'����������checkverify
'����  �ã������̨��˷������Ȩ��                     
'����  ����Ȩ��ֵ
'����  �ڣ�2011/8/8
sub checkverify(str)
Set crs = conn.Execute("select * from Bs_User where adminName='"&request.cookies("buyok")("admin")&"'")
if not (crs.bof and crs.eof) then
manage=crs("admin_Manage")
if instr(manage,str)<=0 then
response.write " disabled"
else
response.Write(" checked ")
end if
else 
response.write "<script language='javascript'>"
response.write "alert('û�е�½������ִ�д˲�����');"
response.write "parent.location.href='Login.asp';"
response.write "</script>"
response.end
end if
set crs=nothing
end sub

%>

