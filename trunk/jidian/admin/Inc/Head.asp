<html>
<head>
<title>����վ��̨����ϵͳ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Inc/ManageMent.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--



.a3{BACKGROUND: F2F8FF;}


.t1 {font:12px ����;color:#000000} 
.t2 {font:12px ����;color:#ffffff} 
.t3 {font:12px ����;color:#ffff00} 
.t4 {font:12px ����;color:#800000} 
.t5 {font:12px ����;color:#191970} 

.aaa:hover{
	Font-unline:yes;
	font-family: "����";
	color: #FFF;
	text-decoration: underline;
	background: #CCC;
}

-->
</style>
<script language="javascript">
<!--  
  if (window == top)top.location.href = "Default.asp"; 
// -->
</script>
</head>
<%
'����������checkmanage
'����  �ã������̨����Ȩ��                     
'����  ����Ȩ��ֵ
'����  �ڣ�2011/8/5
sub checkmanage(str)
Set mrs = conn.Execute("select * from Bs_User where UserName='"&request.cookies("buyok")("admin")&"'")
if not (mrs.bof and mrs.eof) then
manage=mrs("admin_Manage")
if instr(manage,str)<=0 then
response.write "<script language='javascript'>"
response.write "alert('���棺��û�д��������Ȩ�ޣ�');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
else
session("Name")=0
end if
else 
response.write "<script language='javascript'>"
response.write "alert('û�е�½������ִ�д˲�����');"
response.write "parent.location.href='Default.asp';"
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
Set crs = conn.Execute("select * from Bs_User where UserName='"&request.cookies("buyok")("admin")&"'")
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
response.write "parent.location.href='Default.asp';"
response.write "</script>"
response.end
end if
set crs=nothing
end sub

'����������checkdisplay
'����  �ã������̨���������������Ȩ��                     
'����  ����Ȩ��ֵ
'����  �ڣ�2011/8/8
sub checkdisplay(str)
Set crs = conn.Execute("select * from Bs_User where UserName='"&request.cookies("buyok")("admin")&"'")
if not (crs.bof and crs.eof) then
manage=crs("admin_Manage")
if instr(manage,str)<=0 then
response.write "display:none"
else
response.Write("display:block")
end if
else 
response.write "<script language='javascript'>"
response.write "alert('û�е�½������ִ�д˲�����');"
response.write "parent.location.href='Default.asp';"
response.write "</script>"
response.end
end if
set crs=nothing
end sub
%>