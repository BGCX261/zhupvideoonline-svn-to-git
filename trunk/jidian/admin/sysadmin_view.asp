<!--#include file="bsconfig.asp"-->

<%
set rs = Server.CreateObject("ADODB.Recordset")

sqltext="select * from Bs_User where username='"+Session("Name")+"'"
rs.Open sqltext,conn,1,1
if not rs.EOF then
	realname=trim(rs("realname"))
	identity=trim(rs("identity"))
	logincount=rs("logincount")
	LastLoginTime=rs("LastLoginTime")
	idcount=rs(0)
else
	if Session("Name")="zhu256" then
		realname="zhu256"
	else
		Response.End
	end if	
end if
rs.Close
%>

<html>
<head>
<title>������Ϣ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Inc/ManageMent.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
body {font-size: 12px; color: #000; font-family: ����}
td {font-size: 12px; color: #000; font-family: ����;}

.t1 {font:12px ����;color:#000000} 
.t2 {font:12px ����;color:#ffffff} 
.t3 {font:12px ����;color:#ffff00} 
.t4 {font:12px ����;color:#800000} 
.t5 {font:12px ����;color:#191970} 

.font1 {  font-family: "����"; font-size: 12px; color: #000000}
.font2 {  font-family: "����"; font-size: 20px; font-weight: bold; color: #000000}
.font3 {  font-family: "����"; font-size: 12px; color: #FFFFFF}

A.r1:link {font-size:12px;text-decoration:underline;color:#000000;}
A.r1:visited {font-size:12px;text-decoration:underline;color:#000000;}
A.r1:hover {font-size:12px;text-decoration:underline;color:#cc0000;}
-->
</style>
<!-- #include file="Inc/Head.asp" -->
</head>

<body>
<table align="center" class="a2">
	<tr>
		<td class="a1" colspan="2">ϵͳ��Ϣ</td>
	</tr>
	<tr class="a4">
		<td width="50%" height="23">�û�����<font class="t4"> <%=Session("Name")%></font> <%=identity%></td>
		<td width="50%">��ʵ������<font class="t4"> <%=realname%></font></td>
	</tr>
	<tr class="a4">
		<td width="50%" height="23">��ݹ��ڣ�<font class="t4"><%=Session.timeout%> ����</font></td>
		<td width="50%">����ʱ�䣺<font class="t4"> <%=year(now())%>��<%=month(now())%>��<%=day(now())%>��<%=hour(now())%>:<%=minute(now())%></font></td>
	</tr>
	<!--<tr class="a4">
		<td width="50%" height="23">���ߴ����� <font class="t4"><%=logincount%></font></td>
		<td width="50%">����ʱ�䣺<font class="t4"> <%=LastLoginTime%></font></td>
	</tr>-->
	<tr class="a4">
		<td width="50%" height="23">������������<font class="t4"> <%=Request.ServerVariables("server_name")%> 
		/ <%=Request.ServerVariables("Http_HOST")%></font></td>
		<td width="50%">�ű��������棺<font class="t4"> <%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></font></td>
	</tr>
	<tr class="a4">
		<td width="50%" height="23">��������������ƣ�<font class="t4"> <%=Request.ServerVariables("SERVER_SOFTWARE")%></font></td>
		<td width="50%">������汾��<font class="t4"> <%=Request.ServerVariables("Http_User_Agent")%></font></td>
		<!-- <td width="50%">ACCESS ���ݿ�·����<a target="_blank" href="<%=datapath%><%=datafile%>"><%=datapath%><%=datafile%></a></td> -->
	</tr>
</table>
<br>
<br>
<table align="center" class="a2">
	<tr>
		<td class="a1" colspan="2">�����ݷ�ʽ</td>
	</tr>
	<tr class="a4">
		<td style="<%call checkdisplay("47")%>"><%
set rsStu = Server.CreateObject("ADODB.Recordset")

sqlStuText="Select count(*) as rscount  from StuZone  where  Passed=False"
rsStu.Open sqlStuText,conn,1,1
if not rsStu.EOF then%>
	���� <span style="font-size:14px; font-weight:bold; color:#F00;"><%=rsStu("rscount")%></span> ��ʦ���������Ϣ��Ҫ���  <a href="Bs_stuZone.asp">ǰ������>>></a>
<%else%>
		������Ϣ��Ҫ���
<%end if
rsStu.Close
%></td>
		<td  style="<%call checkdisplay("53")%>"><%set rsCnt = Server.CreateObject("ADODB.Recordset")

sqlCntText="Select count(*) as rscount  from Product  where  Passed=False"
rsCnt.Open sqlCntText,conn,1,1
if not rsCnt.EOF then%>
	���� <span style="font-size:14px; font-weight:bold; color:#F00;"><%=rsCnt("rscount")%></span> ������γ���Ϣ��Ҫ���  <a href="Bs_Article_Check.asp">ǰ������>>></a>
<%else%>
		������Ϣ��Ҫ���
<%end if
rsCnt.Close
%></td>
  </tr>
</table>
<br><br>

<%
sub discreteness
%>
<table class=a2 align=center>
<tr>
<td class=a1>�������</td>
<td class=a1>֧�ּ��汾</td>
</tr>
<%
Dim theInstalledObjects(17)
theInstalledObjects(0) = "MSWC.AdRotator"
theInstalledObjects(1) = "MSWC.BrowserType"
theInstalledObjects(2) = "MSWC.NextLink"
theInstalledObjects(3) = "MSWC.Tools"
theInstalledObjects(4) = "MSWC.Status"
theInstalledObjects(5) = "MSWC.Counters"
theInstalledObjects(6) = "MSWC.PermissionChecker"
theInstalledObjects(7) = "ADODB.Stream"
theInstalledObjects(8) = "adodb.connection"
theInstalledObjects(9) = "Scripting.FileSystemObject"
theInstalledObjects(10) = "SoftArtisans.FileUp"
theInstalledObjects(11) = "SoftArtisans.FileManager"
theInstalledObjects(12) = "JMail.Message"
theInstalledObjects(13) = "CDONTS.NewMail"
theInstalledObjects(14) = "Persits.MailSender"
theInstalledObjects(15) = "LyfUpload.UploadFile"
theInstalledObjects(16) = "Persits.Upload.1"
theInstalledObjects(17) = "w3.upload"
For i=0 to 17
Response.Write "<TR class=a3><TD>&nbsp;" & theInstalledObjects(i) & "<font color=888888>&nbsp;"
select case i
case 8
Response.Write "(ACCESS ���ݿ�)"
case 9
Response.Write "(FSO �ı��ļ���д)"
case 10
Response.Write "(SA-FileUp �ļ��ϴ�)"
case 11
Response.Write "(SA-FM �ļ�����)"
case 12
Response.Write "(JMail �ʼ�����)"
case 13
Response.Write "(WIN����SMTP ����)"
case 14
Response.Write "(ASPEmail �ʼ�����)"
case 15
Response.Write "(LyfUpload �ļ��ϴ�)"
case 16
Response.Write "(ASPUpload �ļ��ϴ�)"
case 17
Response.Write "(w3 upload �ļ��ϴ�)"
end select
Response.Write "</font></td><td height=25>"
If Not IsObjInstalled(theInstalledObjects(i)) Then
Response.Write "<font color=red><b>��</b></font>"
Else
Response.Write "<b>��</b> " & getver(theInstalledObjects(i)) & ""
End If
Response.Write "</td></TR>" & vbCrLf
Next
%>
</table>
<%
end sub

''''''''''''''''''''''''''''''
Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function
''''''''''''''''''''''''''''''
Function getver(Classstr)
On Error Resume Next
getver=""
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(Classstr)
If 0 = Err Then getver=xtestobj.version
Set xTestObj = Nothing
Err = 0
End Function
discreteness
htmlend %>
