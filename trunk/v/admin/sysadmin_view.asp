<!--#include file="bsconfig.asp"-->

<%
set rs = Server.CreateObject("ADODB.Recordset")
set rs2 = Server.CreateObject("ADODB.Recordset")

sqltext="select * from Bs_User where adminName='"+Session("adminName")+"'"
rs.Open sqltext,conn,1,1
if not rs.EOF then
	realname=trim(rs("realname"))
	logincount=rs("logincount")
	LastLoginTime=rs("LastLoginTime")
	idcount=rs(0)
else
	if Session("adminName")="xu256" then
		realname="Xu256"
	else
		Response.End
	end if	
end if
rs.Close
%>

<html>
<head>
<title>您的信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Inc/ManageMent.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
body {font-size: 12px; color: #000; font-family: 宋体}
td {font-size: 12px; color: #000; font-family: 宋体;}

.t1 {font:12px 宋体;color:#000000} 
.t2 {font:12px 宋体;color:#ffffff} 
.t3 {font:12px 宋体;color:#ffff00} 
.t4 {font:12px 宋体;color:#800000} 
.t5 {font:12px 宋体;color:#191970} 

.font1 {  font-family: "宋体"; font-size: 12px; color: #000000}
.font2 {  font-family: "黑体"; font-size: 20px; font-weight: bold; color: #000000}
.font3 {  font-family: "宋体"; font-size: 12px; color: #FFFFFF}

A.r1:link {font-size:12px;text-decoration:underline;color:#000000;}
A.r1:visited {font-size:12px;text-decoration:underline;color:#000000;}
A.r1:hover {font-size:12px;text-decoration:underline;color:#cc0000;}
-->
</style>
</head>

<body>
<table align="center" class="a2">
	<tr>
		<td class="a1" colspan="2">系统信息</td>
	</tr>
	<tr class="a4">
		<td width="50%" height="23">用户名：<font class="t4"> <%=Session("adminName")%></font></td>
		<td width="50%">身  份：<font class="t4"> <%=realname%></font></td>
	</tr>
	<tr class="a4">
		<td width="50%" height="23">身份过期：<font class="t4"><%=Session.timeout%> 分钟</font></td>
		<td width="50%">现在时间：<font class="t4"> <%=year(now())%>年<%=month(now())%>月<%=day(now())%>日<%=hour(now())%>:<%=minute(now())%></font></td>
	</tr>
	<!--<tr class="a4">
		<td width="50%" height="23">上线次数： <font class="t4"><%=logincount%></font></td>
		<td width="50%">上线时间：<font class="t4"> <%=LastLoginTime%></font></td>
	</tr>-->
	<tr class="a4">
		<td width="50%" height="23">服务器域名：<font class="t4"> <%=Request.ServerVariables("server_name")%> 
		/ <%=Request.ServerVariables("Http_HOST")%></font></td>
		<td width="50%">脚本解释引擎：<font class="t4"> <%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></font></td>
	</tr>
	<tr class="a4">
		<td width="50%" height="23">服务器软件的名称：<font class="t4"> <%=Request.ServerVariables("SERVER_SOFTWARE")%></font></td>
		<td width="50%">浏览器版本：<font class="t4"> <%=Request.ServerVariables("Http_User_Agent")%></font></td>
		<!-- <td width="50%">ACCESS 数据库路径：<a target="_blank" href="<%=datapath%><%=datafile%>"><%=datapath%><%=datafile%></a></td> -->
	</tr>
</table>
<br>
<br>
<table align="center" class="a2">
	<tr>
		<td class="a1" colspan="2">管理快捷方式</td>
	</tr>
	<tr class="a4">
		<td width="20%" height="23">快捷功能链接</td>
		<td width="80%" height="23"><a href="Bs_Help.asp"><font color="000000">系统帮助</font></a>
		|	<a href="sysadmin_count_list.asp"><font color="000000">网站访问情况</font></a> 
		| <a href="Bs_Admin_Pass_edit.asp?ID=<%=idcount%>"><font color="000000">修改口令</font></a>
		</td>
	</tr>
    <tr>
      <td colspan="2"><%
set rsCntP = Server.CreateObject("ADODB.Recordset")

sqlCntTextP="Select count(*) as rscountP  from Product  where  Passed=False"
rsCntP.Open sqlCntTextP,conn,1,1
if not rsCntP.EOF then%>
	共有 <span style="font-size:14px; font-weight:bold; color:#F00;"><%=rsCntP("rscountP")%></span> 条上传视频需要审核  <a href="Bs_Article_Check.asp">前往处理>>></a>
<%else%>
		暂无视频需要审核
<%end if
rsCntP.Close
%></td>
    </tr>
        <tr>
      <td colspan="2"><%
set rsCntR = Server.CreateObject("ADODB.Recordset")

sqlCntTextR="Select count(*) as rscountR  from buyok_pinglun  where  shenhe=False"
rsCntR.Open sqlCntTextR,conn,1,1
if not rsCntR.EOF then%>
	共有 <span style="font-size:14px; font-weight:bold; color:#F00;"><%=rsCntR("rscountR")%></span> 条评论需要审核  <a href="Bs_Pinglun_Check.asp">前往处理>>></a>
<%else%>
		暂无评论需要审核
<%end if
rsCntR.Close
%></td>
    </tr>
</table>
<br><br>

<%
sub discreteness
%>
<table class=a2 align=center>
<tr>
<td class=a1>组件名称</td>
<td class=a1>支持及版本</td>
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
Response.Write "(ACCESS 数据库)"
case 9
Response.Write "(FSO 文本文件读写)"
case 10
Response.Write "(SA-FileUp 文件上传)"
case 11
Response.Write "(SA-FM 文件管理)"
case 12
Response.Write "(JMail 邮件发送)"
case 13
Response.Write "(WIN虚拟SMTP 发信)"
case 14
Response.Write "(ASPEmail 邮件发送)"
case 15
Response.Write "(LyfUpload 文件上传)"
case 16
Response.Write "(ASPUpload 文件上传)"
case 17
Response.Write "(w3 upload 文件上传)"
end select
Response.Write "</font></td><td height=25>"
If Not IsObjInstalled(theInstalledObjects(i)) Then
Response.Write "<font color=red><b>×</b></font>"
Else
Response.Write "<b>√</b> " & getver(theInstalledObjects(i)) & ""
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
