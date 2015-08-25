<html>
<head>
<title>∷网站后台管理系统∷</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Inc/ManageMent.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--



.a3{BACKGROUND: F2F8FF;}


.t1 {font:12px 宋体;color:#000000} 
.t2 {font:12px 宋体;color:#ffffff} 
.t3 {font:12px 宋体;color:#ffff00} 
.t4 {font:12px 宋体;color:#800000} 
.t5 {font:12px 宋体;color:#191970} 

.aaa:hover{
	Font-unline:yes;
	font-family: "宋体";
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
'│过程名：checkmanage
'│作  用：检验后台管理权限                     
'│参  数：权限值
'│日  期：2011/8/5
sub checkmanage(str)
Set mrs = conn.Execute("select * from Bs_User where UserName='"&request.cookies("buyok")("admin")&"'")
if not (mrs.bof and mrs.eof) then
manage=mrs("admin_Manage")
if instr(manage,str)<=0 then
response.write "<script language='javascript'>"
response.write "alert('警告：您没有此项操作的权限！');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
else
session("Name")=0
end if
else 
response.write "<script language='javascript'>"
response.write "alert('没有登陆，不能执行此操作！');"
response.write "parent.location.href='Default.asp';"
response.write "</script>"
response.end
end if
set mrs=nothing
end sub

'│过程名：checkverify
'│作  用：检验后台审核发表管理权限                     
'│参  数：权限值
'│日  期：2011/8/8
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
response.write "alert('没有登陆，不能执行此操作！');"
response.write "parent.location.href='Default.asp';"
response.write "</script>"
response.end
end if
set crs=nothing
end sub

'│过程名：checkdisplay
'│作  用：检验后台操作帮助发表管理权限                     
'│参  数：权限值
'│日  期：2011/8/8
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
response.write "alert('没有登陆，不能执行此操作！');"
response.write "parent.location.href='Default.asp';"
response.write "</script>"
response.end
end if
set crs=nothing
end sub
%>