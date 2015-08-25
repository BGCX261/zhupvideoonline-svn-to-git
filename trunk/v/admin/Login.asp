<%@language=vbscript codepage=936 %>

<%
option explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
'主要是使随机出现的图片数字随机
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>管理员登录</title>
<link href="Inc/ManageMent.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.title{	background:url(Images/topBar_bg.gif);}
.border{	border: 1px solid #39867B;}
.tdbg{	background:#E1F4EE;	line-height: 120%;}
.topbg{	background:url(Images/topbg.gif);	color: #FFFFFF;}
-->
</style>
<script language=javascript>
<!--
function SetFocus()
{
if (document.Login.adminName.value=="")
	document.Login.adminName.focus();
else
	document.Login.adminName.select();
}
var yz=Math.floor(Math.random()*8999)+1000
function CheckForm()
{
	if(document.Login.adminName.value=="")
	{
		alert("请输入用户名！");
		document.Login.adminName.focus();
		return false;
	}
	if(document.Login.Password.value == "")
	{
		alert("请输入密码！");
		document.Login.Password.focus();
		return false;
	}
	if (document.Login.CheckCode.value==""){
       alert ("请输入您的验证码！");
       document.Login.CheckCode.focus();
       return(false);
    }
}

function CheckBrowser() 
{
  var app=navigator.appName;
  var verStr=navigator.appVersion;
  if (app.indexOf('Netscape') != -1) {
    alert("MY动力友情提示：\n    你使用的是非ie浏览器，可能会导致无法使用后台的部分功能。建议您使用 IE6.0 或以上版本。");
  } 
  else if (app.indexOf('Microsoft') != -1) {
    if (verStr.indexOf("MSIE 3.0")!=-1 || verStr.indexOf("MSIE 4.0") != -1 || verStr.indexOf("MSIE 5.0") != -1 || verStr.indexOf("MSIE 5.1") != -1)
      alert("MY动力友情提示：\n    您的浏览器版本太低，可能会导致无法使用后台的部分功能。建议您使用 IE6.0 或以上版本。");
  }
}
//-->
</script>
</head>
<body style="background-color: #91BCE3;">
<center>
<div style="border:solid #6699CC 1px;border-collapse: collapse; text-align:center; width:450px; margin-top:60px; background:#EFF1F3;"> 
			<table width="100%">
				<tr>
					<td width=150><img src="adminimg/group.jpg" /> </td>
					<td width=280>
<form name="Login" action="Admin_ChkLogin.asp" method="post" target="_parent" onSubmit="return CheckForm();">
	<table width="100%">
		<tr> 
			<td colspan="2" style="height:26px; background:url(adminimg/b1.gif) repeat-x; text-align:center; vertical-align:middle; color:#000; font-size:14px; font-weight:bold;">管理员登录</td>
		</tr>
		<tr> 
			<td nowrap class="iptlab">用户名称：</td>
			<td class="ipt"><input name="adminName"  type="text"  id="adminName" maxlength="20" style="width:160px;border:solid 1px;" onMouseOver="this.style.background='#FDE8FE';" onMouseOut="this.style.background='#FFFFFF'" onFocus="this.select(); "></td>
		</tr>
		<tr> 
			<td class="iptlab">用户密码：</td>
			<td class="ipt"><input name="Password"  type="password" maxlength="20" style="width:160px;border:solid  1px;" onMouseOver="this.style.background='#FDE8FE';" onMouseOut="this.style.background='#FFFFFF'" onFocus="this.select(); "></td>
		</tr>
		<tr> 
			<td class="iptlab">验 证 码：</td>
			<td class="ipt"><input name="CheckCode" size="6" maxlength="4" style="border:solid 1px;" onMouseOver="this.style.background='#FDE8FE';" onMouseOut="this.style.background='#FFFFFF'" onFocus="this.select(); ">
				请在左边输入 <script>
  document.write(yz)
      </script> </td>
		</tr>
		<tr> 
			<td colspan="2" class="iptsubmit"> 
				<input   type="submit" name="Submit" value=" 确 认 " style="background: #FFCCFF; border: 1 solid #336600" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#FFCCFF'">
				<input name="reset" type="reset"  id="reset" value=" 清 除 " style="background: #FFCCFF; border: 1 solid #336600" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#FFCCFF'">
				</td>
		</tr>
	</table>
</form>
					</td>
				</tr>
			</table>
</div>
</center>

<script language="JavaScript" type="text/JavaScript">
CheckBrowser();
SetFocus(); 
</script>
</body>
</html>
