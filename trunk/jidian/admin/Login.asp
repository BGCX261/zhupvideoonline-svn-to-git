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
body{background: #fff;}
.title{height:26px; background:url(adminimg/b1.gif) repeat-x; text-align:center; vertical-align:middle; color:#000; font-size:14px; font-weight:bold;}
</style>
<script language=javascript>
<!--
function SetFocus()
{
if (document.Login.UserName.value=="")
	document.Login.UserName.focus();
else
	document.Login.UserName.select();
}
var yz=Math.floor(Math.random()*8999)+1000
function CheckForm()
{
	if(document.Login.UserName.value=="")
	{
		alert("请输入用户名！");
		document.Login.UserName.focus();
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

//-->
</script>
</head>
<body>
    <div style="height:400px;; text-align:center; vertical-align:middle;"> 
<form name="Login" action="Admin_ChkLogin.asp" method="post" target="_parent" onSubmit="return CheckForm();">
	<table style="width:360px; border:#91BCE3 solid 1px; margin:100px auto;">
		<tr> 
			<td colspan="2" class="title">管理员登录</td>
		</tr>
		<tr> 
			<td nowrap class="iptlab">用户名称：</td>
			<td class="ipt"><input name="UserName"  type="text"  id="UserName" maxlength="20" style="width:160px;border:solid 1px;"></td>
		</tr>
		<tr> 
			<td class="iptlab">用户密码：</td>
			<td class="ipt"><input name="Password"  type="password" maxlength="20" style="width:160px;border:solid  1px;"  ></td>
		</tr>
		<tr> 
			<td class="iptlab">验 证 码：</td>
			<td class="ipt"><input name="CheckCode" size="6" maxlength="4" style="border:solid 1px;" >
				请在左边输入 <script>
  document.write(yz)
      </script> </td>
		</tr>
		<tr> 
			<td colspan="2" class="iptsubmit"> 
				<input   type="submit" name="Submit" value=" 确 认 " style=" border: 1 solid #336600" >
				<input name="reset" type="reset"  id="reset" value=" 清 除 " style=" border: 1 solid #336600">
				</td>
		</tr>
	</table>
</form>
    </div>
</body>
</html>
