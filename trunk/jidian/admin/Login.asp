<%@language=vbscript codepage=936 %>

<%
option explicit
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
'��Ҫ��ʹ������ֵ�ͼƬ�������
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����Ա��¼</title>
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
		alert("�������û�����");
		document.Login.UserName.focus();
		return false;
	}
	if(document.Login.Password.value == "")
	{
		alert("���������룡");
		document.Login.Password.focus();
		return false;
	}
	if (document.Login.CheckCode.value==""){
       alert ("������������֤�룡");
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
			<td colspan="2" class="title">����Ա��¼</td>
		</tr>
		<tr> 
			<td nowrap class="iptlab">�û����ƣ�</td>
			<td class="ipt"><input name="UserName"  type="text"  id="UserName" maxlength="20" style="width:160px;border:solid 1px;"></td>
		</tr>
		<tr> 
			<td class="iptlab">�û����룺</td>
			<td class="ipt"><input name="Password"  type="password" maxlength="20" style="width:160px;border:solid  1px;"  ></td>
		</tr>
		<tr> 
			<td class="iptlab">�� ֤ �룺</td>
			<td class="ipt"><input name="CheckCode" size="6" maxlength="4" style="border:solid 1px;" >
				����������� <script>
  document.write(yz)
      </script> </td>
		</tr>
		<tr> 
			<td colspan="2" class="iptsubmit"> 
				<input   type="submit" name="Submit" value=" ȷ �� " style=" border: 1 solid #336600" >
				<input name="reset" type="reset"  id="reset" value=" �� �� " style=" border: 1 solid #336600">
				</td>
		</tr>
	</table>
</form>
    </div>
</body>
</html>
