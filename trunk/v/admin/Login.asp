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
		alert("�������û�����");
		document.Login.adminName.focus();
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

function CheckBrowser() 
{
  var app=navigator.appName;
  var verStr=navigator.appVersion;
  if (app.indexOf('Netscape') != -1) {
    alert("MY����������ʾ��\n    ��ʹ�õ��Ƿ�ie����������ܻᵼ���޷�ʹ�ú�̨�Ĳ��ֹ��ܡ�������ʹ�� IE6.0 �����ϰ汾��");
  } 
  else if (app.indexOf('Microsoft') != -1) {
    if (verStr.indexOf("MSIE 3.0")!=-1 || verStr.indexOf("MSIE 4.0") != -1 || verStr.indexOf("MSIE 5.0") != -1 || verStr.indexOf("MSIE 5.1") != -1)
      alert("MY����������ʾ��\n    ����������汾̫�ͣ����ܻᵼ���޷�ʹ�ú�̨�Ĳ��ֹ��ܡ�������ʹ�� IE6.0 �����ϰ汾��");
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
			<td colspan="2" style="height:26px; background:url(adminimg/b1.gif) repeat-x; text-align:center; vertical-align:middle; color:#000; font-size:14px; font-weight:bold;">����Ա��¼</td>
		</tr>
		<tr> 
			<td nowrap class="iptlab">�û����ƣ�</td>
			<td class="ipt"><input name="adminName"  type="text"  id="adminName" maxlength="20" style="width:160px;border:solid 1px;" onMouseOver="this.style.background='#FDE8FE';" onMouseOut="this.style.background='#FFFFFF'" onFocus="this.select(); "></td>
		</tr>
		<tr> 
			<td class="iptlab">�û����룺</td>
			<td class="ipt"><input name="Password"  type="password" maxlength="20" style="width:160px;border:solid  1px;" onMouseOver="this.style.background='#FDE8FE';" onMouseOut="this.style.background='#FFFFFF'" onFocus="this.select(); "></td>
		</tr>
		<tr> 
			<td class="iptlab">�� ֤ �룺</td>
			<td class="ipt"><input name="CheckCode" size="6" maxlength="4" style="border:solid 1px;" onMouseOver="this.style.background='#FDE8FE';" onMouseOut="this.style.background='#FFFFFF'" onFocus="this.select(); ">
				����������� <script>
  document.write(yz)
      </script> </td>
		</tr>
		<tr> 
			<td colspan="2" class="iptsubmit"> 
				<input   type="submit" name="Submit" value=" ȷ �� " style="background: #FFCCFF; border: 1 solid #336600" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#FFCCFF'">
				<input name="reset" type="reset"  id="reset" value=" �� �� " style="background: #FFCCFF; border: 1 solid #336600" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#FFCCFF'">
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
