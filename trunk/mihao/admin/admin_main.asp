<!--#include file="../include/head.asp"-->
<%If User_Group<>"admin" Then TurnTo "../"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html;charset=gb2312">
<title>�������</title>
<link href="images/css.css" rel="stylesheet" type="text/css">
<script language="javascript" src="../include/edit.js"></script>
</head>

<body>

<table class="bg1" cellspacing="0" cellpadding="0" border="0" align="center">
	<tr id="topbg">
		<td colspan="2"><div id="topleft">[<%=User_ID%>] �ѵ�¼ | <a href="../include/login.asp?act=logout">��ȫ�˳�</a></div><div id="topright"><a href="admin_main.asp">�������</a> - <a href="../">������ҳ</a> - <a href="http://www.mihao.net/" target="_blank">Mihao.net</a></div></td>
	</tr>
	<tr>
		<td width="600">
			  <div id="sidebg">
			       <div class="sidetop">
				        <a target="show" href="../apost.asp"><b>��������</b></a><br>
				        <a target="show" href="admin_show.asp?target=users"><b>��Ա����</b></a><br>
				   </div>
						<p>- <a target="show" href="admin_show.asp?target=article">���¹���</a> -</p>
						<p>- <a target="show" href="admin_show.asp?target=board">��Ŀ����</a> -</p>
						<p>- <a target="show" href="admin_show.asp?target=subclass">�������</a> -</p>

						<p class="sidebottom">- <a target="show" href="admin_show.asp?target=Comments">���۹���</a> -</p>
						<p>- <a target="show" href="admin_show.asp?target=Submission">Ͷ�����</a> -</p>

						<p class="sidebottom">- <a target="show" href="admin_config.asp">��վ����</a> -</p>
                        <p>- <a target="show" href="admin_show.asp?target=topbg">ҳͷͼƬ</a> -</p>
						<p>- <a target="show" href="admin_show.asp?target=notice">��Ϣ����</a> -</p>
						<p>- <a target="show" href="admin_show.asp?target=friendsite">��������</a> -</p>

						<p class="sidebottom">- <a target="show" href="admin_show.asp?target=SetLog">������־</a> -</p>
              </div>
		</td>
		<td width="100%" rowspan="2" valign="top" class="bg2">
		      <iframe name="show" src="admin_show.asp" width="100%" height="100%" marginwidth="1" marginheight="1" border="0" frameborder="0">�������֧��Ƕ��ʽ��ܣ�������Ϊ����ʾǶ��ʽ��ܡ�</iframe>
	     </td>
	</tr>
</table>

</body>
</html>
