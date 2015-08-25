<!--#include file="../include/head.asp"-->
<%If User_Group<>"admin" Then TurnTo "../"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html;charset=gb2312">
<title>管理面板</title>
<link href="images/css.css" rel="stylesheet" type="text/css">
<script language="javascript" src="../include/edit.js"></script>
</head>

<body>

<table class="bg1" cellspacing="0" cellpadding="0" border="0" align="center">
	<tr id="topbg">
		<td colspan="2"><div id="topleft">[<%=User_ID%>] 已登录 | <a href="../include/login.asp?act=logout">安全退出</a></div><div id="topright"><a href="admin_main.asp">管理面板</a> - <a href="../">返回首页</a> - <a href="http://www.mihao.net/" target="_blank">Mihao.net</a></div></td>
	</tr>
	<tr>
		<td width="600">
			  <div id="sidebg">
			       <div class="sidetop">
				        <a target="show" href="../apost.asp"><b>发表文章</b></a><br>
				        <a target="show" href="admin_show.asp?target=users"><b>会员管理</b></a><br>
				   </div>
						<p>- <a target="show" href="admin_show.asp?target=article">文章管理</a> -</p>
						<p>- <a target="show" href="admin_show.asp?target=board">栏目管理</a> -</p>
						<p>- <a target="show" href="admin_show.asp?target=subclass">分类管理</a> -</p>

						<p class="sidebottom">- <a target="show" href="admin_show.asp?target=Comments">评论管理</a> -</p>
						<p>- <a target="show" href="admin_show.asp?target=Submission">投稿管理</a> -</p>

						<p class="sidebottom">- <a target="show" href="admin_config.asp">网站设置</a> -</p>
                        <p>- <a target="show" href="admin_show.asp?target=topbg">页头图片</a> -</p>
						<p>- <a target="show" href="admin_show.asp?target=notice">消息管理</a> -</p>
						<p>- <a target="show" href="admin_show.asp?target=friendsite">友情链接</a> -</p>

						<p class="sidebottom">- <a target="show" href="admin_show.asp?target=SetLog">操作日志</a> -</p>
              </div>
		</td>
		<td width="100%" rowspan="2" valign="top" class="bg2">
		      <iframe name="show" src="admin_show.asp" width="100%" height="100%" marginwidth="1" marginheight="1" border="0" frameborder="0">浏览器不支持嵌入式框架，或被配置为不显示嵌入式框架。</iframe>
	     </td>
	</tr>
</table>

</body>
</html>
