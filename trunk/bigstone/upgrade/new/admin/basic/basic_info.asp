<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<div class="here">基本信息</div>
<div class="block">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td>当前版本：</td>
      <td><%=get_config("version")%></td>
      <td>最新版本：</td>
      <td width="50%"><div id="ver"><%=get_config("version")%></div></td>
    </tr>
    <tr>
      <td>网站目录：</td>
      <td><%=S_ROOT%></td>
      <td>服务器时间：</td>
      <td><%=now()%></td>
    </tr>
    <tr>
      <td>用户IP：</td>
      <td><%=get_ip()%></td>
      <td>浏览器类型：</td>
      <td><%=request.servervariables("http_user_agent")%></td>
    </tr>
    <tr>
      <td>服务器名：</td>
      <td><%=request.servervariables("server_name")%></td>
      <td>服务器IP：</td>
      <td><%=request.servervariables("local_addr")%></td>
    </tr>
    <tr>
      <td>服务器端口：</td>
      <td><%=request.servervariables("server_port")%></td>
      <td>IIS版本：</td>
      <td><%=request.servervariables("server_software")%></td>
    </tr>
  </table>
</div>
<div class="here">使用帮助</div>
<div class="block">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><a href="http://www.sin<%'njb%>siu.com" target="_blank">新秀官网</a></td>
      <td><a href="safe_set.asp">修改密码</a></td>
      <td><a href="../service/message.asp">留言管理</a></td>
      <td><a href="../file/pic_set.asp">图片管理</a></td>
    </tr>
  </table>
</div>
<script language="javascript">
function get_version()
{
	ajax("post","../system/ajax.asp","cmd=get_version",
	function(data)
	{
		if(data.substr(0,9) == "njb_send:")
		{
			document.getElementById("ver").innerHTML = data.substr(9);
		}
	});
}
get_version();
</script>
</body>
</html>
<%
call db.close_db()
set db = nothing
'新秀
%>