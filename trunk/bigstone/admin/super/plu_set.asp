<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<!-- #include file="function.asp" -->
<%dim arr,i%>
<div class="here">插件管理</div>
<div class="block sheet">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="150px"><b>插件名称</b></td>
      <td><b>插件说明</b></td>
      <td width="150px"><b>操作</b></td>
    </tr>
    <%arr = get_plu_sheet()%>
    <%for i = 0 to ubound(arr)%>
    <tr>
      <td><%=arr(i)%></td>
      <td class="content" style="text-align:left;">
		<%=left(get_plu_text(arr(i)),300)%>
        <a href="<%=S_ROOT%>plugins/<%=arr(i)%>/readme.txt">[阅读原文]</a>
      </td>
      <td>
      <%if not get_plu_config(arr(i)) then%>
        <a onClick="install_plu('<%=S_ROOT%>plugins/<%=arr(i)%>/install.asp')">[安装]</a>
      <%else%>
        <a onClick="uninstall_plu('<%=S_ROOT%>plugins/<%=arr(i)%>/uninstall.asp')">[卸载]</a>
      <%end if%>
      <a href="<%=S_ROOT%>plugins/<%=arr(i)%>/index.asp">[更多操作]</a>
      </td>
    </tr>
    <%next%>
    <tr>
      <td colspan="30" height="30px"></td>
    </tr>
  </table>
</div>
<script language="javascript">
interval = setInterval("get_result()",1000);
function install_plu(fil)
{
	ajax("post",fil,"cmd=install",
	function(data)
	{
		if(data == 1) location.replace(location.href);
	});
}
function uninstall_plu(fil)
{
	ajax("post",fil,"cmd=uninstall",
	function(data)
	{
		if(data == 1) location.replace(location.href);
	});
}
</script>
</body>
</html>
<%
call db.close_db()
set db = nothing
'新秀
%>