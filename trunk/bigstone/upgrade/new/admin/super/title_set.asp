<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<!-- #include file="function.asp" -->
<%
dim title:title = get_title()
dim i
%>
<div class="here">标题设置</div>
<div class="block">
  <form name="form_edit_title" method="post" action="deal.asp" target="deal">
    <input name="cmd" type="hidden" value="edit_title" />
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><b>位置标志</b></td>
        <td><b>标题</b></td>
        <td><b>优先级</b></td>
        <td><b>操作</b></td>
      </tr>
      <%for i = 0 to ubound(title)%>
      <tr>
        <td><input name="site" type="text" class="text" maxlength="50" value="<%=title(i,0)%>" /></td>
        <td><input name="title" type="text" class="text" maxlength="150" value="<%=title(i,1)%>" /></td>
        <td><input name="priority" type="text" class="text" maxlength="150" value="<%=title(i,2)%>" /></td>
        <td><input name="id" type="hidden" value="<%=title(i,3)%>" /><span class="red" onClick="del(<%=title(i,3)%>)">[删除]</span></td>
      </tr>
      <%next%>
      <tr>
        <td class="button" colspan="30">
          <input type="button" onClick="do_submit('form_edit_title')" value="修改" />
          <input type="button" onClick="show_box('add_title',350,185)" value="添加" />
        </td>
      </tr>
    </table>
  </form>
</div>
<div class="here">使用说明</div>
<div class="block content">
  1.这里的标题主要指的是页面标题，但不限于页面标题。例如产品展示页的页面标题中的“产品展示”就是在这里设置的，其页面位置导航中的“产品展示”也是在这里设置。<br />
  2.位置标志指的是页面URL中从问号后面开始的标识字符串，表示该标题所适用的页面。位置标志可以形如product、article_2、article_2_1等。<br />
  3.当位置标志互相包含时，可以通过设定优先级避免冲突。
</div>
<iframe id="deal" name="deal"></iframe>
<!-------------------------- BOX -------------------------->
<div class="box" id="add_title">
  <div class="head">
    <div class="title">添加标题</div>
    <div class="close" onclick="hid_box('add_title')">关闭</div>
  </div>
  <form name="form_add_title" method="post">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>标志：</td>
        <td><input name="site" type="text" class="text" maxlength="50" /></td>
      </tr>
      <tr>
        <td>标题：</td>
        <td><input name="title" type="text" class="text" maxlength="150" /></td>
      </tr>
      <tr>
        <td>优先级：</td>
        <td><input name="priority" type="text" class="text" maxlength="150" /></td>
      </tr>
      <tr>
        <td class="button" colspan="30"><input type="button" onclick="ajax_submit('form_add_title')" value="提交" /></td>
      </tr>
    </table>
  </form>
</div>
<!-------------------------- BOX -------------------------->
<script language="javascript">
interval = setInterval("get_result()",1000);
function do_submit(val)
{
	switch(val)
	{
		case "form_edit_title":document.form_edit_title.submit();break;
	}
	interval = setInterval("get_result()",1000);
}
function ajax_submit(val)
{
	switch(val)
	{
		case "form_add_title":
		var site = document.form_add_title.site.value;
		var title = document.form_add_title.title.value;
		var priority = document.form_add_title.priority.value;
		ajax("post","deal.asp","cmd=add_title&site="+site+"&title="+title+"&priority="+priority,
		function(data)
		{
			if(data == 1) location.replace(location.href);
		});
		break;
		case "form_del":
		var id = document.form_del.id.value;
		ajax("post","deal.asp","cmd=del_title&id="+id,
		function(data)
		{
			if(data == 1) location.replace(location.href);
		});
		break;
	}
}
</script>
</body>
</html>
<%
call db.close_db()
set db = nothing
'新秀
%>