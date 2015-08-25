<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<!-- #include file="function.asp" -->
<div class="here">创建页面</div>
<div class="block">
  <form name="form_create_page" method="post" action="deal2.asp">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td colspan="30">请选择您所要创建的类似页面：</td>
      </tr>
      <tr>
        <td>
        <input name="old_page" type="hidden" value="" />
        <input name="old_page2" type="radio" onclick="set_old_page('about')" />?about.html &nbsp;&nbsp;
        <input name="old_page2" type="radio" onclick="set_old_page('product')" />?product.html &nbsp;&nbsp;
        <input name="old_page2" type="radio" onclick="set_old_page('article')" />?article.html &nbsp;&nbsp;
        <input name="old_page2" type="radio" onclick="set_old_page('recruit')" />?recruit.html &nbsp;&nbsp;
        <input name="old_page2" type="radio" onclick="set_old_page('download')" />?download.html &nbsp;&nbsp;
        </td>
      </tr>
      <tr>
        <td>新页面链接：<input name="new_page" type="text" class="text" maxlength="50" value="" />&nbsp;例如：lianxi 或 ?lianxi.html</td>
      </tr>
      <tr>
        <td>新页面名称：<input name="new_name" type="text" class="text" maxlength="50" value="" />&nbsp;例如：联系我们</td>
      </tr>
      <tr>
        <td class="button">
          <input type="button" onClick="ajax_submit('form_create_page')" value="提交" />
        </td>
      </tr>
    </table>
  </form>
</div>
<div class="here">已创建页面</div>
<div class="block sheet">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><b>页面链接</b></td>
      <td><b>页面名称</b></td>
      <td><b>类似页面</b></td>
      <td width="150px"><b>操作</b></td>
    </tr>
    <%dim arr:arr = get_new_page()%>
    <%for i = 0 to ubound(arr)%>
    <tr>
      <td><a href="<%=S_ROOT%>?<%=arr(i,0)%>.html" target="_blank">?<%=arr(i,0)%>.html</a></td>
      <td><%=arr(i,1)%></td>
      <td><%=arr(i,2)%></td>
      <td>
        <a onClick="del('<%=arr(i,0)%>,<%=arr(i,2)%>')">[删除]</a>
      </td>
    </tr>
    <%next%>
    <tr>
      <td colspan="30" height="30px"></td>
    </tr>
  </table>
</div>
<iframe id="deal" name="deal"></iframe>
<!-------------------------- BOX -------------------------->
<div class="box" id="lack_info">
  <div class="close" onclick="hid_box('lack_info')">关闭</div>
  <div class="main">您所提交的信息不足或有误</div>
</div>
<div class="box" id="new_page_err">
  <div class="close" onclick="hid_box('new_page_err')">关闭</div>
  <div class="main">新页面链接必须是以字母开头、字母数字组合的字符串</div>
</div>
<div class="box" id="new_page_exist">
  <div class="close" onclick="hid_box('new_page_exist')">关闭</div>
  <div class="main">您所要创建的新页面已经存在</div>
</div>
<!-------------------------- BOX -------------------------->
<script language="javascript">
interval = setInterval("get_result()",1000);
function ajax_submit(val)
{
	switch(val)
	{
		case "form_create_page":
		var old_page = document.form_create_page.old_page.value;
		var new_page = document.form_create_page.new_page.value;
		var new_name = document.form_create_page.new_name.value;
		ajax("post","deal2.asp","cmd=create_page&old_page="+old_page+"&new_page="+new_page+"&new_name="+new_name,
		function(data)
		{
			switch(parseInt(data))
			{
				case 1:location.replace(location.href);break;
				case 2:show_box("lack_info",300,95);break;
				case 3:show_box("new_page_err",450,95);break;
				case 4:show_box("new_page_exist",300,95);break;
			}
		});
		break;
		case "form_del":
		var id = document.form_del.id.value;
		var arr = id.split(',');
		ajax("post","deal2.asp","cmd=del_user_page&del_page="+arr[0]+"&del_type="+arr[1],
		function(data)
		{
			if(data == 1) location.replace(location.href);
		});
		break;
	}
}
function set_old_page(val)
{
	document.form_create_page.old_page.value = val;
}
</script>
</body>
</html>
<%
call db.close_db()
set db = nothing
'新秀
%>