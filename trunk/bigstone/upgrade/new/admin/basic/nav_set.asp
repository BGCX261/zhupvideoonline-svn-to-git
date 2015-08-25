<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<!-- #include file="function.asp" -->
<%
dim arr,i,j
dim nav_config:nav_config = get_nav_config("stage")
for i = 0 to ubound(nav_config)
arr = get_nav(nav_config(i,0),1)
%>
<div class="here">导航管理 - <%=nav_config(i,1)%></div>
<div class="block">
  <form name="form_edit_nav_<%=nav_config(i,0)%>" method="post" action="deal.asp" target="deal">
    <input name="cmd" type="hidden" value="edit_nav" />
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><b>文字</b></td>
        <td><b>链接</b></td>
        <td width="45px"><b>排序</b></td>
        <td width="40px"><b>置顶</b></td>
        <td width="40px"><b>显示</b></td>
        <td><b>操作</b></td>
      </tr>
      <%for j = 0 to ubound(arr)%>
      <tr>
        <td><input name="word" type="text" class="text" maxlength="50" value="<%=arr(j,1)%>" /></td>
        <td><input name="url" type="text" class="text" maxlength="150" value="<%=arr(j,2)%>" /></td>
        <td style="text-align:center" onClick="show_edit('index_<%=arr(j,0)%>')">
        <span id="index_<%=arr(j,0)%>_1"><%=arr(j,3)%><img src="../system/images/pencil.gif"></span>
        <span id="index_<%=arr(j,0)%>_2" style="display:none;"><input type="text" id="index_<%=arr(j,0)%>" value="<%=arr(j,3)%>" style="width:30px;" onBlur="set_index(<%=arr(j,0)%>,this.value,'nav')" /></span>
        </td>
        <td style="text-align:center"><input style="border:0;" type="checkbox" <%if arr(j,4) = 1 then%>checked<%end if%> onClick="set_top(<%=arr(j,0)%>,this.checked,'nav')" /></td>
        <td style="text-align:center"><input style="border:0;" type="checkbox" <%if arr(j,5) = 1 then%>checked<%end if%> onClick="set_show(<%=arr(j,0)%>,this.checked,'nav')" /></td>
        <td><input name="id" type="hidden" value="<%=arr(j,0)%>" /><span class="red" onClick="del(<%=arr(j,0)%>)">[删除]</span></td>
      </tr>
      <%next%>
      <tr>
        <td class="button" colspan="30">
          <input type="button" onClick="do_submit('form_edit_nav_<%=nav_config(i,0)%>')" value="修改" />
          <input type="button" onClick="show_box('add_nav_<%=nav_config(i,0)%>',350,150)" value="添加" />
        </td>
      </tr>
    </table>
  </form>
</div>
<%next%>
<iframe id="deal" name="deal"></iframe>
<!-------------------------- BOX -------------------------->
<%for i = 0 to ubound(nav_config)%>
<div class="box" id="add_nav_<%=nav_config(i,0)%>">
  <div class="head">
    <div class="title">添加<%=nav_config(i,1)%></div>
    <div class="close" onclick="hid_box('add_nav_<%=nav_config(i,0)%>')">关闭</div>
  </div>
  <form name="form_add_nav_<%=nav_config(i,0)%>" method="post">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>文字：</td>
        <td><input name="word" type="text" class="text" maxlength="50" /></td>
      </tr>
      <tr>
        <td>链接：</td>
        <td><input name="url" type="text" class="text" maxlength="150" /></td>
      </tr>
      <tr>
        <td class="button" colspan="30"><input type="button" onclick="ajax_submit('form_add_nav_<%=nav_config(i,0)%>')" value="提交" /></td>
      </tr>
    </table>
  </form>
</div>
<%next%>
<!-------------------------- BOX -------------------------->
<script language="javascript">
interval = setInterval("get_result()",1000);
function do_submit(val)
{
	switch(val)
	{
		<%for i = 0 to ubound(nav_config)%>
		case "form_edit_nav_<%=nav_config(i,0)%>":document.form_edit_nav_<%=nav_config(i,0)%>.submit();break;
		<%next%>
	}
	interval = setInterval("get_result()",1000);
}
function ajax_submit(val)
{
	switch(val)
	{
		<%for i = 0 to ubound(nav_config)%>
		case "form_add_nav_<%=nav_config(i,0)%>":
		var word = document.form_add_nav_<%=nav_config(i,0)%>.word.value;
		var url = document.form_add_nav_<%=nav_config(i,0)%>.url.value;
		url = url.replace("&","{njb}");
		ajax("post","deal.asp","cmd=add_nav&men_type=<%=nav_config(i,0)%>&word="+word+"&url="+url,
		function(data)
		{
			if(data == 1) location.replace(location.href);
		});
		break;
		<%next%>
		break;
		case "form_del":
		var id = document.form_del.id.value;
		ajax("post","deal.asp","cmd=del_nav&id="+id,
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