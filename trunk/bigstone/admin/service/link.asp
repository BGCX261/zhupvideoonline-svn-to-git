<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<!-- #include file="function.asp" -->
<%
dim arr:arr = get_link_sheet()
dim i,val
%>
<div class="here">友情链接</div>
<div class="block sheet">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><b>文字</b></td>
          <td><b>网址</b></td>
          <td><b>图片</b></td>
          <td><b>排序</b></td>
          <td><b>描述</b></td>
          <td width="120px"><b>操作</b></td>
        </tr>
        <%for i = 0 to ubound(arr)%>
        <tr>
          <td><%=arr(i,1)%></td>
          <td><%=arr(i,2)%></td>
          <td><%if arr(i,3) <> "none" then%><img class="lin_img" src="<%=arr(i,3)%>" /><%else echo arr(i,3) end if%></td>
          <td><%=arr(i,4)%></td>
          <td><%=arr(i,5)%></td>
          <td>
			<%val = arr(i,0)&"{v}"&arr(i,1)&"{v}"&arr(i,2)&"{v}"&arr(i,3)&"{v}"&arr(i,4)&"{v}"&arr(i,5)%>
            <a onClick="edit_link('<%=val%>')">[编辑]</a>
            <a onClick="del('<%=arr(i,0)%>')">[删除]</a>
          </td>
        </tr>
        <%next%>
        <tr>
          <td class="button" colspan="30">
            <input type="button" onClick="show_box('add_link',350,250)" value="添加" />
          </td>
        </tr>
      </table>
  </form>
</div>
<iframe id="deal" name="deal"></iframe>
<!-------------------------- BOX -------------------------->
<div class="box" id="add_link">
  <div class="head">
    <div class="title">添加友情链接</div>
    <div class="close" onClick="hid_box('add_link')">关闭</div>
  </div>
  <form name="form_add_link" method="post">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>文字：</td>
        <td><input name="lin_word" type="text" class="text" maxlength="40" value="" /></td>
      </tr>
      <tr>
        <td>网址：</td>
        <td><input name="lin_url" type="text" class="text" maxlength="200" value="http://" /></td>
      </tr>
      <tr>
        <td>图片：</td>
        <td><input name="lin_img" type="text" class="text" maxlength="200" value="http://" /></td>
      </tr>
      <tr>
        <td>排序：</td>
        <td><input name="lin_index" type="text" class="text" maxlength="10" value="" /></td>
      </tr>
      <tr>
        <td>描述：</td>
        <td><input name="lin_title" type="text" class="text" maxlength="200" value="" /></td>
      </tr>
      <tr>
        <td class="button" colspan="30"><input type="button" onClick="ajax_submit('form_add_link')" value="提交" /></td>
      </tr>
    </table>
  </form>
</div>
<div class="box" id="edit_link">
  <div class="head">
    <div class="title">编辑友情链接</div>
    <div class="close" onClick="hid_box('edit_link')">关闭</div>
  </div>
  <form name="form_edit_link" method="post">
  <input name="lin_id" type="hidden" value="" />
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>文字：</td>
        <td><input name="lin_word" type="text" class="text" maxlength="40" value="" /></td>
      </tr>
      <tr>
        <td>网址：</td>
        <td><input name="lin_url" type="text" class="text" maxlength="200" value="http://" /></td>
      </tr>
      <tr>
        <td>图片：</td>
        <td><input name="lin_img" type="text" class="text" maxlength="200" value="http://" /></td>
      </tr>
      <tr>
        <td>排序：</td>
        <td><input name="lin_index" type="text" class="text" maxlength="10" value="" /></td>
      </tr>
      <tr>
        <td>描述：</td>
        <td><input name="lin_title" type="text" class="text" maxlength="200" value="" /></td>
      </tr>
      <tr>
        <td class="button" colspan="30"><input type="button" onClick="ajax_submit('form_edit_link')" value="提交" /></td>
      </tr>
    </table>
  </form>
</div>
<!-------------------------- BOX -------------------------->
<script language="javascript">
interval = setInterval("get_result()",1000);
function ajax_submit(val)
{
	switch(val)
	{
		case "form_del":
		var id = document.form_del.id.value;
		ajax("post","deal.asp","cmd=del_link&id="+id,
		function(data)
		{
			if(data == 1) location.replace(location.href);
		});
		break;
		case "form_add_link":
		var lin_word = document.form_add_link.lin_word.value;
		var lin_url = document.form_add_link.lin_url.value;
		var lin_img = document.form_add_link.lin_img.value;
		var lin_index = document.form_add_link.lin_index.value;
		var lin_title = document.form_add_link.lin_title.value;
		var num = parseInt(lin_index) + 0;
		if(parseInt(lin_index) != num) lin_index = 0;
		ajax("post","deal.asp","cmd=add_link&lin_word="+lin_word+"&lin_url="+lin_url+"&lin_img="+lin_img+"&lin_index="+lin_index+"&lin_title="+lin_title,
		function(data)
		{
			if(data == 1) location.replace(location.href);
		});
		break;
		case "form_edit_link":
		var lin_id = document.form_edit_link.lin_id.value;
		var lin_word = document.form_edit_link.lin_word.value;
		var lin_url = document.form_edit_link.lin_url.value;
		var lin_img = document.form_edit_link.lin_img.value;
		var lin_index = document.form_edit_link.lin_index.value;
		var lin_title = document.form_edit_link.lin_title.value;
		var num = parseInt(lin_index) + 0;
		if(parseInt(lin_index) != num) lin_index = 0;
		ajax("post","deal.asp","cmd=edit_link&lin_id="+lin_id+"&lin_word="+lin_word+"&lin_url="+lin_url+"&lin_img="+lin_img+"&lin_index="+lin_index+"&lin_title="+lin_title,
		function(data)
		{
			if(data == 1) location.replace(location.href);
		});
		break;
	}
}
function edit_link(val)
{
	var arr = val.split("{v}");
	document.form_edit_link.lin_id.value = arr[0];
	document.form_edit_link.lin_word.value = arr[1];
	document.form_edit_link.lin_url.value = arr[2];
	document.form_edit_link.lin_img.value = arr[3];
	document.form_edit_link.lin_index.value = arr[4];
	document.form_edit_link.lin_title.value = arr[5];
	show_box("edit_link",350,250);
}
</script>
</body>
</html>
<%
call db.close_db()
set db = nothing
'新秀
%>