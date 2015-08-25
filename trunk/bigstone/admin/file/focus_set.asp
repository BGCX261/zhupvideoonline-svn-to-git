<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<!-- #include file="function.asp" -->
<%
dim arr:arr = get_focus_sheet()
dim i
%>
<div class="here">焦点图片</div>
<div class="block sheet">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><b>图片地址</b></td>
      <td><b>跳转地址</b></td>
      <td><b>文字说明</b></td>
      <td><b>位置标志</b></td>
      <td width="45px"><b>排序</b></td>
      <td width="40px"><b>置顶</b></td>
      <td width="40px"><b>显示</b></td>
      <td width="100px"><b>操作</b></td>
    </tr>
    <%for i = 0 to ubound(arr)%>
    <tr>
      <td><%=arr(i,1)%></td>
      <td><%=arr(i,2)%></td>
      <td><%=arr(i,3)%></td>
      <td><%=arr(i,4)%></td>
      <td onClick="show_edit('index_<%=arr(i,0)%>')">
      <span id="index_<%=arr(i,0)%>_1"><%=arr(i,5)%><img src="../system/images/pencil.gif"></span>
      <span id="index_<%=arr(i,0)%>_2" style="display:none;"><input type="text" id="index_<%=arr(i,0)%>" value="<%=arr(i,5)%>" style="width:30px;" onBlur="set_index(<%=arr(i,0)%>,this.value,'foc')" /></span>
      </td>
      <td><input style="border:0;" type="checkbox" <%if arr(i,6) = 1 then%>checked<%end if%> onClick="set_top(<%=arr(i,0)%>,this.checked,'foc')" /></td>
      <td><input style="border:0;" type="checkbox" <%if arr(i,7) = 1 then%>checked<%end if%> onClick="set_show(<%=arr(i,0)%>,this.checked,'foc')" /></td>
      <td>
        <%val = arr(i,0)&"{v}"&arr(i,1)&"{v}"&arr(i,2)&"{v}"&arr(i,3)&"{v}"&arr(i,4)%>
        <a onClick="edit_focus('<%=val%>')">[编辑]</a>
        <a onClick="del('<%=arr(i,0)%>')">[删除]</a>
      </td>
    </tr>
    <%next%>
    <tr>
      <td class="button" colspan="30">
        <input type="button" onClick="show_box('add_focus',350,210)" value="添加" />
      </td>
    </tr>
  </table>
</div>
<div class="here">使用说明</div>
<div class="block content">
1.焦点图片可以采用相对路径（相对于网站首页）或以http://开头的网络路径。<br />
2.可以在“图片管理”中上传图片后，复制其图片地址到此处添加焦点图片。<br />
3.位置标志指的是页面URL中从问号后面开始的标识字符串，表示该焦点图片所在的页面，default表示全局。
</div>
<iframe id="deal" name="deal"></iframe>
<!-------------------------- BOX -------------------------->
<div class="box" id="add_focus">
  <div class="head">
    <div class="title">添加焦点图片</div>
    <div class="close" onclick="hid_box('add_focus')">关闭</div>
  </div>
  <form name="form_add_focus" method="post">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>图片地址：</td>
        <td><input name="pic_path" type="text" class="text" maxlength="200" value="http://" /></td>
      </tr>
      <tr>
        <td>跳转地址：</td>
        <td><input name="pic_url" type="text" class="text" maxlength="200" value="http://" /></td>
      </tr>
      <tr>
        <td>文字说明：</td>
        <td><input name="pic_title" type="text" class="text" maxlength="200" /></td>
      </tr>
      <tr>
        <td>位置标志：</td>
        <td><input name="pic_site" type="text" class="text" maxlength="200" value="default" /></td>
      </tr>
      <tr>
        <td class="button" colspan="30"><input type="button" onclick="ajax_submit('form_add_focus')" value="提交" /></td>
      </tr>
    </table>
  </form>
</div>
<div class="box" id="edit_focus">
  <div class="head">
    <div class="title">编辑焦点图片</div>
    <div class="close" onclick="hid_box('edit_focus')">关闭</div>
  </div>
  <form name="form_edit_focus" method="post">
  <input name="foc_id" type="hidden" value="" />
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>图片地址：</td>
        <td><input name="pic_path" type="text" class="text" maxlength="200" value="http://" /></td>
      </tr>
      <tr>
        <td>跳转地址：</td>
        <td><input name="pic_url" type="text" class="text" maxlength="200" value="http://" /></td>
      </tr>
      <tr>
        <td>文字说明：</td>
        <td><input name="pic_title" type="text" class="text" maxlength="200" /></td>
      </tr>
      <tr>
        <td>位置标志：</td>
        <td><input name="pic_site" type="text" class="text" maxlength="200" value="default" /></td>
      </tr>
      <tr>
        <td class="button" colspan="30"><input type="button" onclick="ajax_submit('form_edit_focus')" value="提交" /></td>
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
		case "form_edit_focus":document.form_edit_focus.submit();break;
	}
	interval = setInterval("get_result()",1000);
}
function ajax_submit(val)
{
	switch(val)
	{
		case "form_del":
		var id = document.form_del.id.value;
		ajax("post","deal.asp","cmd=del_focus&id="+id,
		function(data)
		{
			if(data == 1) location.replace(location.href);
		});
		break;
		case "form_add_focus":
		var pic_path = document.form_add_focus.pic_path.value;
		var pic_url = document.form_add_focus.pic_url.value;
		var pic_title = document.form_add_focus.pic_title.value;
		var pic_site = document.form_add_focus.pic_site.value;
		ajax("post","deal.asp","cmd=add_focus&pic_path="+pic_path+"&pic_url="+pic_url+"&pic_title="+pic_title+"&pic_site="+pic_site,
		function(data)
		{
			if(data == 1) location.replace(location.href);
		});
		break;
		case "form_edit_focus":
		var foc_id = document.form_edit_focus.foc_id.value;
		var pic_path = document.form_edit_focus.pic_path.value;
		var pic_url = document.form_edit_focus.pic_url.value;
		var pic_title = document.form_edit_focus.pic_title.value;
		var pic_site = document.form_edit_focus.pic_site.value;
		ajax("post","deal.asp","cmd=edit_focus&foc_id="+foc_id+"&pic_path="+pic_path+"&pic_url="+pic_url+"&pic_title="+pic_title+"&pic_site="+pic_site,
		function(data)
		{
			if(data == 1) location.replace(location.href);
		});
		break;
	}
}
function edit_focus(val)
{
	var arr = val.split("{v}");
	document.form_edit_focus.foc_id.value = arr[0];
	document.form_edit_focus.pic_path.value = arr[1];
	document.form_edit_focus.pic_url.value = arr[2];
	document.form_edit_focus.pic_title.value = arr[3];
	document.form_edit_focus.pic_site.value = arr[4];
	show_box("edit_focus",350,210);
}
</script>
</body>
</html>
<%
call db.close_db()
set db = nothing
'新秀
%>