<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<!-- #include file="function.asp" -->
<%
dim arr,i
arr = get_modules_sheet()
arr2 = get_modules_option()
%>
<div class="here">后台模块设置</div>
<div class="block">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="120px">添加新模块配置：</td>
    <td><input type="button" onClick="show_box('add_module',350,245)" value="添加" /></td>
    <td width="120px">添加列表项配置：</td>
    <td><input type="button" onClick="show_box('add_mod_sheet',350,245)" value="添加" /></td>
    <td width="120px">添加表单项配置：</td>
    <td><input type="button" onClick="show_box('add_mod_form',350,270)" value="添加" /></td>
  </tr>
</table>
</div>
<div class="here">后台模块配置列表</div>
<div class="block sheet">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><b>类型</b></td>
      <td><b>标识</b></td>
      <td><b>配置值</b></td>
      <td width="45px"><b>排序</b></td>
      <td width="100px"><b>操作</b></td>
    </tr>
    <%for i = 0 to ubound(arr)%>
    <tr>
      <td><%=arr(i,1)%></td>
      <td><%=arr(i,2)%></td>
      <td><%=arr(i,3)%></td>
      <td onClick="show_edit('index_<%=arr(i,0)%>')">
      <span id="index_<%=arr(i,0)%>_1"><%=arr(i,4)%><img src="../system/images/pencil.gif"></span>
      <span id="index_<%=arr(i,0)%>_2" style="display:none;"><input type="text" id="index_<%=arr(i,0)%>" value="<%=arr(i,4)%>" style="width:30px;" onBlur="set_index(<%=arr(i,0)%>,this.value,'mod')" /></span>
      </td>
      <td>
        <%val = arr(i,0)&"{v}"&arr(i,1)&"{v}"&arr(i,2)&"{v}"&arr(i,3)%>
        <a onClick="edit_module('<%=val%>')">[编辑]</a>
        <a onClick="del('<%=arr(i,0)%>')">[删除]</a>
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
<div class="box" id="add_module">
  <div class="head">
    <div class="title">添加新模块配置</div>
    <div class="close" onClick="hid_box('add_module')">关闭</div>
  </div>
  <form name="form_add_module" method="post">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>英文名称：</td>
        <td><input name="val_1" type="text" class="text" maxlength="200" /></td>
      </tr>
      <tr>
        <td>中文名称：</td>
        <td><input name="val_2" type="text" class="text" maxlength="200" /></td>
      </tr>
      <tr>
        <td>前台导航：</td>
        <td><input name="val_3" type="text" class="text" maxlength="200" value="none" /></td>
      </tr>
      <tr>
        <td>提取图片：</td>
        <td><input name="val_4" type="text" class="text" maxlength="200" value="0" /></td>
      </tr>
      <tr>
        <td>提取导读：</td>
        <td><input name="val_5" type="text" class="text" maxlength="200" value="0" /></td>
      </tr>
      <tr>
        <td class="button" colspan="30"><input type="button" onClick="ajax_submit('form_add_module')" value="提交" /></td>
      </tr>
    </table>
  </form>
</div>
<div class="box" id="add_mod_sheet">
  <div class="head">
    <div class="title">添加列表项配置</div>
    <div class="close" onClick="hid_box('add_mod_sheet')">关闭</div>
  </div>
  <form name="form_add_mod_sheet" method="post">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>所属模块：</td>
        <td>
        <select name="val_1" style="width:225px;">
          <option value="none">请选择</option>
          <%for i = 0 to ubound(arr2)%>
          <option value="<%=arr2(i,0)%>"><%=arr2(i,1)&"&nbsp;"&arr2(i,0)%></option>
          <%next%>
        </select>
        </td>
      </tr>
      <tr>
        <td>字段名：</td>
        <td><input name="val_2" type="text" class="text" maxlength="200" /></td>
      </tr>
      <tr>
        <td>中文提示：</td>
        <td><input name="val_3" type="text" class="text" maxlength="200" /></td>
      </tr>
      <tr>
        <td>宽度：</td>
        <td><input name="val_4" type="text" class="text" maxlength="200" value="auto" /></td>
      </tr>
      <tr>
        <td>处理函数：</td>
        <td>
        <select name="val_5" style="width:225px;">
          <option value="none">请选择</option>
          <option value="befort_sheet_word">文字 befort_sheet_word()</option>
          <option value="befort_sheet_img">图片 befort_sheet_img()</option>
          <option value="befort_sheet_cat">分类 befort_sheet_cat()</option>
          <option value="befort_sheet_link">链接 befort_sheet_link()</option>
          <option value="befort_sheet_index">排序 befort_sheet_index()</option>
          <option value="befort_sheet_top">置顶 befort_sheet_top()</option>
          <option value="befort_sheet_show">显示 befort_sheet_show()</option>
        </select>
        </td>
      </tr>
      <tr>
        <td class="button" colspan="30"><input type="button" onClick="ajax_submit('form_add_mod_sheet')" value="提交" /></td>
      </tr>
    </table>
  </form>
</div>
<div class="box" id="add_mod_form">
  <div class="head">
    <div class="title">添加表单项配置</div>
    <div class="close" onClick="hid_box('add_mod_form')">关闭</div>
  </div>
  <form name="form_add_mod_form" method="post">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>所属模块：</td>
        <td>
        <select name="val_1" style="width:225px;">
          <option value="none">请选择</option>
          <%for i = 0 to ubound(arr2)%>
          <option value="<%=arr2(i,0)%>"><%=arr2(i,1)&"&nbsp;"&arr2(i,0)%></option>
          <%next%>
        </select>
        </td>
      </tr>
      <tr>
        <td>字段名：</td>
        <td><input name="val_2" type="text" class="text" maxlength="200" /></td>
      </tr>
      <tr>
        <td>表单名：</td>
        <td><input name="val_3" type="text" class="text" maxlength="200" /></td>
      </tr>
      <tr>
        <td>中文提示：</td>
        <td><input name="val_4" type="text" class="text" maxlength="200" /></td>
      </tr>
      <tr>
        <td>显示处理：</td>
        <td>
        <select name="val_5" style="width:225px;">
          <option value="none">请选择</option>
          <option value="befort_form_text">简短文本 befort_form_text()</option>
          <option value="befort_form_cat">分类菜单 befort_form_cat()</option>
          <option value="befort_form_editor">编辑器 befort_form_editor()</option>
          <option value="befort_form_upload_img">上传图片 befort_form_upload_img()</option>
          <option value="befort_form_upload_file">上传文件 befort_form_upload_file()</option>
        </select>
        </td>
      </tr>
      <tr>
        <td>提交处理：</td>
        <td>
        <select name="val_6" style="width:225px;">
          <option value="none">请选择</option>
          <option value="deal_form_text">文本 deal_form_text()</option>
          <option value="deal_form_att">属性 deal_form_att()</option>
          <option value="deal_form_editor">编辑器 deal_form_editor()</option>
          <option value="deal_form_img">图片 deal_form_img()</option>
          <option value="deal_form_file">文件 deal_form_file()</option>
        </select>
        </td>
      </tr>
      <tr>
        <td class="button" colspan="30"><input type="button" onClick="ajax_submit('form_add_mod_form')" value="提交" /></td>
      </tr>
    </table>
  </form>
</div>
<div class="box" id="edit_module">
  <div class="head">
    <div class="title">编辑模块配置</div>
    <div class="close" onclick="hid_box('edit_module')">关闭</div>
  </div>
  <form name="form_edit_module" method="post">
  <input name="mod_id" type="hidden" value="" />
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>类型：</td>
        <td><input name="mod_type" type="text" class="text" maxlength="200" value="http://" /></td>
      </tr>
      <tr>
        <td>标识：</td>
        <td><input name="mod_action" type="text" class="text" maxlength="200" value="http://" /></td>
      </tr>
      <tr>
        <td>配置值：</td>
        <td><input name="mod_config" type="text" class="text" maxlength="200" /></td>
      </tr>
      <tr>
        <td class="button" colspan="30"><input type="button" onclick="ajax_submit('form_edit_module')" value="提交" /></td>
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
		case "form_add_module":
		var val_1 = document.form_add_module.val_1.value;
		var val_2 = document.form_add_module.val_2.value;
		var val_3 = document.form_add_module.val_3.value;
		var val_4 = document.form_add_module.val_4.value;
		var val_5 = document.form_add_module.val_5.value;
		ajax("post","deal.asp","cmd=add_module&val_1="+val_1+"&val_2="+val_2+"&val_3="+val_3+"&val_4="+val_4+"&val_5="+val_5,
		function(data)
		{
			if(data == 1) location.replace(location.href);
		});
		break;
		case "form_add_mod_sheet":
		var val_1 = document.form_add_mod_sheet.val_1.value;
		var val_2 = document.form_add_mod_sheet.val_2.value;
		var val_3 = document.form_add_mod_sheet.val_3.value;
		var val_4 = document.form_add_mod_sheet.val_4.value;
		var val_5 = document.form_add_mod_sheet.val_5.value;
		ajax("post","deal.asp","cmd=add_mod_sheet&val_1="+val_1+"&val_2="+val_2+"&val_3="+val_3+"&val_4="+val_4+"&val_5="+val_5,
		function(data)
		{
			if(data == 1) location.replace(location.href);
		});
		break;
		case "form_add_mod_form":
		var val_1 = document.form_add_mod_form.val_1.value;
		var val_2 = document.form_add_mod_form.val_2.value;
		var val_3 = document.form_add_mod_form.val_3.value;
		var val_4 = document.form_add_mod_form.val_4.value;
		var val_5 = document.form_add_mod_form.val_5.value;
		var val_6 = document.form_add_mod_form.val_6.value;
		ajax("post","deal.asp","cmd=add_mod_form&val_1="+val_1+"&val_2="+val_2+"&val_3="+val_3+"&val_4="+val_4+"&val_5="+val_5+"&val_6="+val_6,
		function(data)
		{
			if(data == 1) location.replace(location.href);
		});
		break;
		case "form_del":
		var id = document.form_del.id.value;
		ajax("post","deal.asp","cmd=del_module&id="+id,
		function(data)
		{
			if(data == 1) location.replace(location.href);
		});
		break;
		case "form_edit_module":
		var mod_id = document.form_edit_module.mod_id.value;
		var mod_type = document.form_edit_module.mod_type.value;
		var mod_action = document.form_edit_module.mod_action.value;
		var mod_config = document.form_edit_module.mod_config.value;
		ajax("post","deal.asp","cmd=edit_module&mod_id="+mod_id+"&mod_type="+mod_type+"&mod_action="+mod_action+"&mod_config="+mod_config,
		function(data)
		{
			if(data == 1) location.replace(location.href);
		});
		break;
	}
}
function edit_module(val)
{
	var arr = val.split("{v}");
	document.form_edit_module.mod_id.value = arr[0];
	document.form_edit_module.mod_type.value = arr[1];
	document.form_edit_module.mod_action.value = arr[2];
	document.form_edit_module.mod_config.value = arr[3];
	show_box("edit_module",350,180);
}
</script>
</body>
</html>
<%
call db.close_db()
set db = nothing
'新秀
%>