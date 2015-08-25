<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<!-- #include file="function.asp" -->
<%
dim arr,i,j
dim mod_name:mod_name = get_module_name(module)
dim list:list = get_cat_sheet()
%>
<div class="here"><%=mod_name%>分类</div>
<div class="block sheet">
  <form name="form_edit_cat" method="post" action="deal.asp" target="deal">
    <input name="cmd" type="hidden" value="edit_cat" />
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><b>上级分类</b></td>
          <td><b>分类名称</b></td>
          <td width="30px"><b>级数</b></td>
          <td width="45px"><b>排序</b></td>
          <td width="40px"><b>置顶</b></td>
          <td width="40px"><b>显示</b></td>
          <td><b>操作</b></td>
        </tr>
        <%for i = 0 to ubound(list)%>
        <tr>
          <td>
            <select name="cat_parent_id" class="cat_select">
              <option value="0">无</option>
              <%=set_cat_option(list,0,2,6,list(i,1))%>
            </select>
          </td>
          <td><input name="cat_name" type="text" class="text" maxlength="80" value="<%=list(i,2)%>" /></td>
          <td><%=list(i,6)%></td>
          <td style="text-align:center" onClick="show_edit('index_<%=list(i,0)%>')">
          <span id="index_<%=list(i,0)%>_1"><%=list(i,3)%><img src="../system/images/pencil.gif"></span>
          <span id="index_<%=list(i,0)%>_2" style="display:none;"><input type="text" id="index_<%=list(i,0)%>" value="<%=list(i,3)%>" style="width:30px;" onBlur="set_index(<%=list(i,0)%>,this.value,'cat')" /></span>
          </td>
          <td style="text-align:center"><input style="border:0;" type="checkbox" <%if list(i,4) = 1 then%>checked<%end if%> onClick="set_top(<%=list(i,0)%>,this.checked,'cat')" /></td>
          <td style="text-align:center"><input style="border:0;" type="checkbox" <%if list(i,5) = 1 then%>checked<%end if%> onClick="set_show(<%=list(i,0)%>,this.checked,'cat')" /></td>
          <td>
            <input name="cat_id" type="hidden" value="<%=list(i,0)%>" />
            <%if exist_child(list(i,0)) = 0 then%>
            <span class="red" onClick="del(<%=list(i,0)%>)">[删除]</span>
            <%else%>
            <span class="red" onClick="show_box('no_deal',300,95)">[删除]</span>
            <%end if%>
          </td>
        </tr>
        <%next%>
        <tr>
          <td class="button" colspan="30">
            <input type="button" onClick="do_submit('form_edit_cat')" value="修改" />
            <input type="button" onClick="show_box('add_cat',350,150)" value="添加" />
          </td>
        </tr>
      </table>
  </form>
</div>
<iframe id="deal" name="deal"></iframe>
<!-------------------------- BOX -------------------------->
<div class="box" id="add_cat">
  <div class="head">
    <div class="title">添加<%=mod_name%>分类</div>
    <div class="close" onClick="hid_box('add_cat')">关闭</div>
  </div>
  <form name="form_add_cat" method="post">
    <input name="module" type="hidden" value="<%=module%>" />
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>上级分类：</td>
        <td style="text-align:left">
        <select name="cat_parent_id" >
          <option value="0">无</option>
		  <%=set_cat_option(list,0,2,6,-1)%>
        </select>
        </td>
      </tr>
      <tr>
        <td>分类名称：</td>
        <td style="text-align:left"><input name="cat_name" type="text" class="text" maxlength="150" /></td>
      </tr>
      <tr>
        <td class="button" colspan="30"><input type="button" onClick="ajax_submit('form_add_cat')" value="提交" /></td>
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
		case "form_edit_cat":document.form_edit_cat.submit();break;
	}
	interval = setInterval("get_result()",1000);
}
function ajax_submit(val)
{
	switch(val)
	{
		case "form_del":
		var id = document.form_del.id.value;
		ajax("post","deal.asp","cmd=del_cat&id="+id,
		function(data)
		{
			if(data == 1) location.replace(location.href);
		});
		break;
		case "form_add_cat":
		var module = document.form_add_cat.module.value;
		var cat_parent_id = document.form_add_cat.cat_parent_id.value;
		var cat_name = document.form_add_cat.cat_name.value;
		ajax("post","deal.asp","cmd=add_cat&module="+module+"&cat_parent_id="+cat_parent_id+"&cat_name="+cat_name,
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