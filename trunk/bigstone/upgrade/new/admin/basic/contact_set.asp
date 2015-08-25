<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<!-- #include file="function.asp" -->
<%
dim contact:contact = get_contact()
dim i
%>
<div class="here">联系方式</div>
<div class="block">
  <form name="form_edit_contact" method="post" action="deal.asp" target="deal">
    <input name="cmd" type="hidden" value="edit_contact" />
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><b>提示词</b></td>
        <td><b>内容</b></td>
        <td><b>操作</b></td>
      </tr>
      <%for i = 0 to ubound(contact)%>
      <tr>
        <td><input name="cue" type="text" class="text" maxlength="50" value="<%=contact(i,0)%>" /></td>
        <td><input name="content" type="text" class="text" maxlength="150" value="<%=contact(i,1)%>" /></td>
        <td><input name="id" type="hidden" value="<%=contact(i,2)%>" /><span class="red" onClick="del(<%=contact(i,2)%>)">[删除]</span></td>
      </tr>
      <%next%>
      <tr>
        <td class="button" colspan="30">
          <input type="button" onClick="do_submit('form_edit_contact')" value="修改" />
          <input type="button" onClick="show_box('add_contact',350,150)" value="添加" />
        </td>
      </tr>
    </table>
  </form>
</div>
<iframe id="deal" name="deal"></iframe>
<!-------------------------- BOX -------------------------->
<div class="box" id="add_contact">
  <div class="head">
    <div class="title">添加联系方式</div>
    <div class="close" onclick="hid_box('add_contact')">关闭</div>
  </div>
  <form name="form_add_contact" method="post">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>提示词：</td>
        <td><input name="cue" type="text" class="text" maxlength="50" /></td>
      </tr>
      <tr>
        <td>内容：</td>
        <td><input name="content" type="text" class="text" maxlength="150" /></td>
      </tr>
      <tr>
        <td class="button" colspan="30"><input type="button" onclick="ajax_submit('form_add_contact')" value="提交" /></td>
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
		case "form_edit_contact":document.form_edit_contact.submit();break;
	}
	interval = setInterval("get_result()",1000);
}
function ajax_submit(val)
{
	switch(val)
	{
		case "form_add_contact":
		var cue = document.form_add_contact.cue.value;
		var content = document.form_add_contact.content.value;
		ajax("post","deal.asp","cmd=add_contact&cue="+cue+"&content="+content,
		function(data)
		{
			if(data == 1) location.replace(location.href);
		});
		break;
		case "form_del":
		var id = document.form_del.id.value;
		ajax("post","deal.asp","cmd=del_contact&id="+id,
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