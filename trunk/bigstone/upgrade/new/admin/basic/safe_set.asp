<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<!-- #include file="function.asp" -->
<div class="here">管理员管理</div>
<div class="block">
  <form name="form_admin" method="post">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>所有管理员：</td>
        <td><div id="admins"></div></td>
      </tr>
      <tr>
        <td>旧用户名：</td>
        <td><input name="old_user_name" type="text" class="text" maxlength="30" /></td>
      </tr>
      <tr>
        <td>旧密码：</td>
        <td><input name="old_password" type="text" class="text" maxlength="30" /></td>
      </tr>
      <tr>
        <td>新用户名：</td>
        <td><input name="new_user_name" type="text" class="text" maxlength="30" /></td>
      </tr>
      <tr>
        <td>新密码：</td>
        <td><input name="new_password" type="text" class="text" maxlength="30" /></td>
      </tr>
      <tr>
        <td class="button" colspan="30"><input type="button" onClick="ajax_submit('form_admin')" value="提交" /></td>
      </tr>
    </table>
  </form>  
</div>
<div class="here">使用说明</div>
<div class="block content">
  1.如果旧用户和新用户都有填，则把旧用户换成新用户。<br />
  2.如果只填旧用户，则删除旧用户。<br />
  3.如果只填新用户，则添加新用户。<br />
</div>
<!-------------------------- BOX -------------------------->
<div class="box" id="user_err">
  <div class="close" onclick="hid_box('user_err')">关闭</div>
  <div class="main">您所输入的用户信息有误</div>
</div>
<div class="box" id="being">
  <div class="close" onclick="hid_box('being')">关闭</div>
  <div class="main">您所输入的用户名已经存在</div>
</div>
<div class="box" id="del_err">
  <div class="close" onclick="hid_box('del_err')">关闭</div>
  <div class="main">您不能删除自己的帐号</div>
</div>
<!-------------------------- BOX -------------------------->
<script language="javascript">
admins();
function ajax_submit(val)
{
	var old_user_name = get_value("old_user_name");
	var old_password = get_value("old_password");
	var new_user_name = get_value("new_user_name");
	var new_password = get_value("new_password");
	ajax("post","deal.asp","cmd=edit_safe&old_user_name="+old_user_name+"&old_password="+old_password+"&new_user_name="+new_user_name+"&new_password="+new_password,
	function(data)
	{
		switch(parseInt(data))
		{
			case 1:location.replace(location.href);break;
			case 2:admins();break;
			case 3:show_box("user_err",300,95);break;
			case 4:admins();break;
			case 5:show_box("being",300,95);break;
			case 6:admins();break;
			case 7:show_box("user_err",300,95);break;
			case 8:show_box("del_err",300,95);break;
		}
		interval = setInterval("get_result()",1000);
	});
}
function admins()
{
	ajax("post","../system/ajax.asp","cmd=get_admins",
	function(data)
	{
		document.getElementById("admins").innerHTML = data;
		set_value("old_user_name","");
		set_value("old_password","");
		set_value("new_user_name","");
		set_value("new_password","");
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