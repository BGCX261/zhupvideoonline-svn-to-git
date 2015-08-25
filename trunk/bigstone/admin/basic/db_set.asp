<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<!-- #include file="function.asp" -->
<div class="here">备份数据库</div>
<div class="block">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td>备份数据库：</td>
      <td><input type="button" onClick="ajax_submit('backup')" value="立即备份" /></td>
    </tr>
  </table>
</div>
<div class="here">备份管理</div>
<div class="block">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><b>已备份数据库</b></td>
    <td><b>操作</b></td>
  </tr>
  <%
  dim path:path = server.mappath("../../data/backup")
  dim fso:set fso = server.createobject("scripting.filesystemobject")    
  dim folder:set folder = fso.getfolder(path)
  dim files:set files = folder.files
  dim fil
  for each fil in files
  %>
  <tr>
    <td><%=fil.name%></td>
    <td>
      <span class="red" onClick="restore('<%=fil.name%>')">[还原]</span>
      <span class="red" onClick="del_db('<%=fil.name%>')">[删除]</span>
    </td>
  </tr>
  <%
  next
  set files = nothing
  set folder = nothing
  set fso=nothing
  %>
</table>
</div>
<!-------------------------- BOX -------------------------->
<div class="box" id="del_db">
  <div class="close" onclick="hid_box('del')">关闭</div>
  <div class="main">您确定要删除该备份文件吗？
    <div class="button">
      <form name="form_del_db" method="post">
        <input name="id" id="del_db_id" type="hidden" value="" />
        <input type="button" onclick="ajax_submit('form_del_db')" value="确定" />
        <input type="button" onclick="hid_box('del_db')" value="取消" />
      </form>
    </div>
  </div>
</div>
<div class="box" id="restore">
  <div class="close" onclick="hid_box('restore')">关闭</div>
  <div class="main">您确定要还原该备份文件吗？
    <div class="button">
      <form name="form_restore" method="post">
        <input name="id" id="restore_id" type="hidden" value="" />
        <input type="button" onclick="ajax_submit('form_restore')" value="确定" />
        <input type="button" onclick="hid_box('restore')" value="取消" />
      </form>
    </div>
  </div>
</div>
<!-------------------------- BOX -------------------------->
<script language="javascript">
interval = setInterval("get_result()",1000);
function ajax_submit(val)
{
	switch(val)
	{
		case "backup":
		ajax("post","deal.asp","cmd=backup_db",
		function(data)
		{
			if(data == 1) location.replace(location.href);
		});
		break;
		case "form_del_db":
		var id = document.form_del_db.id.value;
		ajax("post","deal.asp","cmd=del_db&id="+id,
		function(data)
		{
			if(data == 1) location.replace(location.href);
		});
		break;
		case "form_restore":
		var id = document.form_restore.id.value;
		ajax("post","deal.asp","cmd=restore_db&id="+id,
		function(data)
		{
			if(data == 1) location.replace(location.href);
		});
		break;
	}
}
function restore(id)
{
	document.getElementById("restore_id").value = id;
	var st = window.pageYOffset || document.documentElement.scrollTop || document.body.scrollTop || 0;
	cw = document.body.clientWidth;
	t = st + 150;
	l = Math.floor((cw - 260) / 2);
	document.getElementById("restore").style.display = "block";
	document.getElementById("restore").style.width = "300px";
	document.getElementById("restore").style.height = "125px";
	document.getElementById("restore").style.top = t + "px";
	document.getElementById("restore").style.left = l + "px";
}
function del_db(id)
{
	document.getElementById("del_db_id").value = id;
	var st = window.pageYOffset || document.documentElement.scrollTop || document.body.scrollTop || 0;
	cw = document.body.clientWidth;
	t = st + 150;
	l = Math.floor((cw - 260) / 2);
	document.getElementById("del_db").style.display = "block";
	document.getElementById("del_db").style.width = "300px";
	document.getElementById("del_db").style.height = "125px";
	document.getElementById("del_db").style.top = t + "px";
	document.getElementById("del_db").style.left = l + "px";
}
</script>
</body>
</html>
<%
call db.close_db()
set db = nothing
'新秀
%>