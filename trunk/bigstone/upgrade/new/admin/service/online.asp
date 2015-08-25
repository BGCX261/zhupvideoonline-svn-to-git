<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<%
call db.reset()
call db.sql_query("select art_id,art_text from "&S_DB_PREFIX&"article where art_type = 'service'")
if db.get_record_count() > 0 then
  dim id:id = db.sql_result(0,"art_id")
  dim online:online = db.sql_result(0,"art_text")
end if
%>
<div class="here">在线客服</div>
<div class="block">
  <form name="form_edit_online" method="post" action="deal.asp" target="deal">
    <input name="cmd" type="hidden" value="edit_online" />
    <input name="id" type="hidden" value="<%=id%>" />
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>获取代码：</td>
        <td class="content">
          请登录&nbsp;<a href="http://wp.qq.com/" target="_blank">http://wp.qq.com/</a>&nbsp;获取QQ在线客服代码<br />
        </td>
      </tr>
      <tr>
        <td>客服代码：</td>
        <td><div class="editor"><%call interface("admin_editor",online)%></div></td>
      </tr>
      <tr>
        <td class="button" colspan="30"><input type="submit" onclick="do_submit('form_edit_online')" value="修改" /></td>
      </tr>
    </table>
  </form>
</div>
<div class="here">使用说明</div>
<div class="block content">
  1.填写客服代码之前先把编辑器切换到“源代码”模式，“源代码”模式一般对应编辑器上的“<>”按钮。<br />
  2.此处可填写QQ、阿里旺旺等多种在线客服代码，但请注意所填写的代码是否会影响前台页面布局。<br />
</div>
<iframe id="deal" name="deal"></iframe>
<script language="javascript">
function do_submit(val)
{
	switch(val)
	{
		case "form_edit_online":break;
	}
	interval = setInterval("get_result()",1000);
}
</script>
</body>
</html>
<%
call db.close_db()
set db = nothing
'新秀
%>