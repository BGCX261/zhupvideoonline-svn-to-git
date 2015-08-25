<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<%
call db.reset()
call db.sql_query("select art_id,art_text from "&S_DB_PREFIX&"article where art_type = 'notice'")
if db.get_record_count() > 0 then
  dim id:id = db.sql_result(0,"art_id")
  dim notice:notice = db.sql_result(0,"art_text")
end if
%>
<div class="here">网站公告</div>
<div class="block">
  <form name="form_edit_notice" method="post" action="deal.asp" target="deal">
    <input name="cmd" type="hidden" value="edit_notice" />
    <input name="id" type="hidden" value="<%=id%>" />
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>网站公告：</td>
        <td><div class="editor"><%call interface("admin_editor",notice)%></div></td>
      </tr>
      <tr>
        <td class="button" colspan="30"><input type="submit" onclick="do_submit('form_edit_notice')" value="修改" /></td>
      </tr>
    </table>
  </form>
</div>
<iframe id="deal" name="deal"></iframe>
<script language="javascript">
function do_submit(val)
{
	switch(val)
	{
		case "form_edit_notice":break;
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