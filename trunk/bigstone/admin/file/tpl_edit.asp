<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<%
dim fil:fil = gets("fil","")
dim path:path = server.mappath(S_TPL_PATH&fil)
dim fso:set fso = server.createobject("scripting.filesystemobject")
if fso.fileexists(path) then
  dim stm:set stm = server.createobject("adodb.stream")
  with stm
	.type = 2
	.mode = 3
	.charset = "utf-8"
	.open
	.loadfromfile path
	dim txt:txt = .readtext
	.close
  end with
  set stm = nothing
else
  txt = "none"
end if
set fso = nothing
dim tpl:tpl = txt
%>
<div class="here">编辑模板文件 - <%=fil%></div>
<div class="block sheet">
  <form name="form_edit_tpl" method="post" action="deal.asp" target="deal">
    <input name="cmd" type="hidden" value="edit_tpl" />
    <input name="fil" type="hidden" value="<%=fil%>" />
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>
          <textarea name="tpl" rows="20" cols="100"><%=tpl%></textarea>
        </td>
      </tr>
      <tr>
        <td class="button">
          <input type="submit" onclick="do_submit('form_edit_tpl')" value="提交" />
          <input type="button" onclick="location.replace(document.referrer)" value="返回" />
        </td>
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
		case "form_edit_tpl":break;
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