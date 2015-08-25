<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<!-- #include file="function.asp" -->
<%dim arr,i%>
<div class="here">语言包</div>
<div class="block sheet">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><b>文件名</b></td>
      <td width="150px"><b>操作</b></td>
    </tr>
    <%arr = get_lang_sheet()%>
    <%for i = 0 to ubound(arr)%>
    <tr>
      <td><%=arr(i)%></td>
      <td>
        <a href="lang_edit.asp?fil=<%=arr(i)%>">[编辑]</a>
        <%if arr(i) <> S_LANG&".asp" then%><a onClick="sel_lang('<%=arr(i)%>')">[使用]</a><%else%><span class="red">[使用中]</span><%end if%>
      </td>
    </tr>
    <%next%>
    <tr>
      <td colspan="30" height="30px"></td>
    </tr>
  </table>
</div>
<!-------------------------- BOX -------------------------->
<div class="box" id="sel_lang">
  <div class="close" onclick="hid_box('sel_lang')">关闭</div>
  <div class="main">您确定使用该语言包吗？
    <div class="button">
      <form name="form_sel_lang" method="post">
        <input name="lang" type="hidden" value="" />
        <input type="button" onclick="ajax_submit('form_sel_lang')" value="确定" />
        <input type="button" onclick="hid_box('sel_lang')" value="取消" />
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
		case "form_sel_lang":
		var lang = document.form_sel_lang.lang.value;
		ajax("post","deal.asp","cmd=sel_lang&lang="+lang,
		function(data)
		{
			if(data != "") location.replace(location.href);
		});
		break;
	}
}
function sel_lang(val)
{
	document.form_sel_lang.lang.value = val;
	show_box("sel_lang",300,125);
}
</script>
</body>
</html>
<%
call db.close_db()
set db = nothing
'新秀
%>