<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<!-- #include file="function.asp" -->
<div class="here">网站设置</div>
<div class="block">
  <form name="form_edit_site" method="post" action="deal.asp" target="deal">
    <input name="cmd" type="hidden" value="edit_site" />
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>网站标题：</td>
        <td><input name="sit_title" type="text" class="text" maxlength="200" value="<%=get_config("sit_title")%>" /></td>
      </tr>
      <tr>
        <td>企业名称：</td>
        <td><input name="sit_name" type="text" class="text" maxlength="200" value="<%=get_config("sit_name")%>" /></td>
      </tr>
      <tr>
        <td>网站域名：</td>
        <td><input name="sit_domain" type="text" class="text" maxlength="200" value="<%=get_config("sit_domain")%>" /></td>
      </tr>
      <tr>
        <td>ICP备案号：</td>
        <td><input name="sit_code" type="text" class="text" maxlength="200" value="<%=get_config("sit_code")%>" /></td>
      </tr>
      <tr>
        <td>备案号链接：</td>
        <td><input name="sit_code_url" type="text" class="text" maxlength="200" value="<%=get_config("sit_code_url")%>" /></td>
      </tr>
      <tr>
        <td>技术支持：</td>
        <td><input name="sit_tec" type="text" class="text" maxlength="200" value="<%=get_config("sit_tec")%>" /></td>
      </tr>
      <tr>
        <td>技术支持链接：</td>
        <td><input name="sit_tec_url" type="text" class="text" maxlength="200" value="<%=get_config("sit_tec_url")%>" /></td>
      </tr>
      <tr>
        <td>网站关键字：</td>
        <td><input name="sit_keywords" type="text" class="text" maxlength="200" value="<%=get_config("sit_keywords")%>" /></td>
      </tr>
      <tr>
        <td>网站描述：</td>
        <td><textarea class="area" name="sit_description"><%=get_config("sit_description")%></textarea></td>
      </tr>
      <tr>
        <td>统计代码：</td>
        <td><textarea class="area" name="sit_count_code"><%=get_config("sit_count_code")%></textarea></td>
      </tr>
      <tr>
        <td class="button" colspan="30"><input type="button" onclick="do_submit('form_edit_site')" value="修改" /></td>
      </tr>
    </table>
  </form>
</div>
<iframe id="deal" name="deal"></iframe>
<!-------------------------- BOX -------------------------->
<div class="box" id="too_len">
  <div class="close" onclick="hid_box('too_len')">关闭</div>
  <div class="main">统计代码长度不能超过250字符</div>
</div>
<!-------------------------- BOX -------------------------->
<script language="javascript">
function do_submit(val)
{
	switch(val)
	{
		case "form_edit_site":
		  var count_code = document.form_edit_site.sit_count_code.value;
		  if(count_code.length > 250)
		  {
			  show_box("too_len",300,95);
		  }else{
			  document.form_edit_site.submit();
		  }
		break;
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