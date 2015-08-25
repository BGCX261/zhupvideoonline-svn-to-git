<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<!-- #include file="function.asp" -->
<%dim arr,i%>
<div class="here">选择模板</div>
<div class="block sheet">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><b>模板预览</b></td>
      <td width="100px"><b>模板名称</b></td>
      <td><b>模板说明</b></td>
      <td width="150px"><b>操作</b></td>
    </tr>
    <%arr = get_tpl_sheet()%>
    <%for i = 0 to ubound(arr)%>
    <tr>
      <td class="img">
        <div class="unit">
          <table cellspacing="0" cellpadding="0">
            <tr><td>
              <img src="<%=S_ROOT%>templates/<%=arr(i)%>/images/this.jpg" onload="picresize(this,50,50)"/>
            </td></tr>
          </table>
        </div>
      </td>
      <td><a href="<%=S_ROOT%>templates/<%=arr(i)%>/images/this.jpg"><%=arr(i)%></a></td>
      <td class="content" style="text-align:left;">
		<%=left(get_tpl_text(arr(i)),300)%>
        <a href="<%=S_ROOT%>templates/<%=arr(i)%>/readme.txt">[阅读原文]</a>
      </td>
      <td>
        <%if S_TPL_PATH <> S_ROOT&"templates/"&arr(i)&"/" then%>
          <a onClick="sel_tpl('<%=arr(i)%>')">[使用该模板]</a>
        <%else%>
          <span class="red">[当前模板]</span>
        <%end if%>
      </td>
    </tr>
    <%next%>
    <tr>
      <td colspan="30" height="30px"></td>
    </tr>
  </table>
</div>
<!-------------------------- BOX -------------------------->
<div class="box" id="sel_tpl">
  <div class="close" onclick="hid_box('sel_tpl')">关闭</div>
  <div class="main">您确定使用该模板吗？
    <div class="button">
      <form name="form_sel_tpl" method="post">
        <input name="tpl" type="hidden" value="" />
        <input type="button" onclick="ajax_submit('form_sel_tpl')" value="确定" />
        <input type="button" onclick="hid_box('sel_tpl')" value="取消" />
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
		case "form_sel_tpl":
		var tpl = document.form_sel_tpl.tpl.value;
		ajax("post","deal.asp","cmd=sel_tpl&tpl="+tpl,
		function(data)
		{
			if(data != "") location.replace(location.href);
		});
		break;
	}
}
function sel_tpl(val)
{
	document.form_sel_tpl.tpl.value = val;
	show_box("sel_tpl",300,125);
}
</script>
</body>
</html>
<%
call db.close_db()
set db = nothing
'新秀
%>