<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<!-- #include file="function.asp" -->
<%
dim i
dim show:show = get_show()
%>
<div class="here">模块启闭</div>
<div class="block">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <%for i = 0 to ubound(show,2)%>
      <td width="15%"><%=show(0,i)%>：</td>
      <td width="35%">
        <input type="radio" name="<%=show(1,i)%>" id="<%=show(1,i)%>-1" <%=show(2,i)%> onclick="set_show(this.id)" />开启
        <input type="radio" name="<%=show(1,i)%>" id="<%=show(1,i)%>-0" <%=show(3,i)%> onclick="set_show(this.id)" />关闭
      </td>
      <%
      if i mod 2 = 1 then
      %></tr><tr><%
      end if
      next
      if ubound(show,2) mod 2 = 0 then
      %>
      <td width="15%">&nbsp;</td>
      <td width="35%">&nbsp;</td>
      <%end if%>
    </tr>
  </table>
</div>
<iframe id="deal" name="deal"></iframe>
<script language="javascript">
function do_submit(val)
{
	switch(val)
	{
		case "form_edit_show":document.form_edit_show.submit();break;
	}
	interval = setInterval("get_result()",1000);
}
function set_show(id)
{
	ajax("post","deal.asp","cmd=edit_show&id="+id,
	function(data)
	{
		if(data == 1) result();
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