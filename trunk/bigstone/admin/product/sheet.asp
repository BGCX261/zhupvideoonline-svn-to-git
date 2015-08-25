<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<!-- #include file="function.asp" -->
<%
dim search:search = gets("search","")
dim field_:field_ = gets("field_","")
dim page:page = gets("page","")
dim page_sum:page_sum = 0
dim mod_name:mod_name = get_module_name(module)
dim arr:arr = get_art_sheet(search,field_,page,page_sum)
dim rank:rank = get_module_config(module&"_sheet")
dim i,j
%>
<div class="here"><%=mod_name%>列表</div>
<div class="block sheet">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <%
	  for i = 0 to ubound(rank)
		call set_sheet_head(rank(i,1))
	  next
	  %>
      <td width="150px"><b>操作</b></td>
    </tr>
    <%for i = 0 to ubound(arr)%>
    <tr>
      <%
	  for j = 0 to ubound(rank)
	    call set_sheet_rank(rank(j,1),arr(i,0),arr(i,j+1))
	  next
	  %>
      <td>
        <a href="<%=S_ROOT%>?<%=module%>_<%=arr(i,0)%>.html" target="_blank">[查看]</a>&nbsp;
        <a href="edit.asp?id=<%=arr(i,0)%>&page=<%=page%>&module=<%=module%>">[编辑]</a>&nbsp;
        <a onClick="del('<%=arr(i,0)%>,<%=page%>')">[删除]</a>
      </td>
    </tr>
    <%next%>
    <tr>
      <td colspan="30"><%call page_link("sheet.asp",page,page_sum,field_,search)%></td>
    </tr>
  </table>
</div>
<div class="here"><%=mod_name%>搜索</div>
<div class="block">
<form name="form_search" method="get">
<input name="module" type="hidden" value="<%=module%>" />
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="right" width="150px">
      <select name="field_">
        <option value="art_title"><%=mod_name%>标题</option>
        <option value="art_attributes"><%=mod_name%>属性</option>
        <option value="art_text"><%=mod_name%>描述</option>
      </select>
    </td>
    <td>
      <input name="search" type="text" />
      <input type="submit" value="搜索<%=mod_name%>" />
    </td>
  </tr>
</table>
</form>
</div>
<script language="javascript">
interval = setInterval("get_result()",1000);
function ajax_submit(val)
{
	switch(val)
	{
		case "form_del":
		var id = document.form_del.id.value;
		ajax("post","deal.asp","cmd=del_art&id="+id,
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