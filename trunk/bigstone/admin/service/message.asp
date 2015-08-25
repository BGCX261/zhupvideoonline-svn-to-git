<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<!-- #include file="function.asp" -->
<%
dim page:page = gets("page","")
dim page_sum:page_sum = 0
dim arr:arr = get_mes_sheet(page,page_sum)
dim i
%>
<div class="here">留言列表</div>
<div class="block sheet">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><b>留言内容</b></td>
      <td width="80px"><b>通过审核</b></td>
      <td width="120px"><b>操作</b></td>
    </tr>
    <%for i = 0 to ubound(arr)%>
    <tr>
      <td style="text-align:left;line-height:20px;">
        <div><b><%=arr(i,1)%></b>&nbsp;<%=arr(i,2)%>&nbsp;说：</div>
        <div><%=arr(i,3)%></div>
		<%if arr(i,4) <> "" then%>
        <div><span class="red">管理员回复：</span><%=arr(i,4)%></div>
        <%end if%>
      </td>
      <td>
		<%if arr(i,5) <> 2 then%>
        <input type="checkbox" <%if arr(i,5) = 1 then%>checked<%end if%> onClick="verify(<%=arr(i,0)%>,<%=arr(i,5)%>)" />
		<%else%>悄悄话<%end if%>
      </td>
      <td>
        <a onClick="reply(<%=arr(i,0)%>,<%=page%>,'<%=arr(i,4)%>')">[回复]</a>
        <a onClick="del('<%=arr(i,0)%>,<%=page%>')">[删除]</a>
      </td>
    </tr>
    <%next%>
    <tr>
      <td colspan="30"><%call page_link("message.asp",page,page_sum,"","")%></td>
    </tr>
  </table>
</div>
<!-------------------------- BOX -------------------------->
<div class="box" id="reply">
  <div class="head">
    <div class="title">回复留言</div>
    <div class="close" onclick="hid_box('reply')">关闭</div>
  </div>
  <form name="form_reply_mes" method="post">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td>
        <input name="id" type="hidden" value="" />
        <input name="page" type="hidden" value="" />
        <textarea class="area" name="mes_reply"></textarea>
      </td>
    </tr>
    <tr>
      <td class="button">
        <input type="button" onclick="ajax_submit('form_reply_mes')" value="提交" />
      </td>
    </tr>
  </table>
  </form>
</div>
<!-------------------------- BOX -------------------------->
<script language="javascript">
interval = setInterval("get_result()",1000);
function ajax_submit(val)
{
	switch(val)
	{
		case "form_del":
		var id = document.form_del.id.value;
		ajax("post","deal.asp","cmd=del_mes&id="+id,
		function(data)
		{
			if(data != "") location.replace(location.href);
		});
		break;
		case "form_reply_mes":
		var id = document.form_reply_mes.id.value;
		var page = document.form_reply_mes.page.value;
		var mes_reply = document.form_reply_mes.mes_reply.value;
		ajax("post","deal.asp","cmd=reply_mes&id="+id+"&page="+page+"&mes_reply="+mes_reply,
		function(data)
		{
			if(data != "") location.replace(location.href);
		});
		break;
	}
}
function reply(id,page,text)
{
	show_box("reply",380,180);
	document.form_reply_mes.id.value = id;
	document.form_reply_mes.page.value = page;
	document.form_reply_mes.mes_reply.value = text;
}
function verify(id,val)
{
	if(val == 0) val = 1;
	else val = 0;
	ajax("post","deal.asp","cmd=verify_mes&id="+id+"&mes_show="+val,
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