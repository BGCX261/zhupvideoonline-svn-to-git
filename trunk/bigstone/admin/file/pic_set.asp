<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<!-- #include file="function.asp" -->
<%
dim ser_nam:ser_nam = request.servervariables("server_name")
dim arr,i,id
%>
<div class="block fil_upl">
<form name="upload" method="post" enctype="multipart/form-data" target="deal">
  <span>上传图片</span>
  <input name="upl_file" type="file" onChange="do_upl()" />
</form>
</div>
<div class="here">公共图片</div>
<div class="block sheet">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="150px"><b>图片预览</b></td>
      <td><b>图片地址</b></td>
      <td width="120px"><b>数据库相关</b></td>
      <td width="120px"><b>操作</b></td>
    </tr>
    <%arr = get_file_sheet(S_ROOT&"images/")%>
    <%for i = 0 to ubound(arr)%>
    <tr>
      <td><img src="<%=S_ROOT%>images/<%=arr(i)%>" height="30px" /></td>
      <td>http://<%=ser_nam&S_ROOT&"images/"&arr(i)%></td>
      <td>
		<%id = sel_file("images/"&arr(i))%>
        <%if id <> false then%>ID:<%=id%><%else%><span class="red">无关</span><%end if%>
      </td>
      <td>
        <a href="<%=S_ROOT&"images/"&arr(i)%>">[查看]</a>
        <%if id = false then%><a onClick="del('<%=S_ROOT&"images/"&arr(i)%>')">[删除]</a><%end if%>
      </td>
    </tr>
    <%next%>
    <tr>
      <td colspan="30" height="30px">
      </td>
    </tr>
  </table>
</div>
<div class="here">产品图片</div>
<div class="block sheet">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><b>图片目录</b></td>
      <td width="120px"><b>操作</b></td>
    </tr>
	<%arr = get_folder_sheet(S_ROOT&"images/article/")%>
    <%for i = 0 to ubound(arr)%>
    <tr>
      <td><%=S_ROOT&"images/article/"&arr(i)&"/"%></td>
      <td>
        <a href="pic_show.asp?dir=<%=S_ROOT&"images/article/"&arr(i)&"/"%>">[进入查看]</a>
      </td>
    </tr>
    <%next%>
    <tr>
      <td colspan="30" height="30px">
      </td>
    </tr>
  </table>
</div>
<div class="here">编辑器图片</div>
<div class="block sheet">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><b>图片目录</b></td>
      <td width="120px"><b>操作</b></td>
    </tr>
	<%arr = get_folder_sheet(S_ROOT&"images/editor/")%>
    <%for i = 0 to ubound(arr)%>
    <tr>
      <td><%=S_ROOT&"images/editor/"&arr(i)&"/"%></td>
      <td>
        <a href="pic_show.asp?dir=<%=S_ROOT&"images/editor/"&arr(i)&"/"%>">[进入查看]</a>
      </td>
    </tr>
    <%next%>
    <tr>
      <td colspan="30" height="30px">
      </td>
    </tr>
  </table>
</div>
<iframe id="deal" name="deal"></iframe>
<script language="javascript">
var dir = "<%=S_ROOT%>images/";
interval = setInterval("get_result()",1000);
function ajax_submit(val)
{
	switch(val)
	{
		case "form_del":
		var id = document.form_del.id.value;
		ajax("post","deal.asp","cmd=del_pic&id="+id,
		function(data)
		{
			if(data != "") location.replace(location.href);
		});
		break;
	}
}
function do_upl()
{
	var fil = get_value("upl_file");
	if(fil != "")
	{
		var site = fil.lastIndexOf(".");
		if(site != -1)
		{
			site = fil.lastIndexOf("\\");
			fil = fil.substr(site + 1,fil.length - site);
			ajax("post","../system/ajax.asp","cmd=get_name&dir="+dir+"&fil="+fil,
			function(data)
			{
				fil = data;
				document.upload.action = "upload.asp?dir=" + dir + "&fil=" + fil;
				document.upload.submit();
				njb = setInterval("get_pic_ajax_b()",1000);
			});
		}
	}
}
function get_pic_ajax_b()
{
	ajax("post","../system/ajax.asp","cmd=get_fil",
	function(data)
	{
		if(data != "") location.replace(location.href);
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