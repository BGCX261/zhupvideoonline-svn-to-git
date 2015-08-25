<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<!-- #include file="function.asp" -->
<%
dim ser_nam:ser_nam = request.servervariables("server_name")
dim arr,i,id
%>
<div class="block fil_upl">
<form name="upload" method="post" enctype="multipart/form-data" target="deal">
  <span>上传文件</span>
  <input name="upl_file" type="file" onChange="do_upl()" />
</form>
</div>
<div class="here">下载资源</div>
<div class="block sheet">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><b>资源地址</b></td>
      <td width="120px"><b>数据库相关</b></td>
      <td width="120px"><b>操作</b></td>
    </tr>
    <%arr = get_file_sheet(S_ROOT&"download/")%>
    <%for i = 0 to ubound(arr)%>
    <tr>
      <td>http://<%=ser_nam&S_ROOT&"download/"&arr(i)%></td>
      <td>
		<%id = sel_file("download/"&arr(i))%>
        <%if id <> false then%>ID:<%=id%><%else%><span class="red">无关</span><%end if%>
      </td>
      <td>
        <%if id = false then%><a onclick="del('<%=S_ROOT&"download/"&arr(i)%>')">[删除]</a><%else%>无<%end if%>
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
var dir = "<%=S_ROOT%>download/";
interval = setInterval("get_result()",1000);
function ajax_submit(val)
{
	switch(val)
	{
		case "form_del":
		var id = document.form_del.id.value;
		ajax("post","deal.asp","cmd=del_file&id="+id,
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