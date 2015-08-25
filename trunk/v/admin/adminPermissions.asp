<!--#include file="conn.asp"-->
<%call checkmanage("02")%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../public/css/adminStyle.css" type="text/css" rel="stylesheet" />
<script language="javascript" src="../../public/js/ctrl_checkbox.js"></script>
<title>后台用户权限管理</title>
<%
if request("deladm")<>"" then
deladm=request("deladm")
call deladmin()
end if

admin=request("admin")
if admin="" then admin=request.cookies("buyok")("admin")
if request("edit")<>"ok" then
%>
</head>

<body>

<div id="sysMain">
<div id="sysNavPath">您的位置：后台用户权限管理列表页</div>
<div id="sysSch"><span>后台用户列表</span>
  <table class="sysSchTable">
    <tr>
      <td>
      <div>
<%
		Set rs = conn.Execute("select * from Bs_User")
		do while not rs.eof
		response.write "<img border=0 src=../images/topBar_bg.gif> "
		if admin=rs("adminName") then
		response.write "<a href=adminPermissions.asp?admin="&rs("adminName")&"><font color=red><b>"&rs("adminName")&"</b></font></a>"
		else
		response.write "<a href=adminPermissions.asp?admin="&rs("adminName")&">"&rs("adminName")&"</a>"
	end if
%>
	<img border=0 src=../images/delete.gif alt="删除此管理员" style="cursor:hand" onClick="{if(confirm('该操作不可恢复！\n\n确实要删除这个管理员吗？ ')){location.href='adminPermissions.asp?deladm=<%=rs("adminName")%>';}}"></div>
	<%
	rs.movenext
	loop
	%>
    
    </td>
      </tr>
  </table>
</div>
<div class="comnDiv">
<font color=red><b><%=admin%></b></font> 的管理权限：
<form action="adminPermissions.asp" method=post name=manage>
	<%
	Set rs = conn.Execute("select * from Bs_User where adminName='"&admin&"'")
		  if rs.eof and rs.bof then
			response.write "<script language='javascript'>"
			response.write "alert('没有此管理员！');"
			response.write "location.href='javascript:history.go(-1)';"
			response.write "</script>"
			response.end
		  end if
 		  manage=rs("admin_Manage")
		%>
  <table class="sysListTable">
    <tr class="sysListTitleTr">
      <td>模块</td>
      <td>功能</td>
      <td>权限名称</td>
      <td>备注</td>
      </tr>
    <tbody id="goaler">
    <tr class="sysListSingleTr">
      <td>系统管理</td>
      <td>管理员管理 </td>
      <td>
       <input type="checkbox" value="yes" name="manage01" <%if instr(manage,"01")>0 then%>checked<%end if%>> 
       列表       </td>
      <td>&nbsp;</td>
      </tr>
    <tr class="sysListSingleTr">
      <td>&nbsp;</td>
      <td>管理员权限管理</td>
      <td><input type="checkbox" value="yes" name="manage02" <%if instr(manage,"02")>0 then%>checked<%end if%>>
        列表 </td>
      <td>&nbsp;</td>
    </tr>
    <tr class="sysListSingleTr">
      <td>&nbsp;</td>
      <td>视频上传用户批量增加</td>
      <td><input type="checkbox" value="yes" name="manage03" <%if instr(manage,"03")>0 then%>checked<%end if%>>
        新增</td>
      <td>&nbsp;</td>
    </tr>
    <tr class="sysListSingleTr">
      <td>&nbsp;</td>
      <td>管理员资料管理 </td>
      <td><input type="checkbox" value="yes" name="manage04" <%if instr(manage,"04")>0 then%>checked<%end if%>>
列表  </td>
      <td>&nbsp;</td>
    </tr>
    <tr class="sysListSingleTr">
      <td>&nbsp;</td>
      <td>用户修改密码</td>
      <td><input type="checkbox" value="yes" name="manage05" <%if instr(manage,"05")>0 then%>checked<%end if%>>
修改  </td>
      <td>&nbsp;</td>
    </tr>
    <tr class="sysListSingleTr">
      <td>&nbsp;</td>
      <td>数据库备份</td>
      <td><input type="checkbox" value="yes" name="manage06" <%if instr(manage,"06")>0 then%>checked<%end if%>>
备份 </td>
      <td>&nbsp;</td>
    </tr>
    <tr class="sysListSingleTr">
      <td>&nbsp;</td>
      <td>帮助</td>
      <td><input type="checkbox" value="yes" name="manage07" <%if instr(manage,"07")>0 then%>checked<%end if%>>
帮助设置  </td>
      <td>&nbsp;</td>
    </tr>
    <tr class="sysListSingleTr">
      <td>&nbsp;</td>
      <td>页头图片</td>
      <td><input type="checkbox" value="yes" name="manage08" <%if instr(manage,"08")>0 then%>checked<%end if%>>
设置</td>
      <td>&nbsp;</td>
    </tr>
    <tr class="sysListSingleTr">
      <td>&nbsp;</td>
      <td>广告设置</td>
      <td><input type="checkbox" value="yes" name="manage09" <%if instr(manage,"09")>0 then%>checked<%end if%>>
设置  </td>
      <td>&nbsp;</td>
    </tr>
    <tr class="sysListSingleTr">
      <td>&nbsp;</td>
      <td>上传文件管理</td>
      <td><input type="checkbox" value="yes" name="manage10" <%if instr(manage,"10")>0 then%>checked<%end if%>>
管理  </td>
      <td>&nbsp;</td>
    </tr>
    <tr class="sysListSingleTr">
      <td>&nbsp;</td>
      <td>视频文件管理</td>
      <td><input type="checkbox" value="yes" name="manage11" <%if instr(manage,"11")>0 then%>checked<%end if%>>
管理  </td>
      <td>&nbsp;</td>
    </tr>
    <tr class="sysListSingleTr">
      <td>&nbsp;</td>
      <td>访问情况</td>
      <td><input type="checkbox" value="yes" name="manage12" <%if instr(manage,"12")>0 then%>checked<%end if%>>
管理  </td>
      <td>&nbsp;</td>
    </tr>
    <tr class="sysListSingleTr">
      <td>基本信息管理</td>
      <td>基本资料设置</td>
      <td><input type="checkbox" value="yes" name="manage13" <%if instr(manage,"13")>0 then%>checked<%end if%>>
设置  </td>
      <td>&nbsp;</td>
    </tr>
    <tr class="sysListSingleTr">
      <td>&nbsp;</td>
      <td>网站声明</td>
      <td><input type="checkbox" value="yes" name="manage14" <%if instr(manage,"14")>0 then%>checked<%end if%>>
设置  </td>
      <td>&nbsp;</td>
    </tr>
    <tr class="sysListSingleTr">
      <td>&nbsp;</td>
      <td>基本介绍</td>
      <td><input type="checkbox" value="yes" name="manage15" <%if instr(manage,"15")>0 then%>checked<%end if%>>
设置  </td>
      <td>&nbsp;</td>
    </tr>
    <tr class="sysListSingleTr">
      <td>视频文件上传管理</td>
      <td>视频类别设置</td>
      <td><input type="checkbox" value="yes" name="manage16" <%if instr(manage,"16")>0 then%>checked<%end if%>>
管理  </td>
      <td>&nbsp;</td>
    </tr>
    <tr class="sysListSingleTr">
      <td>&nbsp;</td>
      <td>视频作品 </td>
      <td><input type="checkbox" value="yes" name="manage17" <%if instr(manage,"17")>0 then%>checked<%end if%>>
查询
  <input type="checkbox" value="yes" name="manage18" <%if instr(manage,"18")>0 then%>checked<%end if%>>
修改
<input type="checkbox" value="yes" name="manage19" <%if instr(manage,"19")>0 then%>checked<%end if%> />
审核 <input type="checkbox" value="yes" name="manage20" <%if instr(manage,"20")>0 then%>checked<%end if%> /> 
删除 
<input type="checkbox" value="yes" name="manage21" <%if instr(manage,"21")>0 then%>checked<%end if%> /> 
新增</td>
      <td>&nbsp;</td>
    </tr>
    <tr class="sysListSingleTr">
      <td>视频教程管理</td>
      <td>教程类别管理</td>
      <td><input type="checkbox" value="yes" name="manage22" <%if instr(manage,"22")>0 then%>checked<%end if%> /> 管理</td>
      <td>&nbsp;</td>
    </tr>
    <tr class="sysListSingleTr">
      <td>&nbsp;</td>
      <td>教程管理</td>
      <td><input type="checkbox" value="yes" name="manage23" <%if instr(manage,"23")>0 then%>checked<%end if%> /> 
        管理 
          <input type="checkbox" value="yes" name="manage24" <%if instr(manage,"24")>0 then%>checked<%end if%> />
          删除 <input type="checkbox" value="yes" name="manage25" <%if instr(manage,"25")>0 then%>checked<%end if%> />
          修改 <input type="checkbox" value="yes" name="manage26" <%if instr(manage,"26")>0 then%>checked<%end if%> />
          新增</td>
      <td>&nbsp;</td>
    </tr>
    <tr class="sysListSingleTr">
      <td>新闻动态管理</td>
      <td>新闻管理</td>
      <td><input type="checkbox" value="yes" name="manage27" <%if instr(manage,"27")>0 then%>checked<%end if%> />
        管理 <input type="checkbox" value="yes" name="manage28" <%if instr(manage,"28")>0 then%>checked<%end if%> />
        删除 <input type="checkbox" value="yes" name="manage29" <%if instr(manage,"29")>0 then%>checked<%end if%> />
        修改 <input type="checkbox" value="yes" name="manage30" <%if instr(manage,"30")>0 then%>checked<%end if%> />
        新增</td>
      <td>&nbsp;</td>
    </tr>
    <tr class="sysListSingleTr">
      <td>评论用户管理</td>
      <td>注册用户管理</td>
      <td><input type="checkbox" value="yes" name="manage31" <%if instr(manage,"31")>0 then%>checked<%end if%> /> 
        管理</td>
      <td>&nbsp;</td>
    </tr>
    <tr class="sysListSingleTr">
      <td>&nbsp;</td>
      <td>用户评论</td>
      <td><input type="checkbox" value="yes" name="manage32" <%if instr(manage,"32")>0 then%>checked<%end if%> />
        管理 <input type="checkbox" value="yes" name="manage33" <%if instr(manage,"33")>0 then%>checked<%end if%> />
        审核 <input type="checkbox" value="yes" name="manage34" <%if instr(manage,"34")>0 then%>checked<%end if%> />
        删除</td>
      <td>&nbsp;</td>
    </tr>
    <tr class="sysListSingleTr">
      <td>其他管理</td>
      <td>友情链接管理</td>
      <td><input type="checkbox" value="yes" name="manage35" <%if instr(manage,"35")>0 then%>checked<%end if%> /> 
        管理</td>
      <td>&nbsp;</td>
    </tr>
    <tr class="sysListSingleTr">
      <td>&nbsp;</td>
      <td>广告宣传片设置</td>
      <td><input type="checkbox" value="yes" name="manage36" <%if instr(manage,"36")>0 then%>checked<%end if%> /> 
        设置</td>
      <td>&nbsp;</td>
    </tr>
    <tr class="sysListSingleTr">
      <td>&nbsp;</td>
      <td>邮件服务器设置</td>
      <td><input type="checkbox" value="yes" name="manage37" <%if instr(manage,"37")>0 then%>checked<%end if%> /> 
        设置</td>
      <td>&nbsp;</td>
    </tr>
    </tbody>
  </table>
      <div class="sysSchSmtTd"><input type="submit" name="Submit" value=" 保 存 ">
		<input type="hidden" name="edit" value="ok">
		<input type="hidden" name="admin" value="<%=admin%>"> <input name="input" type="button" value=" 返 回 " onClick="window.history.go(-1)" />
      </div>
</form>
<%   
else

	manage=""
	if request.form("manage01")="yes" then manage=manage+"|01"
	if request.form("manage02")="yes" then manage=manage+"|02"
	if request.form("manage03")="yes" then manage=manage+"|03"
	if request.form("manage04")="yes" then manage=manage+"|04"
	if request.form("manage05")="yes" then manage=manage+"|05"
	if request.form("manage06")="yes" then manage=manage+"|06"
	if request.form("manage07")="yes" then manage=manage+"|07"
	if request.form("manage08")="yes" then manage=manage+"|08"
	if request.form("manage09")="yes" then manage=manage+"|09"
	if request.form("manage10")="yes" then manage=manage+"|10"
	if request.form("manage11")="yes" then manage=manage+"|11"
	if request.form("manage12")="yes" then manage=manage+"|12"
	if request.form("manage13")="yes" then manage=manage+"|13"
	if request.form("manage14")="yes" then manage=manage+"|14"
	if request.form("manage15")="yes" then manage=manage+"|15"
	if request.form("manage16")="yes" then manage=manage+"|16"
	if request.form("manage17")="yes" then manage=manage+"|17"
	if request.form("manage18")="yes" then manage=manage+"|18"
	if request.form("manage19")="yes" then manage=manage+"|19"
	if request.form("manage20")="yes" then manage=manage+"|20"
	if request.form("manage21")="yes" then manage=manage+"|21"
	if request.form("manage22")="yes" then manage=manage+"|22"
	if request.form("manage23")="yes" then manage=manage+"|23"
	if request.form("manage24")="yes" then manage=manage+"|24"
	if request.form("manage25")="yes" then manage=manage+"|25"
	if request.form("manage26")="yes" then manage=manage+"|26"
	if request.form("manage27")="yes" then manage=manage+"|27"
	if request.form("manage28")="yes" then manage=manage+"|28"
	if request.form("manage29")="yes" then manage=manage+"|29"
	if request.form("manage30")="yes" then manage=manage+"|30"
	if request.form("manage31")="yes" then manage=manage+"|31"
	if request.form("manage32")="yes" then manage=manage+"|32"
	if request.form("manage33")="yes" then manage=manage+"|33"
	if request.form("manage34")="yes" then manage=manage+"|34"
	if request.form("manage35")="yes" then manage=manage+"|35"
	if request.form("manage36")="yes" then manage=manage+"|36"
	if request.form("manage37")="yes" then manage=manage+"|37"

	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select * from Bs_User where adminName='"&admin&"'"
	rs.open sql,conn,1,3
		rs("admin_Manage")=manage
		rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing

	response.write "<script language='javascript'>"
	response.write "alert('管理权限设置成功！');"
	response.write "location.href='adminPermissions.asp?admin="&admin&"';"
	response.write "</script>"
end if


sub deladmin()

	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select * from Bs_User where adminName='"&deladm&"'"
	rs.open sql,conn,1,3
		if (rs.bof and rs.eof) then
			response.write "<script language='javascript'>"
			response.write "alert('操作失败，没有此管理员！');"
			response.write "location.href='adminPermissions.asp?admin="&admin&"';"
			response.write "</script>"
			response.end
		elseif deladm=request.cookies("buyok")("admin") then
			response.write "<script language='javascript'>"
			response.write "alert('您不能删除自己！');"
			response.write "location.href='adminPermissions.asp?admin="&admin&"';"
			response.write "</script>"
			response.end
		else
		rs.delete
		rs.update
			response.write "<script language='javascript'>"
			response.write "alert('您已成功删除管理员"&deladm&"');"
			response.write "location.href='adminPermissions.asp?admin="&admin&"';"
			response.write "</script>"
			response.end
		end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing

end sub
%>
</div>

</div>

</body>
</html>
<script>
//列表表格隔行变色
var TbRow = document.getElementById("goaler");
if (TbRow != null){
   for (var i=0;i<TbRow.rows.length ;i++ )
     {
        if (TbRow.rows[i].rowIndex%2==1){
            TbRow.rows[i].className="sysListSingleTr";
        }
        else{
            TbRow.rows[i].className="sysListDoubleTr";
        }
     }
}
</script>
<%
call closeConnDB()
%>