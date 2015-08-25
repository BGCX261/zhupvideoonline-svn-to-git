<!--#include file="bsconfig.asp"-->
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("04")%>
<%
if request("deladm")<>"" then
deladm=request("deladm")
call deladmin()
end if

admin=request("admin")
if admin="" then admin=request.cookies("buyok")("admin")
if request("edit")<>"ok" then
%>

  <table class="a2" align="center">
  <tr>
   <td class="a1" colspan="3">后台用户列表(请点下面具体的用户查看所拥有的权限)</td>
  </tr>
    
<%
		Set rs = conn.Execute("select * from Bs_User")
		do while not rs.eof
		response.write "<tr><td>"
		if admin=rs("UserName") then
		response.write "<img border=0 src=../images/topBar_bg.gif> <a href=Bs_adminPermissions.asp?admin="&rs("UserName")&"><font color=red><b>"&rs("UserName")&"</b></font></a> </td><td>" & rs("realname")
		else
		response.write "<a href=Bs_adminPermissions.asp?admin="&rs("UserName")&">"&rs("UserName")&"</a> </td><td>" & rs("realname")
	end if
%>
	</td><td><img border=0 src=../images/delete.gif alt="删除此管理员" style="cursor:hand" onClick="{if(confirm('该操作不可恢复！\n\n确实要删除这个管理员吗？ ')){location.href='Bs_adminPermissions.asp?deladm=<%=rs("UserName")%>';}}">
        </td>
      </tr>
	<%
	rs.movenext
	loop
	%> 
  </table>

<form action="Bs_adminPermissions.asp" method=post name=manage>
	<%
	Set rs = conn.Execute("select * from Bs_User where UserName='"&admin&"'")
		  if rs.eof and rs.bof then
			response.write "<script language='javascript'>"
			response.write "alert('没有此管理员！');"
			response.write "location.href='javascript:history.go(-1)';"
			response.write "</script>"
			response.end
		  end if
 		  manage=rs("admin_Manage")
		%>
  <table align="center" class="a2">
	<tr>
		<td class="a1"><font color=red><b><%=admin%></b></font> 的管理权限</td>
	</tr>
	<tr class="a4">
		<td>
  <table width="100%" style=" border-collapse:collapse;">
    <tr class="opttitle">
      <td>模块</td>
      <td>功能点</td>
      <td>权限名称</td>
      <td>备注</td>
      </tr>
    <tbody id="goaler">
    <tr>
      <td rowspan="13">系统管理</td>
      <td>管理员管理 </td>
      <td>
       <input type="checkbox" value="yes" name="manage01" <%if instr(manage,"01")>0 then%>checked<%end if%>>
       增加
       <input type="checkbox" value="yes" name="manage02" <%if instr(manage,"02")>0 then%>checked<%end if%>>
       修改密码 
       <input type="checkbox" value="yes" name="manage03" <%if instr(manage,"03")>0 then%>checked<%end if%>>删除</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>权限管理</td>
      <td><input type="checkbox" value="yes" name="manage04" <%if instr(manage,"04")>0 then%>checked<%end if%>>
管理</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>数据库备份</td>
      <td><input type="checkbox" value="yes" name="manage05" <%if instr(manage,"05")>0 then%>checked<%end if%>>
管理</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>基本资料</td>
      <td><input type="checkbox" value="yes" name="manage06" <%if instr(manage,"06")>0 then%>checked<%end if%>>
管理</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>系统帮助</td>
      <td><input type="checkbox" value="yes" name="manage07" <%if instr(manage,"07")>0 then%>checked<%end if%>>
管理</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>上传文件管理</td>
      <td><input type="checkbox" value="yes" name="manage08" <%if instr(manage,"08")>0 then%>checked<%end if%>>
列表
  <input type="checkbox" value="yes" name="manage09" <%if instr(manage,"09")>0 then%>checked<%end if%>>
删除</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>首页浮动广告</td>
      <td><input type="checkbox" value="yes" name="manage10" <%if instr(manage,"10")>0 then%>checked<%end if%>>
管理</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>用户修改自己的密码</td>
      <td><input type="checkbox" value="yes" name="manage11" <%if instr(manage,"11")>0 then%>checked<%end if%>>
管理</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>友情链接</td>
      <td><input type="checkbox" value="yes" name="manage12" <%if instr(manage,"12")>0 then%>checked<%end if%>>
管理</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>首页广告</td>
      <td><input type="checkbox" value="yes" name="manage13" <%if instr(manage,"13")>0 then%>checked<%end if%>>
管理</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>页头图片</td>
      <td><input type="checkbox" value="yes" name="manage14" <%if instr(manage,"14")>0 then%>checked<%end if%>>
管理</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>主导航设置</td>
      <td><input type="checkbox" value="yes" name="manage55" <%if instr(manage,"55")>0 then%>checked<%end if%>>
        管理</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>页脚信息设置</td>
      <td><input type="checkbox" value="yes" name="manage56" <%if instr(manage,"56")>0 then%>checked<%end if%>>
        管理</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>学院概况管理</td>
      <td>学院简介信息设置</td>
      <td align="left"><input type="checkbox" value="yes" name="manage33" <%if instr(manage,"33")>0 then%>checked<%end if%>>
        列表
        <input type="checkbox" value="yes" name="manage54" <%if instr(manage,"54")>0 then%>checked<%end if%>>
        增加
        <input type="checkbox" value="yes" name="manage23" <%if instr(manage,"23")>0 then%>checked<%end if%>>
        修改,
        删除</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>专业设置管理</td>
      <td>专业信息设置</td>
      <td><input type="checkbox" value="yes" name="manage15" <%if instr(manage,"15")>0 then%>checked<%end if%>>
        列表
        <input type="checkbox" value="yes" name="manage16" <%if instr(manage,"16")>0 then%>checked<%end if%>>
        增加
        <input type="checkbox" value="yes" name="manage18" <%if instr(manage,"18")>0 then%>checked<%end if%>>
        修改，删除</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td rowspan="4">教学科研管理</td>
      <td>教师团队</td>
      <td><input type="checkbox" value="yes" name="manage19" <%if instr(manage,"19")>0 then%>checked<%end if%>>
        列表
        <input type="checkbox" value="yes" name="manage20" <%if instr(manage,"20")>0 then%>checked<%end if%>>
        增加
        <input type="checkbox" value="yes" name="manage21" <%if instr(manage,"21")>0 then%>checked<%end if%>>
        修改
        
        <input type="checkbox" value="yes" name="manage22" <%if instr(manage,"22")>0 then%>checked<%end if%>>
        删除</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>教师修改自己的信息</td>
      <td><input type="checkbox" value="yes" name="manage57" <%if instr(manage,"57")>0 then%>checked<%end if%>>
        管理</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>教学科研分类管理</td>
      <td>
<input type="checkbox" value="yes" name="manage17" <%if instr(manage,"17")>0 then%>checked<%end if%>>
管理</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>教学科研文章管理</td>
      <td>
        <input type="checkbox" value="yes" name="manage30" <%if instr(manage,"30")>0 then%>checked<%end if%>>
        列表
         <input type="checkbox" value="yes" name="manage40" <%if instr(manage,"40")>0 then%>checked<%end if%>>
         增加 
         
        <input type="checkbox" value="yes" name="manage26" <%if instr(manage,"26")>0 then%>checked<%end if%>>
        修改，删除</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td rowspan="2">党建工作管理</td>
      <td>栏目分类</td>
      <td><input type="checkbox" value="yes" name="manage59" <%if instr(manage,"59")>0 then%>checked<%end if%>>
管理</td>
      <td rowspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td>党建工作</td>
      <td><input type="checkbox" value="yes" name="manage24" <%if instr(manage,"24")>0 then%>checked<%end if%>>
列表
  <input type="checkbox" value="yes" name="manage25" <%if instr(manage,"25")>0 then%>checked<%end if%>>
增加
<input type="checkbox" value="yes" name="manage27" <%if instr(manage,"27")>0 then%>checked<%end if%>>
修改，删除</td>
    </tr>
    <tr>
      <td rowspan="2">实践教学管理</td>
      <td>实践教学管理</td>
      <td><input type="checkbox" value="yes" name="manage28" <%if instr(manage,"28")>0 then%>checked<%end if%>>
        列表
          <input type="checkbox" value="yes" name="manage29" <%if instr(manage,"29")>0 then%>checked<%end if%>>
          增加
          <input type="checkbox" value="yes" name="manage31" <%if instr(manage,"31")>0 then%>checked<%end if%>>
修改，删除
<!--<input type="checkbox" value="yes" name="manage32" <%if instr(manage,"32")>0 then%>checked<%end if%>>
审核--></td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>实践教学分类</td>
      <td><input type="checkbox" value="yes" name="manage58" <%if instr(manage,"58")>0 then%>checked<%end if%>>
        管理</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>新闻公告管理</td>
      <td>新闻公告信息</td>
      <td><input type="checkbox" value="yes" name="manage34" <%if instr(manage,"34")>0 then%>checked<%end if%>>
        列表
          <input type="checkbox" value="yes" name="manage35" <%if instr(manage,"35")>0 then%>checked<%end if%>>
          增加
          <input type="checkbox" value="yes" name="manage36" <%if instr(manage,"36")>0 then%>checked<%end if%>>
          修改
          
          <input type="checkbox" value="yes" name="manage37" <%if instr(manage,"37")>0 then%>checked<%end if%>>
删除</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>招生就业管理</td>
      <td>招生就业</td>
      <td><input type="checkbox" value="yes" name="manage38" <%if instr(manage,"38")>0 then%>checked<%end if%>>
        列表
          <input type="checkbox" value="yes" name="manage39" <%if instr(manage,"39")>0 then%>checked<%end if%>>
        增加
        <input type="checkbox" value="yes" name="manage41" <%if instr(manage,"41")>0 then%>checked<%end if%>>
        修改，删除        </td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td rowspan="2">师生天地管理</td>
      <td>分类管理</td>
      <td>
        <input type="checkbox" value="yes" name="manage42" <%if instr(manage,"42")>0 then%>checked<%end if%>>
        管理</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>文章管理</td>
      <td>
        <input type="checkbox" value="yes" name="manage43" <%if instr(manage,"43")>0 then%>checked<%end if%>> 
        列表
         <input type="checkbox" value="yes" name="manage44" <%if instr(manage,"44")>0 then%>checked<%end if%>> 
         增加 
        <input type="checkbox" value="yes" name="manage45" <%if instr(manage,"45")>0 then%>checked<%end if%>> 
        修改 
        <input type="checkbox" value="yes" name="manage46" <%if instr(manage,"46")>0 then%>checked<%end if%>> 
        删除
        <input type="checkbox" value="yes" name="manage47" <%if instr(manage,"47")>0 then%>checked<%end if%>>
        审核</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td rowspan="2">网络课程管理</td>
      <td>分类管理</td>
      <td><input type="checkbox" value="yes" name="manage48" <%if instr(manage,"48")>0 then%>checked<%end if%>>
        管理</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>课程管理</td>
      <td><input type="checkbox" value="yes" name="manage49" <%if instr(manage,"49")>0 then%>checked<%end if%>>
列表
  <input type="checkbox" value="yes" name="manage50" <%if instr(manage,"50")>0 then%>checked<%end if%>>
增加
<input type="checkbox" value="yes" name="manage51" <%if instr(manage,"51")>0 then%>checked<%end if%>>
修改

<input type="checkbox" value="yes" name="manage52" <%if instr(manage,"52")>0 then%>checked<%end if%>>
删除
<input type="checkbox" value="yes" name="manage53" <%if instr(manage,"53")>0 then%>checked<%end if%>>
审核</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td rowspan="2">创新专题管理</td>
      <td>创新专题管理</td>
      <td><input type="checkbox" value="yes" name="manage61" <%if instr(manage,"61")>0 then%>checked<%end if%>>
        列表
          <input type="checkbox" value="yes" name="manage62" <%if instr(manage,"62")>0 then%>checked<%end if%>>
          增加
          <input type="checkbox" value="yes" name="manage63" <%if instr(manage,"63")>0 then%>checked<%end if%>>
修改，删除
      </td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>创新专题分类</td>
      <td><input type="checkbox" value="yes" name="manage60" <%if instr(manage,"60")>0 then%>checked<%end if%>>
        管理</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>大文件上传管理</td>
      <td>大文件上传和列表</td>
      <td><input type="checkbox" value="yes" name="manage64" <%if instr(manage,"64")>0 then%>checked<%end if%>>
        管理</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td colspan="4" align="center"><input type="submit" name="Submit" value=" 保 存 ">
		<input type="hidden" name="edit" value="ok">
		<input type="hidden" name="admin" value="<%=admin%>"> <input name="input" type="button" value=" 返 回 " onClick="window.history.go(-1)" /></td>
      </tr>
    </tbody>
  </table>
  		</td>
	</tr>
</table>
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
	if request.form("manage38")="yes" then manage=manage+"|38"
	if request.form("manage39")="yes" then manage=manage+"|39"
	if request.form("manage40")="yes" then manage=manage+"|40"
	if request.form("manage41")="yes" then manage=manage+"|41"
	if request.form("manage42")="yes" then manage=manage+"|42"
	if request.form("manage43")="yes" then manage=manage+"|43"
	if request.form("manage44")="yes" then manage=manage+"|44"
	if request.form("manage45")="yes" then manage=manage+"|45"
	if request.form("manage46")="yes" then manage=manage+"|46"
	if request.form("manage47")="yes" then manage=manage+"|47"
	if request.form("manage48")="yes" then manage=manage+"|48"
	if request.form("manage49")="yes" then manage=manage+"|49"
	if request.form("manage50")="yes" then manage=manage+"|50"
	if request.form("manage51")="yes" then manage=manage+"|51"
	if request.form("manage52")="yes" then manage=manage+"|52"
	if request.form("manage53")="yes" then manage=manage+"|53"
	if request.form("manage54")="yes" then manage=manage+"|54"
	if request.form("manage55")="yes" then manage=manage+"|55"
	if request.form("manage56")="yes" then manage=manage+"|56"
	if request.form("manage57")="yes" then manage=manage+"|57"
	if request.form("manage58")="yes" then manage=manage+"|58"
	if request.form("manage59")="yes" then manage=manage+"|59"
	if request.form("manage60")="yes" then manage=manage+"|60"
	if request.form("manage61")="yes" then manage=manage+"|61"
	if request.form("manage62")="yes" then manage=manage+"|62"
	if request.form("manage63")="yes" then manage=manage+"|63"
	if request.form("manage64")="yes" then manage=manage+"|64"
	
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select * from Bs_User where UserName='"&admin&"'"
	rs.open sql,conn,1,3
		rs("admin_Manage")=manage
		rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing

	response.write "<script language='javascript'>"
	response.write "alert('管理权限设置成功！');"
	response.write "location.href='Bs_adminPermissions.asp?admin="&admin&"';"
	response.write "</script>"
end if


sub deladmin()

	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select * from Bs_User where UserName='"&deladm&"'"
	rs.open sql,conn,1,3
		if (rs.bof and rs.eof) then
			response.write "<script language='javascript'>"
			response.write "alert('操作失败，没有此管理员！');"
			response.write "location.href='Bs_adminPermissions.asp?admin="&admin&"';"
			response.write "</script>"
			response.end
		elseif deladm=request.cookies("buyok")("admin") then
			response.write "<script language='javascript'>"
			response.write "alert('您不能删除自己！');"
			response.write "location.href='Bs_adminPermissions.asp?admin="&admin&"';"
			response.write "</script>"
			response.end
		else
		rs.delete
		rs.update
			response.write "<script language='javascript'>"
			response.write "alert('您已成功删除管理员"&deladm&"');"
			response.write "location.href='Bs_adminPermissions.asp?admin="&admin&"';"
			response.write "</script>"
			response.end
		end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing

end sub
%>

<%htmlend%>
<script>
//列表表格隔行变色
var TbRow = document.getElementById("goaler");
if (TbRow != null){
   for (var i=0;i<TbRow.rows.length ;i++ )
     {
        if (TbRow.rows[i].rowIndex%2==1){
            TbRow.rows[i].className="TdOver";
        }
        else{
            TbRow.rows[i].className="TdDown";
        }
     }
}
</script>
