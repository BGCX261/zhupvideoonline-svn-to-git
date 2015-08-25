<!--#include file="Conn.asp"-->
<html><head>
<meta http-equiv=Content-Type content=text/html;charset=gb2312>
<link href="Inc/bs.css" rel="stylesheet" type="text/css">
<script>

function aa(Dir)
{tt.doScroll(Dir);Timer=setTimeout('aa("'+Dir+'")',100)}//这里100为滚动速度
function StopScroll(){if(Timer!=null)clearTimeout(Timer)}



function initIt(){
divColl=document.all.tags("DIV");
for(i=0; i<divColl.length; i++) {
whichEl=divColl(i);
if(whichEl.className=="child")whichEl.style.display="none";}
}
function expands(el) {
whichEl1=eval(el+"Child");
if (whichEl1.style.display=="none"){
initIt();
whichEl1.style.display="block";
}else{whichEl1.style.display="none";}
}
var tree= 0;
function loadThreadFollow(){
if (tree==0){
document.frames["hiddenframe"].location.replace("LeftTree.asp");
tree=1
}
}
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
imgmenu = eval("imgmenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
imgmenu.background="adminimg/menuup.gif";
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
imgmenu.background="adminimg/menudown.gif";
}
}
function loadingmenu(id){
var loadmenu =eval("menu" + id);
if (loadmenu.innerText=="Loading..."){
document.frames["hiddenframe"].location.replace("LeftTree.asp?menu=menu&id="+id+"");
}
}
top.document.title="∷网站后台管理系统∷"; 
</script>
<style type="text/css">
BODY {
	BACKGROUND:799ae1; MARGIN: 0px;
}

.sec_menu {
	BORDER-RIGHT: white 1px solid; BACKGROUND: #d6dff7; OVERFLOW: hidden; BORDER-LEFT: white 1px solid; BORDER-BOTTOM: white 1px solid
}

.menu_title SPAN {
	FONT-WEIGHT: bold; LEFT: 10px; COLOR: #215dc6; POSITION: relative; TOP: 2px 
}

.menu_title2 SPAN {
	FONT-WEIGHT: bold; LEFT: 10px; COLOR: #428eff; POSITION: relative; TOP: 2px
}
</style>
</head>
<%
Set mrs = conn.Execute("select * from Bs_User where adminName='"&request.cookies("buyok")("admin")&"'")
if not (mrs.eof and mrs.bof) then
manage=mrs("admin_Manage")
end if
set mrs=nothing
%>
<BODY>
<table cellspacing="0" cellpadding="0" width="158" align="center">
	<tr>
		<td valign="bottom"><img src="adminimg/title.gif" border="0"></td>
	</tr>
	<tr>
		<td class="menu_title" onMouseOver="this.className='menu_title2';" onMouseOut="this.className='menu_title';" background="adminimg/title_bg_quit.gif" height="25">
		<span><b><a target="main" href="sysadmin_view.asp"><font color="215DC6">回到首页</font></a></b> 
		| <a target="_top" href="Loginout.asp"><font color="215DC6"><b>退出
		</font></a></span></td>
	</tr>
	<tr>
		<td align="center" onMouseOver="aa('up')" onMouseOut="StopScroll()">
		<font face="Webdings" color="#FFFFFF" style=cursor:hand>5</font> </td>
	</tr>
</table>
<script>
var he=document.body.clientHeight-105
document.write("<div id=tt style=height:"+he+";overflow:hidden>")
</script>
<table cellspacing="0" cellpadding="0" width="158" align="center">
	<tr>
		<td id="imgmenu1" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(1);" onMouseOut="this.className='menu_title';" style="cursor:hand" background=adminimg/menudown.gif height="25">
		<span>系统管理 </span></td>
	</tr>

	<tr>
		<td id="submenu1" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
                      <% if instr(manage,"01")>0 then %>
				<tr><td><a target="main" href="Bs_Admin.asp">管理员管理</a></td></tr>
                      <%end if 
					  if instr(manage,"02")>0 then%>
                <tr><td><a target="main" href="adminPermissions.asp">管理员权限设置</a></td></tr>
                <tr><td><a target="main" href="Bs_Role.asp">身份角色设置</a></td></tr>
                      <%end if 
					  if instr(manage,"03")>0 then%>
                      <tr><td><a target="main" href="Bs_uBatchok_Add.asp">ok上传用户批量增加</a></td></tr>
                <tr><td><a target="main" href="Bs_uBatch_Add.asp">上传用户批量增加</a></td></tr>
                      <%end if 
					  if instr(manage,"04")>0 then%>
                <tr><td><a target="main" href="Bs_Admin_List.asp">管理员资料管理</a></td></tr>
                      <%end if 
					  if instr(manage,"05")>0 then%>
                <tr><td><a target="main" href='Bs_Admin_Pass_edit.asp?id=<%=session("adminId")%>'>修改登录密码</a></td></tr>
                      <%end if 
					  if instr(manage,"06")>0 then%>
				<tr><td><a target="main" href="Bs_backup.asp">数据库备份</a></td></tr>
                      <%end if 
					  if instr(manage,"07")>0 then%>
				<tr><td><a target="main" href="Bs_Help.asp">系统帮助</a></td></tr>
                      <%end if 
					  if instr(manage,"08")>0 then%>
                <tr><td><a target="main" href="Bs_topPic.asp">页头图片管理</a></td></tr>
                      <%end if 
					  if instr(manage,"09")>0 then%>
                <tr><td><a target="main" href="Bs_adv.asp">广告管理</a></td></tr>
                      <%end if 
					  if instr(manage,"10")>0 then%>
				<tr><td><a target="main" href="Bs_Upload_File.asp">上传文件管理</a></td></tr>
                      <%end if 
					  if instr(manage,"11")>0 then%>
				<tr><td><a target="main" href="Bs_Upload_FileDown.asp">视频作品管理</a></td></tr>
                      <%end if 
					  if instr(manage,"12")>0 then%>
				<tr><td><a target="main" href="sysadmin_count_list.asp">网站访问情况</a></td></tr>
                      <%end if%>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu2" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(2)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="adminimg/menudown.gif" height="25">
		<span>基本信息管理 </span></td>
	</tr>
	<tr>
		<td id="submenu2" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
                <%if instr(manage,"13")>0 then%>
				<tr><td><a target="main" href="Bs_CoData.asp">基本资料</a></td></tr>
                <%end if 
				 if instr(manage,"14")>0 then%>
				<tr><td><a target="main" href="Bs_Declare.asp">网站声明信息</a></td></tr>
                <%end if 
				if instr(manage,"15")>0 then%>
				<tr><td><a target="main" href="Bs_CoProfile.asp">基本介绍设置</a></td></tr>
                <%end if%>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu4" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(4)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="adminimg/menudown.gif" height="25">
		<span>视频文件上传管理 </span></td>
	</tr>
	<tr>
		<td id="submenu4" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
            <%if instr(manage,"16")>0 then%>
				<tr><td><a target="main" href="Bs_Class.asp">视频类别设置</a></td></tr>
            <%end if 
			if instr(manage,"17")>0 then%>
				<tr><td><a target="main" href="Bs_Article.asp">视频作品管理</a></td></tr>
            <%end if 
			if instr(manage,"21")>0 then%>
				<tr><td><a target="main" href="Bs_Article_Add.asp">添加上传视频</a></td></tr>
            <%end if 
			if instr(manage,"19")>0 then%>
				<tr><td><a target="main" href="Bs_Article_Check.asp">审核视频作品</a></td></tr>
            <%end if%>
			</table>
		</div>
		<br>
		</td>
	</tr>
	
	<tr>
		<td id="imgmenu5" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(5)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="adminimg/menudown.gif" height="25">
		<span>视频教程管理 </span></td>
	</tr>
	<tr>
		<td id="submenu5" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
                <%if instr(manage,"22")>0 then%>
				<tr><td><a target="main" href="Bs_nbType.asp">教程类别</a></td></tr>
                <%end if 
				if instr(manage,"23")>0 then%>
				<tr><td><a target="main" href="Bs_willbook.asp">教程管理</a></td></tr>
                <%end if 
				if instr(manage,"26")>0 then%>
				<tr><td><a target="main" href="Bs_willbook_add.asp">添加教程</a></td></tr>
                <%end if%>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu7" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(7)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="adminimg/menudown.gif" height="25">
		<span>新闻动态管理</span></td>
	</tr>
	<tr>
		<td id="submenu7" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
                <%if instr(manage,"27")>0 then%>
				<tr><td><a target="main" href="Bs_News.asp">新闻动态管理</a></td></tr>
                <%end if 
				if instr(manage,"30")>0 then%>
				<tr><td><a target="main" href="Bs_News_Add.asp">添加新闻动态</a></td></tr>
                <%end if%>
			</table>
		</div>
		<br>
		</td>
	</tr>
    
    	<tr>
		<td id="imgmenu8" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(8)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="adminimg/menudown.gif" height="25">
		<span>评论用户管理 </span></td>
	</tr>
	<tr>
		<td id="submenu8" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
            <%if instr(manage,"31")>0 then%>
				<tr><td><a target="main" href="Bs_User.asp">注册用户管理</a></td></tr>
                <%end if 
				if instr(manage,"32")>0 then%>
				<tr><td><a target="main" href="Bs_Pinglun_Check.asp">用户评论管理</a></td></tr>
                <%end if%>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu9" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(9)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="adminimg/menudown.gif" height="25">
		<span>其它信息管理 </span></td>
	</tr>
	<tr>
		<td id="submenu9" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
             <%if instr(manage,"35")>0 then%>
				<tr><td><a target="main" href="Bs_Link.asp">友情链接管理</a></td></tr>
                <%end if 
				if instr(manage,"36")>0 then%>
				<tr><td><a target="main" href="Bs_ADSheet.asp">新手上路设置</a></td></tr>
                <%end if
                if instr(manage,"37")>0 then%>
                <tr><td><a target="main" href="Bs_Setmail.asp">邮件服务器设置</a></td></tr>
                <%end if%>
			</table>
		</div>
		<br>
		</td>
	</tr>

	
</table>
</div>
<table cellspacing="0" cellpadding="0" width="158" align="center">
	<tr>
		<td align="center" onMouseOver="aa('Down')" onMouseOut="StopScroll()" valign="bottom">
		<font face="Webdings" color="#FFFFFF" style=cursor:hand>6</font> </td>
	</tr>
</table>
</BODY>
</HTML>
