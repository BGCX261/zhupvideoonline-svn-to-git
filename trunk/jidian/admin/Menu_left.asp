<!--#include file="bsconfig.asp"-->
<%
Set mrs = conn.Execute("select * from Bs_User where UserName='"&request.cookies("buyok")("admin")&"'")
if not (mrs.eof and mrs.bof) then
manage=mrs("admin_Manage")
end if
set mrs=nothing
%>
<html><head>
<meta http-equiv=Content-Type content=text/html;charset=gb2312>
<link href="Inc/bs.css" rel="stylesheet" type="text/css">
<script>

function aa(Dir)
{tt.doScroll(Dir);Timer=setTimeout('aa("'+Dir+'")',100)}//����100Ϊ�����ٶ�
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
top.document.title="����վ��̨����ϵͳ��"; 
</script>
<style type="text/css">
BODY              {BACKGROUND:799ae1; MARGIN: 0px;}

.sec_menu         {BORDER-RIGHT: white 1px solid; BACKGROUND: #eff7fd; OVERFLOW: hidden; BORDER-LEFT: white 1px solid; BORDER-BOTTOM: white 1px solid; width:100%;}
.sec_box          { width:100%;padding:1px;}
.sec_box td       {BORDER-BOTTOM: white 1px solid;  padding:4px 4px 4px 24px;background:url(adminimg/menuico.gif) no-repeat left;}
.menuTitle        { height:28px; background:#a4b6d7;text-align:center;}
.menu_title       { height:28px; background:#ebeffc url(adminimg/menudown.gif) no-repeat left;cursor:hand; padding:0 0 0 18px;}
.menu_title SPAN  {FONT-WEIGHT: bold; LEFT: 10px; COLOR: #215dc6; POSITION: relative; TOP: 2px }
.menu_title2      { height:28px;background:#c7d4f7 url(adminimg/menuup.gif) no-repeat left;cursor:hand;padding:0 0 0 18px;}
.menu_title2 SPAN {FONT-WEIGHT: bold; LEFT: 10px; COLOR: #428eff; POSITION: relative; TOP: 2px}

.sysTitle         { background:#039; height:32px; line-height:32px; font-size:14px; font-weight:bold; color:#a4b6d7; text-align:center;}
.menuArrow        {background:#a4b6d7;cursor:hand; color:#fff; font-family:Webdings; text-align:center; height:24px;}
</style>
</head>

<BODY>
<table cellspacing="0" cellpadding="0" width="100%" align="center">
	<tr>
		<td class="sysTitle">��վ��̨����˵�</td>
	</tr>
	<tr>
		<td class="menuTitle" >
		<span><a target="_top" href="default.asp">������ҳ</a> | <a target="_top" href="Loginout.asp">�˳�
		</a></span></td>
	</tr>
	<tr>
		<td onMouseOver="aa('up')" onMouseOut="StopScroll()" class="menuArrow">
		5</td>
	</tr>
</table>
<script>
var he=document.body.clientHeight-105
document.write("<div id=tt style=height:"+he+";overflow:hidden>")
</script>
<table cellspacing="0" cellpadding="0" width="100%" align="center">
	<tr>
		<td id="imgmenu1" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(1);" onMouseOut="this.className='menu_title';">
		<span>ϵͳ���� </span></td>
	</tr>
	<tr>
		<td id="submenu1" style="DISPLAY: none">
		<div class="sec_menu">
			<table class="sec_box">
            <%
if instr(manage,"01")>0 then
%>
				<tr><td><a target="main" href="Bs_Admin.asp">����Ա����</a></td></tr>
                <%
end if
if instr(manage,"04")>0 then
%>
                <tr><td><a target="main" href="Bs_adminPermissions.asp">����ԱȨ�޹���</a></td></tr>
                <%
end if
if instr(manage,"05")>0 then
%>
				<tr><td><a target="main" href="Bs_backup.asp">���ݿⱸ��</a></td></tr>
                <%
end if
if instr(manage,"06")>0 then
%>              <tr><td><a target="main" href="Bs_CoData.asp">��������</a></td></tr>
                <%
end if
if instr(manage,"07")>0 then
%>
              <tr><td><a target="main" href="Bs_notice.asp">��̨��������</a></td></tr>
                 <%
end if
if instr(manage,"11")>0 then
%>
              <tr><td><a target="main" href="Bs_Self_Pass_edit.asp?ID=<%=session("adminid")%>">�޸�����</a></td></tr>
                 <%
end if
if instr(manage,"12")>0 then
%>
              <tr><td><a target="main" href="Bs_Link.asp">�������ӹ���</a></td></tr>
                 <%
end if
if instr(manage,"13")>0 then
%>
              <tr><td><a target="main" href="Bs_adv.asp">��ҳ������</a></td></tr>
                 <%
end if
if instr(manage,"08")>0 then
%>
				<tr><td><a target="main" href="Bs_Upload_File.asp">�ϴ��ļ�����</a></td></tr>
                 <%
end if
if instr(manage,"10")>0 then
%>
				<tr><td><a target="main" href="Bs_advFloat.asp">����������</a></td></tr>
                 <%
end if
if instr(manage,"14")>0 then
%>
                <tr><td><a target="main" href="Bs_topPic.asp">ҳͷͼƬ����</a></td></tr>
                                 <%
end if
if instr(manage,"55")>0 then
%>
                <tr><td><a target="main" href="Bs_menuNav.asp">�������˵�����</a></td></tr>
                                                 <%
end if
if instr(manage,"56")>0 then
%>
                <tr><td><a target="main" href="Bs_copyRight.asp">ҳ����Ϣ����</a></td></tr>
                <%end if%>
			</table>
		</div>
		<br>
		</td>
	</tr>
    
    <tr>
		<td id="imgmenu2" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(2)" onMouseOut="this.className='menu_title';">
		<span>ѧԺ�ſ����� </span></td>
	</tr>
	<tr>
		<td id="submenu2" style="DISPLAY: none">
		<div class="sec_menu">
			<table class="sec_box">
                            <%
if instr(manage,"33")>0 then
%>
                <tr><td><a target="main" href="Bs_Profile.asp">�ſ����ù���</a></td></tr>
                            <%
end if
if instr(manage,"54")>0 then
%>
                <tr><td><a target="main" href="Bs_Profile_Add.asp">��Ӹſ���Ϣ</a></td></tr>
                             <%
end if
%>
			</table>
		</div>
		<br>
		</td>
	</tr>

    <tr>
		<td id="imgmenu3" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(3)" onMouseOut="this.className='menu_title';">
		<span>רҵ���ù��� </span></td>
	</tr>
	<tr>
		<td id="submenu3" style="DISPLAY: none">
		<div class="sec_menu">
			<table class="sec_box">
                            <%
if instr(manage,"15")>0 then
%>
                <tr><td><a target="main" href="Bs_proSet.asp">רҵ���ù���</a></td></tr>
                                <%
end if
if instr(manage,"16")>0 then
%>
				<tr><td><a target="main" href="Bs_proSet_Add.asp">���רҵ��Ϣ</a></td></tr>
                <%end if%>
			</table>
		</div>
		<br>
		</td>
	</tr>
    
    	<tr>
		<td id="imgmenu4" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(4)" onMouseOut="this.className='menu_title';">
		<span>��ѧ���й��� </span></td>
	</tr>
	<tr>
		<td id="submenu4" style="DISPLAY: none">
		<div class="sec_menu">
			<table class="sec_box">
                            <%
if instr(manage,"19")>0 then
%>
                <tr><td><a target="main" href="Bs_coursTeacher.asp">��ѧ�Ŷ�</a></td></tr>
                            <%
end if
if instr(manage,"20")>0 then
%>
				<tr><td><a target="main" href="Bs_coursTeacher_add.asp">��ӽ�ʦ</a></td></tr> 
                 <%
end if
if instr(manage,"57")>0 then
%>
				<tr><td><a target="main" href="Bs_self_coursTeacher_edit.asp?rName=<%=session("realname")%>">�޸Ľ�ʦ����</a></td></tr>
                            <%
end if
if instr(manage,"17")>0 then
%>        
                <tr><td><a target="main" href="Bs_scType.asp">������Ŀ���</a></td></tr>
                                            <%
end if
if instr(manage,"30")>0 then
%>   
				<tr><td><a target="main" href="Bs_Scientific.asp">�������¹���</a></td></tr>
                                                            <%
end if
if instr(manage,"40")>0 then
%> 
				<tr><td><a target="main" href="Bs_Scientific_add.asp">��ӿ�������</a></td></tr>
				<%end if%>
			</table>
		</div>
		<br>
		</td>
	</tr>
    
    	<tr>
		<td id="imgmenu5" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(5)" onMouseOut="this.className='menu_title';">
		<span>ʵ����ѧ���� </span></td>
	</tr>
	<tr>
		<td id="submenu5" style="DISPLAY: none">
		<div class="sec_menu">
			<table class="sec_box">
                            <%
if instr(manage,"58")>0 then
%>
                <tr><td><a target="main" href="Bs_facType.asp">ʵ��������</a></td></tr>
                                                <%
end if
if instr(manage,"28")>0 then
%>
				<tr><td><a target="main" href="Bs_facResource.asp">ʵ����ѧ��Ϣ����</a></td></tr>
                                <%
end if
if instr(manage,"29")>0 then
%>
				<tr><td><a target="main" href="Bs_facResource_add.asp">���ʵ����ѧ��Ϣ</a></td></tr>
                                
                <%end if%>
			</table>
		</div>
		<br>
		</td>
	</tr>  
	
	<tr>
		<td id="imgmenu6" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(6)" onMouseOut="this.className='menu_title';">
		<span>������ҵ���� </span></td>
	</tr>
	<tr>
		<td id="submenu6" style="DISPLAY: none">
		<div class="sec_menu">
			<table class="sec_box">
                            <%
if instr(manage,"38")>0 then
%>
                <tr><td><a target="main" href="Bs_recruit.asp">������ҵ��Ϣ����</a></td></tr>
                                <%
end if
if instr(manage,"39")>0 then
%>
                <tr><td><a target="main" href="Bs_recruit_Add.asp">���������ҵ��Ϣ</a></td></tr>
                <%end if%>
			</table>
		</div>
		<br>
		</td>
	</tr>
    
    <tr>
		<td id="imgmenu7" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(7)" onMouseOut="this.className='menu_title';">
		<span>������������ </span></td>
	</tr>
	<tr>
		<td id="submenu7" style="DISPLAY: none">
		<div class="sec_menu">
			<table class="sec_box">
<%
if instr(manage,"59")>0 then
%>
                <tr><td><a target="main" href="Bs_SubjCType.asp">������Ŀ����</a></td></tr>
<%end if
if instr(manage,"24")>0 then
%>
                <tr><td><a target="main" href="Bs_subjConstruct.asp">������������</a></td></tr>
<%
end if
if instr(manage,"25")>0 then
%>
				<tr><td><a target="main" href="Bs_subjConstruct_Add.asp">��ӵ���������Ϣ</a></td></tr>
<%
end if
%>
			</table>
		</div>
		<br>
		</td>
	</tr>
    
    <tr>
		<td id="imgmenu8" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(8)" onMouseOut="this.className='menu_title';">
		<span>רҵ������� </span></td>
	</tr>
	<tr>
		<td id="submenu8" style="DISPLAY: none">
		<div class="sec_menu">
			<table class="sec_box">
                <%
if instr(manage,"34")>0 then
%>
                <tr><td><a target="main" href="Bs_News.asp">רҵ�������</a></td></tr>
                                <%
end if
if instr(manage,"35")>0 then
%>
				<tr><td><a target="main" href="Bs_News_Add.asp">�����¹���</a></td></tr>
                <%end if%>
			</table>
		</div>
		<br>
		</td>
	</tr>
    
    <tr>
		<td id="imgmenu9" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(9)" onMouseOut="this.className='menu_title';">
		<span>ʦ����ع��� </span></td>
	</tr>
	<tr>
		<td id="submenu9" style="DISPLAY: none">
		<div class="sec_menu">
			<table class="sec_box">
                            <%
if instr(manage,"42")>0 then
%>
				<tr><td><a target="main" href="Bs_nbType.asp">��Ŀ���</a></td></tr>
                                <%
end if
if instr(manage,"43")>0 then
%>
				<tr><td><a target="main" href="Bs_stuZone.asp">���¹���</a></td></tr>
                                <%
end if
if instr(manage,"44")>0 then
%>
				<tr><td><a target="main" href="Bs_stuZone_add.asp">�������</a></td></tr>
                <%end if%>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu10" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(10)" onMouseOut="this.className='menu_title';">
		<span>����γ̹��� </span></td>
	</tr>
	<tr>
		<td id="submenu10" style="DISPLAY: none">
		<div class="sec_menu">
			<table class="sec_box">
                            <%
if instr(manage,"48")>0 then
%>
				<tr><td><a target="main" href="Bs_Class.asp">�γ��������</a></td></tr>
                                <%
end if
if instr(manage,"49")>0 then
%>
				<tr><td><a target="main" href="Bs_Article.asp">����γ̹���</a></td></tr>
                                <%
end if
if instr(manage,"50")>0 then
%>
				<tr><td><a target="main" href="Bs_Article_Add.asp">�������γ�</a></td></tr>
                                <%
end if
if instr(manage,"53")>0 then
%>
				<tr><td><a target="main" href="Bs_Article_Check.asp">�������γ�</a></td></tr>
                <%end if%>
			</table>
		</div>
		<br>
		</td>
	</tr>
    
    <tr>
		<td id="imgmenu11" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(11)" onMouseOut="this.className='menu_title';">
		<span>����ר����� </span></td>
	</tr>
	<tr>
		<td id="submenu11" style="DISPLAY: none">
		<div class="sec_menu">
			<table class="sec_box">
                            <%
if instr(manage,"60")>0 then
%>
                <tr><td><a target="main" href="Bs_innovType.asp">����������</a></td></tr>
                                                <%
end if
if instr(manage,"61")>0 then
%>
				<tr><td><a target="main" href="Bs_innovation.asp">������Ϣ����</a></td></tr>
                                <%
end if
if instr(manage,"62")>0 then
%>
				<tr><td><a target="main" href="Bs_innovation_add.asp">��Ӵ�����Ϣ</a></td></tr>
                                
                <%end if%>
			</table>
		</div>
		<br>
		</td>
	</tr> 

    <tr>
		<td id="imgmenu12" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(12)" onMouseOut="this.className='menu_title';">
		<span>���ļ��ϴ� </span></td>
	</tr>
	<tr>
		<td id="submenu12" style="DISPLAY: none">
		<div class="sec_menu">
			<table class="sec_box">
                <%
if instr(manage,"64")>0 then
%>
                <tr><td><a target="main" href="Bs_Upload_FileDown.asp">���ļ�����</a></td></tr>
				<tr><td><a target="main" href="Bs_Upload_File_Add.asp">�ϴ����ļ�</a></td></tr>
                <%end if%>
			</table>
		</div>
		<br>
		</td>
	</tr>
	
</table>
</div>
<table cellspacing="0" cellpadding="0" width="100%" align="center">
	<tr>
		<td onMouseOver="aa('Down')" onMouseOut="StopScroll()" class="menuArrow">
		6
        </td>
	</tr>
</table>
</BODY>
</HTML>
