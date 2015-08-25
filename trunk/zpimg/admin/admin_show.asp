<!--#include file="../include/head.asp"-->
<!--#include file="../include/ofunctions.asp"-->
<%
'文件说明：后台管理显示模块
If User_Group<>"admin" Then Call ShowError("此页只有管理员才能访问！",1)
dim bgcolor,m,ii,rs2,sql2,rs3,sql3,target,page,Pages,PageSize,AllCount,ShowCount,PastCount,sqlstr,j,orderstr,wherestr
target=Request.QueryString("target")
if target="" then target="info"
Page=Max(Cnum(Request("page")),1)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
<script language="javascript" src="../include/edit.js"></script>
<script language="javascript">

function askdelete()
{
	if(confirm("是否确定删除？"))
		return(true);
	else
		return(false);
}

function checkform()
{
	if(form1.boardid.value==""){
		alert(" 栏目ID不能为空！");
		form1.boardid.focus();
		return(false);
	}
	if(form1.boardname.value==""){
		alert("栏目名称不能为空！");
		form1.boardname.focus();
		return(false);
	}
	return(true);
}


function unselectall()
{
    if(document.myform.chkAll.checked){
	document.myform.chkAll.checked = document.myform.chkAll.checked&0;
    } 	
}

function CheckAll(form)
{
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
    }
}

function setclass()
{
var boardCount=<%=BoardCount%>
var subclass='<%=Application(SystemKey&"subclassAll")%>';
var pro_Class=new Array();
var pro_SubClass=subclass.split(',');

pro_Class[0]='--分类--';
pro_Class[boardCount+1]='<%=Application(Systemkey&"subclass999")%>';

for(z=1;z<=boardCount;z++) {
pro_Class[z]=pro_SubClass[z-1];
}

var nowSelectIndex=document.all("classid").selectedIndex+1;
for(i=document.all("subclass").length-1;i>=0;i--){document.all("subclass").options.remove(i);}

var Array_Class=pro_Class[nowSelectIndex].split('|');
   if(Array_Class.length>1){
   for(j=0;j<Array_Class.length;j++){
   document.all("subclass").options.add(new Option(Array_Class[j],j));
   document.all("subclass")[0].selected=true;
   }
  }else{
  document.all("subclass").options.add(new Option(pro_Class[nowSelectIndex],0));
  }
}

</script>
</head>

<body>

<%Select Case target%> <%
'===========================================
'基本信息
     Case "info"
     dim sitename%>
<div class="DB">
    <div class="L1">当前位置：管理面板 &gt;</div>
	<div class="D1">
		 <div class="T1">我的操作信息</div>
		 <ul class="U1">
		 <%
		 Sql="select * from Users where UserName='"&User_ID&"'"
		 OpenRs(Sql)
		 Response.write("<li>上次登录时间："&Rs("LastLogonTime")&"</li>")
		 Response.write("<li>上次登录IP："&Rs("Ip")&"</li>")
		 Rs.Close
		 %>
			  <li class="border-top"><a href="?target=SetLog">最近操作日志</a>：</li>
		 <%
		 Sql="select * from SetLog where AdminName='"&User_ID&"' order by id desc"
		 OpenRs(Sql)
		 If Rs.RecordCount>3 Then ii=3 Else ii=Rs.RecordCount
		 For i=1 To ii
		 Response.write("<li>"&Rs("SetTxt")&"（"&Rs("SetTime")&"）</li>")
		 Rs.Movenext
		 Next
		 Rs.Close
		 %>
	     </ul>
    </div>
	<div class="G1">&nbsp;</div>
	<div class="D1">
		 <div class="T1">快捷操作</div>
		 <ul class="U2">
		      <li><a href="admin_show.asp?target=ShengCheng">更新网站首页</a></li>
		      <li><a href="javascript:opennotice('ks')">发布滚动消息</a></li>
			  <li><a href="javascript:openhotkw('kw')">设置搜索热词</a></li>
			  <li><a href="javascript:openjifen('')">修改会员积分</a></li>
			  <li><a href="admin_all.asp?target=remove">更新网站缓存</a></li>
			  <li><a href="http://www.baidu.com/s?wd=site:<%=request.ServerVariables("HTTP_HOST")%>" target="_blank">百度收录情况</a></li>
	     </ul>
    </div>
	<div class="G2">&nbsp;</div>
	<div class="D1">
		 <div class="T1"><a href="admin_show.asp?target=Submission">最新投稿</a></div>
		 <ul class="U1">
		<%sql="select * from Content where Submission=-1 order by Verify asc,id desc"
    openrs(sql)
If Rs.RecordCount>6 Then ii=6 Else ii=Rs.RecordCount
For i=1 To ii
Response.write("<li>["&Rs("Poster")&"] <a href=""../context.asp?id="&rs("id")&""" target=""_blank"">"&Rs("Title")&"</a>")
If Rs("Verify")=-1 Then
Response.write(" <font color=""#FF0000"">未审核</font>")
%>
【<a href="admin_all.asp?target=article&act=deletetougao&id=<%=rs("id")%>" onClick="javascript:return confirm('确定删除该主题?(包括他的一切回复)');">删除</a> - <a href="../apost.asp?act=edit&id=<%=rs("id")%>">编辑</a>】
<%
End If
Response.write("</li>")
Rs.Movenext
Next
Rs.Close
%>
	     </ul>
    </div>
	<div class="G1">&nbsp;</div>
	<div class="D1">
		 <div class="T1">米号官方动态</div>
		 <ul class="U1"><iframe src="http://net.mihao.net/notice/index.html" frameborder="0" scrolling="no" height="160px" width="350px" noresize="noresize"></iframe></ul>
    </div>
</div>
<%

Function GetTotalSize(GetLocal,GetType)
	Dim FSO
	Set FSO=Server.CreateObject("Scripting.FileSystemObject")
	IF Err<>0 Then
		Err.Clear
		GetTotalSize="服务器关闭FSO，查看占用空间失败"
	Else
		Dim SiteFolder
		IF GetType="Folder" Then
			Set SiteFolder=FSO.GetFolder(GetLocal) 
		Else
			Set SiteFolder=FSO.GetFile(GetLocal) 
		End IF
		GetTotalSize=SiteFolder.Size
		IF GetTotalSize>1024*1024 Then
		GetTotalSize=GetTotalSize/1024/1024
		IF inStr(GetTotalSize,".") Then GetTotalSize = Left(GetTotalSize,inStr(GetTotalSize,".")+2)
			GetTotalSize=GetTotalSize&" MB"
		Else
			GetTotalSize=Fix(GetTotalSize/1024)&" KB"
		End IF
		Set SiteFolder=Nothing
	End IF
	Set FSO=Nothing
End Function

'==========================================
'网站设置
	Case "setup"
	dim txt
	txt=Read("../include/config.asp")
	%>
<p align="center"><font color="#FF0000"><b>网站基本设置</b></font> </p>
<form name="forumsetup" method="POST" action="admin_all.asp?target=setup&act=save"></form>
<table cellspacing="1" cellpadding="5" border="0" class="a23" bgcolor="#FFFFFF" align="center" width="631">
	<tr class="a1">
		<td>设置网站基本参数</td>
	</tr>
	<tr class="a4">
		<td><textarea rows="28" name="setup" cols="85"><%=txt%></textarea></td>
	</tr>
	<tr class="a1">
		<td height="95">
		<p align="center">
		<input type="submit" value="保存设置" name="savesetup" class="button"></p>
		</td>
	</tr>
</table>
<%'==========================================
'栏目管理
     Case "board"%>
<div class="DB">
    <div class="L1">当前位置：<a href="admin_show.asp">管理面板</a> &gt; 栏目管理</div>
    <form name="form1" method="POST" action="admin_all.asp?target=board&act=new" onSubmit="javascript:return checkform();">
	<table cellspacing="1" cellpadding="5" border="0" class="a23" bgcolor="#FFFFFF" align="center">
		<tr class="BoardTitle">
			<td width="30" class="center bold">ID</td>
			<td width="150" class="center bold">栏目名称</td>
			<td width="150" class="center bold">栏目类型</td>
			<td width="70" class="center bold">文章数量</td>
			<td width="150" class="center bold">操作</td>
		</tr>
		<%sql="select * from Board where BoardID<>999 order by BoardID"
    openrs(sql)
	For i=1 To Rs.RecordCount
	If i mod 2 = 0 Then 
          bgcolor = "UsersTr1"
        Else
         bgcolor = "UsersTr2" 
     End If
     %>
		<tr onMouseOver="this.className='ahover'" onMouseOut="this.className='<%=bgcolor%>'" class="<%=bgcolor%>">
			<td class="center"><%=rs("boardid")%></td>
			<td class="center"><%=rs("boardname")%></td>
			<td class="center">
<%
If rs("BoardType")=0 Then Response.Write("新闻")
If rs("BoardType")=1 Then Response.Write("图片")
If rs("BoardType")=2 Then Response.Write("视频")
%></td>
			<td class="center"><font style="font-size: 8pt"><%'显示该版文章信息
		sql2="select count(*) from content where classid="&rs("boardid")&"and parent=0"
        Set Rs2=Server.CreateObject("Adodb.Recordset")
        Rs2.Open sql2,Conn,1,3
        Response.write rs2(0)
        rs2.close

%></font> </td>
			<td class="center">
			<a href="admin_all.asp?target=board&act=edit&boardid=<%=rs("boardid")%>">
			修改</a> <span lang="en-us">&nbsp; </span><%If rs("boardid")<>999 Then%><a href="admin_all.asp?target=board&act=delete&boardid=<%=rs("boardid")%>" onClick="javascript:return askdelete();"> 
			删除</a> <%End If%>
			<a href="admin_all.asp?target=board&act=clear&boardid=<%=rs("boardid")%>" onClick="javascript:return confirm('确定清空本栏目吗？！！！');">&nbsp;清空</a>
			</td>
		</tr>
		<%
	rs.movenext
	Next
	%>
		<tr class="UsersTr2">
			<td class="center"><input type="hidden" name="save" value="yes">
			<input type="text" name="boardid" size="2" maxlength="4" class="input"></td>
			<td class="center"><input type="text" name="boardname" maxlength="20" class="IT2"></td>
			<td class="center">
			<select class="Sel" name="BoardType" id="BoardType">
			<option value="0">选择类型</option>
			<option value="0">新闻</option>
			<option value="1">图片</option>
            <option value="2">视频</option>
			</select>
			</td>
			<td class="center"></td>
			<td class="center" colspan="3"><span lang="en-us">&nbsp;<input type="submit" value="确定" name="addforum" class="btn2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<a href="admin_all.asp?target=board&act=lineup">
			<font color="#FF0000">重新排序</font></a></span></td>
		</tr>
	</table>
</form>
<form name="form2" method="POST" action="admin_all.asp?target=board&act=move">
	<p class="BoardFoot">转移<select class="Sel" name="sourceboard" id="sourceboard">
	<option value="998">选择栏目</option> <%for i=1 to BoardCount
	Response.write "<option value='"&i&"'>"&Application(systemkey&i)&"</option>"
	Next
	%></option>
	</select> 所有文章到 <select class="Sel" name="targetboard" id="targetboard">
	<option value="998">选择栏目</option> <%for i=1 to BoardCount
Response.write "<option value='"&i&"'>"&Application(systemkey&i)&"</option>"
Next
%></option>
	</select>&nbsp;&nbsp;<input type="submit" name="Submit1" value="执行" class="btn2"></p>
</form>
<p class="BoardFoot2">注意：栏目ID必须从1开始，并且是连续数字，如不连续可以点“重新排序”。栏目一经设定且其中含有内容时请不要轻易改动栏目ID！</p>
</div>
<%
'==========================================
'二级栏目管理
 Case "subclass"
 %>
<div class="DB">
    <div class="L1">当前位置：<a href="admin_show.asp">管理面板</a> &gt; 分类管理</div>
<table cellspacing="1" cellpadding="5" border="0" class="a23" bgcolor="#FFFFFF" align="center" width="631">
	<tr class="T3">
		<td width="29" class="center bold">ID</td>
		<td width="195" class="center bold">分类名称</td>
		<td width="89" class="center bold">发表/回复</td>
		<td width="122" class="center bold">父栏目名称</td>
		<td width="134" class="center bold">操作</td>
	</tr>
	<%FOR j=1 to boardcount
		If j mod 2 = 0 Then 
        bgcolor = "ClassBg2"
        Else
         bgcolor = "ClassBg1" 
	    End If
%>
	<tr>
		<td colspan="5" class="<%=bgcolor%>"><b><%=Application(SystemKey&j)%></b></td>
	</tr>
	<%sql="select * from subclass where subparent="&j&" order by subnum"
	    openrs(sql)
		If Rs.RecordCount>0 Then
		For i=1 To Rs.RecordCount
	     %>
		<tr onMouseOver="this.className='ahover'" onMouseOut="this.className='<%=bgcolor%>'" class="<%=bgcolor%>">
		<td width="29" class="center"><%=rs("subnum")%></td>
		<td width="195" class="center"><%=rs("subname")%></td>
		<td width="89" class="center"><%'显示该分类信息
		sql2="select count(*) from content where classid="&rs("subparent")&" and parent=0 and subclass="&Rs("subnum")&""
        Set Rs2=Server.CreateObject("Adodb.Recordset")
        Rs2.Open sql2,Conn,1,3
        Response.write rs2(0)
        rs2.close
        
        sql2="select sum(replynum) from content where classid="&rs("subparent")&"and parent=0 and subclass="&Rs("subnum")&""
        Set Rs2=Server.CreateObject("Adodb.Recordset")
        Rs2.Open sql2,Conn,1,3
        Response.write "/"&rs2(0)
        rs2.close

%></td>
		<td width="122" class="center"><%=Application(SystemKey&rs("subparent"))%></td>
		<td width="134" class="center">
		<a href="admin_all.asp?target=subclass&act=edit&boardid=<%=rs("subparent")%>&subnum=<%=Rs("subnum")%>">
		修改</a> <span lang="en-us">&nbsp; </span>
		<a href="admin_all.asp?target=subclass&act=delete&boardid=<%=rs("subparent")%>&subnum=<%=Rs("subnum")%>" onClick="javascript:return askdelete();">
		删除</a>&nbsp; </td>
	</tr>
	<%
rs.movenext
Next
Else
%>
	<tr class="<%=bgcolor%>">
		<td colspan="5" class="center">该栏目下无分类，请新建分类！</td>
	</tr>
	<%End If
Rs.Close
%>
	<form name="form<%=j%>" method="POST" action="admin_all.asp?target=subclass&act=new">
		<tr class="<%=bgcolor%>">
			<td width="29" class="center">
			<input type="hidden" name="save" value="yes">
			<input type="hidden" name="subparent" value="<%=j%>">
			<input type="text" name="subid" size="3" maxlength="4" class="IT3"></td>
			<td width="195" class="center">
			<input type="text" name="subname" maxlength="20" class="IT2">
			</td>
			<td colspan="3" class="center"><input type="submit" value="确认添加" name="addforum" class="btn3">&nbsp;&nbsp;<a href="admin_all.asp?target=subclass&act=lineup&subparent=<%=j%>">
			<font color="#FF0000">重新排序</font></a></td>
		</tr>
	</form>
	<form name="form_move" method="POST" action="admin_all.asp?target=subclass&act=move">
		<tr class="<%=bgcolor%>">
			<td colspan="5">
			<p align="center">转移<select class="Sel" name="sourceclass" id="sourceclass">
			<%
	If instr(Application(SystemKey&"subclass"&j),"|")>1 Then
	subclassArray=split(Application(SystemKey&"subclass"&j),"|")
	Else
	subclassArray=split("分类","|")
	End If
	for i=0 to ubound(subclassArray)
	Response.write "<option value='"&i&"'>"&subclassArray(i)&"</option>"
	Next
	%></select> 分类中到 <select class="Sel" name="targetclass" id="targetclass">
			<%for i=0 to ubound(subclassArray)
Response.write "<option value='"&i&"'>"&subclassArray(i)&"</option>"
Next
%></select> <input type="submit" name="Submit1" value="执行" class="btn2">
			</p>
			</td>
		</tr>
		<input type="hidden" name="subparent" value="<%=j%>">
	</form>
	<%
NEXT%>
</table>
<p class="div2">注意：分类ID表示的是分类的顺序，如不连续可以点“重新排序”。分类一经设定且其中含有内容时请不要轻易改动分类ID！</p>
</form>
</div>
<%
'=========================================================
'消息管理
Case "notice"
dim xx,SetStyle
xx=Request.QueryString("xx")

If xx="ks" Then
SetStyle="style=""display:none"""
End If
%>
<div class="DB">
<div class="L1" <%=SetStyle%>>当前位置：<a href="admin_show.asp">管理面板</a> &gt; 消息管理</div>
<form name="form1" method="POST" action="admin_all.asp?target=notice&act=new" >
<table cellspacing="1" cellpadding="5" border="0" class="a23" bgcolor="#FFFFFF" align="center">
<tr class="T4">
<td width="50" class="center bold">排序</td>
<td width="275" class="center bold">标题</td>
<td width="275" class="center bold">链接地址</td>
<td width="100" class="center bold">是否显示</td>
<td width="100" class="center bold">操作</td>
</tr>
<%
Sql="select * from notice order by OrderID DESC"
Openrs(sql)

PageSize=15
AllCount=Rs.RecordCount
Pages=Fix(AllCount/PageSize)
If Pages*PageSize<AllCount Then
Pages=Pages+1	
End If
PastCount=(Page-1)*PageSize
If PastCount>=AllCount Then
Page=1
PastCount=0
End If
If AllCount>PastCount Then
Rs.Move PastCount
End If
ShowCount=Min(AllCount-PastCount,PageSize)

For i=1 To ShowCount
If i mod 2 = 0 Then 
bgcolor = "Tr1"
Else
bgcolor = "Tr2" 
End If
%>
<tr onMouseOver="this.className='ahover'" onMouseOut="this.className='<%=bgcolor%>'" class="<%=bgcolor%>" <%=SetStyle%>>
<td class="center"><%=rs("OrderID")%></td>
<td>&nbsp;&nbsp;<a href="../<%=rs("Nurl")%>" target="_blank" title="消息发表时间：<%=rs("NPostTime")%>"><%=rs("NTitle")%></a></td>
<td>&nbsp;&nbsp;<a href="../<%=rs("Nurl")%>" target="_blank"><%=rs("Nurl")%></a></td>
<td class="center"><%If Rs("EnableNotice")=True Then%>
<a href="admin_all.asp?target=notice&act=disable&id=<%=rs("id")%>">开</a></p>
<%Else%>
<a href="admin_all.asp?target=notice&act=enable&id=<%=rs("id")%>">关</a>
<%End If%>
</td>
<td class="center"><a href="admin_all.asp?target=notice&act=edit&id=<%=rs("id")%>">修改</a> - <a href="admin_all.asp?target=notice&act=delete&id=<%=rs("id")%>" onClick="javascript:return askdelete();">删除</a></td>
</tr>
<%
Rs.Movenext
Next
%>
<tr class="Tr3">
<td class="center"><input type="text" name="OrderID" class="IT3"></td>
<td>&nbsp;&nbsp;<input type="text" name="NTitle" class="IT5"></td>
<td>&nbsp;&nbsp;<input type="text" name="Nurl" class="IT5"></td>
<td class="center">
<input type="checkbox" name="EnableNotice" value="ON" checked></td>
<td class="center"><input type="submit" value="确认添加" name="addurl" class="btn3"></td>
</tr>
<%
Rs.Close
%>
</table>

<p class="Pages3" <%=SetStyle%>>
<span>提示：排序数值越大，显示位置越靠前；点击标题或链接地址可查看链接内容</span>
<%If page>1 Then%>
<a href="?target=notice&page=<%=page-1%>">上一页</a>&nbsp;
<%End If%>
<%
Response.write ShowPages(Pages,Page,"?target=notice")
%>
<%If page<pages Then%>
&nbsp;<a href="?target=notice&page=<%=page+1%>">下一页</a>
<%End If%>
</p>
</div>
<%
'========================================================
'用户管理
Case "users"%> <a name="#up"></a>
<div class="DB">
    <div class="L1">当前位置：<a href="admin_show.asp">管理面板</a> &gt; 会员管理</div>
<form name="myform" method="POST" action="admin_all.asp?target=users">
	<table cellspacing="1" cellpadding="5" border="0" class="a23" bgcolor="#FFFFFF" align="center" width="631">
		<tr class="UsersTitle">
			<td width="50" class="center">　</td>
			<td width="104" class="center bold"><a href="admin_show.asp?target=users&sqlstr=name">用户名</a></td>
			<td width="69" class="center bold"><a href="admin_show.asp?target=users&sqlstr=usergroup">用户组</a></td>
			<td width="69" class="center bold"><a href="admin_show.asp?target=users&sqlstr=regtime">注册时间</a></td>
			<td width="33" class="center bold"><a href="admin_show.asp?target=users&sqlstr=postcount">发表</a></td>
			<td width="110" class="center bold"><a href="admin_show.asp?target=users&sqlstr=logintime">登录时间</a>/<a href="admin_show.asp?target=users&sqlstr=logincount">次数</a></td>
			<td width="37" class="center bold">状态</td>
			<td width="100" class="center bold">操作</td>
		</tr>
		<%
    sqlstr=Request.Querystring("sqlstr")
    Select Case sqlstr
    case "name" 
    	sql="select * from users order by username"
    case "regtime"
    	sql="select * from users order by regtime desc"
    case "postcount" 
    	sql="select * from users order by postnum desc"
    case "logintime" 
    	sql="select * from users order by lastlogontime desc"
    case "logincount"
    	sql="select * from users order by logoncount desc"
    case "usergroup" 
    	sql="select * from users order by usergroup"
    case "admin" 
    	sql="select * from users where usergroup='admin' order by username"
    case "master"
    	sql="select * from users where usergroup<>'admin' and usergroup<>'member' order by username"
    case "vip"
    	sql="select * from users where vip=True order by username"
    case "lock" 
		sql="select * from users where lockuser=True order by username"
    case "search" 
    	sql="select * from users where username like '%"&Request("searchname")&"%' order by regtime"
    	dim searchusername
    	searchusername=Request("searchname")
	case else
    	sql="select * from users order by regtime desc"
    End Select

    openrs(sql)
    PageSize=15
    AllCount=Rs.RecordCount
	Pages=Fix(AllCount/PageSize)
	If Pages*PageSize<AllCount Then
		Pages=Pages+1	
	End If
	PastCount=(Page-1)*PageSize
	If PastCount>=AllCount Then
		Page=1
		PastCount=0
	End If
	If AllCount>PastCount Then
		Rs.Move PastCount
	End If
	ShowCount=Min(AllCount-PastCount,PageSize)
	%>
		<input type="hidden" name="page" value="<%=page%>"><%For i=1 To ShowCount
	If i mod 2 = 0 Then 
         bgcolor = "UsersTr1"
       Else
         bgcolor = "UsersTr2" 
     End If
     %>
		<tr onMouseOver="this.className='ahover'" onMouseOut="this.className='<%=bgcolor%>'" class="<%=bgcolor%>">
			<td class="center">
			<input name="UserID" type="checkbox" onClick="unselectall()" id="UserID" value="<%=rs("UserName")%>"></td>
			<td width="104" class="center">
			<a href="javascript:memberinfo2('<%=rs("username")%>')"><%
		Response.write Rs("UserName")
		Response.write "</font></a>"
%></a></td>
			<td width="69" class="center"><%
			
			If Cnum(Rs("usergroup"))>0 Then 
				Response.write "<font color=red>"&Application(systemkey&Cnum(rs("usergroup")))&"</font>"  
				Else 
				Response.write Rs("usergroup") 
			End If
			If Rs("vip")=True Then
			Response.write "<br><font color=red>*VIP</font>"
			End If%> </td>
			<td width="69" class="center"><%=rs("regtime")%></td>
			<td width="33" class="center"><%=rs("PostNum")%></td>
			<td width="158" class="center"><%=rs("lastlogontime")%>[<font color="red"><%=rs("logoncount")%></font>]</td>
			<td class="center" width="29"><%if rs("lockuser")=TRUE then Response.write "<font color=red>锁定</font>" else Response.write "<font color=green>正常</font>"%></b></td>
			<td class="center">
			<a href="javascript:openjifen('<%=rs("UserName")%>')">积分</a> - <a href="admin_all.asp?target=users&act=delete&username=<%=rs("UserName")%>" onClick="javascript:return confirm('确定删除该用户?(包括他的一切痕迹)');">删除</a> - <%if rs("lockuser")=False Then
      Response.write "<a href=""admin_all.asp?target=users&sqlstr="&sqlstr&"&act=lock&page="&page&"&username="&rs("UserName")&""">锁定</a></td></tr>" 
      else 
      Response.write "<a href=""admin_all.asp?target=users&sqlstr="&sqlstr&"&act=unlock&page="&page&"&username="&rs("UserName")&""">解锁</a></td></tr>"
      End If
      rs.movenext
    Next%> </td>
		</tr>
		<tr class="a10">
			<td class="center">
			<input name="chkAll" type="checkbox" id="chkAll" onClick="CheckAll(this.form)" value="checkbox"></td>
			<td colspan="7" class="center">【操作】
			<input name="Action" type="radio" value="delete">删除&nbsp;
			<input name="Action" type="radio" value="unlock">解锁&nbsp;
			<input name="Action" type="radio" value="lock">锁定&nbsp;
			<input name="Action" type="radio" value="initpass">重设密码&nbsp;
			<input name="Action" type="radio" value="vip">设为VIP&nbsp;
			<input name="Action" type="radio" value="removevip">解除VIP&nbsp;
			<input name="Action" type="radio" value="move">移动到 
			<select class="Sel" name="UserLevel" id="UserLevel">
			<option value="member">普通会员</option>
			<option value="admin">管理人员</option>
			<%
                rs.close%></select>&nbsp;&nbsp;
			<input type="submit" name="Submit" value="执行" class="btn2">
			</td>
		</tr>
	</table>
</form>
<%'显示注册用户总数
    sql="select count(*) from users"
    openrs sql%>
<p class="UsersUp">
<span>共有位<font color="red"><%=rs(0)%></font>注册用户</span>
<a href="admin_show.asp?target=users&page=<%=page-1%>&sqlstr=<%=sqlstr%>">上一页</a>&nbsp;<%
Response.write ShowPages(Pages,Page,"admin_show.asp?target=users&sqlstr="&sqlstr&"")
%>&nbsp;<a href="admin_show.asp?target=users&page=<%=page+1%>&sqlstr=<%=sqlstr%>&searchname=<%=searchusername%>">下一页</a>
</p>
<form method="POST" action="admin_show.asp?target=users&sqlstr=search#up">
<p class="div2"><a href="admin_show.asp?target=users&sqlstr=">所有用户列表</a> | <a href="admin_show.asp?target=users&sqlstr=admin#up">管理员列表</a> | <a href="admin_show.asp?target=users&sqlstr=vip#up">VIP会员列表</a> | <a href="admin_show.asp?target=users&sqlstr=lock#up">锁定用户列表</a>&nbsp;&nbsp;<input type="text" name="searchname" size="20" class="IT2">&nbsp;<input type="submit" value="搜索" name="B1" class="btn2"></p>
</form>
<p class="div2">提示：点击列名称可以对该列进行排序；点击用户名可查看该会员的更多资料。</p>
<%
'========================================================
'文章管理
	Case "article"
	forumid=Request.QueryString("forumid")
	sqlstr=Request.Querystring("sqlstr")
	subclassid=Request.QueryString("subclassid")
	'读取子栏目
	tempClass=Application(SystemKey&"subclass"&ForumID)
	If instr(tempClass,"|")>0 Then
	subclassArray=split(tempClass,"|")
	Else
	subclassArray=split("分类","|")
	End If
	if forumid="" then
        If sqlstr="replynum" Then orderStr=" order by replynum desc"
    	If sqlstr="clicknum" Then orderStr=" order by clicknum desc"
		If sqlstr="" Then orderStr=" order by settop desc,PostTime desc"
		whereStr=" where parent=0" 
	Else
    	If sqlstr="replynum" Then orderStr= " order by replynum desc"
    	If sqlstr="clicknum" Then orderStr= " order by clicknum desc"
    	If sqlstr="classid" Or sqlstr="" Then orderStr= "  order by PostTime desc"
		whereStr=" where parent=0 and classid="&forumid&"" 
    End If
	If IsNumeric(subclassID)=True and Cnum(subclassID)>0 and Cnum(subclassID)<999 Then
	whereStr=whereStr&" and subclass="&subclassID&""
	End If
	If Cnum(subclassID)=999 Then
	whereStr=whereStr&" and subclass<1"
	End If
	sql="select * from content "&whereStr&orderStr
	openrs(sql)
		%>
<div class="DB">
    <div class="L1">当前位置：<a href="admin_show.asp">管理面板</a> &gt; 文章管理</div>
    <form name="myform" method="POST" action="admin_all.asp?target=article">
	<div class="D3">
		 <div class="T2">
		 <%If forumID="" Then%>
		 <a href="admin_show.asp?target=article">全部文章</a> <%For j=1 To BoardCount
				Print "| <a href=admin_show.asp?target=article&sqlstr=classid&forumID=" &j&  ">"&Application(systemkey&j)&"</a> "
			Next%>
	<%Else
	Response.write("&nbsp;<a href=admin_show.asp?target=article>全部栏目</a>&nbsp;|&nbsp;")
	for i=1 to ubound(subclassArray)
	If i=cnum(subclassID) Then 
	Response.write("&nbsp;-&nbsp;<b>"&subclassArray(i)&"</b>")
	Else
	response.write("&nbsp;-&nbsp;<a class='classnametitle' href=admin_show.asp?target=article&forumID="&forumID&"&subclassID="&i&">"&subclassArray(i)&"</a>")
	End If
	next
	response.write("&nbsp;-&nbsp;<a class='classnametitle' href=admin_show.asp?target=article&forumID="&forumID&"&subclassID=999>未分类</a>")
	End If%>
	</div>
	<table cellspacing="0" cellpadding="4" border="0" class="a23" bgcolor="#FFFFFF" align="center">
		<% 
	'设定每页显示文章数
	PageSize=18
    AllCount=Rs.RecordCount
	Pages=Fix(AllCount/PageSize)
	If Pages*PageSize<AllCount Then
		Pages=Pages+1	
	End If
	PastCount=(Page-1)*PageSize
	If PastCount>=AllCount Then
		Page=1
		PastCount=0
	End If
	If AllCount>PastCount Then
		Rs.Move PastCount
	End If
	ShowCount=Min(AllCount-PastCount,PageSize)
	%>
		<input type="hidden" name="page" value="<%=page%>"><%For i=1 To ShowCount
	If i mod 2 = 0 Then 
         bgcolor = "ArticleTr1"
       Else
         bgcolor = "ArticleTr2" 
     End If
     %>
		<tr onMouseOver="this.className='ahover'" onMouseOut="this.className='<%=bgcolor%>'" class="<%=bgcolor%>">
			<td width="33" align="center">&nbsp;&nbsp;&nbsp;&nbsp;<input name="artID" type="checkbox" onClick="unselectall()" id="artID" value="<%=rs("ID")%>"></td>
			<td width="680" align="left"><%
dim title,strlength
title=Rs("Title")
if forumid="" then strlength=40 else strlength=50
if strlen(title)>strlength then
title=CutStr(title,strlength)
end if
If Rs("TitleBold") = True Then title="<b>"&title&"</b>"
If Rs("TitleColor") <>"" Then
title="<font color="""&Rs("TitleColor")&""">"&title&"</font>"
Else
title=""&title&""
End If

'空栏目前面显示栏目名称，否则显示子栏目名称
If forumID=""  Then    
Response.write "&nbsp;<a href=admin_show.asp?target=article&sqlstr=classid&forumID=" & Rs("classID") &  ">" & "<font color=#666666>[" & Application(systemkey&Rs("classID")) &"]</font>" & "</a>"& "&nbsp;"
Else
If Rs("subclass")>0 and Rs("classid")=Cnum(forumID) Then Response.write("<font color=#666666>["&subclassArray(Rs("subclass"))&"]</font>&nbsp;")
End If
Response.write "<a target=""_blank"" href=""../context.asp?id=" & Rs("ID") & """>"  &  title & "</a>"
If Rs("SetTop")=1 Then
	Response.write "&nbsp;&nbsp;<img align=absmiddle border=""0"" src=""images/thread_top.gif"" width=""15"" height=""15"" alt=""专栏热点"">"
End If
If Rs("Classic")=1 Then
	Response.write "&nbsp;&nbsp;<img align=absmiddle border=""0"" src=""images/thread_digest.gif"" width=""15"" height=""15"" alt=""精华推荐"">"
End If
If Rs("lock")=true Then
	Response.write "&nbsp;&nbsp;<img align=absmiddle border=""0"" src=""images/thread_swf.gif"" width=""15"" height=""15"" alt=""热图幻灯片"">"
End If
			%></a></td>
			<td align="center" width="80">
			<a href="admin_all.asp?target=article&act=deletesingle&id=<%=rs("id")%>" onClick="javascript:return confirm('确定删除该主题?(包括他的一切回复)');">
			删除</a> - <a href="../apost.asp?act=edit&id=<%=rs("id")%>">编辑</a></td>
			<%rs.movenext%>
			</tr>
    <%Next
	%> 
		
		<tr class="a10">
			<td>&nbsp;&nbsp;&nbsp;&nbsp;<input name="chkAll" type="checkbox" id="chkAll" onClick="CheckAll(this.form)" value="checkbox"></td>
			<td colspan="3">&nbsp;&nbsp;<strong>操作：</strong>
			<input name="Action" type="radio" value="delete">删除 |&nbsp;
			<input name="Action" type="radio" value="resetbold">解除 |&nbsp;
			<input name="Action" type="radio" id="move" value="move" checked>移动到 
			<select onChange="javascript:setclass();" class="Sel" name="classid" id="classid"><%for i=1 to BoardCount
				Response.write "<option value='"&i&"'"
				If i=Cnum(forumID) Then
				Response.write " selected"
				End If
				Response.write ">"&Application(systemkey&i)&"</option>"
				Next
				Response.write "<option value='999'>"&Application(systemkey&999)&"</option>"
				rs.close%></select> <%
				Response.write("&nbsp;<select onchange=""javascript:document.all('move').checked=1;"" class=""Sel"" id=""subclass"" name=""subclass"">")
				For i=0 to ubound(subclassArray)
				Response.write("<option value="&i&">"&subclassArray(i)&"</option>")
				Next
				Response.write("</select>&nbsp;&nbsp;")
			%><input type="submit" name="Submit" value="执行" class="btn2">
			</td>
		</tr>
	</table>
</form>
</div>
<%
'显示文章总数
Sql="select count(*) from content where parent=0"
Openrs(sql)
Response.Write("<p class=""ArticleUp""><span>共有<font color=""red"">"&rs(0)&"</font>篇主题&nbsp;&nbsp;</span>")

If page>1 Then
Response.Write("<a href=""admin_show.asp?target=article&forumid="&forumid&"&page="&page-1&"&subclassID="&subclassID&""">上一页</a>")
End If

Response.write ShowPages(Pages,Page,"admin_show.asp?target=article&forumid="&forumid&"&sqlstr="&sqlstr&"&subclassID="&subclassID&"")


If page<pages Then
Response.Write("<a href=""admin_show.asp?target=article&forumid="&forumid&"&page="&page+1&"&subclassID="&subclassID&""">下一页</a>")
End If
Response.Write("</p></div>")
'========================================================
'评论管理
	Case "Comments"
	forumid=Request.QueryString("forumid")
	sqlstr=Request.Querystring("sqlstr")
	subclassid=Request.QueryString("subclassid")
	'读取子栏目
	tempClass=Application(SystemKey&"subclass"&ForumID)
	If instr(tempClass,"|")>0 Then
	subclassArray=split(tempClass,"|")
	Else
	subclassArray=split("分类","|")
	End If
	if forumid="" then
		if sqlstr="" then    	
			orderStr=" order by settop desc,lastupdatetime desc"
    	end if
		whereStr=" where ReplyNum>0" 
	else
    	if sqlstr="classid" or sqlstr="" then
		orderStr= "  order by PostTime desc"
		end if		
		whereStr=" where ReplyNum>0 and classid="&forumid&"" 
    end if 
	If IsNumeric(subclassID)=True and Cnum(subclassID)>0 and Cnum(subclassID)<999 Then
	whereStr=whereStr&" and subclass="&subclassID&""
	End If
	If Cnum(subclassID)=999 Then
	whereStr=whereStr&" and subclass<1"
	End If
	sql="select * from content "&whereStr&orderStr
	openrs(sql)
		%>
<div class="DB">
    <div class="L1">当前位置：<a href="admin_show.asp">管理面板</a> &gt; 评论管理</div>
	<div class="D3">
		 <div class="T2">
	<%If forumID="" Then%>
	<a href="admin_show.asp?target=Comments">全部评论</a> <%For j=1 To BoardCount
				Print "| <a href=admin_show.asp?target=Comments&sqlstr=classid&forumID=" &j&  ">"&Application(systemkey&j)&"</a> "
			Next%>
	<%Else
	Response.write("&nbsp;<a href=admin_show.asp?target=Comments>全部评论</a>&nbsp;|&nbsp;")
	for i=1 to ubound(subclassArray)
	If i=cnum(subclassID) Then 
	Response.write("&nbsp;-&nbsp;<b>"&subclassArray(i)&"</b>")
	Else
	response.write("&nbsp;-&nbsp;<a class='classnametitle' href=admin_show.asp?target=Comments&forumID="&forumID&"&subclassID="&i&">"&subclassArray(i)&"</a>")
	End If
	next
	response.write("&nbsp;-&nbsp;<a class='classnametitle' href=admin_show.asp?target=Comments&forumID="&forumID&"&subclassID=999>未分类</a>")
	End If%>
	</div>
	<table cellspacing="0" cellpadding="4" border="0" class="a23" bgcolor="#FFFFFF" align="center">
		<% 
	'设定每页显示文章数
	PageSize=18
    AllCount=Rs.RecordCount
	Pages=Fix(AllCount/PageSize)
	If Pages*PageSize<AllCount Then
		Pages=Pages+1	
	End If
	PastCount=(Page-1)*PageSize
	If PastCount>=AllCount Then
		Page=1
		PastCount=0
	End If
	If AllCount>PastCount Then
		Rs.Move PastCount
	End If
	ShowCount=Min(AllCount-PastCount,PageSize)
	%>
		<input type="hidden" name="page" value="<%=page%>"><%For i=1 To ShowCount
	If i mod 2 = 0 Then 
         bgcolor = "ArticleTr1"
       Else
         bgcolor = "ArticleTr2" 
     End If
     %>
		<tr onMouseOver="this.className='ahover'" onMouseOut="this.className='<%=bgcolor%>'" class="<%=bgcolor%>">
			<td width="440" align="left">&nbsp;&nbsp;<%
title=Rs("Title")
if forumid="" then strlength=40 else strlength=50
if strlen(title)>strlength then
title=CutStr(title,strlength)
end if
title=""&title&""

'空栏目前面显示栏目名称，否则显示子栏目名称
If forumID=""  Then    
Response.write "&nbsp;<a href=admin_show.asp?target=Comments&sqlstr=classid&forumID=" & Rs("classID") &  ">" & "<font color=#666666>[" & Application(systemkey&Rs("classID")) &"]</font>" & "</a>"& "&nbsp;"
Else
If Rs("subclass")>0 and Rs("classid")=Cnum(forumID) Then Response.write("<font color=#666666>["&subclassArray(Rs("subclass"))&"]</font>&nbsp;")
End If

Response.write "<a target=""_blank"" href=""../context.asp?id=" & Rs("ID") & """>"&title&"</a>"
			%></a></td>
			<%
      rs.movenext
    Next%>
		</tr>
</table>
</form>
</div>
<%
Response.Write("<p class=""ArticleUp"">")
If page>1 Then
Response.Write("<a href=""admin_show.asp?target=Comments&forumid="&forumid&"&page="&page-1&"&subclassID="&subclassID&""">上一页</a>")
End If
Response.write ShowPages(Pages,Page,"admin_show.asp?target=Comments&forumid="&forumid&"&sqlstr="&sqlstr&"&subclassID="&subclassID&"")
If page<pages Then
Response.write("<a href=""admin_show.asp?target=Comments&forumid="&forumid&"&page="&page+1&"&subclassID="&subclassID&""">下一页</a>")
End If
Response.write("</p></div>")
'========================================================
'链接管理
Case "friendsite"
%>
<form name="FrontPage_Form9" method="POST" action="admin_all.asp?target=friendsite&act=new" onSubmit="return FrontPage_Form9_Validator(this)" language="JavaScript">
<div class="DB">
    <div class="L1">当前位置：<a href="admin_show.asp">管理面板</a> &gt; 友情链接管理</div>
	<table cellspacing="1" cellpadding="5" border="0" class="a5" id="table1">
		<tr class="T4">
			<td width="29" class="center bold">ID</td>
			<td width="160" class="center bold">站点名称</td>
			<td width="281" class="center bold">站点地址</td>
			<td width="116" class="center bold">操作</td>
		</tr>
		<%sql="select * from friendsite order by siteorder"
    openrs(sql)
	For i=1 To Rs.RecordCount
	If i mod 2 = 0 Then 
          bgcolor = "Tr1"
        Else
         bgcolor = "Tr2" 
     End If
     %>
		<tr onMouseOver="this.className='ahover'" onMouseOut="this.className='<%=bgcolor%>'" class="<%=bgcolor%>">
			<td width="29" class="center"><%=rs("siteorder")%></td>
			<td width="160" class="center"><%=rs("sitename")%></td>
			<td width="281" class="center">
			<a href="<%=rs("siteurl")%>" target="_blank"><%=rs("siteurl")%></a></td>
			<td class="center" width="116">
			<p>
			<a href="admin_all.asp?target=friendsite&act=edit&siteorder=<%=rs("siteorder")%>">
			修改</a> <span lang="en-us">&nbsp; </span>
			<a href="admin_all.asp?target=friendsite&act=delete&siteorder=<%=rs("siteorder")%>" onClick="javascript:return askdelete();">
			删除</a></p>
			</td>
		</tr>
		<%
	rs.movenext
	Next
	%>
		<tr class="a10">
			<td width="29" class="center"><input type="text" name="siteorder" size="3" maxlength="4" class="IT3"></td>
			<td width="160" class="center"><input type="text" name="sitename" size="21" maxlength="30" class="IT2"></td>
			<td width="281" class="center"><input type="text" name="siteurl" size="38" maxlength="100" class="IT5" value="http://"></td>
			<td class="center"><input type="submit" value="确认添加" name="addurl" class="btn3"></td>
		</tr>
	</table></div>
</form>
<%
'========================================================
'页头图片管理
Case "topbg"
%>
<div class="DB">
  <div class="L1">当前位置：<a href="admin_show.asp">管理面板</a> &gt; 页头图片管理</div>
	<table cellspacing="1" cellpadding="5" border="0" class="a5" id="table1">
		<tr>
			<td width="120" class="center bold">图片预览</td>
			<td width="500"><img src="../images/topbg.jpg" width="500"></td>
		</tr>
        <tr>
			<td colspan="2">注意：<font color="#FF0000"> * 图片名必须为：topbg.jpg，请勿上传其它的,新上传的将覆盖旧的。图片大小1250px * 680px，<a href="../images/topbg.psd">点此下载psd模板</a>)</font></td>
		</tr>
        <tr> 
		<td colspan="2" align="center"><iframe name="ad" frameborder=0 width=100% height=60 scrolling=auto src=uploadc.asp></iframe></td>
	</tr>
	</table></div>

<%
'=========================================================
'查看操作日志
Case "SetLog"
%>
<div class="DB">
<div class="L1">当前位置：<a href="admin_show.asp">管理面板</a> &gt; 操作日志</div>
<table cellspacing="1" cellpadding="3" bgcolor="#cccccc" width="650">
<tr class="T4">
<td width="100" class="center bold">管理员名称</td>
<td width="400" class="center bold">操作内容</td>
<td width="150" class="center bold">操作时间</td>
</tr>
<%
Dim AdminName
sqlstr=Request.Querystring("sqlstr")
AdminName=Request.Querystring("AdminName")
Select Case sqlstr
case "name" 
sql="select * from SetLog where AdminName='"&AdminName&" order by id desc'"
case else
sql="select * from SetLog order by id desc"
End Select
Openrs(sql)

PageSize=18
AllCount=Rs.RecordCount
Pages=Fix(AllCount/PageSize)
If Pages*PageSize<AllCount Then
Pages=Pages+1	
End If
PastCount=(Page-1)*PageSize
If PastCount>=AllCount Then
Page=1
PastCount=0
End If
If AllCount>PastCount Then
Rs.Move PastCount
End If
ShowCount=Min(AllCount-PastCount,PageSize)

For i=1 To ShowCount
If i mod 2 = 0 Then 
bgcolor = "Tr1"
Else
bgcolor = "Tr2" 
End If
%>
<tr onMouseOver="this.className='ahover'" onMouseOut="this.className='<%=bgcolor%>'" class="<%=bgcolor%>">
<td width="100" class="center"><a href="?target=SetLog&sqlstr=name&AdminName=<%=rs("AdminName")%>"><%=rs("AdminName")%></a></td>
<td width="400">&nbsp;&nbsp;<%=rs("SetTxt")%></td>
<td width="150" class="center"><%=rs("SetTime")%></td>
</tr>
<%
Rs.Movenext
Next
Rs.Close
%>
</table>

<p class="Pages2">
<span>提示：点击管理员名称可以查看该管理员的操作日志</span>
<%If page>1 Then%>
<a href="?target=SetLog&page=<%=page-1%>&sqlstr=<%=sqlstr%>&AdminName=<%=AdminName%>">上一页</a>&nbsp;
<%End If%>
<%
Response.write ShowPages(Pages,Page,"?target=SetLog&sqlstr="&sqlstr&"&AdminName="&AdminName&"")
%>
<%If page<pages Then%>
&nbsp;<a href="?target=SetLog&page=<%=page+1%>&sqlstr=<%=sqlstr%>&AdminName=<%=AdminName%>">下一页</a>
<%End If%>
</p>
</div>
<%
'=========================================================
'投稿管理
Case "Submission"
%>
<div class="DB">
<div class="L1">当前位置：<a href="admin_show.asp">管理面板</a> &gt; 投稿管理</div>
<table cellspacing="1" cellpadding="4" bgcolor="#cccccc" width="650">
<tr class="T4">
<td width="100" class="center bold">投稿会员</td>
<td width="300" class="center bold">文章标题</td>
<td width="150" class="center bold">投稿时间</td>
<td width="100" class="center bold">操作</td>
</tr>
<%
sqlstr=Request.Querystring("sqlstr")
AdminName=Request.Querystring("AdminName")
Select Case sqlstr
case "name" 
sql="select * from Content"
case else
sql="select * from Content where Submission=-1 order by Verify asc,id desc"
End Select
Openrs(sql)

PageSize=18
AllCount=Rs.RecordCount
Pages=Fix(AllCount/PageSize)
If Pages*PageSize<AllCount Then
Pages=Pages+1	
End If
PastCount=(Page-1)*PageSize
If PastCount>=AllCount Then
Page=1
PastCount=0
End If
If AllCount>PastCount Then
Rs.Move PastCount
End If
ShowCount=Min(AllCount-PastCount,PageSize)

For i=1 To ShowCount
If i mod 2 = 0 Then 
bgcolor = "Tr1"
Else
bgcolor = "Tr2" 
End If
%>
<tr onMouseOver="this.className='ahover'" onMouseOut="this.className='<%=bgcolor%>'" class="<%=bgcolor%>">
<td width="100" class="center"><%=rs("Poster")%></td>
<td width="300">&nbsp;&nbsp;<a href="../context.asp?id=<%=rs("id")%>" target="_blank"><%=rs("Title")%></a></td>
<td width="150" class="center"><%=rs("PostTime")%></td>
<td width="100" class="center">
<%If Rs("Verify")=-1 Then%>
<a href="admin_all.asp?target=article&act=deletesingle&id=<%=rs("id")%>" onClick="javascript:return confirm('确定删除该主题?(包括他的一切回复)');">删除</a> - <a href="../apost.asp?act=edit&id=<%=rs("id")%>">编辑</a>
<%Else%>
<font color="#cccccc">已审核</font>
<%End If%>
</td></tr>
<%
Rs.Movenext
Next
Rs.Close
%>
</table>

<p class="Pages2">
<span>提示：已经通过审核的文章，请到文章管理进行管理</span>
<%If page>1 Then%>
<a href="?target=Submission&page=<%=page-1%>">上一页</a>&nbsp;
<%End If%>
<%
Response.write ShowPages(Pages,Page,"?target=Submission")
%>
<%If page<pages Then%>
&nbsp;<a href="?target=Submission&page=<%=page+1%>">下一页</a>
<%End If%>
</p>
</div>
<%
'=========================================================
'生成网站首页
Case "ShengCheng"
dim objXmlHttp,binFileData,objAdoStream 
set objXmlHttp = Server.CreateObject("Microsoft.XMLHTTP") 
objXmlHttp.open "GET","http://"&request.ServerVariables("HTTP_HOST")&"/index.asp",false 
objXmlHttp.send() 
binFileData = objXmlHttp.responseBody 
set objAdoStream = Server.CreateObject("ADODB.Stream") 
objAdoStream.Type = 1 
objAdoStream.Open() 
objAdoStream.Write(binFileData) 
objAdoStream.SaveToFile server.MapPath("../index.html"),2 
objAdoStream.Close() 
set objAdoStream=nothing 
set objXmlHttp=nothing 
'创建管理员操作日志
Sql="select * from SetLog"
OpenRs(Sql)
Rs.AddNew
Rs("AdminName")=User_ID
Rs("SetTxt")="更新网站首页"
Rs("SetTime")=Now
Rs.Update
Rs.Close
'生成完成返回
Response.write"<script language=javascript>alert('首页更新完成!');this.location.href='admin_show.asp';</script>"

Case Else

End Select 
Call CloseAll()
%>
</body>
</html>
