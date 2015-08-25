<!--#include file="../include/head.asp"-->
<!--#include file="../include/ofunctions.asp"-->
<!--#include file="../include/md5.asp"-->
<%
If User_Group<>"admin" Then
	Call ShowError("此页只有管理员才能访问！",1)
End If
dim target,boardid,boardid2,boardname,BoardType,ubb,boardmaster,issave,m,username,page,j,userlevel,upload,upfilelist,classid,ParentID
target=Request.QueryString("target")
act=Request.QueryString("act")%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
<body>

<%
'基本信息
'=================================================
If target="setup" Then
   Select Case act
   Case "save"
   dim txt
   txt=Request.Form("setup")
   If txt<>"" Then
   Call Save("../include/config.asp",Txt)
	Call ShowError("设置成功！",1)
   Else
	Call ShowError("设置出错！",1)
   End If
   End Select
End If
'栏目管理
'==================================================
If target="board"  Then
	Select Case act
	Case "new"
		boardid=Request.form("boardid")
		boardname=Request.Form("boardname")
		BoardType=Request.Form("BoardType")
		If boardid="" or boardname="" Then
           Call ShowError("栏目ID或栏目名称不能为空！",1)
		End If
		'判断栏目ID不能与以前的重复
		Sql="select * from board where boardid="&boardid&""
		OpenRs(sql)
		If Rs.recordcount>0 then
		   Call ShowError("栏目ID不能与现有栏目ID重复！",1)
		End If
		rs.close
		'增加新栏目
		Sql="select * from board where boardid=null"
		OpenRs(Sql)
		Rs.AddNew
		Rs("boardid")=boardid
		Rs("boardname")=boardname
		Rs("BoardType")=BoardType
		Rs.Update
		Application(SystemKey&"Loaded")=""
		Turnto "admin_show.asp?target=board"
	Case "edit"
	boardid=Request.QueryString("boardid")
	issave=Request.Form("issave")
	If issave="yes" Then
		boardid2=Request.form("boardid")
		If boardid2<>boardid Then
		sql="select * from board where boardid="&boardid2&""
		openrs(sql)
		If rs.recordcount>0 Then
				Call ShowError("栏目ID不能与现有栏目ID重复！",1)
		End If
		Rs.close
		End If
		boardname=Request.Form("boardname")
		BoardType=Request.Form("BoardType")
		sql="select * from board where boardid="&boardid&""
		openrs(sql)
		Rs("boardid")=boardid2
		Rs("boardname")=boardname
		Rs("BoardType")=BoardType
		Rs.Update
		Rs.Close
		If boardid<>boardid2 Then
		sql="update content set classid="&boardid2&" where classid="&boardid
		openrs(sql)
		'关联设置分类
		sql="update subclass set subparent="&boardid2&" where subparent="&boardid&""
		openrs(sql)
		'关联设置分类完毕
		End If
		Application(SystemKey&"Loaded")=""
		Turnto "admin_show.asp?target=board"
        Else
        sql="select * from board where boardid="&boardid&""
		openrs(sql)%>
<form method="POST" action="admin_all.asp?target=board&act=edit&boardid=<%=rs("boardid")%>">
	<table cellspacing="1" cellpadding="5" border="0" class="a23" bgcolor="#FFFFFF" align="center" width="606" height="64">
		<tr class="a1">
			<td width="50" height="26" align="center"><b>栏目ID</b></td>
			<td width="157" height="26" align="center"><b>栏目名称</b></td>
			<td width="157" height="26" align="center"><b>栏目类型</b></td>
			<td width="114" height="26" align="center"><b>操作</b></td>
		</tr>
		<tr class="a4">
			<td height="30" align="center">
			<p>
			<input type="hidden" name="issave" value="yes">
			<input type="text" name="boardid" size="3" class="input" value="<%=rs("boardid")%>"></p>
			</td>
			<td height="30" align="center">
			<input type="text" name="boardname" size="20" class="input" value="<%=rs("boardname")%>"></td>
			<td class="center">
			<select class="input" name="BoardType" id="BoardType">
			<option value="<%=rs("BoardType")%>"><%
			If rs("BoardType")=0 Then Response.Write("新闻")
			If rs("BoardType")=1 Then Response.Write("图片")
			If rs("BoardType")=2 Then Response.Write("视频")
			%></option>
			<option value="0">新闻</option>
			<option value="1">图片</option>
            <option value="2">视频</option>
			</select>
			</td>
			<td height="30" align="center">
			<input type="submit" value="确认修改" name="B1" class="button"></td>
		</tr>
	</table>
	
</form><%rs.close
End If%>
<%Case "delete"
		boardid=Request.QueryString("boardid")
		sql="select *  from content where classid="&boardid&""
		openrs(sql)
		If Rs.recordcount>0 then
			Call ShowError("栏目中还有文章不能删除！",1)
		End If
		rs.close
		sql="delete * from board where boardid="&boardid&""
		openrs(sql)
		'关联删除分类
		sql="delete * from subclass where subparent="&boardid&""
		openrs(sql)
		'关联删除完毕
		Application(SystemKey&"Loaded")=""
		Turnto "admin_show.asp?target=board"
	Case "lineup"
		dim m2,sql2,rs2
		sql="select * from board where boardid<998 order by boardid,id"
		openrs(sql)
		if rs.recordcount>0 then
		for	m=1 to rs.recordcount
			m2=rs("boardid")
			If m2<>m Then
			rs("boardid")=m
			sql2="update content set classid="&m&" where classid="&m2
	        Set Rs2=Server.CreateObject("Adodb.Recordset")
			Rs2.Open sql2,Conn,1,3
			'分类关联设置
			sql2="update subclass set subparent="&m&" where subparent="&m2
	        Set Rs2=Server.CreateObject("Adodb.Recordset")
			Rs2.Open sql2,Conn,1,3
			'分类关联设置完毕
			End If
			rs.movenext
		next
		end if
		Application(SystemKey&"Loaded")=""
		Turnto "admin_show.asp?target=board"
	Case "clear"
		dim adminpass
		adminpass=trim(request.form("adminpass"))
		boardid=Request.QueryString("boardid")
		if boardid="" then 
		boardid=request.form("boardid")
		end if
		
		if adminpass="" then%>
		<form method="POST" action="admin_all.asp?target=board&act=clear">
		<input type="hidden" name="boardid" value="<%=boardid%>">
		<p align="center">请再次输入管理员密码:<input type="password" name="adminpass" size="14" class="input">&nbsp;
		<input type="submit" value="确认" name="confirmadmin" class="button">
		</p>
		</form>
		<%
		response.end
		end if
		sql="select * from users where username='admin'"
		openrs sql
		if rs("userpass") = md5(adminpass) then
		rs.close
		sql="delete * from content where classid = "&boardid&""
		openrs sql
		Application(SystemKey&"Loaded")=""
		Response.write "<script language='javascript'>alert('栏目已清空！');history.go(-2);</script>"
		response.end
		else
		Call ShowError("密码错误！",1)
		end if
	Case "move"
		dim sourceboard,targetboard
		sourceboard=request.form("sourceboard")
		targetboard=request.form("targetboard")
		sql="update content set classid="&targetboard&" where classid="&sourceboard&""
		openrs sql
		Application(SystemKey&"Loaded")=""
		Call ShowError("移动成功！",1)
End Select
End If



'分类管理
'==================================================
If target="subclass"  Then
	Select Case act
	Case "new"
		dim subid,subname,subparent
		subid=Request.form("subid")
		subname=Request.Form("subname")
		subparent=Request.Form("subparent")
		if subid="" or subname="" then
			Call ShowError("分类ID或名称不能为空！",1)
		end if

		'判断分类ID不能与以前的重复
		Sql="select * from subclass where subnum="&subid&" and subparent="&subparent&""
		OpenRs(sql)
		If Rs.recordcount>0 then
			Call ShowError("分类ID不能与现有分类ID重复！",1)
		End If
		rs.close
		'增加新栏目
		Sql="select * from subclass where id=null"
		OpenRs(Sql)
		Rs.AddNew
		Rs("subnum")=subid
		Rs("subname")=subname
		RS("subparent")=subparent
		Rs.Update
		Application(SystemKey&"Loaded")=""
		Turnto "admin_show.asp?target=subclass"

	Case "edit"
		dim subnum,subnum2
		boardid=Request.QueryString("boardid")
		subnum=Request.QueryString("subnum")
		issave=Request.Form("issave")

		If issave="yes" Then
		subid=Request.form("id")
		subnum2=Request.form("subnum")
		subname=Request.form("subname")
		subparent=Request.form("subparent")

		If subnum2<>subnum Then
		sql="select * from subclass where subparent="&subparent&" and subnum="&subnum2&""
		openrs(sql)
		If rs.recordcount>0 Then
		Call ShowError("分类ID不能与现有分类ID重复！",1)
		End If
		Rs.close
		End If
		sql="select * from subclass where id="&subid&""
		openrs sql
		If rs.recordcount>0 Then
		Rs("subnum")=subnum2
		Rs("subname")=subname
		Rs.Update
		End If

		If subnum<>subnum2 Then
		sql="update content set subclass="&subnum2&" where subclass="&subnum&" and classid="&subparent&""
		openrs(sql)
		End If
		Application(SystemKey&"Loaded")=""
		Turnto "admin_show.asp?target=subclass"
        Else
        sql="select * from subclass where subparent="&boardid&" and subnum="&subnum&""
		openrs(sql)%>
		<form method="POST" action="admin_all.asp?target=subclass&act=edit&boardid=<%=rs("subparent")%>&subnum=<%=rs("subnum")%>">
		<table cellspacing="1" cellpadding="5" border="0" class="a23" bgcolor="#FFFFFF" align="center" width="500" height="64">
		<tr class="a1">
			<td width="44" height="26">
			<p align="center"><b>分类ID</b></p>
			</td>
			<td width="157" height="26">
			<p align="center"><b>分类名称</b></p>
			</td>
			<td align="center" width="122"><b>父栏目名称</b></td>
			<td width="96" height="26" align="center"><b>操作</b></td>
		</tr>
		<tr class="a4">
			<td height="30" align="center">
			<p>
			<input type="hidden" name="issave" value="yes">
			<input type="text" name="subnum" size="6" class="input" value="<%=rs("subnum")%>"></p>
			</td>
			<td height="30" align="center">
			<input type="text" name="subname" size="26" class="input" value="<%=rs("subname")%>"></td>
			<td align="center" width="122"><%=Application(SystemKey&rs("subparent"))%></td>
			<input type="hidden" name="subparent" value="<%=Rs("subparent")%>">
			<input type="hidden" name="id" value="<%=Rs("id")%>">
			<td height="30" align="center">
			<input type="submit" value="确认修改" name="B1" class="button"></td>
		</tr>
		<tr class="a3">
			<td colspan="7"></td>
		</tr>
	</table>
</form><%rs.close
End If

Case "delete"
		boardid=Request.QueryString("boardid")
		subnum=Request.QueryString("subnum")
		sql="select top 1 * from content where classid="&boardid&" and subclass="&subnum&""
		openrs(sql)
		If Rs.recordcount>0 then
			Call ShowError("分类中还有文章不能删除！请移动文章后再删除！",1)
		End If
		rs.close
		sql="delete * from subclass where subparent="&boardid&" and subnum="&subnum&""
		openrs(sql)
		Application(SystemKey&"Loaded")=""
		Turnto "admin_show.asp?target=subclass"
	Case "lineup"
		subparent=Request.QueryString("subparent")
		sql="select * from subclass where subparent="&subparent&" order by subnum"
		openrs(sql)
		if rs.recordcount>0 Then
		for	m=1 to rs.recordcount
			m2=rs("subnum")
			If m2<>m Then
			rs("subnum")=m
			sql2="update content set subclass="&m&" where classid="&subparent&" and subclass="&m2
	        Set Rs2=Server.CreateObject("Adodb.Recordset")
			Rs2.Open sql2,Conn,1,3
			End If
			rs.movenext
		next
		end if
		Application(SystemKey&"Loaded")=""
		Turnto "admin_show.asp?target=subclass"

	Case "move"
		dim sourceclass,targetclass
		sourceclass=request.form("sourceclass")
		targetclass=request.form("targetclass")
		subparent=request.form("subparent")
		sql="update content set subclass="&targetclass&" where subclass="&sourceclass&" and classid="&subparent&""
		openrs sql
		Application(SystemKey&"Loaded")=""
		Call ShowError("移动成功！",1)
End Select
End If

'=================================================
'用户管理
If target="users" Then
   If act="" Then act=request.form("action")
   username=request.querystring("username")
   page=request.querystring("page")
   if page="" then page=request.form("page")
   Select Case Act
   Case "delete"
   		if username="admin" then
   		Call ShowError("不能删除管理员！",1)
		End If
   		if username<>"" then
   		Sql="delete * from Content where poster='"&SqlShow(UserName)&"'"
		OpenRs Sql
		Sql="delete * from Users where username='" & SqlShow(UserName) & "'"
		OpenRs Sql
   		Sql="delete * from Ssb where poster='"&SqlShow(UserName)&"'"
		OpenRs Sql
		Turnto "admin_show.asp?target=users"
		end if
   		if len(Request("UserID")) then
		For i=1 to Request("UserID").count
		username=Request("UserID")(i)
   		if username="admin" then
   		Call ShowError("不能删除管理员！",1)
		End If
   		Sql="delete * from Content where poster='"&SqlShow(UserName)&"'"
		OpenRs Sql
		Sql="delete * from Users where username='" & SqlShow(UserName) & "'"
		OpenRs Sql
   		Sql="delete * from Ssb where poster='"&SqlShow(UserName)&"'"
		OpenRs Sql
		Next
		End If

        '创建管理员操作日志
		Sql="select * from SetLog"
		OpenRs(Sql)
		Rs.AddNew
		Rs("AdminName")=User_ID
		Rs("SetTxt")="删除用户["&SqlShow(UserName)&"]"
		Rs("SetTime")=Now
		Rs.Update
		Rs.Close

	Case "lock"
		if username="admin" then
   		Call ShowError("不能锁定管理员！",1)
		End If
		If Username<>"" then
		sql="update users set lockuser=1 where username='" & SqlShow(UserName) & "'"
		openrs sql

        '创建管理员操作日志
		Sql="select * from SetLog"
		OpenRs(Sql)
		Rs.AddNew
		Rs("AdminName")=User_ID
		Rs("SetTxt")="锁定用户["&SqlShow(UserName)&"]"
		Rs("SetTime")=Now
		Rs.Update
		Rs.Close

		Turnto "admin_show.asp?target=users&page="&page
		End If
		if len(Request("UserID")) then
		For i=1 to Request("UserID").count
		username=Request("UserID")(i)
		if username="admin" then
   		Call ShowError("不能锁定管理员！",1)
		End If
		sql="update users set lockuser=1 where username='" & SqlShow(UserName) & "'"
		openrs sql
		Next
		End If

	Case "unlock"
		if username<>"" then
		sql="update users set lockuser=0 where username='" & SqlShow(UserName) & "'"
		openrs sql

        '创建管理员操作日志
		Sql="select * from SetLog"
		OpenRs(Sql)
		Rs.AddNew
		Rs("AdminName")=User_ID
		Rs("SetTxt")="解除用户["&SqlShow(UserName)&"]锁定"
		Rs("SetTime")=Now
		Rs.Update
		Rs.Close

		Turnto "admin_show.asp?target=users&page="&page
		End If
		if len(Request("UserID")) then
		For i=1 to Request("UserID").count
		username=Request("UserID")(i)
		sql="update users set lockuser=0 where username='" & SqlShow(UserName) & "'"
		openrs sql
		Next
		end if

	Case "move"
		UserLevel=trim(request("UserLevel"))
		if len(Request("UserID")) then
		For i=1 to Request("UserID").count
		username=Request("UserID")(i)
			if username="admin" then
   				Call ShowError("不能重设管理员权限！",1)
			End If
		sql="update users set usergroup='"&UserLevel&"'  where username='" & SqlShow(UserName) & "'"
		openrs sql
		Next
		end if
		Application(SystemKey&"Loaded")=""
		Turnto "admin_show.asp?target=users&page="&page
	Case "initpass"
		if len(Request("UserID")) then
		For i=1 to Request("UserID").count
		username=Request("UserID")(i)
		sql="update users set userpass='49ba59abbe56e057' where username='"&SqlShow(UserName)&"'"
		openrs sql
		Next
		end if

        '创建管理员操作日志
		Sql="select * from SetLog"
		OpenRs(Sql)
		Rs.AddNew
		Rs("AdminName")=User_ID
		Rs("SetTxt")="重设用户["&SqlShow(UserName)&"]密码"
		Rs("SetTime")=Now
		Rs.Update
		Rs.Close

		Call ShowError("密码重设成功，初始密码为123456！",1)

	Case "vip"
		if len(Request("UserID")) then
		For i=1 to Request("UserID").count
		username=Request("UserID")(i)
		sql="update users set vip=1 where username='" & SqlShow(UserName) & "'"
		openrs sql
		Next
		end If
		
        '创建管理员操作日志
		Sql="select * from SetLog"
		OpenRs(Sql)
		Rs.AddNew
		Rs("AdminName")=User_ID
		Rs("SetTxt")="提升用户["&SqlShow(UserName)&"]VIP权限"
		Rs("SetTime")=Now
		Rs.Update
		Rs.Close

	Case "removevip"
		if len(Request("UserID")) then
		For i=1 to Request("UserID").count
		username=Request("UserID")(i)
		sql="update users set vip=0 where username='" & SqlShow(UserName) & "'"
		openrs sql
		Next
		end if
		
        '创建管理员操作日志
		Sql="select * from SetLog"
		OpenRs(Sql)
		Rs.AddNew
		Rs("AdminName")=User_ID
		Rs("SetTxt")="解除用户["&SqlShow(UserName)&"]VIP权限"
		Rs("SetTime")=Now
		Rs.Update
		Rs.Close

 End Select
		Response.write("<script language='javascript'>alert('设定成功！');window.location='admin_show.asp?target=users&page="&page&"';</script>")
		Response.End
End If
'========================================
'文章管理
Dim contextID
If target="article" Then
Act=Request.QueryString("act")
If act="" Then act=request.form("action")
ContextID=Cnum(Request.QueryString("id"))

Select Case Act
Case "deletetougao"
		Sql="delete * from Content where ID=" & ContextID & " or Parent=" & ContextID
		OpenRs(Sql)
		Turnto "admin_show.asp"
Case "deletesingle"
		Sql="delete * from Content where ID=" & ContextID & " or Parent=" & ContextID
		OpenRs(Sql)
		Call ShowError("删除成功！",1)
Case "delete"
		dim i2
		if len(Request("artID")) then
		For i=1 to Request("artID").count
			ContextID=Request("artID")(i)
			Sql="delete * from Content where ID=" & ContextID & " or Parent=" & ContextID
			OpenRs(Sql)
		Next
		End If
		Call ShowError("删除文章成功！",1)
Case "move"
		classid=trim(request("classid"))
		subclassid=trim(request("subclass"))
		if len(Request("artID")) then
		For i=1 to Request("artID").count
				ContextID=Request("artID")(i)
				Sql="select * from  Content where ID = " & ContextID & ""
				OpenRs(Sql)
				ParentID = Rs("Parent")
				Rs.Close
		sql="update content set classid="&classid&" where ID=" & ContextID & " or Parent=" & ContextID
		openrs sql
		Sql="update Content set subclass="&subclassid&" where ID = " & ContextID & ""
		OpenRs(Sql)
		Next
		end if
		Call ShowError("文章移动成功！",1)
Case "setbold"
		if len(Request("artID")) then
		For i=1 to Request("artID").count
				ContextID=Request("artID")(i)
				sql="update content set titleBold=1 where ID=" & ContextID & ""
		openrs sql
		Next
		end if
		Call ShowError("标题加粗成功！",1)
Case "resetbold"
		if len(Request("artID")) then
		For i=1 to Request("artID").count
				ContextID=Request("artID")(i)
				sql="update content set titleBold=0 where ID=" & ContextID & ""
		openrs sql
		Next
		end if
		Call ShowError("文章操作成功！",1)
Case "setcolor"
		classid=trim(request("colorid"))
		if len(Request("artID")) then
		For i=1 to Request("artID").count
			ContextID=Request("artID")(i)
			sql="update content set titlecolor='"&classid&"' where ID=" & ContextID & ""
		openrs sql
		Next
		end if
		Call ShowError("标题加色成功！",1)
Case "setclass"
		classid=trim(request("subclass"))
		if len(Request("artID")) then
		For i=1 to Request("artID").count
				ContextID=Request("artID")(i)
				Sql="update Content set subclass="&classid&" where ID = " & ContextID & ""
				OpenRs(Sql)
		Next
		end if
		Call ShowError("分类设定成功！",1)
End Select
End If
'========================================
'公告管理
If target="notice" Then
  	dim id,NTitle,Nurl,NPostTime,EnableNotice,OrderID
	id=trim(request.form("id"))
	NTitle=trim(request.form("NTitle"))
	Nurl=trim(request.form("Nurl"))
	NPostTime=trim(request.form("NPostTime"))	
	EnableNotice=trim(request.form("EnableNotice"))
	OrderID=trim(request.form("OrderID"))
	If EnableNotice="" Then EnableNotice=False Else EnableNotice=True
    Select Case act
	Case "new"
	Sql="select * from notice where id=null"
    openrs sql
	rs.addnew
	rs("NTitle")=NTitle
	rs("Nurl")=Nurl
	rs("OrderID")=Cnum(OrderID)
	rs("EnableNotice")=EnableNotice
	rs("NPostTime")=Now()
	rs.update
	Rs.close
   	Call NoticeSite()
	Application(SystemKey&"Loaded")=""
    'response.write"<script language=javascript>alert('消息发表成功！');window.opener.document.location.reload();window.close();</script>"
	Turnto "admin_show.asp?target=notice"
	
	Case "edit"
	issave=Request.Form("issave")
	If issave="yes" Then
	sql="select * from notice where id="&id&""
	OpenRs(sql)
	rs("NTitle")=NTitle
	rs("Nurl")=Nurl
	rs("OrderID")=Cnum(OrderID)
	rs("EnableNotice")=EnableNotice
	rs("NPostTime")=Now()
	Rs.update
	Rs.close
   	Application(SystemKey&"Loaded")=""
   	Call NoticeSite()
	Turnto "admin_show.asp?target=notice"
    else
	id=SqlShow(request.querystring("id"))
    sql="select * from notice where id="&id&""
    openrs(sql)
    
%>
<div class="DB">
<div class="L1">当前位置：<a href="admin_show.asp">管理面板</a> &gt; <a href="admin_show.asp?target=notice">消息管理</a> &gt; 修改消息</div>
<form name="FrontPage_Form3" method="POST" action="admin_all.asp?target=notice&act=edit" >
<input type="hidden" name="issave" value="yes">
<input type="hidden" name="id" value="<%=id%>">
<table cellspacing="1" cellpadding="5" border="0" class="a23" bgcolor="#FFFFFF" align="center">
<tr class="T4">
<td width="50" class="center">排序</td>
<td width="275" class="center">标题</td>
<td width="275" class="center">链接地址</td>
<td width="100" class="center">是否显示</td>
<td width="100" class="center">操作</td>
</tr>
<tr class="Tr1">
<td class="center"><input type="text" name="OrderID" class="IT3" value="<%=Rs("OrderID")%>"></td>
<td>&nbsp;&nbsp;<input type="text" name="NTitle" class="IT5" value="<%=Rs("NTitle")%>"></td>
<td>&nbsp;&nbsp;<input type="text" name="Nurl" class="IT5" value="<%=Rs("Nurl")%>"></td>
<td class="center">
<input type="checkbox" name="EnableNotice" 
<%If Rs("EnableNotice")=True Then Response.write " value=""ON"" checked"%>></td>
<td class="center"><input type="submit" value="确认添加" name="addurl" class="btn3"></td>
</tr>
</table>
<form>
</div>
<%
End If
	Case "delete"
	id=SqlShow(request.querystring("id"))
	Sql="delete * from notice where id = " & id & ""
	OpenRs(Sql)
	Case "enable"
	id=SqlShow(request.querystring("id"))
	Sql="update notice set enablenotice=1 where id = " & id & ""
	OpenRs(Sql)
	Case "disable"
	id=SqlShow(request.querystring("id"))
	Sql="update notice set enablenotice=0 where id = " & id & ""
	OpenRs(Sql)
    End Select
   	Call NoticeSite()
	Turnto "admin_show.asp?target=notice"
End If
'========================================
'友情链接管理
If target="friendsite" Then
  Select Case act
	Case "new"
	dim siteorder,sitename,siteurl
	siteorder=trim(request.form("siteorder"))
	sitename=trim(request.form("sitename"))
	siteurl=trim(request.form("siteurl"))
    sql="select * from friendsite where siteorder="&siteorder&""
	OpenRs(sql)
		If Rs.recordcount>0 then
			Call ShowError("ID不能与现有ID重复！",1)
		End If
	rs.close
	'增加新链接
	Sql="select * from friendsite where siteorder=null"
    openrs sql
	rs.addnew
	rs("siteorder")=siteorder
	rs("sitename")=sitename
	rs("siteurl")=siteurl
	rs.update
    Call friendsite()
	Turnto "admin_show.asp?target=friendsite"
	Case "edit"
	issave=Request.Form("issave")
	If issave="yes" Then
	dim siteid
	siteid=request.form("id")
	siteorder=request.form("siteorder")
	sitename=request.form("sitename")
	siteurl=request.form("siteurl")
    sql="select * from friendsite where id="&siteid&""
	OpenRs(sql)
	Rs("siteorder")=siteorder
	Rs("sitename")=sitename
	Rs("siteurl")=siteurl
	Rs.update
	Rs.close
   	Call friendsite()
	Turnto "admin_show.asp?target=friendsite"
    else
	siteorder=request.querystring("siteorder")
    sql="select * from friendsite where siteorder="&siteorder&""
    openrs(sql)
    
%>
<form name="FrontPage_Form3" method="POST" action="admin_all.asp?target=friendsite&act=edit" onSubmit="return FrontPage_Form3_Validator(this)" language="JavaScript">
	<table cellspacing="1" cellpadding="5" border="0" class="a23" bgcolor="#FFFFFF" align="center" width="631" id="table1">
		<tr class="a1">
			<td width="29" align="center"><b>ID</b></td>
			<td width="160" align="center"><b>站点名称</b></td>
			<td width="347" align="center"><b>站点地址</b></td>
			<td align="center" width="50"><b>操作</b></td>
		</tr>
		<tr class="a3">
			<input type="hidden" name="issave" value="yes">
			<td width="29" align="center">
			<input type="hidden" name="id" value="<%=rs("id")%>">
			<input name="siteorder" value="<%=rs("siteorder")%>" size="3" class="input"></td>
			
			<td width="160" name="sitename" align="center">
			<input type="text" name="sitename" size="20" maxlength="30" class="input"value="<%=trim(rs("sitename"))%>"></td>
			<td width="347" align="center">
			<input type="text" value="<%=trim(rs("siteurl"))%>" size="47" maxlength="100" class="input" name="siteurl">
			</td>
			<td align="center" width="50">
			<input type="submit" value="确认" name="confirmadmin" class="button">
			</td>
		</tr>
	</table>
</form>
<%
    End If
	Case "delete"
	siteorder=request.querystring("siteorder")
	Sql="delete * from friendsite where siteorder = " & siteorder & ""
	OpenRs(Sql)
	Call friendsite()
	Call ShowError("删除成功！",1)
  End Select

End If

'清除缓存
If target="remove" Then
Application.Contents.RemoveAll()
Call ShowError("缓存对象更新完毕！",1)
End If
%>
</p>
</body></html>