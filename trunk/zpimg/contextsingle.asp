<!--#include file="include/head.asp"-->
<!--#include file="include/ubb2html.asp"-->
<%
Dim ContextID,Form_Title,Form_Context,formID,i2,title,Page,Pages,AllCount,ShowCount,PastCount,listMethod,ip,title2,status,bgcolor,listpage,csspath,j,ismobile,lastupdatetime,tmpJs,isLock,threadmaster,showOrder,keywords,tags
ContextID=Cnum(Request.QueryString("id"))
Page=Max(Cnum(sqlshow(Request("page"))),1)
listpage=sqlshow(Request.QueryString("listpage"))
formID=ContextID
csspath=left(css_style,instr(css_style,".")-1)
showOrder=0  '1为正序 0为逆序

If ContextID=0 Then
  turnto ""&strUrl&""
End If
Enable_Anonymous=False
Sql="select * from Content where id=" & ContextID & " or Parent=" & ContextID & " order by Parent,PostTime"
OpenRs(Sql)
'ID号错误
If Rs.RecordCount=0 Then
  turnto ""&strUrl&""
End If
forumID=Rs("classid")
subclassID=Rs("subclass")
lastupdatetime=Rs("LastUpdateTime")
If rs("tags")<>"" Then 
keywords=replace(rs("tags"),"|",",")
Tags=replace(rs("tags"),"|","%' or Title like '%")
Else
Tags=""
End if
'定义楼主
threadmaster=Rs("Poster")
title=HTMLEncode(ChkBadWords(Rs("Title")))
if strlen(title)>56 then
title=CutStr(title,56)
end if
status=Application(systemkey&forumID)&"→"&title
Form_Title=title
If Rs.RecordCount=0 Then
	TurnTo ""&strUrl&""
End If

tempForum=ForumID
If ForumID="" Then tempForum=DefaultPost
tempClass=Application(SystemKey&"subclass"&tempForum)
If instr(tempClass,"|")>0 Then
subclassArray=split(tempClass,"|")
Else
subclassArray=split("分类","|")
End If

Sub HtmlTltle '浏览器标题
Response.write(""&Title&"_"&subclassArray(subclassid)&" - "&strSiteName&"")
End sub

Sub HtmlKeywords '关键字
Response.write keywords
End Sub

Sub HtmlDescription '描述
Response.write CutStr(delHtml(Rs("Content")),100)
End Sub

Sub getart()
'显示文章内容
num=Rs.RecordCount
Response.write ubb2html(Rs("Content"))
End Sub

'评论
Sub pinglun()
If Instr(1,Session("ReadIndex"),"[" & ContextID & "]")=0 Then
Rs("ClickNum")=Rs("ClickNum")+1
Rs.Update
Session("ReadIndex")=Session("ReadIndex") & "[" & ContextID & "]"
End If
Rs.Close
'点击数加1,不计同一人浏览结束
If ShowOrder=1 Then
Sql="select * from Content where Parent=" & ContextID & " order by PostTime"
Else
Sql="select * from Content where Parent=" & ContextID & " order by PostTime DESC"
End If
OpenRs Sql

'cookie记忆分页模式
If sqlShow(Request.QueryString("listMethod"))="all" Then
	Response.cookies(systemkey)("listMethod")="all"
	Response.Cookies(systemkey).Expires=Date+365
End If
If sqlShow(Request.QueryString("listMethod"))="page" Then
	Response.cookies(systemkey)("listMethod")="page"
	Response.Cookies(systemkey).Expires=Date+365
End If
listMethod=Request.cookies(systemkey)("listMethod")
AllCount=Rs.RecordCount
If listMethod = "" or listMethod= "page" Then
	Pages=Fix(AllCount/contextPagesize)
	If Pages*contextPagesize<AllCount Then
		Pages=Pages+1	
	End If
	PastCount=(Page-1)*contextPagesize
	If PastCount>=AllCount Then
		Page=1
		PastCount=0
	End If
	If AllCount>PastCount Then
		Rs.Move PastCount
	End If
	ShowCount=Min(AllCount-PastCount,contextPagesize)
Else
	ShowCount=AllCount
End If
'分页结束

dim Cnumber
Cnumber=contextPagesize*(Page-1)+contextPagesize
If Pages=1 Then Cnumber=ShowCount
If Page=Pages and Page>1 Then Cnumber=Allcount
'===================================评论内容
If Allcount>0 Then
For i=1 To ShowCount
If ShowOrder=1 Then
i2=contextPagesize*(Page-1)+i
Else
i2= Allcount-(contextPagesize*(Page-1)+i-1)
End If

Response.write("<div class=""plzz"">")
If User_Group="admin" Then
Response.write("<span><a href=""manage.asp?act=DeleteContext&id="&Rs("ID")&"&ContextID="&ContextID&""" onclick='javascript:return askdelete();' class=""red"">删除</a></span>"&i2&".&nbsp;<a title=""显示["&Rs("poster")&"]的基本信息"" href=""javascript:memberinfo('"&Rs("Poster")&"')"" class=""blue"">"&Rs("Poster")&"</a>&nbsp;("&Rs("PostTime")&")&nbsp;&nbsp;IP:"&Rs("IP")&"</div>")
Else
Response.write(""&i2&".&nbsp;<a title=""显示["&Rs("poster")&"]的基本信息"" href=""javascript:memberinfo('"&Rs("Poster")&"')"" class=""blue"">"&Rs("Poster")&"</a>&nbsp;<span class=""gray"">("&Rs("PostTime")&")</span></div>")
End If
Response.write("<div class=""plzw"">"&HTMLEncode(Rs("Content"))&"</div>")
Rs.MoveNext
Next
End If
Rs.Close
End Sub

Sub contextPages()
Response.write("<span>"&ShowPages(Pages,Page,"context.asp?id="&contextID&"")&"</span>[本主题共<font color='#cc0000'>"&allcount&"</font>条评论 | 每页显示<font color='#cc0000'>"&contextPagesize&"</font>条评论]")
End Sub

Sub pinglunfb() '发表评论
Response.write("<div id=""pinglunfbt"">")
If User_ID="" then 
Response.write("评论前，请先 <a href=""#"" class=""red"">登录</a> ！")
Else
Response.write("用户："&User_ID&"")
End If
Response.write("</div>")
Response.write("<form name=""form1"" method=""POST"" action=""edit.asp?act=reply""><input type=""hidden"" name=""artid"" value="""&ContextID&"""><input type=""hidden"" name=""forumid"" value="""&forumid&""">")
Response.write("<textarea rows=""12"" name=""content"" cols=""165"" class=""pltextarea"" id=""message""></textarea>")
Response.write("<div class=""qdfb"">验证码：<input name=""vl_code"" class=""button"" type=""text"" id=""vl_code"" size=""4"" maxlength=""4""> <img src=""include/GetCode.asp"">")
Response.write("<input type=""submit"" value="" 确认发表 "" name=""B3"" id=""postbtn"" class=""button"" ")
If User_ID<>"" Then Response.write "enabled" Else Response.write "disabled"
Response.write(" onclick=""javascript:mysubmit();""></div><div id=""waiting""></div>")
End sub
%>
<!--#include file="html/Header.html"-->
<!--#include file="html/ContextSingle.html"-->
<!--#include file="html/Footer.html"-->
<%Call CloseAll()%>