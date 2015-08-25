<!--#include file="include/head.asp"-->
<%
Dim Page,Pages,AllCount,ShowCount,PastCount,classicID,bgcolor,Form_Title,Form_Context,forumName,j,title,strlength,poster,orderStr,whereStr,lastupdateuser,status,postUrl,sepLine,BoardType
forumID=sqlshow(Request.QueryString("forumID"))
subclassID=sqlshow(Request.QueryString("subclassID"))
Page=Max(Cnum(sqlshow(Request.form("page"))),1)
Page=Max(Cnum(sqlshow(Request("page"))),1)
tempForum=ForumID
If ForumID="" Then tempForum=DefaultPost
tempClass=Application(SystemKey&"subclass"&tempForum)
If instr(tempClass,"|")>0 Then subclassArray=split(tempClass,"|")

Sql="select * from Board where BoardID="&forumID&" "
OpenRs(Sql)
BoardType=Rs("BoardType")
Rs.Close

If BoardType=1 Then ListPageSize=15  '图片列表分页数
If BoardType=2 Then ListPageSize=15  '视频列表分页数

Sub HtmlTltle '浏览器标题
Response.write(""&application(systemkey&forumid)&"_"&strSiteName&"")
End Sub

Sub HtmlKeywords '关键字
Response.write(""&application(systemkey&forumid)&","&strSiteName&"")
End Sub

Sub HtmlDescription '描述
Response.write strAbout
End Sub

Sub LieBiao '列表
'设定分类条件
If Cnum(subclassID)>0 Then whereStr="and subclass="&subclassID&"" Else whereStr=""

Sql="select * from Content where parent=0 and Verify=0 and classID="&forumID&" "&whereStr&" order by PostTime desc"
OpenRs(Sql)

'显示页数
AllCount=Rs.RecordCount
Pages=Fix(AllCount/ListPageSize)
If Pages*ListPageSize<AllCount Then Pages=Pages+1	
PastCount=(Page-1)*ListPageSize
If AllCount>PastCount Then Rs.Move PastCount
ShowCount=Min(AllCount-PastCount,ListPageSize)

If BoardType=1 Then Response.write("<div id=""PicList""><ul>")
If BoardType=2 Then Response.write("<div id=""VideoList""><ul>")

For i=1 To ShowCount
If BoardType<>1 and BoardType<>2 Then
	
	Response.write("<div class=""list""><div class=""lb4"">("&OnlyDate(Rs("PostTime"))&")</div><div class=""lb1""><a target=""_blank"" href=""context.asp?id="&Rs("ID")&""" title=""新窗口打开"">"&title&"</a></div><div class=""lb2"">")
	title=Rs("Title")
	If Rs("TitleColor")<>"" Then title="<font color="""&Rs("TitleColor")&""">"&title&"</font>"
	If Rs("TitleBold")=True Then title="<b>"&title&"</b>"
	'显示分类名称
	Response.write("[<a href=""list.asp?forumID="&Rs("classID")&"&subclassID="&Rs("subclass")&""">"&subclassArray(Rs("subclass"))&"</a>]")
	Response.write("</div>")
	Response.write("<div class=""lb3""><a href=""context.asp?id="&Rs("ID")&""">"&title&"</a>")
	If Rs("isPic")=True Then Response.write(" <span class=""green"">(图)</span>")
	Response.write("</div></div>")
	If i mod 10=0 Then Response.write("<div class=""br"">&nbsp;</div>")
	Rs.MoveNext
Else

	If BoardType=1 Then
	Dim pic_dy,cid
	'=======================================================================
	pic_dy=lcase(HTMLEncode(Rs("Content")))
	pic_dy=left(pic_dy,instr(pic_dy,"[/img]")-1)
	pic_dy=right(pic_dy,len(pic_dy)-4-instr(pic_dy,"[img]"))
	Response.Write"<li><a href=""contextSingle.asp?id="&Rs("ID")&"""><img src="""&pic_dy&""" alt="""&Rs("title")&"""><br>"&Rs("title")&"</a></li>"
	Rs.movenext
	
	Else
	'=======================================================================
	pic_dy=lcase(HTMLEncode(Rs("Content")))
	pic_dy=left(pic_dy,instr(pic_dy,"[/img]")-1)
	pic_dy=right(pic_dy,len(pic_dy)-4-instr(pic_dy,"[img]"))
	Response.Write"<li><a href=""context.asp?id="&Rs("ID")&"""><img src="""&pic_dy&""" alt="""&Rs("title")&"""><br>"&Rs("title")&"</a></li>"
	Rs.movenext
    End If

'=======================================================================
End If
Next
If BoardType=1 Then Response.write("</ul></div>")
If BoardType=2 Then Response.write("</ul></div>")
Rs.Close
End sub
%>

<!--#include file="html/Header.html"-->
<!--#include file="html/List.html"-->
<!--#include file="html/Footer.html"-->
<%Call CloseAll()%>
