<!--#include file="include/head.asp"-->
<!--#include file="include/dypicsub.asp"-->
<%
Dim j

Sub HtmlTltle '浏览器标题
Response.write strSiteName
End Sub

Sub HtmlKeywords '关键字
Response.write strKeyword
End Sub

Sub HtmlDescription '描述
Response.write strAbout
End Sub

Sub Photo '热门图片
dim pic_dy,cid
Sql="Select Top 5 * From content WHERE Parent=0 and ClassId<>3 and  content like '%[[]img]%.jpg[[]/img]%' order by PostTime DESC,id DESC"
Openrs(sql)
If Rs.eof and Rs.bof then
Response.Write"<img src=""images/showerr.gif"" width=""138"" height=""138"">"
Else
do while not Rs.eof
pic_dy=lcase(HTMLEncode(Rs("Content")))
pic_dy=left(pic_dy,instr(pic_dy,"[/img]")-1)
pic_dy=right(pic_dy,len(pic_dy)-4-instr(pic_dy,"[img]"))
If Rs("parent")=0 Then cid=Rs("ID") Else cid=Rs("parent")&CHR(38)&"listMethod=search#"&Rs("ID")
Response.Write"<li><a href=""contextsingle.asp?id="&cid&"""><img src="""&pic_dy&""" width=""160"" height=""126"" alt="""&Rs("title")&"""><br>"&Rs("title")&"</a></li>"
Rs.movenext
loop
End If
Rs.Close
End Sub

Sub NewsHot(number) '新闻热点
dim wherestr,orderstr,title,BoardType,showdate,dyurl,dylm,dydao,neworold,neworcommon,poster,tempClass,subclassArray

Sql="Select Top "&Cnum(number)&" Content.Id,Content.Title,Content.Content,Content.ClassID,Content.TitleColor From content WHERE hide=True order by posttime desc,id desc"
Openrs(sql)

If Rs.eof and Rs.bof Then
	Response.Write "<center>该栏目下没有内容！</center>"
	Rs.Close
	Exit Sub
End If

For i = 1 to Rs.Recordcount
title=Rs("Title")
BoardType=Rs("ClassId")
if strlen(title)>48 Then title=CutStr(title,48)
If Rs("TitleColor")<>"" Then title="<font color="""&Rs("TitleColor")&""">"&title&"</font>"
  If BoardType=2 Then
    Response.write "<p class=""IndexHotTitle""><a href=""contextsingle.asp?id="&Rs("ID")&""">"&title&"</a></p><p class=""IndexHotText"">"&CutStr(delHtml(Rs("Content")),100)&"</p>"
  else
    Response.write "<p class=""IndexHotTitle""><a href=""context.asp?id="&Rs("ID")&""">"&title&"</a></p><p class=""IndexHotText"">"&CutStr(delHtml(Rs("Content")),100)&"</p>"
  end if
Rs.Movenext
Next
Rs.Close
End Sub
%>
<!--#include file="html/Header.html"-->
<!--#include file="html/Index.html"-->
<!--#include file="html/Footer.html"-->
<%Call CloseAll()%>
