<!--#include file="include/head.asp"-->
<%
Dim Page,Pages,PastCount,ShowCount,j,PageSize,AllCount,PageEndTime,PageRunTime,kw,fid,suserid,content,content_preview,title,BoardType,poster,lastupdateuser,strWhere,strOrder,status,strLength,allreplies,tempWords,strwords,startwords,strSql,zn
zn=SqlShow(Request("zn"))
kw=SqlShow(Request("kw"))
Page=Max(Cnum(Request("page")),1)
PageSize=15

Sub HtmlKeywords '�ؼ���
Response.write kw
End Sub

Sub HtmlDescription '����
Response.write kw
End Sub

If kw="" Then
Call ShowError("δ�����������ݣ����������룡",1)
Else
strSql="and title like '%"&kw&"%'"
If zn="1" Then strSql="and (title like '%"&kw&"%' or Content like '%"&kw&"%') "
End If

Sub HtmlTltle '���������
Response.write(""&kw&" - �������")
End sub

Sql="select * from content where parent=0 "&strSql&" order by id desc"
OpenRs(Sql)

AllCount=Rs.RecordCount
Pages=Fix(AllCount/PageSize)
If Pages*PageSize<AllCount Then Pages=Pages+1	
PastCount=(Page-1)*PageSize
If AllCount>PastCount Then Rs.Move PastCount
ShowCount=Min(AllCount-PastCount,PageSize) 

Sub liebiao() '�б�

If rs.recordcount = 0 Then response.write "<div style='text-align:center;'>δ�ҵ���������鹴ѡ�����������ٽ���ȫվ����!</div>"

For i=1 To ShowCount
title=Rs("Title")
BoardType=Rs("ClassId")
If Rs("TitleColor")<>"" Then title="<font color="""&Rs("TitleColor")&""">"&title&"</font>"
 
  If BoardType=2 Then
	 Response.write("<div class=""list""><div class=""lb4"">("&OnlyDate(Rs("PostTime"))&")</div><div class=""lb1""><a target=""_blank"" href="""&strUrl2s&"?id="&Rs("ID")&""" title=""�´��ڴ�"">"&title&"</a></div><div class=""lb2"">&nbsp;")
  else
	 Response.write("<div class=""list""><div class=""lb4"">("&OnlyDate(Rs("PostTime"))&")</div><div class=""lb1""><a target=""_blank"" href="""&strUrl2&"?id="&Rs("ID")&""" title=""�´��ڴ�"">"&title&"</a></div><div class=""lb2"">&nbsp;")
  end if

Response.write("</div>")
  If BoardType=2 Then
     Response.write("<div class=""lb3""><a href="""&strUrl2s&"?id="&Rs("ID")&""" target=""_blank"">"&title&"</a>")
  else
     Response.write("<div class=""lb3""><a href="""&strUrl2&"?id="&Rs("ID")&""" target=""_blank"">"&title&"</a>")
  end if

If Rs("isPic")=True Then Response.write(" <span class=""green"">(ͼ)</span>")
Response.write("</div></div>")
If i mod 5=0 Then Response.write("<div class=""br"">&nbsp;</div>")
Rs.MoveNext
Next
Rs.Close
End sub
%>
<!--#include file="html/Header.html"-->
<!--#include file="html/search.html"-->
<!--#include file="html/Footer.html"-->
<%Call CloseAll()%>
