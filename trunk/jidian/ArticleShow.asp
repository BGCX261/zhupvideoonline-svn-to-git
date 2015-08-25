<!--#include file="Inc/syscode.asp"-->
<!--#include file="Inc/top.asp"-->
<%
ShowSmallClassType=ShowSmallClassType_Article
dim ArticleID
ArticleID=trim(request("ArticleID"))
if ArticleId="" then
	response.Redirect("article.asp")
end if

sql="select * from Product where ArticleID=" & ArticleID & ""
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,3
if rs.bof and rs.eof then
	response.write"<SCRIPT language=JavaScript>alert('找不到此内容！');"
  response.write"javascript:history.go(-1)</SCRIPT>"
else	
	rs("Hits")=rs("Hits")+1
	rs.update
	if rs("hits")>=HitsOfHot then
		rs("Hot")=True
		rs.update
	end if
	BigClassName=rs("BigClassName")
	SmallClassName=rs("SmallClassName")
%>
<%
dim selclass
selclass=rs("BigClassName")
sqlClassWrd="select * from BigClass where BigClassName='" & selclass & "'"
Set rsClassWrd= Server.CreateObject("ADODB.Recordset")
rsClassWrd.open sqlClassWrd,conn,1,1
%>
<!-- #include file="Inc/Head.asp" -->
<div id="body">
  <table class="bodytable">
    <tr>
      <td class="bodyleft">
          <!-- #include file="L_article.asp" -->
      </td>
      <td class="bodyline"></td>
      <td class="bodyright">
      <!-- #include file="Inc/headAd.asp" -->
         <div class="mbox">
            <div class="artpath"><%
response.write "<img src='img/title_ico.gif' />  | 网络课程&nbsp;&gt;&gt;&nbsp;<a href='article.asp?BigClassName=" & rs("BigClassName") & "'>" & rs("BigClassName") & "</a>&nbsp;&gt;&gt;&nbsp;"
if rs("SmallClassName") & ""<>"" then
response.write "<a href='article.asp?BigClassName=" & rs("BigClassName")&"&SmallClassName=" & rs("SmallClassName") & "'>" & rs("SmallClassName") & "</a>&nbsp;&gt;&gt;&nbsp;"
end if
response.write rs("Title")
%></div>
            <div class="contentbox">
                  <div class="chptitletd"><%=rs("Title")%></div>
                  <div class="chphittd">点击数：<%=rs("Hits")%>&nbsp; 录入时间：<%= FormatDateTime(rs("UpdateTime"),2) %></div>
<%if rsClassWrd("Wrd1")<>"" then%>
                  <div class="uncertainfield"><%=trim(rsClassWrd("Wrd1"))%><%=rs("Zi1")%></div>
<%end if
if rsClassWrd("Wrd2")<>"" then%>	
                  <div class="uncertainfield"><%=trim(rsClassWrd("Wrd2"))%><%=rs("Zi2")%></div>
<%end if
if rsClassWrd("Wrd3")<>"" then%>	
                  <div class="uncertainfield"><%=trim(rsClassWrd("Wrd3"))%><%=rs("Zi3")%></div>
<%end if
if rsClassWrd("Wrd4")<>"" then%>	
                  <div class="uncertainfield"><%=trim(rsClassWrd("Wrd4"))%><%=rs("Zi4")%></div>
<%end if
if rsClassWrd("Wrd5")<>"" then%>	
                  <div class="uncertainfield"><%=trim(rsClassWrd("Wrd5"))%><%=rs("Zi5")%></div>
<%end if
if rsClassWrd("Wrd6")<>"" then%>	
                  <div class="uncertainfield"><%=trim(rsClassWrd("Wrd6"))%><%=rs("Zi6")%></div>
<%end if
if rsClassWrd("Wrd7")<>"" then%>	
                  <div class="uncertainfield"><%=trim(rsClassWrd("Wrd7"))%><%=rs("Zi7")%></div>		
<%end if
rsClassWrd.close%>
                  <div class="contentview"><%call ShowArticleContent()%> </div>
              </div>
            <div class="prt">【<a href='javascript:window.print()'>打印此页</a>】&nbsp;【<a href='javascript:history.back()'>返回</a>】 </div>
          </div>
        </td>
    </tr>
  </table>
</div>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>
<%
rsClassWrd.close
set rsClassWrd=nothing
%>
<%
end if
rs.close
set rs=nothing
call CloseConn()
%>
