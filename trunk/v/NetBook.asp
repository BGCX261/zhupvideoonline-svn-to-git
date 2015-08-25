<!--#include file="Inc/syscode.asp" -->
<script language="JavaScript" src="js/mxxCalendar.js" type="text/javascript"></script>
<%If Session("UserName")="" Then
response.redirect"Server.asp"
end if%>
<%
Id=Session("UserName")
set Rst = Server.CreateObject("ADODB.recordset")
sql="select * from User where UserName='"&Id&"'"
rst.open sql,conn,1,1
%>
<!--#include file="Inc/Top.asp" -->
<!-- #include file="Inc/Head.asp" -->
<div id="body">
  <table class="bodyTable">
    <tr>
      <td class="bodyLeft">
        <!-- #include file="L_vip.asp" -->
      </td>
      <td class="bodyLine"></td>
      <td class="bodyRight">
          <div class="navPath">个人评论信息中心</div>
              <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td><div align="center"><img src="img/newsico.gif" width="12" height="13" />&nbsp;&nbsp;<b><%=Session("UserName")%></b> 的评论信息</div></td>
                </tr>
                <tr>
                  <td>&nbsp;&nbsp;&nbsp;&nbsp;<%=Session("UserName")%>，这里是您的评论信息空间，您可在此查看您的留言审核状态，我们会尽力在最快的时间审核，如遇特殊情况请联系管理员！
                   </td>
                </tr>
              </table>
              <%

name=Session("UserName")
set rs=server.createobject("adodb.recordset")
sql="select * from buyok_pinglun where name='"&name&"' order by id desc"
rs.open sql,conn,1,1

dim MaxPerPageB
MaxPerPageB=10

'假如没有数据时
'If rs.eof and rs.bof then
'call showpages
'response.write "<p align='center'><font color='#ff0000'>您还没任何留言资料</font></p>"
'response.end
'End if

'取得页数,并判断留言输入的是否数字类型的数据，如不是将以第一页显示
dim text,checkpage
text="0123456789"
Rs.PageSize=MaxPerPageB
for i=1 to len(request("page"))
checkpage=instr(1,text,mid(request("page"),i,1))
if checkpage=0 then
exit for
end if
next

If checkpage<>0 then
If NOT IsEmpty(request("page")) Then
CurrentPage=Cint(request("page"))
If CurrentPage < 1 Then CurrentPage = 1
If CurrentPage > Rs.PageCount Then CurrentPage = Rs.PageCount
Else
CurrentPage= 1
End If
If not Rs.eof Then Rs.AbsolutePage = CurrentPage end if
Else
CurrentPage=1
End if

call showpages
call list

If Rs.recordcount > MaxPerPageB then
call showpages
end if

'显示帖子的子程序
Sub list()

%>
              <div class="turnPage"><%
Response.write "信息列表"
Response.write "<span class='pageTxt'>共</span>" & "<span class='redTxt'>" & Cstr(Rs.RecordCount) & "</span>" & "<span class='pageTxt'>条留言</span>"
Response.write "<span class='pageTxt'>第</span>" & "<span class='redTxt'>" & Cstr(CurrentPage) &  "</span>" & "/<span class='pageTxt'>" & Cstr(rs.pagecount) & "</span>"
If currentpage > 1 Then
response.write "<span class='light'><a href='NetBook.asp?&page="+cstr(1)+"'>首页</a></span>"
Response.write "<span class='light'><a href='NetBook.asp?page="+Cstr(currentpage-1)+"'>上一页</a></span>"
Else
Response.write "<span class='pageTxt'>上一页 </span>"
End if
If currentpage < Rs.PageCount Then
Response.write "<span class='light'><a href='NetBook.asp?page="+Cstr(currentPage+1)+"'>下一页</a></span>"
Response.write "<span class='light'><a href='NetBook.asp?page="+Cstr(Rs.PageCount)+"'>尾页</a></span>"
Else
Response.write "<span class='pageTxt'>下一页</span>"
End if
'response.write "<font color='#000000' >转到：</font><input type='text' name='page' size=4 maxlength=4 class=smallInput value="&Currentpage&">"
'response.write "<input class=buttonface type='submit'  value='Go'  name='cndok'>"
%>
                </div>
              <%
If rs.eof and rs.bof then
call showpages
response.write ""
response.end
End if
%>
              <table class="viewListTable">
                <tr class="viewListTitle">
                  <td>内容</td>
                  <td>时间</td>
                  <td>审核状态</td>
                </tr>
                <%
if not rs.eof then
i=0
do while not rs.eof
%>
                <tr class="viewList">
                  <td><a href="ArticleShow.asp?ArticleID=<%=rs("prodid")%>" 
					title="<%=rs("nr")%>"> <%=replace(left(nohtml(rs("nr")),200),chr(34),"") & "……"%></a></td>
                  <td><%=rs("adddate")%></td>
                  <td>
                    <%If rs("shenhe")=true Then%>
                    已审核
                    <%else%>
                    未审核
                    <%End If%>
                    </td>
                </tr>
                <%
i=i+1
if i >= MaxPerPageB then exit do
rs.movenext
loop
end if
%>
                <%end sub%>
              </table>
      </td>
    </tr>
  </table>
</div>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>
