<!--#include file="Inc/syscode.asp"-->
<!-- #include file="Inc/Head.asp" -->
<div id="body">
  <table class="bodytable">
    <tr>
      <td class="bodyleft">
          <!-- #include file="L_news.asp" -->
      </td>
      <td class="bodyline"></td>
      <td class="bodyright">
      <!-- #include file="Inc/headAd.asp" -->
        <div class="mbox">
          <div id="newslist">
          <div class="positiontitle"><img src="img/title_ico.gif" /> 新闻公告 </div>
          <%
flag="尚未处理"
set rs=server.createobject("adodb.recordset")
sqltext="select * from news Order by id desc"
rs.open sqltext,conn,1,1
dim PerPage
PerPage=10
'假如没有数据时
If rs.eof and rs.bof then
call showpages
response.write "<div align='center'>还没任何数据</div>"
response.end
End if
'取得页数,并判断用户输入的是否数字类型的数据，如不是将以第一页显示
text="0123456789"
Rs.PageSize=PerPage
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

If Rs.recordcount > PerPage then
call showpages
end if

'显示帖子的子程序
Sub list()
%>
              <ul>
                <%
if not rs.eof then
do while not rs.eof
%>
                <li><span class="artlnk">[<%=rs("typed")%>] <a href="NewsInfo.asp?id=<%=rs("id")%>"><%=rs("title")%></a> </span>
            <% if datediff("d",rs("mdy_time"),date())<3 then%>
            <span class='newest' title='最新发表'></span>&nbsp;  
            <%end if%>
             <span class="artlistupdate"><%=formatdatetime(rs("mdy_time"),2)%></span> 
                </li>
                <%
i=i+1
if i >= Perpage then exit do
rs.movenext
loop
end if
%>
              </ul>
          </div>  
		  <div id="tpage">
              <img src="img/arrow.gif" />
                <% 
Response.write "共" & "<span id='numnews'>" & Cstr(Rs.RecordCount) & "</span>" & "篇 "
Response.write "第" & "<span id='numnews'>" & Cstr(CurrentPage) &  "</span>" & "/" & Cstr(rs.pagecount) 
If currentpage > 1 Then
response.write "<span id='light'><a href='News.asp?&page="+cstr(1)+"'>首页</a></span>"
Response.write "<span id='light'><a href='News.asp?page="+Cstr(currentpage-1)+"'>上一页</a></span>"
Else
Response.write "<span id='gray'>上一页 </span>"
End if
If currentpage < Rs.PageCount Then
Response.write "<span id='light'><a href='News.asp?page="+Cstr(currentPage+1)+"'>下一页</a></span>"
Response.write "<span id='light'><a href='News.asp?page="+Cstr(Rs.PageCount)+"'>尾页</a></span>"
Else
Response.write "<span id='gray'>下一页</span>"
End if
'response.write "<span id='goto' >转到:<input type='text' name='page' size=4 maxlength=4 class=smallInput value="&Currentpage&">"
'response.write "<input class=buttonface type='submit'  value='Go'  name='cndok'></span>"
%>
              <%
End sub
rs.close
%>            </div>
        </div>
      </td>
    </tr>
  </table>
</div>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>
