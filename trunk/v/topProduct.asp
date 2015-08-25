<!--#include file="Inc/syscode.asp"-->
<!-- #include file="Inc/Head.asp" -->
<%Function DateStringFromNow(Byval sTheDate)
' 格式化显示时间为几个月,几天前,几小时前,几分钟前,或几秒前
    Dim iSeconds, iMinutes, iHours, iDays

    iSeconds =DateDiff("s", sTheDate, Now())  'd/h/n/s
    iMinutes =Int(iSeconds/60)
    iHours =Int(iSeconds/3600)
    iDays =Int(iSeconds/86400)

    If iDays >60Then
        DateStringFromNow= sTheDate
    ElseIf iDays >30Then
        DateStringFromNow="1个月前"
   ElseIf iDays >14Then
        DateStringFromNow="2周前"
   ElseIf iDays >7Then
        DateStringFromNow="1周前"
   ElseIf iDays >1Then
        DateStringFromNow= iDays &"天前"
   ElseIf iHours >1Then
        DateStringFromNow= iHours &"小时前"
   ElseIf iMinutes >1Then
        DateStringFromNow= iMinutes &"分钟前"
   ElseIf iSeconds >=1Then
        DateStringFromNow= iSeconds &"秒前"
   Else
        DateStringFromNow="1秒前"
   End If
End Function
%>
<div id="body">
  <table class="bodytable">
    <tr>
      <td class="bodyleft"><div id="leftbg">
          <!-- #include file="L_news.asp" -->
        </div>
        <img src="img/1.gif" /></td>
      <td class="bodyline"></td>
      <td class="bodyright"><table class="mbox">
          <tr>
            <td><div class="rtitle"><img src="img/title_ico.gif" /> 视 频 列 表 </div></td>
          </tr>
          <tr>
            <td id="newslist" class="cttd"><%
flag="尚未处理"
set rs=server.createobject("adodb.recordset")

if request("topPro")=1 then
sqltext="select * from Product where showMov=False and Passed=True Order BY UpdateTime desc"
elseif request("topPro")=2 then
sqltext="select * from Product where showMov=False and Passed=True Order BY Hits desc"
elseif request("topPro")=3 then
sqltext="select * from Product where showMov=False and Passed=True Order BY Votes desc"
elseif request("topPro")=4 then
sqltext="select * from Product where showMov=False and Passed=True Order BY Author desc"
else
sqltext="select * from Product where showMov=False and Passed=True Order BY UpdateTime desc"
end if


'sqltext="select * from Product where showMov=False and Passed=True Order BY UpdateTime desc"
rs.open sqltext,conn,1,1
dim PerPage
PerPage=20
'假如没有数据时
If rs.eof and rs.bof then
call showpages
response.write "<p align='center'><font color='#ff0000'>还没任何数据</font></p>"
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
                <li><a href="ArticleShow.asp?ArticleID=<%=rs("articleid")%>"> <%=rs("Title")%> </a> <font color="#999999"> &nbsp; <%=rs("Author")%>&nbsp; 访：<%=rs("Hits")%>&nbsp;评：<%=rs("Votes")%>&nbsp;<%=DateStringFromNow(rs("crtTime"))%></font>
                </li>
                <%
i=i+1
if i >= Perpage then exit do
rs.movenext
loop
end if
%>
              </ul>			</td>
          </tr>  
          <tr>
		    <td id="tpage">
              <img src="img/arrow.gif" />
                <%
Response.write "共" & "<span id='numnews'>" & Cstr(Rs.RecordCount) & "</span>" & "篇"
Response.write "第" & "<span id='numnews'>" & Cstr(CurrentPage) &  "</span>" & "/" & Cstr(rs.pagecount) 
If currentpage > 1 Then
response.write "<span id='light'><a href='topProduct.asp?&page="+cstr(1)+"'>首页</a></span>"
Response.write "<span id='light'><a href='topProduct.asp?page="+Cstr(currentpage-1)+"'>上一页</a></span>"
Else
Response.write "<span id='gray'>上一页 </span>"
End if
If currentpage < Rs.PageCount Then
Response.write "<span id='light'><a href='topProduct.asp?page="+Cstr(currentPage+1)+"'>下一页</a></span>"
Response.write "<span id='light'><a href='topProduct.asp?page="+Cstr(Rs.PageCount)+"'>尾页</a></span>"
Else
Response.write "<span id='gray'>下一页</span>"
End if
'response.write "<span id='goto' >转到:<input type='text' name='page' size=4 maxlength=4 class=smallInput value="&Currentpage&">"
'response.write "<input class=buttonface type='submit'  value='Go'  name='cndok'></span>"
%>
              <%
End sub
rs.close
%>            </td>
          </tr>
        </table></td>
    </tr>
  </table>
</div>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>
