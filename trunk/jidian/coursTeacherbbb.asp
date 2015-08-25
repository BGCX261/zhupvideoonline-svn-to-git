<!--#include file="Inc/syscode.asp"-->
<!-- #include file="Inc/Head.asp" -->
<SCRIPT language=JavaScript>
<!--
function high(which2){
   theobject=which2
   highlighting=setInterval("highlightit(theobject)",100)
   }
function low(which2){
	clearInterval(highlighting)
    which2.filters.alpha.opacity=15
	}
function highlightit(cur2){
    if (cur2.filters.alpha.opacity<200)
    cur2.filters.alpha.opacity+=20
    else if (window.highlighting)
    clearInterval(highlighting)
	}
//-->
</SCRIPT>

<div id="body">
  <table class="bodytable">
    <tr>
      <td class="bodyleft">
        <!-- #include file="L_teacher.asp" -->
        <%
set rs=server.createobject("adodb.recordset")
sqltext="select * from Teacher order by time desc"
rs.open sqltext,conn,1,1

dim MaPerPage
MaPerPage=12

dim text,checkpage
text="0123456789"
Rs.PageSize=MaPerPage
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

If Rs.recordcount > MaPerPage then
call showpages
end if

'显示帖子的子程序
Sub list()%>
 <%i=1 %>
 <table>
                      <%i=1 %>
                      <tr>
                <% Do While Not rs.EOF%>
                  <td><table width="100%">
                      <tr>
                        <td align="center" style="background:#999;"><a href='coursTeacher.asp?id=<%=rs("id")%>'><img src="<%=rs("Honor")%>" width="66" border="0" onmouseout=low(this) onmouseover=high(this) alt="<%=rs("teachName")%>" style="FILTER: alpha(opacity=30)" /></a></td>
                      </tr>
					  <tr>
                        <td align="center"><%=rs("teachName")%></td>
                      </tr>
                    </table></td>
                        <%if i mod 3 <>0 then%>
                        <%end if%>
                        <% if i mod 3 =0 then%>
                      </tr>
                      <tr>
                        <%end if%>
                        <%  
'rs.MoveNext      
'i=i+1      
'Loop
'end sub   			      
%>
                        <%
i=i+1 
if i >= 12 then exit do 
rs.MoveNext 
loop
end sub
%>
                      </tr>
                    </table>
        <div>
                <%
Response.write "<strong>共" & "<font color=#FF0000>" & Cstr(Rs.RecordCount) & "</font>" & "<font color='#000000'>位教师</font></strong> "
Response.write "<strong><font color='#000000'>第</font>" & "<font color=#FF0000>" & Cstr(CurrentPage) &  "</font>" & "<font color='#000000'>/" & Cstr(rs.pagecount) & "</font></strong>&nbsp;"
If currentpage > 1 Then
response.write "<strong><a href='coursTeacher.asp?&page="+cstr(1)+"'><font color='#000000'>首页</font></a><font color='#ffffff'> </font></strong>"
Response.write "<strong><a href='coursTeacher.asp?page="+Cstr(currentpage-1)+"'><font color='#000000'>上一页</font></a><font color='#ffffff'> </font></strong>"
Else
Response.write "<strong><font color='#000000'>上一页 </font></strong>"
End if
If currentpage < Rs.PageCount Then
Response.write "<strong><a href='coursTeacher.asp?page="+Cstr(currentPage+1)+"'><font color='#000000'>下一页</font></a><font color='#ffffff'> </font>"
Response.write "<a href='coursTeacher.asp?page="+Cstr(Rs.PageCount)+"'><font color='#000000'>尾页</font></a></strong>&nbsp;&nbsp;"
Else
Response.write ""
Response.write "<strong><font color='#000000'>下一页</font></strong>&nbsp;&nbsp;"
End if
%>
              </div>
        </td>
      <td class="bodyline"></td>
      <td class="bodyright"><table class="mbox">
	      <!-- #include file="Inc/headAd.asp" -->
          <tr>
            <td><div class="positiontitle"><img src="img/title_ico.gif" /> 教学管理 - 教学团队</div></td>
          </tr>
          <tr>
            <td class="cttd">
            <table>

                <tr>
                  <td><%
if not isEmpty(request.QueryString("id")) then
id=request.QueryString("id")
else
id=2
end if

set rs=server.createobject("adodb.recordset")
sqltext="select * from Teacher where id="&id
rs.open sqltext,conn,1,1
%>
              <div class="contentbox">
                  <div class="chptitletd"><%=rs("teachName")%></div>
                  <div class="contentview"><p><img src="<%=rs("Honor")%>" width="150" hspace="8" border="0" align="left" /><span class="contentviewLebal">简介：</span><%=rs("teachInfo")%></p><p><span class="contentviewLebal">主讲课程：</span><%=rs("teachSubj")%></p><p><span class="contentviewLebal">座右铭：</span><%=rs("teachSay")%><%rs.close%></p></div>
                  <div class="prt">【<a href='javascript:window.print()'>打印此页</a
>】【<a href='javascript:history.back()'>返回</a
>】【<a href="javascript:window.scroll(0,-360)">顶部</a
>】【<a href="javascript:self.close()">关闭</a>】</div>
              </div></td>
                </tr>
              </table>
              </td>
          </tr>
        </table></td>
    </tr>
  </table>
</div>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>
