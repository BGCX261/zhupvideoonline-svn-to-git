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
      </td>
      <td class="bodyline"></td>
      <td class="bodyright"><table class="mbox">
          <tr>
            <td><div class="positiontitle"><img src="img/title_ico.gif" /> 教学管理 >> 教学团队</div></td>
          </tr>
          <tr>
            <td class="cttd">
            <!-- #include file="Inc/headAd.asp" -->
            <table>
			<%
set rs=server.createobject("adodb.recordset")
sqltext="select * from Teacher order by sortId asc"
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
                <tr>
                  <td><table cellspacing="4">
                      <%i=1 %>
                      <tr>
                <% Do While Not rs.EOF%>
                  <td><table width="100%" style="background:#e8e8e8;">
                      <tr>
                        <td align="center"><a href='coursTeacherInfo.asp?id=<%=rs("sortId")%>'><img src="<%=rs("Honor")%>" width="150" border="0" onmouseout=low(this) onmouseover=high(this) alt="<%=rs("teachName")%>" style="FILTER: alpha(opacity=30)" /></a></td>
                      </tr>
					  <tr>
                        <td align="center"><%=rs("teachName")%></td>
                      </tr>
                    </table></td>
                        <%if i mod 4 <>0 then%>
                        <%end if%>
                        <% if i mod 4 =0 then%>
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
if i >= 13 then exit do 
rs.MoveNext 
loop
end sub
%>
                      </tr>
                    </table></td>
                </tr>
              </table>
              <div>
                <%
Response.write "<strong><font color='#000000'>-> 全部-</font>"
Response.write "共</font>" & "<font color=#FF0000>" & Cstr(Rs.RecordCount) & "</font>" & "<font color='#000000'>条信息</font></strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
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
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
</div>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>
