<!--#include file="Inc/syscode.asp"-->
<!-- #include file="Inc/Head.asp" -->
<%Function DateStringFromNow(Byval sTheDate)
' ��ʽ����ʾʱ��Ϊ������,����ǰ,��Сʱǰ,������ǰ,����ǰ
    Dim iSeconds, iMinutes, iHours, iDays

    iSeconds =DateDiff("s", sTheDate, Now())  'd/h/n/s
    iMinutes =Int(iSeconds/60)
    iHours =Int(iSeconds/3600)
    iDays =Int(iSeconds/86400)

    If iDays >60Then
        DateStringFromNow= sTheDate
    ElseIf iDays >30Then
        DateStringFromNow="1����ǰ"
   ElseIf iDays >14Then
        DateStringFromNow="2��ǰ"
   ElseIf iDays >7Then
        DateStringFromNow="1��ǰ"
   ElseIf iDays >1Then
        DateStringFromNow= iDays &"��ǰ"
   ElseIf iHours >1Then
        DateStringFromNow= iHours &"Сʱǰ"
   ElseIf iMinutes >1Then
        DateStringFromNow= iMinutes &"����ǰ"
   ElseIf iSeconds >=1Then
        DateStringFromNow= iSeconds &"��ǰ"
   Else
        DateStringFromNow="1��ǰ"
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
            <td><div class="rtitle"><img src="img/title_ico.gif" /> �� Ƶ �� �� </div></td>
          </tr>
          <tr>
            <td id="newslist" class="cttd"><%
flag="��δ����"
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
'����û������ʱ
If rs.eof and rs.bof then
call showpages
response.write "<p align='center'><font color='#ff0000'>��û�κ�����</font></p>"
response.end
End if
'ȡ��ҳ��,���ж��û�������Ƿ��������͵����ݣ��粻�ǽ��Ե�һҳ��ʾ
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

'��ʾ���ӵ��ӳ���
Sub list()
%>
              <ul>
                <%
if not rs.eof then
do while not rs.eof
%>
                <li><a href="ArticleShow.asp?ArticleID=<%=rs("articleid")%>"> <%=rs("Title")%> </a> <font color="#999999"> &nbsp; <%=rs("Author")%>&nbsp; �ã�<%=rs("Hits")%>&nbsp;����<%=rs("Votes")%>&nbsp;<%=DateStringFromNow(rs("crtTime"))%></font>
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
Response.write "��" & "<span id='numnews'>" & Cstr(Rs.RecordCount) & "</span>" & "ƪ"
Response.write "��" & "<span id='numnews'>" & Cstr(CurrentPage) &  "</span>" & "/" & Cstr(rs.pagecount) 
If currentpage > 1 Then
response.write "<span id='light'><a href='topProduct.asp?&page="+cstr(1)+"'>��ҳ</a></span>"
Response.write "<span id='light'><a href='topProduct.asp?page="+Cstr(currentpage-1)+"'>��һҳ</a></span>"
Else
Response.write "<span id='gray'>��һҳ </span>"
End if
If currentpage < Rs.PageCount Then
Response.write "<span id='light'><a href='topProduct.asp?page="+Cstr(currentPage+1)+"'>��һҳ</a></span>"
Response.write "<span id='light'><a href='topProduct.asp?page="+Cstr(Rs.PageCount)+"'>βҳ</a></span>"
Else
Response.write "<span id='gray'>��һҳ</span>"
End if
'response.write "<span id='goto' >ת��:<input type='text' name='page' size=4 maxlength=4 class=smallInput value="&Currentpage&">"
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
