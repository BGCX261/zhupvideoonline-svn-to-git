<!--#include file="Inc/syscode.asp"-->
<!-- #include file="Inc/Head.asp" -->
<div id="body">
  <table class="bodytable">
    <tr>
      <td class="bodyleft">
          <!-- #include file="L_proSet.asp" -->
      </td>
      <td class="bodyline"></td>
      <td class="bodyright">
      <!-- #include file="Inc/headAd.asp" -->
        <div class="mbox">
          <div class="positiontitle"><img src="img/title_ico.gif" /> רҵ����</div>
          <div  class="artlist"><%
flag="��δ����"
set rs=server.createobject("adodb.recordset")
sqltext="select * from ProSet Order BY id desc"
rs.open sqltext,conn,1,1
dim PerPage
PerPage=20
'����û������ʱ
If rs.eof and rs.bof then
call showpages
response.write "<div align='center'>��������</div>"
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
                <li><span class="artlnk"><a href="proSetInfo.asp?id=<%=rs("id")%>"><%=cutstr(rs("title"),50)%></a></span>
                    <span class="artlistupdate"><%=rs("time")%></span>
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
		  <div id="tpage"><img src="img/arrow.gif" />
                <%
Response.write "�ϼ�"
Response.write "<span id='numnews'>" & Cstr(Rs.RecordCount) & "</span> " 
Response.write "<span id='numnews'>" & Cstr(CurrentPage) &  "</span>" & "/" & Cstr(rs.pagecount) 
If currentpage > 1 Then
response.write "<span id='light'><a href='proSet.asp?&page="+cstr(1)+"'>��ҳ</a></span>"
Response.write "<span id='light'><a href='proSet.asp?page="+Cstr(currentpage-1)+"'>��һҳ</a></span>"
Else
Response.write "<span id='gray'>��һҳ </span>"
End if
If currentpage < Rs.PageCount Then
Response.write "<span id='light'><a href='proSet.asp?page="+Cstr(currentPage+1)+"'>��һҳ</a></span>"
Response.write "<span id='light'><a href='proSet.asp?page="+Cstr(Rs.PageCount)+"'>βҳ</a></span>"
Else
Response.write "<span id='gray'>��һҳ</span>"
End if
'response.write "<span id='goto' >Goto��<input type='text' name='page' size=4 maxlength=4 class=smallInput value="&Currentpage&">"
'response.write "<input class=buttonface type='submit'  value='Go'  name='cndok'></span>"
%>
              
<%
End sub
rs.close
%>
          </div>
        </div>
      </td>
    </tr>
  </table>
</div>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>