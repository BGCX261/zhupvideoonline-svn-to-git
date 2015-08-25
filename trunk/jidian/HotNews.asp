<ul> 
				<%
t=0
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="SELECT * from news where typed='通知公告' order by id desc,mdy_time desc" 
rs.Open sql,conn,1,1
if not Rs.eof then
do while not rs.eof
t=t+1
%>
                    
                      <li <%if rs("toTop")=true then response.write "class='toTop'" else response.write "class=''" end if%> ><a href="NewsInfo.asp?id=<%=rs("id")%>" title="<%=rs("title")%>" target="_blank"><%=cutstr(rs("title"),24)%></a><span class="noticeTime"><%=DateValue(rs("mdy_time"))%></span></li>
                    
					<% 
if t>=4 then exit do 
rs.movenext 
loop 
else 
response.write "<div align=center>暂无新闻!</div>" 
end if 
rs.close 
%>
  </ul>