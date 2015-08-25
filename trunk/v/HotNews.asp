<div id="newslist">
  <ul> 
				<%
t=0
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="SELECT * from news order by id desc" 
rs.Open sql,conn,1,1
if not Rs.eof then
do while not rs.eof
t=t+1
%>
                    
                      <li>¡¡
                        <a href="NewsInfo.asp?id=<%=rs("id")%>" title="<%=rs("title")%>" target="_blank"><%=left(rs("title"),10)%></a></li>
                    
					<% 
if t>=6 then exit do 
rs.movenext 
loop 
else 
response.write "<tr><td align=center bgcolor=#E8E8F4>ÔÝÎÞÐÂÎÅ!</td></tr>" 
end if 
rs.close 
%>
  </ul>
</div>