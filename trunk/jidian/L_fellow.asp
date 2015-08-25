<%
set rs_links=server.createobject("adodb.recordset")
sqltext4="select top 10 * from links order by id desc"
rs_links.open sqltext4,conn,1,1
%>
<table class="leftbox">
  <tr>
    <td class="leftboxtitle">”—«È¡¨Ω”</td>
  </tr>
  <tr>
    <td><div id="mainlist">
        <ul>
          <%
i=0
do while not rs_links.eof
%>
          <li id="one"><a href="<%=rs_links("link")%>" title="<%=rs_links("note")%>"target="_blank"><%=rs_links("name")%></a></li>
          <%
rs_links.movenext
i=i+1
if i=10 then exit do
loop
rs_links.close 
%>
        </ul>
      </div></td>
  </tr>
  <tr>
    <td class="leftboxbottom"></td>
  </tr>
</table>
