<div class="leftbox">
    <div class="leftboxtitle">ѧԺ�ſ�</div>
    <div class="leftboxcenter">
<% 
set rspro=server.createobject("adodb.recordset")
sqlpro="select top 9 *  from Profile  order by mdy_time desc"
rspro.open sqlpro,conn,1,1
  If rspro.eof and rspro.bof then
    response.write "��������"
  else 
  dim strnewsName
  rspro.movefirst
  response.write strnewsName & "<ul>"
  do while not rspro.eof
response.write strnewsName & "<li id='one'><a href='profileInfo.asp?id=" & rspro("id") & "'>" & cutstr(rspro("title"),14) & "</a></li>"
     rspro.movenext
  loop
  response.write strnewsName & "</ul>"
  rspro.close
  set rspro=nothing
end if
 %>
    </div>
    <div class="leftboxbottom"></div>
</div>
