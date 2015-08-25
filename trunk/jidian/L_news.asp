<div class="leftbox">
    <div class="leftboxtitle">最近动态</div>
    <div class="leftboxcenter">
    
<%
        
        
set rsnews=server.createobject("adodb.recordset")
sqltext1="select top 10 *  from news  order by mdy_time desc"
rsnews.open sqltext1,conn,1,1

  If rsnews.eof and rsnews.bof then
    response.write "暂无内容"
  else 
  dim strnewsName
  rsnews.movefirst
  response.write strnewsName & "<ul>"
  do while not rsnews.eof
     response.write strnewsName & "<li id='one'><a href='NewsInfo.asp?id=" & rsnews("id") & "'>" & cutstr(rsnews("title"),12) & "</a></li>"
     rsnews.movenext
  loop
  response.write strnewsName & "</ul>"
  rsnews.close
  set rsnews=nothing
end if 
%>
      </div>
    <div class="leftboxbottom"></div>
</div>
