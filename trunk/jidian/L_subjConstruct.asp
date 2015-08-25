<div class="leftbox">
    <div class="leftboxtitle">党建工作</div>
    <div class="leftboxcenter">
<%

set rssubjc=server.createobject("adodb.recordset")
sqlsubjc="select *  from SubjCType  order by id desc"
rssubjc.open sqlsubjc,conn,1,1

  If rssubjc.eof and rssubjc.bof then
    response.write "暂无设置"
  else 
  dim strSubjc
  rssubjc.movefirst
  response.write strSubjc & "<ul>"
  do while not rssubjc.eof
response.write strSubjc & "<li id='one'><a href='subjConstructs.asp?nbtyped=" & rssubjc("nbtyped") & "'>" & rssubjc("nbtyped") & "</a></li>"
     rssubjc.movenext
  loop
  response.write strSubjc & "</ul>"
  rssubjc.close
  set rssubjc=nothing
end if

%>
      </div>
    <div class="leftboxbottom"></div>
</div>
