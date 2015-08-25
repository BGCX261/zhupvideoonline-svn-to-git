<table class="leftbox">
  <tr>
    <td class="leftboxtitle">学习教程分类</td>
  </tr>
  <tr>
    <td class="leftboxcenter">
        <% 
		sqltype="select * from newbooktype Order by id desc"
		Set rstype= Server.CreateObject("ADODB.Recordset")
		rstype.open sqltype,conn,1,1
	if rstype.bof and rstype.eof then 
		response.Write "没有数据"
	else
		rstype.movefirst
		response.write "<ul>"
		do while not rstype.eof
			response.write "<li id='one'><a href='willbook.asp?nbtyped="&rstype("nbtyped")&"'>"&rstype("nbtyped")&"</a>"
			response.write "</li>"
			rstype.movenext
		loop
		response.write "</ul>"
		rstype.close
		set rstype=nothing
	end if
 %>
      </td>
  </tr>
  <tr>
    <td class="leftboxbottom"></td>
  </tr>
</table>
