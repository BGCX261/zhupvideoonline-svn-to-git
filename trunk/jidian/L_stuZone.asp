<div class="leftbox">
    <div class="leftboxtitle">师生天地</div>
    <div class="leftboxcenter">
        <% 
		sqltype="select * from StuZoneType Order by id desc"
		Set rstype= Server.CreateObject("ADODB.Recordset")
		rstype.open sqltype,conn,1,1
	if rstype.bof and rstype.eof then 
		response.Write "没有数据"
	else
		rstype.movefirst
		response.write "<ul>"
		do while not rstype.eof
			response.write "<li id='one'><a href='stuZone.asp?nbtyped="&rstype("nbtyped")&"'>"&rstype("nbtyped")&"</a>"
			response.write "</li>"
			rstype.movenext
		loop
		response.write "</ul>"
		rstype.close
		set rstype=nothing
	end if
 %>
    </div>
</div>
