<div class="leftbox">
    <div class="leftboxtitle">ʦ�����</div>
    <div class="leftboxcenter">
        <% 
		sqltype="select * from StuZoneType Order by id desc"
		Set rstype= Server.CreateObject("ADODB.Recordset")
		rstype.open sqltype,conn,1,1
	if rstype.bof and rstype.eof then 
		response.Write "û������"
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
