<div class="leftbox">
    <div class="leftboxtitle">实践教学</div>
    <div class="leftboxcenter">
<% 
		sqltype="select * from FacResourceType Order by id desc"
		Set rstype= Server.CreateObject("ADODB.Recordset")
		rstype.open sqltype,conn,1,1
	if rstype.bof and rstype.eof then 
		response.Write "没有信息"
	else
		rstype.movefirst
		response.write "<ul>"
		do while not rstype.eof
			response.write "<li id='one'><a href='facResource.asp?frtype="&rstype("frtype")&"'>"&rstype("frtype")&"</a>"
			response.write "</li>"
			rstype.movenext
		loop
		response.write "</ul>"
		rstype.close
		set rstype=nothing
	end if
 %>
    </div>
    <div class="leftboxbottom"></div>
</div>
