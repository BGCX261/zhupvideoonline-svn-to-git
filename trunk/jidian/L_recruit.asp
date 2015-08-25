<div class="leftbox">
    <div class="leftboxtitle">招生就业</div>
    <div class="leftboxcenter">      
<% 
		sqltype="select * from RecruitType Order by id asc"
		Set rstype= Server.CreateObject("ADODB.Recordset")
		rstype.open sqltype,conn,1,1
	if rstype.bof and rstype.eof then 
		response.Write "没有数据"
	else
		rstype.movefirst
		response.write "<ul>"
		do while not rstype.eof
			response.write "<li id='one'><a href='recruit.asp?typed="&rstype("typed")&"'>"&rstype("typed")&"</a>"
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
