<div class="leftbox">
    <div class="leftboxtitle">��ѧ����</div>
    <div class="leftboxcenter">
      <ul>
        <li id='one'><a href='coursTeacher.asp'>��ѧ�Ŷ�</a></li>

      <% 
		sqltype="select * from scType Order by id desc"
		Set rstype= Server.CreateObject("ADODB.Recordset")
		rstype.open sqltype,conn,1,1
	if rstype.bof and rstype.eof then 
		response.Write ""
	else
		rstype.movefirst
		do while not rstype.eof
			response.write "<li id='one'><a href='Scientific.asp?nbtyped="&rstype("nbtyped")&"'>"&rstype("nbtyped")&"</a>"
			response.write "</li>"
			rstype.movenext
		loop
		rstype.close
		set rstype=nothing
	end if
 %>
       </ul>
    </div>
    <div class="leftboxbottom"></div>
</div>
