<!--#include file="bsconfig.asp"-->
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("04")%>
<%
if request("deladm")<>"" then
deladm=request("deladm")
call deladmin()
end if

admin=request("admin")
if admin="" then admin=request.cookies("buyok")("admin")
if request("edit")<>"ok" then
%>

  <table class="a2" align="center">
  <tr>
   <td class="a1" colspan="3">��̨�û��б�(������������û��鿴��ӵ�е�Ȩ��)</td>
  </tr>
    
<%
		Set rs = conn.Execute("select * from Bs_User")
		do while not rs.eof
		response.write "<tr><td>"
		if admin=rs("UserName") then
		response.write "<img border=0 src=../images/topBar_bg.gif> <a href=Bs_adminPermissions.asp?admin="&rs("UserName")&"><font color=red><b>"&rs("UserName")&"</b></font></a> </td><td>" & rs("realname")
		else
		response.write "<a href=Bs_adminPermissions.asp?admin="&rs("UserName")&">"&rs("UserName")&"</a> </td><td>" & rs("realname")
	end if
%>
	</td><td><img border=0 src=../images/delete.gif alt="ɾ���˹���Ա" style="cursor:hand" onClick="{if(confirm('�ò������ɻָ���\n\nȷʵҪɾ���������Ա�� ')){location.href='Bs_adminPermissions.asp?deladm=<%=rs("UserName")%>';}}">
        </td>
      </tr>
	<%
	rs.movenext
	loop
	%> 
  </table>

<form action="Bs_adminPermissions.asp" method=post name=manage>
	<%
	Set rs = conn.Execute("select * from Bs_User where UserName='"&admin&"'")
		  if rs.eof and rs.bof then
			response.write "<script language='javascript'>"
			response.write "alert('û�д˹���Ա��');"
			response.write "location.href='javascript:history.go(-1)';"
			response.write "</script>"
			response.end
		  end if
 		  manage=rs("admin_Manage")
		%>
  <table align="center" class="a2">
	<tr>
		<td class="a1"><font color=red><b><%=admin%></b></font> �Ĺ���Ȩ��</td>
	</tr>
	<tr class="a4">
		<td>
  <table width="100%" style=" border-collapse:collapse;">
    <tr class="opttitle">
      <td>ģ��</td>
      <td>���ܵ�</td>
      <td>Ȩ������</td>
      <td>��ע</td>
      </tr>
    <tbody id="goaler">
    <tr>
      <td rowspan="13">ϵͳ����</td>
      <td>����Ա���� </td>
      <td>
       <input type="checkbox" value="yes" name="manage01" <%if instr(manage,"01")>0 then%>checked<%end if%>>
       ����
       <input type="checkbox" value="yes" name="manage02" <%if instr(manage,"02")>0 then%>checked<%end if%>>
       �޸����� 
       <input type="checkbox" value="yes" name="manage03" <%if instr(manage,"03")>0 then%>checked<%end if%>>ɾ��</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>Ȩ�޹���</td>
      <td><input type="checkbox" value="yes" name="manage04" <%if instr(manage,"04")>0 then%>checked<%end if%>>
����</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>���ݿⱸ��</td>
      <td><input type="checkbox" value="yes" name="manage05" <%if instr(manage,"05")>0 then%>checked<%end if%>>
����</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>��������</td>
      <td><input type="checkbox" value="yes" name="manage06" <%if instr(manage,"06")>0 then%>checked<%end if%>>
����</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>ϵͳ����</td>
      <td><input type="checkbox" value="yes" name="manage07" <%if instr(manage,"07")>0 then%>checked<%end if%>>
����</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>�ϴ��ļ�����</td>
      <td><input type="checkbox" value="yes" name="manage08" <%if instr(manage,"08")>0 then%>checked<%end if%>>
�б�
  <input type="checkbox" value="yes" name="manage09" <%if instr(manage,"09")>0 then%>checked<%end if%>>
ɾ��</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>��ҳ�������</td>
      <td><input type="checkbox" value="yes" name="manage10" <%if instr(manage,"10")>0 then%>checked<%end if%>>
����</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>�û��޸��Լ�������</td>
      <td><input type="checkbox" value="yes" name="manage11" <%if instr(manage,"11")>0 then%>checked<%end if%>>
����</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>��������</td>
      <td><input type="checkbox" value="yes" name="manage12" <%if instr(manage,"12")>0 then%>checked<%end if%>>
����</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>��ҳ���</td>
      <td><input type="checkbox" value="yes" name="manage13" <%if instr(manage,"13")>0 then%>checked<%end if%>>
����</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>ҳͷͼƬ</td>
      <td><input type="checkbox" value="yes" name="manage14" <%if instr(manage,"14")>0 then%>checked<%end if%>>
����</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>����������</td>
      <td><input type="checkbox" value="yes" name="manage55" <%if instr(manage,"55")>0 then%>checked<%end if%>>
        ����</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>ҳ����Ϣ����</td>
      <td><input type="checkbox" value="yes" name="manage56" <%if instr(manage,"56")>0 then%>checked<%end if%>>
        ����</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>ѧԺ�ſ�����</td>
      <td>ѧԺ�����Ϣ����</td>
      <td align="left"><input type="checkbox" value="yes" name="manage33" <%if instr(manage,"33")>0 then%>checked<%end if%>>
        �б�
        <input type="checkbox" value="yes" name="manage54" <%if instr(manage,"54")>0 then%>checked<%end if%>>
        ����
        <input type="checkbox" value="yes" name="manage23" <%if instr(manage,"23")>0 then%>checked<%end if%>>
        �޸�,
        ɾ��</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>רҵ���ù���</td>
      <td>רҵ��Ϣ����</td>
      <td><input type="checkbox" value="yes" name="manage15" <%if instr(manage,"15")>0 then%>checked<%end if%>>
        �б�
        <input type="checkbox" value="yes" name="manage16" <%if instr(manage,"16")>0 then%>checked<%end if%>>
        ����
        <input type="checkbox" value="yes" name="manage18" <%if instr(manage,"18")>0 then%>checked<%end if%>>
        �޸ģ�ɾ��</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td rowspan="4">��ѧ���й���</td>
      <td>��ʦ�Ŷ�</td>
      <td><input type="checkbox" value="yes" name="manage19" <%if instr(manage,"19")>0 then%>checked<%end if%>>
        �б�
        <input type="checkbox" value="yes" name="manage20" <%if instr(manage,"20")>0 then%>checked<%end if%>>
        ����
        <input type="checkbox" value="yes" name="manage21" <%if instr(manage,"21")>0 then%>checked<%end if%>>
        �޸�
        
        <input type="checkbox" value="yes" name="manage22" <%if instr(manage,"22")>0 then%>checked<%end if%>>
        ɾ��</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>��ʦ�޸��Լ�����Ϣ</td>
      <td><input type="checkbox" value="yes" name="manage57" <%if instr(manage,"57")>0 then%>checked<%end if%>>
        ����</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>��ѧ���з������</td>
      <td>
<input type="checkbox" value="yes" name="manage17" <%if instr(manage,"17")>0 then%>checked<%end if%>>
����</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>��ѧ�������¹���</td>
      <td>
        <input type="checkbox" value="yes" name="manage30" <%if instr(manage,"30")>0 then%>checked<%end if%>>
        �б�
         <input type="checkbox" value="yes" name="manage40" <%if instr(manage,"40")>0 then%>checked<%end if%>>
         ���� 
         
        <input type="checkbox" value="yes" name="manage26" <%if instr(manage,"26")>0 then%>checked<%end if%>>
        �޸ģ�ɾ��</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td rowspan="2">������������</td>
      <td>��Ŀ����</td>
      <td><input type="checkbox" value="yes" name="manage59" <%if instr(manage,"59")>0 then%>checked<%end if%>>
����</td>
      <td rowspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td>��������</td>
      <td><input type="checkbox" value="yes" name="manage24" <%if instr(manage,"24")>0 then%>checked<%end if%>>
�б�
  <input type="checkbox" value="yes" name="manage25" <%if instr(manage,"25")>0 then%>checked<%end if%>>
����
<input type="checkbox" value="yes" name="manage27" <%if instr(manage,"27")>0 then%>checked<%end if%>>
�޸ģ�ɾ��</td>
    </tr>
    <tr>
      <td rowspan="2">ʵ����ѧ����</td>
      <td>ʵ����ѧ����</td>
      <td><input type="checkbox" value="yes" name="manage28" <%if instr(manage,"28")>0 then%>checked<%end if%>>
        �б�
          <input type="checkbox" value="yes" name="manage29" <%if instr(manage,"29")>0 then%>checked<%end if%>>
          ����
          <input type="checkbox" value="yes" name="manage31" <%if instr(manage,"31")>0 then%>checked<%end if%>>
�޸ģ�ɾ��
<!--<input type="checkbox" value="yes" name="manage32" <%if instr(manage,"32")>0 then%>checked<%end if%>>
���--></td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>ʵ����ѧ����</td>
      <td><input type="checkbox" value="yes" name="manage58" <%if instr(manage,"58")>0 then%>checked<%end if%>>
        ����</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>���Ź������</td>
      <td>���Ź�����Ϣ</td>
      <td><input type="checkbox" value="yes" name="manage34" <%if instr(manage,"34")>0 then%>checked<%end if%>>
        �б�
          <input type="checkbox" value="yes" name="manage35" <%if instr(manage,"35")>0 then%>checked<%end if%>>
          ����
          <input type="checkbox" value="yes" name="manage36" <%if instr(manage,"36")>0 then%>checked<%end if%>>
          �޸�
          
          <input type="checkbox" value="yes" name="manage37" <%if instr(manage,"37")>0 then%>checked<%end if%>>
ɾ��</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>������ҵ����</td>
      <td>������ҵ</td>
      <td><input type="checkbox" value="yes" name="manage38" <%if instr(manage,"38")>0 then%>checked<%end if%>>
        �б�
          <input type="checkbox" value="yes" name="manage39" <%if instr(manage,"39")>0 then%>checked<%end if%>>
        ����
        <input type="checkbox" value="yes" name="manage41" <%if instr(manage,"41")>0 then%>checked<%end if%>>
        �޸ģ�ɾ��        </td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td rowspan="2">ʦ����ع���</td>
      <td>�������</td>
      <td>
        <input type="checkbox" value="yes" name="manage42" <%if instr(manage,"42")>0 then%>checked<%end if%>>
        ����</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>���¹���</td>
      <td>
        <input type="checkbox" value="yes" name="manage43" <%if instr(manage,"43")>0 then%>checked<%end if%>> 
        �б�
         <input type="checkbox" value="yes" name="manage44" <%if instr(manage,"44")>0 then%>checked<%end if%>> 
         ���� 
        <input type="checkbox" value="yes" name="manage45" <%if instr(manage,"45")>0 then%>checked<%end if%>> 
        �޸� 
        <input type="checkbox" value="yes" name="manage46" <%if instr(manage,"46")>0 then%>checked<%end if%>> 
        ɾ��
        <input type="checkbox" value="yes" name="manage47" <%if instr(manage,"47")>0 then%>checked<%end if%>>
        ���</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td rowspan="2">����γ̹���</td>
      <td>�������</td>
      <td><input type="checkbox" value="yes" name="manage48" <%if instr(manage,"48")>0 then%>checked<%end if%>>
        ����</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>�γ̹���</td>
      <td><input type="checkbox" value="yes" name="manage49" <%if instr(manage,"49")>0 then%>checked<%end if%>>
�б�
  <input type="checkbox" value="yes" name="manage50" <%if instr(manage,"50")>0 then%>checked<%end if%>>
����
<input type="checkbox" value="yes" name="manage51" <%if instr(manage,"51")>0 then%>checked<%end if%>>
�޸�

<input type="checkbox" value="yes" name="manage52" <%if instr(manage,"52")>0 then%>checked<%end if%>>
ɾ��
<input type="checkbox" value="yes" name="manage53" <%if instr(manage,"53")>0 then%>checked<%end if%>>
���</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td rowspan="2">����ר�����</td>
      <td>����ר�����</td>
      <td><input type="checkbox" value="yes" name="manage61" <%if instr(manage,"61")>0 then%>checked<%end if%>>
        �б�
          <input type="checkbox" value="yes" name="manage62" <%if instr(manage,"62")>0 then%>checked<%end if%>>
          ����
          <input type="checkbox" value="yes" name="manage63" <%if instr(manage,"63")>0 then%>checked<%end if%>>
�޸ģ�ɾ��
      </td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>����ר�����</td>
      <td><input type="checkbox" value="yes" name="manage60" <%if instr(manage,"60")>0 then%>checked<%end if%>>
        ����</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>���ļ��ϴ�����</td>
      <td>���ļ��ϴ����б�</td>
      <td><input type="checkbox" value="yes" name="manage64" <%if instr(manage,"64")>0 then%>checked<%end if%>>
        ����</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td colspan="4" align="center"><input type="submit" name="Submit" value=" �� �� ">
		<input type="hidden" name="edit" value="ok">
		<input type="hidden" name="admin" value="<%=admin%>"> <input name="input" type="button" value=" �� �� " onClick="window.history.go(-1)" /></td>
      </tr>
    </tbody>
  </table>
  		</td>
	</tr>
</table>
</form>
<%   
else

	manage=""
	if request.form("manage01")="yes" then manage=manage+"|01"
	if request.form("manage02")="yes" then manage=manage+"|02"
	if request.form("manage03")="yes" then manage=manage+"|03"
	if request.form("manage04")="yes" then manage=manage+"|04"
	if request.form("manage05")="yes" then manage=manage+"|05"
	if request.form("manage06")="yes" then manage=manage+"|06"
	if request.form("manage07")="yes" then manage=manage+"|07"
	if request.form("manage08")="yes" then manage=manage+"|08"
	if request.form("manage09")="yes" then manage=manage+"|09"
	if request.form("manage10")="yes" then manage=manage+"|10"
	if request.form("manage11")="yes" then manage=manage+"|11"
	if request.form("manage12")="yes" then manage=manage+"|12"
	if request.form("manage13")="yes" then manage=manage+"|13"
	if request.form("manage14")="yes" then manage=manage+"|14"
	if request.form("manage15")="yes" then manage=manage+"|15"
	if request.form("manage16")="yes" then manage=manage+"|16"
	if request.form("manage17")="yes" then manage=manage+"|17"
	if request.form("manage18")="yes" then manage=manage+"|18"
	if request.form("manage19")="yes" then manage=manage+"|19"
	if request.form("manage20")="yes" then manage=manage+"|20"
	if request.form("manage21")="yes" then manage=manage+"|21"
	if request.form("manage22")="yes" then manage=manage+"|22"
	if request.form("manage23")="yes" then manage=manage+"|23"
	if request.form("manage24")="yes" then manage=manage+"|24"
	if request.form("manage25")="yes" then manage=manage+"|25"
	if request.form("manage26")="yes" then manage=manage+"|26"
	if request.form("manage27")="yes" then manage=manage+"|27"
	if request.form("manage28")="yes" then manage=manage+"|28"
	if request.form("manage29")="yes" then manage=manage+"|29"
	if request.form("manage30")="yes" then manage=manage+"|30"
	if request.form("manage31")="yes" then manage=manage+"|31"
	if request.form("manage32")="yes" then manage=manage+"|32"
	if request.form("manage33")="yes" then manage=manage+"|33"
	if request.form("manage34")="yes" then manage=manage+"|34"
	if request.form("manage35")="yes" then manage=manage+"|35"
	if request.form("manage36")="yes" then manage=manage+"|36"
	if request.form("manage37")="yes" then manage=manage+"|37"
	if request.form("manage38")="yes" then manage=manage+"|38"
	if request.form("manage39")="yes" then manage=manage+"|39"
	if request.form("manage40")="yes" then manage=manage+"|40"
	if request.form("manage41")="yes" then manage=manage+"|41"
	if request.form("manage42")="yes" then manage=manage+"|42"
	if request.form("manage43")="yes" then manage=manage+"|43"
	if request.form("manage44")="yes" then manage=manage+"|44"
	if request.form("manage45")="yes" then manage=manage+"|45"
	if request.form("manage46")="yes" then manage=manage+"|46"
	if request.form("manage47")="yes" then manage=manage+"|47"
	if request.form("manage48")="yes" then manage=manage+"|48"
	if request.form("manage49")="yes" then manage=manage+"|49"
	if request.form("manage50")="yes" then manage=manage+"|50"
	if request.form("manage51")="yes" then manage=manage+"|51"
	if request.form("manage52")="yes" then manage=manage+"|52"
	if request.form("manage53")="yes" then manage=manage+"|53"
	if request.form("manage54")="yes" then manage=manage+"|54"
	if request.form("manage55")="yes" then manage=manage+"|55"
	if request.form("manage56")="yes" then manage=manage+"|56"
	if request.form("manage57")="yes" then manage=manage+"|57"
	if request.form("manage58")="yes" then manage=manage+"|58"
	if request.form("manage59")="yes" then manage=manage+"|59"
	if request.form("manage60")="yes" then manage=manage+"|60"
	if request.form("manage61")="yes" then manage=manage+"|61"
	if request.form("manage62")="yes" then manage=manage+"|62"
	if request.form("manage63")="yes" then manage=manage+"|63"
	if request.form("manage64")="yes" then manage=manage+"|64"
	
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select * from Bs_User where UserName='"&admin&"'"
	rs.open sql,conn,1,3
		rs("admin_Manage")=manage
		rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing

	response.write "<script language='javascript'>"
	response.write "alert('����Ȩ�����óɹ���');"
	response.write "location.href='Bs_adminPermissions.asp?admin="&admin&"';"
	response.write "</script>"
end if


sub deladmin()

	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select * from Bs_User where UserName='"&deladm&"'"
	rs.open sql,conn,1,3
		if (rs.bof and rs.eof) then
			response.write "<script language='javascript'>"
			response.write "alert('����ʧ�ܣ�û�д˹���Ա��');"
			response.write "location.href='Bs_adminPermissions.asp?admin="&admin&"';"
			response.write "</script>"
			response.end
		elseif deladm=request.cookies("buyok")("admin") then
			response.write "<script language='javascript'>"
			response.write "alert('������ɾ���Լ���');"
			response.write "location.href='Bs_adminPermissions.asp?admin="&admin&"';"
			response.write "</script>"
			response.end
		else
		rs.delete
		rs.update
			response.write "<script language='javascript'>"
			response.write "alert('���ѳɹ�ɾ������Ա"&deladm&"');"
			response.write "location.href='Bs_adminPermissions.asp?admin="&admin&"';"
			response.write "</script>"
			response.end
		end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing

end sub
%>

<%htmlend%>
<script>
//�б�����б�ɫ
var TbRow = document.getElementById("goaler");
if (TbRow != null){
   for (var i=0;i<TbRow.rows.length ;i++ )
     {
        if (TbRow.rows[i].rowIndex%2==1){
            TbRow.rows[i].className="TdOver";
        }
        else{
            TbRow.rows[i].className="TdDown";
        }
     }
}
</script>
