<!--#include file="Inc/syscode.asp" -->
<!--#include file="inc/md5.asp"-->
<%
dim Action,UserName
dim rsUser,sqlUser
Action=trim(request("Action"))
UserName=trim(request("UserName"))
if Action="" and session("UserName")="" then
	response.redirect "Server.asp"
end if
if Action="Modify" and UserName<>"" then
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from [User] where UserName='" & UserName & "'"
	rsUser.Open sqlUser,conn,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br>�Ҳ���ָ�����û���"
	else
		dim OldPassword,Password,PwdConfirm
		OldPassword=trim(request("OldPassword"))
		Password=trim(request("Password"))
		PwdConfirm=trim(request("PwdConfirm"))
		if OldPassword="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br>����������룡"
		else
			if Instr(OldPassword,"=")>0 or Instr(OldPassword,"%")>0 or Instr(OldPassword,chr(32))>0 or Instr(OldPassword,"?")>0 or Instr(OldPassword,"&")>0 or Instr(OldPassword,";")>0 or Instr(OldPassword,",")>0 or Instr(OldPassword,"'")>0 or Instr(OldPassword,",")>0 or Instr(OldPassword,chr(34))>0 or Instr(OldPassword,chr(9))>0 or Instr(OldPassword,"��")>0 or Instr(OldPassword,"$")>0 then
				errmsg=errmsg+"<br>�������к��зǷ��ַ�"
				founderr=true
			else
				if md5(OldPassword)<>rsUser("Password") then
					FoundErr=True
					ErrMsg=ErrMsg & "<br>������ľ����벻�ԣ�û��Ȩ���޸ģ�"
				end if
			end if
		end if
		if strLength(Password)>12 or strLength(Password)<6 then
			founderr=true
			errmsg=errmsg & "<br>������������(���ܴ���12С��6)��"
		else
			if Instr(Password,"=")>0 or Instr(Password,"%")>0 or Instr(Password,chr(32))>0 or Instr(Password,"?")>0 or Instr(Password,"&")>0 or Instr(Password,";")>0 or Instr(Password,",")>0 or Instr(Password,"'")>0 or Instr(Password,",")>0 or Instr(Password,chr(34))>0 or Instr(Password,chr(9))>0 or Instr(Password,"��")>0 or Instr(Password,"$")>0 then
				errmsg=errmsg+"<br>�������к��зǷ��ַ�"
				founderr=true
			end if
		end if
		if PwdConfirm="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br>������ȷ�����룡"
		else
			if PwdConfirm<>Password then
				FoundErr=True
				ErrMsg=ErrMsg & "<br>ȷ�������������벻һ�£�"
			end if
		end if
		if FoundErr<>true then
			rsUser("Password")=md5(Password)
			rsUser.update
		end if
	end if
	rsUser.close
	set rsUser=nothing
	if FoundErr=True then
		call WriteErrMsg()
	else
	    response.write"<SCRIPT language=JavaScript>alert('�ɹ��޸����룡');"
        response.write"javascript:history.go(-1)</SCRIPT>"
		'call WriteSuccessMsg("�ɹ��޸����룡")
	end if
else
%>
<!-- #include file="Inc/Head.asp" -->
<div id="body">
  <table class="bodyTable">
    <tr>
      <td  class="bodyLeft"><!-- ��Ա���� -->
        <!-- #include file="L_vip.asp" -->
      </td>
      <td class="bodyLine"></td>
      <td class="bodyRight">
          <!-- �޸Ļ�Ա���� -->
          <div class="navPath">�޸�����</div>
          <div>
            <form action="ManagePwd.asp" method="post" name="Form1" id="Form1">
              <table width="400" border="0" align="center" cellpadding="2" cellspacing="2" class='border'>
                  <tr align="center" class='title'>
                    <td height="30" colspan="2"><font color="#FF6600" size="2" class="en"><b>�޸��û�����</b></font></td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="120" align="right"><b>�� �� ����</b></td>
                    <td><%=session("UserName")%>
                      <input name="UserName" type="hidden" value="<%=session("UserName")%>" /></td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="120" align="right"><b>�� �� �룺</b></td>
                    <td><input   type="password" maxlength="16" size="30" name="OldPassword" /></td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="120" align="right"><b>�� �� �룺</b></td>
                    <td><input   type="password" maxlength="16" size="30" name="Password" /></td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="120" align="right"><strong>ȷ�����룺</strong></td>
                    <td><input name="PwdConfirm" type="password" id="PwdConfirm" size="30" maxlength="16" /></td>
                  </tr>
                  <tr align="center" class="tdbg" >
                    <td height="40" colspan="2"><input name="Action" type="hidden" id="Action" value="Modify" />
                      <input name="Submit" type="submit" id="Submit" value=" �� �� " /></td>
                  </tr>
                </table>
            </form>
          </div>
      </td>
    </tr>
  </table>
</div>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>
<%
end if
call CloseConn()
%>