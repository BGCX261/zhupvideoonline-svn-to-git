<!--#include file="Inc/syscode.asp" -->
<!-- #include file="Inc/Head.asp" -->
<!--#include file="Inc/md5.asp"-->
<style type="text/css">
a:link {
	color: #333333;
	TEXT-DECORATION: none
}
a:visited {
	color: #333333;
	TEXT-DECORATION: none
}
a:active {
	color: #003300;
	TEXT-DECORATION: none
}
a:hover {
	color: #003300;
	TEXT-DECORATION: underline overline
}
.navtrail {
	COLOR: #eeeeee;
	FONT-SIZE: 12px;
	LINE-HEIGHT: 12px
}
a.navtrail:link {
	COLOR: #eeeeee;
	CURSOR: w-resize
}
a.navtrail:visited {
	COLOR: #eeeeee;
	CURSOR: w-resize
}
a.navtrail:active {
	COLOR: #eeeeee;
	CURSOR: w-resize
}
a.navtrail:hover {
	COLOR: #eeeeee;
	CURSOR: e-resize
}
input {
	BORDER-TOP-WIDTH: 1px;
	PADDING-RIGHT: 1px;
	PADDING-LEFT: 1px;
	BORDER-LEFT-WIDTH: 1px;
	FONT-SIZE: 9pt;
	BORDER-LEFT-COLOR: #cccccc;
	BORDER-BOTTOM-WIDTH: 1px;
	BORDER-BOTTOM-COLOR: #cccccc;
	PADDING-BOTTOM: 1px;
	BORDER-TOP-COLOR: #cccccc;
	PADDING-TOP: 1px;
	HEIGHT: 18px;
	BORDER-RIGHT-WIDTH: 1px;
	BORDER-RIGHT-COLOR: #cccccc
}
<!--
td {
	font-family: "����";
	font-size: 9pt;
	color: #333333;
	text-decoration: none
}
-->
</style>
<%

'����Ķ����������д���
ShowSmallClassType=ShowSmallClassType_Default
MaxPerPage=MaxPerPage_Default

dim UserName,Password,PwdConfirm,Question,Answer,Sex,Email,HomePage,Comane,Add,Zip,Phone,Fox
UserName=trim(request("UserName"))
Password=trim(request("Password"))
PwdConfirm=trim(request("PwdConfirm"))
Question=trim(request("Question"))
Answer=trim(request("Answer"))
Sex=trim(Request("Sex"))
Email=trim(request("Email"))
HomePage=trim(request("HomePage"))
Comane=trim(request("Comane"))
Add=trim(request("Add"))
Zip=trim(request("Zip"))
Phone=trim(request("Phone"))
Fox=trim(request("Fox"))
if UserName="" or strLength(UserName)>25 or strLength(UserName)<4 then
	founderr=true
	errmsg=errmsg & "<br><li>�������û���(���ܴ���25С��4)</li>"
else
  	if Instr(UserName,"=")>0 or Instr(UserName,"%")>0 or Instr(UserName,chr(32))>0 or Instr(UserName,"?")>0 or Instr(UserName,"&")>0 or Instr(UserName,";")>0 or Instr(UserName,",")>0 or Instr(UserName,"'")>0 or Instr(UserName,",")>0 or Instr(UserName,chr(34))>0 or Instr(UserName,chr(9))>0 or Instr(UserName,"��")>0 or Instr(UserName,"$")>0 then
		errmsg=errmsg+"<br><li>�û����к��зǷ��ַ�</li>"
		founderr=true
	end if
end if
if Password="" or strLength(Password)>12 or strLength(Password)<6 then
	founderr=true
	errmsg=errmsg & "<br><li>����������(���ܴ���12С��6)</li>"
else
	if Instr(Password,"=")>0 or Instr(Password,"%")>0 or Instr(Password,chr(32))>0 or Instr(Password,"?")>0 or Instr(Password,"&")>0 or Instr(Password,";")>0 or Instr(Password,",")>0 or Instr(Password,"'")>0 or Instr(Password,",")>0 or Instr(Password,chr(34))>0 or Instr(Password,chr(9))>0 or Instr(Password,"��")>0 or Instr(Password,"$")>0 then
		errmsg=errmsg+"<br><li>�����к��зǷ��ַ�</li>"
		founderr=true
	end if
end if
if PwdConfirm="" then
	founderr=true
	errmsg=errmsg & "<br><li>������ȷ������(���ܴ���12С��6)</li>"
else
	if Password<>PwdConfirm then
		founderr=true
		errmsg=errmsg & "<br><li>�����ȷ�����벻һ��</li>"
	end if
end if
if Question="" then
	founderr=true
	errmsg=errmsg & "<br><li>������ʾ���ⲻ��Ϊ��</li>"
end if
if Answer="" then
	founderr=true
	errmsg=errmsg & "<br><li>����𰸲���Ϊ��</li>"
end if
if Sex="" then
	founderr=true
	errmsg=errmsg & "<br><li>�Ա���Ϊ��</li>"
else
	sex=cint(sex)
	if Sex<>0 and Sex<>1 then
		Sex=1
	end if
end if
if Email="" then
	founderr=true
	errmsg=errmsg & "<br><li>Email����Ϊ��</li>"
else
	if IsValidEmail(Email)=false then
		errmsg=errmsg & "<br><li>����Email�д���</li>"
   		founderr=true
	end if
end if
if Add="" then
	founderr=true
	errmsg=errmsg & "<br><li>רҵ�༶����Ϊ��</li>"
end if
if homepage="" then
	founderr=true
	errmsg=errmsg & "<br><li>��ѧ��ݲ���Ϊ��</li>"
end if
if Phone="" then
	founderr=true
	errmsg=errmsg & "<br><li>��ϵ�绰����Ϊ��</li>"
end if
if Comane="" then
	founderr=true
	errmsg=errmsg & "<br><li>��ʵ��������Ϊ��</li>"
end if

if founderr=false then
	dim sqlReg,rsReg
	sqlReg="select * from [User] where UserName='" & Username & "'"
	set rsReg=server.createobject("adodb.recordset")
	rsReg.open sqlReg,conn,1,3
	
	if not(rsReg.bof and rsReg.eof) then
		founderr=true
		errmsg=errmsg & "<br><li>��ע����û��Ѿ����ڣ��뻻һ���û��������ԣ�</li>"
	else
		rsReg.addnew
		rsReg("UserName")=UserName
		rsReg("Password")=md5(Password)
		rsReg("Question")=Question
		rsReg("Answer")=md5(Answer)
		rsReg("Sex")=Sex
		rsReg("Email")=Email
		rsReg("HomePage")=HomePage
		rsReg("Comane")=Comane
		rsReg("Add")=Add
		rsReg("Zip")=Zip
		rsReg("Phone")=Phone
		rsReg("Fox")=Fox
		rsReg.update
		founderr=false
	end if
	rsReg.close
	set rsReg=nothing
end if		
%>
<div id="body">
<table class="bodyTable">
  <tr>
    <td class="bodyLeft">
            <!-- #include file="L_login.asp" -->
    </td>
    <td class="bodyLine"></td>
    <td class="bodyRight"><table width="100%" border="0" align="right" cellpadding="0" cellspacing="0">
        <tr>
          <td>
            <div class="rtitle"><img src="img/title_ico.gif" /> �� Ա �� ��</div>
          </td>
        </tr>
        <tr>
          <td height="1"><%
if founderr=false then
call RegSuccess()
else
call WriteErrmsg()
end if
%></td>
        </tr>
      </table></td>
  </tr>
</table>
</div>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>
<%
call CloseConn

sub WriteErrMsg()
    response.write "<table align='center' width='300' border='0' cellpadding='2' cellspacing='0' class='border'>"
    response.write "<tr class='title'><td align='center' height='15'>�������µ�ԭ����ע���û���</td></tr>"
    response.write "<tr class='tdbg'><td align='left' height='100'><br>" & errmsg & "<p align='center'>��<a href='javascript:onclick=history.go(-1)'>�� ��</a>��<br></p></td></tr>"
	response.write "</table>" 
end sub

sub RegSuccess()
    response.write "<table align='center' width='300' border='0' cellpadding='2' cellspacing='0' class='border'>"
    response.write "<tr class='title'><td align='center' height='15'>�ɹ�ע���û���</td></tr>"
    response.write "<tr class='tdbg'><td align='left' height='100'><br>��ע����û�����" & UserName & "<p align='center'>��<a href='javascript:onclick=window.close()'>�� ��</a>��<br></p></td></tr>"
	response.write "</table>" 
end sub
%>
