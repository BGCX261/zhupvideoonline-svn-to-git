<!--#include file="include/head.asp"-->
<!--#include file="include/ofunctions.asp"-->
<!--#include file="include/md5.asp"-->
<%
'** �ļ�˵������Աע��
Dim Form_UserName,Form_Password,Form_Email,Form_Userweb,Form_Sign,Form_userPicUrl,vl_code
Dim Msg
Act=Request.QueryString("act")
If Act<>"new" And Act<>"edit" Then
	If User_ID="" Then
		Act="new"
		If Enable_Register = 0 Then Call ShowError("Ŀǰ��ͣע�ᣡ",0)
	Else
		Act="edit"
		Call LoadUser
	End If
Else
	If Act="new" Then
		If Enable_Register = 0 Then Call ShowError("Ŀǰ��ͣע�ᣡ",0)
		Call SaveUser
	Else
		Call EditUser
	End If
End If

Sub LoadUser
	Sql="select * from Users where UserName='" & SqlShow(User_ID) & "'"
	OpenRs(Sql)
	Form_UserName=User_ID
	Form_Password="v8KWs5Uj"
	Form_Email=Rs("Email")
	Form_Userweb=Rs("Userweb")
	Form_Sign=Rs("Sign")
	Form_userPicUrl=Rs("userPicUrl")
	Rs.Close
End Sub

Sub EditUser
	Form_Password=Request.Form("userpass")
	Form_Email=Request.Form("email")
	Form_Userweb=Request.Form("userweb")
	Form_Sign=Request.Form("Sign")
	vl_code=trim(request.form("vl_code"))
	Form_userPicUrl=trim(request.form("userPicUrl"))
	If Form_userPicUrl="" Then Form_userPicUrl="images/face/user.gif"
	If isValidEmail(Form_Email)=False  Then
	    Msg="�����ʼ���ַ����"
	   Exit Sub
	End If
	If cstr(vl_code)="" or cstr(Session("GetCode"))<>cstr(vl_code) then
	Call ShowError("��������ȷ����֤�룡",1)
	Exit Sub
	End If

	Sql="select * from Users where UserName='" & SqlShow(User_ID) & "'"
	OpenRs(Sql)
	If Form_Password<>"v8KWs5Uj" Then
	Rs("UserPass")=MD5(Form_Password)
	Response.cookies(systemkey)("userpassword")=MD5(Form_Password)
	End If
	Rs("Email")=Form_Email
	Rs("Userweb")=Form_Userweb
	Rs("Sign")=Form_Sign
	Rs("userPicUrl")=Form_userPicUrl
	Rs.Update
	Rs.Close
	Act="Save"
	Call ShowError("�޸ĳɹ���",0)
End Sub

Sub SaveUser
	vl_code=trim(request.form("vl_code"))
	Form_UserName=Request.Form("username")
	Form_Password=Request.Form("userpass")
	Form_Email=Request.Form("email")
	Form_Userweb=Request.Form("userweb")
	Form_Sign=Request.Form("Sign")
	Form_userPicUrl=trim(request.form("userPicUrl"))
	If Form_userPicUrl="" Then Form_userPicUrl="images/face/user.gif"
	If IsSpecial(Form_UserName)=True or CCEmpty(Form_UserName)<>Form_UserName or Form_UserName="" Then
		Msg="�û����в����԰��������ַ�!"
		Exit Sub
	End If
	Sql="select * from Users where UserName='" & SqlShow(Form_UserName) & "'"
	OpenRs(Sql)
	If Rs.RecordCount>0 Then
		Rs.Close
		Msg="���û����Ѿ����ڣ�������ѡ���û���"
		Exit Sub
	End If
	If isValidEmail(Form_Email)=False  Then
	        Rs.Close
	    Msg="�����ʼ���ַ����"
	    Exit Sub
	End If
	If strlen(Form_UserName)<3 Then
	        Rs.Close
		Msg="�û�������С��3���ַ�"
		Exit Sub
	End If
	If cstr(vl_code)="" or cstr(Session("GetCode"))<>cstr(vl_code) then
	Msg="��������ȷ����֤�룡"
	Exit Sub
	End If
	Rs.AddNew
	Rs("UserName")=Form_UserName
	Rs("UserPass")=MD5(Form_Password)
	Rs("UserGroup")="member"
	Rs("Email")=Form_Email
	Rs("Userweb")=Form_Userweb
	Rs("Sign")=Form_Sign
	Rs("userPicUrl")=Form_userPicUrl
	Rs("RegTime")=Date
	Rs("vip")=0

	Rs.Update
	Rs.Close
	If Enable_Cookies = 1 Then
			Response.cookies(systemkey)("userid")=Form_UserName
			Response.cookies(systemkey)("userpassword")=MD5(Form_Password)
			Response.cookies(systemkey)("usergroup")="member"
       Else
		Session(SystemKey & "User_ID")=Form_UserName
		Session(SystemKey & "User_Group")=Rs("UserGroup")
    End If

	Msg="�ѳɹ�ע�ᣬ�����Ѿ���ע���û���ݵ�¼"
	Act="Save"
End Sub
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%If Act="new" Then%>ע���û�<%Else%>�޸�����<%End If%></title>
<link rel="stylesheet" type="text/css" href="images/<%=strStyle%>">
<script language="javascript" src="include/js/edit.js"></script>
<script language="javascript">
function checkform()
{
	if(form1.username.value==""){
		form1.username.focus();
		return(false);
	}
	if(form1.userpass.value==""){
		form1.userpass.focus();
		return(false);
	}
	if(form1.passcheck.value==""){
		form1.passcheck.focus();
		return(false);
	}
	if(form1.passcheck.value!=form1.userpass.value){
		alert("������������벻һ��!");
		form1.userpass.focus();
		return(false);
	}
}
<%
If Msg<>"" Then
	Printl "alert(""" & Msg & """);"
End If
If Act="Save" Then
	Act="edit"
%>
window.opener.document.location.reload();
window.close();
<%
End If
%>
</script>
<body >
<form name="form1" action="register.asp?act=<%=Act%>" method="post" onSubmit="javascript:return checkform();">

<div class="reg_">
	<div class="reg_1">�û�����</div>

	<div class="reg_2">
		<span class="red_color">������</span><input <%if act="edit" then%>readonly <%end if%>maxlength="9" type="text" name="username" size="20" class="input" value="<%=HtmlEncode(Form_UserName)%>">
	</div>

	<div class="reg_2">
         <span class="red_color">���룺</span><input maxlength="20" type="password" name="userpass" size="20" class="input" value="<%=HtmlEncode(Form_Password)%>">
	</div>

	<div class="reg_2">
         <span class="red_color">ȷ�����룺</span><input maxlength="20" type="password" name="passcheck" size="20" class="input" value="<%=HtmlEncode(Form_Password)%>">
	</div>

	<div class="reg_2">
        <span class="red_color"> Email��</span><input maxlength="48" type="text" name="email" size="28" class="input" value="<%=HtmlEncode(Form_Email)%>">
	</div>
	
<%
dim Form_userpicurl2
If Form_userpicurl<>"" Then Form_userpicurl2=Form_userpicurl Else Form_userpicurl2="images/face/user.gif"
%>
    <div style="display:none;">
	     <span>�û�ͷ��</span><INPUT TYPE="text" size=42 NAME="userPicUrl"  id="userPicUrl" value="<%=Form_userPicUrl2%>" readonly>
	</div>

    <div class="display:none;">
		 <table title="ѡ��ͷ��" cellspacing=0 cellpadding=0 border=0 style="display:none;" onclick="javascript:open('INClude/Face.htm','','width=500,height=500,resizable,scrollbars')">
			  <tr>
			       <td>
			            <img src="<%=Form_userpicurl2%>" width=48 height=48 id="tus" name="tus"  border="0">
			       </td>
			  </tr>
	     </table>

	</div>

	<div class="reg_2">
	     <span class="red_color">��֤�룺</span>
		 <input name="vl_code" class="input" type="text" id="vl_code" size="8" maxlength="4" onKeyPress="if ((event.keyCode<48 || event.keyCode>57)) event.returnValue=false"><img src="include/GetCode.asp" alt= "�翴�������֤������ˢ��" style="cursor : pointer;" onClick="this.src='include/GetCode.asp'">
	</div>

    <div class="reg_4">
	  <input type="submit" value=" ȷ �� " name="B1" onMouseOut="javascript:this.style.backgroundColor='';" class="reg_button">&nbsp;&nbsp;
	  <input type="reset" value=" �� �� " name="B2" onMouseOut="javascript:this.style.backgroundColor='';" class="reg_button">
	</div>

</div>
</form>

</body>
</html>
<%CloseAll%>