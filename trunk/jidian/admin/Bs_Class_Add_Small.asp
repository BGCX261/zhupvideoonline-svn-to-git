<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/Function.asp"-->

<%
dim Action,BigClassName,SmallClassName,rs,FoundErr,ErrMsg
Action=trim(Request("Action"))
BigClassName=trim(request("BigClassName"))
SmallClassName=trim(request("SmallClassName"))

sqlBig="select * from BigClass where BigClassName='" & BigClassName & "'"
Set rsBig= Server.CreateObject("ADODB.Recordset")
rsBig.open sqlBig,conn,1,1
BigClassName=rsBig("BigClassName")
rsBig.close

if Action="Add" then
	if BigClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>����������Ϊ�գ�</li>"
	end if
	if SmallClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>С��������Ϊ�գ�</li>"
	end if
	if FoundErr<>True then
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.open "Select * From SmallClass Where BigClassName='" & BigClassName & "' AND SmallClassName='" & SmallClassName & "'",conn,1,3
		if not rs.EOF then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>��" & BigClassName & "�����Ѿ����ڸ�С�ࡰ" & SmallClassName & "����</li>"
		else
     		rs.addnew
			rs("BigClassName")=BigClassName
    	 	rs("SmallClassName")=SmallClassName
     		rs.update
	     	rs.Close
    	 	set rs=Nothing
     		call CloseConn()
			Response.Redirect "Bs_Class.asp"  
		end if
	end if
end if
if FoundErr=True then
	call WriteErrMsg()
else
%>
<script language="JavaScript" type="text/JavaScript">
function checkSmall()
{
  if (document.form2.BigClassName.value=="")
  {
    alert("������Ӵ������ƣ�");
	document.form1.BigClassName.focus();
	return false;
  }

  if (document.form2.SmallClassName.value=="")
  {
    alert("С�����Ʋ���Ϊ�գ�");
	document.form2.SmallClassName.focus();
	return false;
  }
}
</script>
<!-- #include file="Inc/Head.asp" -->
<table align="center" class="a2">
	<tr>
		<td class="a1">����γ��������</td>
	</tr>
	<tr class="a4">
		<td>
			<form name="form2" method="post" action="Bs_Class_Add_Small.asp" onSubmit="return checkSmall()">
			<table width="100%">
				<tr> 
					<td colspan="2" class="iptsubmit">���С��</td>
				</tr>
				<tr> 
					<td class="iptlab">�������ࣺ</td>
					<td class="ipt"> 
<select name="BigClassName">
<%
dim rsBigClass
set rsBigClass=server.CreateObject("adodb.recordset")
rsBigClass.open "Select * From BigClass",conn,1,1
if rsBigClass.bof and rsBigClass.bof then
	response.write "<option>������Ӵ���</option>"
else
	do while not rsBigClass.eof
		if rsBigClass("BigClassName")=BigClassName then
			response.write "<option value='"& rsBigClass("BigClassName") & "' selected>" & rsBigClass("BigClassName") & "</option>"
		else
			response.write "<option value='"& rsBigClass("BigClassName") & "'>" & rsBigClass("BigClassName") & "</option>"
		end if
		rsBigClass.movenext
	loop
end if
rsBigClass.close
set rsBigClass=nothing
%>
</select>					</td>
				</tr>
				<tr> 
					<td class="iptlab">С�����ƣ�</td>
					<td class="ipt"><input name="SmallClassName" type="text" size="20" maxlength="50"></td>
				</tr>
				<tr> 
					<td colspan="2" class="iptsubmit"> 
<input name="Action" type="hidden" id="Action3" value="Add"><input name="Add" type="submit" value=" �� �� "></td>
				</tr>
			</table>
			</form>
		</td>
	</tr>
</table>
<%htmlend%>
<%
end if
call CloseConn()
%>