<!--#include file="bsconfig.asp"-->

<%
dim sql,rsBigClass,rsSmallClass,ErrMsg
set rsBigClass=server.CreateObject("adodb.recordset")
rsBigClass.open "Select * From BigClass",conn,1,3
%>
<script language="JavaScript" type="text/JavaScript">
function checkBig()
{
  if (document.form1.BigClassName.value=="")
  {
    alert("�������Ʋ���Ϊ�գ�");
    document.form1.BigClassName.focus();
    return false;
  }
}
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
function ConfirmDelBig()
{
   if(confirm("ȷ��Ҫɾ���˴�����ɾ���˴���ͬʱ��ɾ����������С�࣬���Ҳ��ָܻ���"))
     return true;
   else
     return false;
	 
}

function ConfirmDelSmall()
{
   if(confirm("ȷ��Ҫɾ����С����һ��ɾ�������ָܻ���"))
     return true;
   else
     return false;
	 
}
</script>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("48")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">����γ��������</td>
	</tr>
	<tr class="a4">
		<td>����ѡ�<a href="Bs_Class_Add_Big.asp">��Ӵ���</a><br>
      <table width="100%">
        <tr bgcolor="#999999"> 
          <td height="24" align="center"><strong>��Ŀ����</strong></td>
          <td align="center"><strong>����ѡ��</strong></td>
        </tr>
<%
do while not rsBigClass.eof
%>
        <tr bgcolor="#E3E3E3"> 
          <td height="20" bgcolor="#C0C0C0"><img src="../Images/tree_folder4.gif" width="15" height="15"><%=rsBigClass("BigClassName")%></td>
          <td align="center" bgcolor="#C0C0C0"><a href="Bs_Class_Add_Small.asp?BigClassName=<%=rsBigClass("BigClassName")%>">�������Ŀ</a> 
            | <a href="Bs_Class_Big_edit.asp?BigClassID=<%=rsBigClass("BigClassID")%>">�޸�</a> 
            | <a href="Bs_Class_Del_Big.asp?BigClassName=<%=Server.URLEncode(rsBigClass("BigClassName"))%>" onClick="return ConfirmDelBig();">ɾ��</a></td>
        </tr>
<%
set rsSmallClass=server.CreateObject("adodb.recordset")
rsSmallClass.open "Select * From SmallClass Where BigClassName='" & rsBigClass("BigClassName") & "'",conn,1,3
if not(rsSmallClass.bof and rsSmallClass.eof) then
do while not rsSmallClass.eof
%>
        <tr bgcolor="#E3E3E3"> 
          <td height="22">&nbsp;&nbsp;<img src="../Images/tree_folder3.gif" width="15" height="15"><%=rsSmallClass("SmallClassName")%></td>
          <td align="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
            <a href="Bs_Class_Small_edit.asp?SmallClassID=<%=rsSmallClass("SmallClassID")%>">�޸�</a> 
            | <a href="Bs_Class_Del_Small.asp?SmallClassName=<%=Server.URLEncode(rsSmallClass("SmallClassName"))%>" onClick="return ConfirmDelSmall();">ɾ��</a></td>
        </tr>
<%
rsSmallClass.movenext
loop
end if
rsSmallClass.close
set rsSmallClass=nothing	
rsBigClass.movenext
loop
%>
      </table><BR>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>

<%
rsBigClass.close
set rsBigClass=nothing
call CloseConn()
%>