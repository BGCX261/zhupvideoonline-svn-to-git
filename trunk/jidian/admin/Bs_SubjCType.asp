<!--#include file="bsconfig.asp"-->

<%
set rs=server.CreateObject("adodb.recordset")
rs.open "Select * From SubjCType",conn,1,3

if Request.QueryString("no")="yes" then
nbtyped= Trim(Request("nbtyped"))
Brief= Trim(Request("Brief"))
if nbtyped<>"" then
sql="delete from SubjCType where nbtyped='" & nbtyped & "'"
conn.Execute  sql
sql="delete from SubjConstruct where nbtyped='" & nbtyped & "'"
conn.Execute sql
end if


Set rs= Nothing
Set Conn = Nothing
Response.Redirect "Bs_SubjCType.asp"
end if



if Request.QueryString("no")="eshop" then
nbtyped=request.form("nbtyped")
Brief=request.form("Brief")

If nbtyped="" Then
response.write "SORRY <br>"
response.write "������!"
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from SubjCType "
rs.open sql,conn,1,3
rs.addnew
rs("nbtyped")=nbtyped
rs("Brief")=Brief

rs.update
Response.Redirect "Bs_SubjCType.asp"
end if
%>
<script>
function CheckForm(){
  if (document.myform.Brief.value=="")
  {
    alert("���ݲ���Ϊ�գ�");
	return false;
  }
  if (document.myform.typed.value=="")
  {
    alert("���Ͳ���Ϊ�գ�");
	myform.typed.focus();
	return false;
  }
  return true;  
}
</script>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("59")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">������</td>
	</tr>
	<tr class="a4">
		<td>
          <form method="post" action="Bs_SubjCType.asp?no=eshop" name="myform" onSubmit="return CheckForm()">
			<table width="100%">
                <tr> 
                  <td class="iptlab"> �������</td>
                  <td class="ipt"><input type="text" name="nbtyped" size="35" maxlength="40"></td>
                </tr>
                <tr> 
                  <td class="iptlab"> ����˵��</td>
                  <td class="ipt"><textarea name="Brief" id="Brief" cols="50" rows="10" style="width:770px;"></textarea>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" class="iptsubmit">  
                      <input type="submit" name="Submit" value="ȷ��">
                      <input type="reset" name="Submit2" value="����">
                    </td>
                </tr>
              </table>
              <br> 
            <table width="100%">
                <tr class="opttitle"> 
                  <td>�������</td>
                  <td>����</td>
                </tr>
                <% do while not rs.eof %>
                <tr class="opt"> 
                  <td><%=rs("nbtyped")%></td>
                  <td><a href="Bs_SubjCType_edit.asp?id=<%=rs("id")%>">�޸�</a> | <a href="Bs_SubjCType.asp?nbtyped=<%=rs("nbtyped")%>&amp;no=yes">ɾ��</a></td>
                </tr>
<%
rs.MoveNext 
loop 
%>
              </table>
<%
Set rs = Nothing 
Set Conn = Nothing 
%>
          </form>
		</td>
	</tr>
</table>
<script type="text/javascript" src="../xheditor/jquery.js"></script>
<script type="text/javascript" src="../xheditor/xheditor.js"></script>
<script type="text/javascript" src="../xheditor/xheditor_plugins/ubb.min.js"></script>
<script type="text/javascript">$('#Brief').xheditor({upLinkUrl:"upload.asp?immediate=1",upLinkExt:"zip,rar,txt",upImgUrl:"upload.asp?immediate=1",upImgExt:"jpg,jpeg,gif,png",upFlashUrl:"upload.asp?immediate=1",upFlashExt:"swf",upMediaUrl:"upload.asp?immediate=1",upMediaExt:"avi"/*,�Ƿ�ʹ��ubb����beforeSetSource:ubb2html,beforeGetSource:html2ubb*/});</script>
<%htmlend%>
