<!--#include file="bsconfig.asp"-->

<%if Request.QueryString("no")="eshop" then

mailsrv=request("mailsrv")
mailincept=request("mailincept")
mailuser=request("mailuser")
mailpwd=request("mailpwd")


If mailsrv="" Then
response.write "SORRY <br>"
response.write "�������ʼ���������ַ!"
response.end
end if
If mailincept="" Then
response.write "SORRY <br>"
response.write "�������ռ����ַ!"
response.end
end if
If mailuser="" Then
response.write "SORRY <br>"
response.write "������stmp�������ַ!"
response.end
end if
If mailpwd="" Then
response.write "SORRY <br>"
response.write "������stmp����������!"
response.end
end if


Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from mailset where code='99'"
rs.open sql,conn,1,3

	rs("mailsrv")=mailsrv
	rs("mailincept")=mailincept
	rs("mailuser")=mailuser
	rs("mailpwd")=mailpwd
rs.update
rs.close
response.redirect "Bs_Setmail.asp"
end if
%>
<%
'code=request.querystring("code")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="Select * From mailset where code='99'"
rs.Open sql,conn,3,3
%>
<%call checkmanage("37")%>
<!-- #include file="Inc/Head.asp" -->
<table  class="bsOutTable">
	<tr>
		<td class="bsNavTitle">�ʼ���������</td>
	</tr>
	<tr>
		<td class="bsMainTd">
          <form method="post" name="myform" action="Bs_Setmail.asp?no=eshop" onSubmit="return ckform(this)">
            <input type=hidden name=id value="101">
			<table class="bsIptTable">
            								<tr> 
									<td class="iptLab"> �ռ������ַ</td>
									<td class="iptFom"><input type="text" name="mailincept" maxlength="80" size="50" value="<%=rs("mailincept")%>">
									��д��Ҫ���ʼ����Ǹ�����</td>
								</tr>
                                <tr> 
									<td colspan="2" style="border-bottom:#F00 solid 1px"> 
								   </td>
								</tr>
								<tr> 
									<td class="iptLab">�ʼ�����������</td>
									<td class="iptFom"><input type="text" name="mailsrv" maxlength="80" size="50" value="<%=rs("mailsrv")%>"></td>
								</tr>

                                <tr> 
									<td class="iptLab"> stmp���������ַ</td>
									<td class="iptFom"><input type="text" name="mailuser" maxlength="80" size="50" value="<%=rs("mailuser")%>"></td>
								</tr>
								<tr> 
									<td class="iptLab"> stmp������������</td>
									<td class="iptFom"><input type="text" name="mailpwd" maxlength="80" size="50" value="<%=rs("mailpwd")%>"></td>
								</tr>
								<tr> 
									<td colspan="2" class="iptSubmit"> 
											<input name="submit" type="submit" value="ȷ���޸�"> 
											<input name="reset" type="reset" value=" ȡ �� ">
								   </td>
								</tr>
							</table>
          </form>
		</td>
	</tr>
</table>
<%htmlend%>
<script language=Javascript>
	
	function ckform(myform)
	{
  if (myform.email.value == "")
  {
    alert("����д��������!");
    myform.email.focus();
    return (false);
  }
  var checkOK = "0123456789abcdefghijklmnopqrstuvwxyz@.";
  var checkStr = myform.email.value;
  var allValid = true;
  var decPoints = 0;
  var allNum = "";
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
    if (ch != ".")
      allNum += ch;
  }
  if (!allValid)
  {
    alert("���������ַ���зǷ��ַ�!");
    myform.email.focus();
    return (false);
  }
  

  if (myform.email.value !== "")
  {
     var checkOK2 = myform.email.value;
     var checkStr2 = "@.";
     var allValid2 = true;
     var decPoints2 = 0;
     var allNum2 = "";
     for (i = 0;  i < checkStr2.length;  i++)
     {
       ch2 = checkStr2.charAt(i);
       for (j = 0;  j < checkOK2.length;  j++)
         if (ch2 == checkOK2.charAt(j))
           break;
       if (j == checkOK2.length)
       {
         allValid2 = false;
         break;
       }
       if (ch2 != ".")
         allNum2 += ch2;
     }
     if (!allValid2)
     {
       alert("���������ַ��ȱ����Ч�ַ�!");
       myform.email.focus();
       return (false);
     }
  }
	}		
</script>
