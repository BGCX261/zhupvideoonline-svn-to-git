<!--#include file="../include/head.asp"-->
<!--#include file="../include/ofunctions.asp"-->

<%
'** �ļ�˵������վ����
Dim Form_strRegTsEmail,Form_strSubTsEmail,Form_strReTsEmail,Form_strPostTsEmail,Form_strReSsbTsEmail
dim j,id,SetStyle
id=Request.QueryString("id")

If id="kw" Then
SetStyle="style=""display:none"""
End If

If User_Group<>"admin" Then
	Call ShowError("��ҳֻ�й���Ա���ܷ��ʣ�",1)
End If

Sub ShowObjectInstalled(strObjName)
	If IsObjInstalled(strObjName) Then
		Response.Write "<font color='#0066FF'><b>��</b></font>"
	Else
		Response.Write "<font color='red'><b>��</b>(��)</font>"
	End If
End Sub

Function IsObjInstalled(strClassString)
    On Error Resume Next
    IsObjInstalled = False
    Err = 0
    Dim xTestObj
    Set xTestObj = Server.CreateObject(strClassString)
    If 0 = Err Then IsObjInstalled = True
    Set xTestObj = Nothing
    Err = 0
End Function

dim ObjInstalled,Action
ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")
Action=trim(request("Action"))
if Action="" then
	Action="ShowConfig"
end If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<%
If Action="SaveConfig" Then
	Call SaveConfig()
Else
	Call ShowConfig()
End If

Call CloseAll()

sub ShowConfig()
%>
<form method="POST" action="?Action=SaveConfig" name="form">
<div class="DB">
    <div class="L1" <%=SetStyle%>>��ǰλ�ã�<a href="admin_show.asp">�������</a> &gt; ��վ��������</div>
	<table cellspacing="1" cellpadding="8" border="0" class="a23" bgcolor="#FFFFFF" align="center">

         <tr class="a10" <%=SetStyle%>> 
              <td colspan="2" class="bold red PadLeft">ע�⣺��ҳ�������ݲ�����ʹ��Ӣ��˫���ţ�&quot;����������ʹ�õ������õ������滻��</td>
         </tr>

         <tr class="a4" <%=SetStyle%>>
              <td width="30%" height="32" class="PadLeft"><b>��վ���ƣ�</b></td>
              <td width="69%" height="32" class="PadLeft2"> 
                   <input name="strSiteName" type="text" id="WebmasterEmail14" value="<%=strSiteName%>" size="57" maxlength="100" class="IT5">
		      </td>
    	 </tr>

		<tr class="a4" <%=SetStyle%>>
              <td width="30%" height="32" class="PadLeft"><b>��վ��ַ��</b></td>
              <td width="69%" height="32" class="PadLeft2"> 
                   <input name="strSiteUrl" type="text" id="WebmasterEmail16" value="<%=strSiteUrl%>" size="57" maxlength="100" class="IT5"> ������/����
              </td>
    	 </tr>

		 <tr class="a4" <%=SetStyle%>>
              <td width="30%" height="32" class="PadLeft"><b>������վ���:</b></td>
              <td width="69%" height="32" class="PadLeft2"> 
                   <input name="strStyle" type="text" id="SiteName1" value="<%=strStyle%>" size="13" maxlength="50" class="IT2"> Ĭ�Ϸ��default.css
		      </td>
    	 </tr>

		 <tr class="a4" <%=SetStyle%>>
              <td width="30%" height="32" class="PadLeft"><b>�Ƿ񿪷Ż�Աע��:</b></td>
              <td width="69%" height="32" class="PadLeft2"> 
                   <input name="Enable_Register" type="text" id="LogoUrl2" value="<%=Enable_Register%>" size="4" maxlength="255" class="IT3"> 0�رգ�1����
			  </td>
    	 </tr>

		 <tr class="a4" <%=SetStyle%>>
              <td width="30%" height="32" class="PadLeft"><b>�Ƿ���cookie��¼��ʽ��</b></td>
              <td width="69%" height="32" class="PadLeft2"> 
                   <input name="Enable_Cookies" type="text" id="LogoUrl0" value="<%=Enable_Cookies%>" size="4" maxlength="255" class="IT3"> 0�رգ�1����
			  </td>
    	 </tr>

		 <tr class="a4" <%=SetStyle%>>
              <td width="30%" height="32" class="PadLeft"><b>�ظ���ҳ����</b></td>
              <td width="69%" height="32" class="PadLeft2"> 
                   <input name="contextPageSize" type="text" id="WebmasterEmail1" value="<%=contextPageSize%>" size="4" maxlength="100" class="IT3"> ���ٸ��ظ�����ʼ��ҳ
			  </td>
    	 </tr>

		 <tr class="a4" <%=SetStyle%>>
              <td width="30%" height="32" class="PadLeft"><b>�б��ҳ����</b></td>
              <td width="69%" height="32" class="PadLeft2"> 
                   <input name="ListPageSize" type="text" id="WebmasterEmail2" value="<%=ListPageSize%>" size="4" maxlength="100" class="IT3"> ÿҳ�б���ʾ���������º󼴷�ҳ
			  </td>
    	 </tr>

		 <tr class="a4">
              <td width="30%" height="32" class="PadLeft"><b>�����ȴ�:<br></b>����|�ŷָ�����</td>
              <td width="69%" height="32" >
			       <textarea rows="4" name="Hotkw" cols="88" class="IT4"><%=Hotkw%></textarea>
			  </td>
    	 </tr>

		 <tr class="a4" <%=SetStyle%>>
              <td width="30%" height="32" class="PadLeft"><b>�����������:<br></b>����|�ŷָ���������</td>
              <td width="69%" height="32" >
			       <textarea rows="4" name="strBadWords" cols="88" class="IT4"><%=strBadWords%></textarea>
			  </td>
    	 </tr>

         <tr class="a4" <%=SetStyle%>>
              <td width="30%" height="32" class="PadLeft"><b>��Ȩ��Ϣ��</b><br>ע���滻��ȥ��˫����</td>
              <td width="69%" height="32" >
			       <textarea rows="4" name="strCopyright" cols="88" class="IT4"><%=strCopyright%></textarea>
			  </td>
    	 </tr>
    	
         <tr class="a4" <%=SetStyle%>> 
              <td width="30%" height="32" class="PadLeft"><b>������Ϣ��</b><br>ע���滻��ȥ��˫����</td>
              <td width="69%" height="32" > 
                   <textarea rows="4" name="strBeiAn" cols="88" class="IT4"><%=strBeiAn%></textarea>
			  </td>
         </tr>

         <tr class="a4" <%=SetStyle%>> 
              <td width="30%" height="32" class="PadLeft"><b>ҳ��ײ���Ϣ��</b><br>ע���滻��ȥ��˫����</td>
              <td width="69%" height="32" > 
                   <textarea rows="4" name="strfooterin" cols="88" class="IT4"><%=strfooterin%></textarea>
			  </td>
         </tr>

		 <tr class="a4" <%=SetStyle%>>
              <td width="30%" height="32" class="PadLeft"><b>��վ��飺</b><br>����վ������ʱ������ʾ�����ݡ�</td>
              <td width="69%" height="32" > 
                   <textarea rows="4" name="strAbout" cols="88" class="IT4"><%=strAbout%></textarea>
		      </td>
    	 </tr>

         <tr class="a4" <%=SetStyle%>> 
              <td width="30%" height="32" class="PadLeft"><b>�ؼ�����Ϣ��</b><br>������վ������������¼�����ŷָ�</td>
              <td width="69%" height="32" > 
                   <textarea rows="4" name="strKeyword" cols="88" class="IT4"><%=strKeyword%></textarea>
			  </td>
         </tr>

         <tr class="a4" <%=SetStyle%>> 
              <td width="30%" height="32" class="PadLeft"><b>��վͳ�ƴ��룺</b><br>ע���滻��ȥ��˫����</td>
              <td width="69%" height="32"> 
                   <textarea rows="4" name="strtongji" cols="88" class="IT4"><%=strtongji%></textarea>
			  </td>
         </tr>

		 <tr class="a4">
              <td height="39" colspan="2" class="center"> 
                   <input name="Action" type="hidden" id="Action" value="SaveConfig">
                   <input name="cmdSave" type="submit" id="cmdSave"  <%If ObjInstalled=false Then response.write "value=��������֧��FSO���ܱ������� disabled" else response.write "value=��������"%> class="btn">
		      </td>
         </tr>
<%
If ObjInstalled=false Then
	Response.Write "<tr><td height='50' colspan='3'>"
	Response.Write "<b><font color=red>��ķ�����δ����FSO(Scripting.FileSystemObject)���֧���ı���д�����ܱ������á�<br>��ֱ���޸ġ�include/config.asp���ļ��е����ݡ�</font>"
	Response.Write "<BR><a href='http://www.baidu.com/s?ie=gb2312&wd=FSO�������/�رշ���'>��˲鿴FSO�������/�رշ���</a></b></td></tr>"
End If
%>

    </table>
	</div>
<%end sub%>

</form>
</body>
</html>
<%
sub SaveConfig()
	dim tempTxt
	tempTxt= "<" & "%" & vbcrlf & vbcrlf

	tempTxt=tempTxt& "'++++++++++++++++��վ�����趨+++++++++++++++" & vbcrlf & vbcrlf

	tempTxt=tempTxt&"'�������=��֮��Ĳ��֣����ϸ���˵������" & vbcrlf & vbcrlf

	tempTxt=tempTxt&"'++++++++++++++++++++++++++++++++++++++++++" & vbcrlf & vbcrlf
		
	tempTxt=tempTxt&"strSiteName=" & chr(34) & trim(request("strSiteName")) & chr(34) &"     'վ������" & vbcrlf & vbcrlf
	
	tempTxt=tempTxt&"strSiteUrl=" & chr(34) & trim(request("strSiteUrl")) & chr(34) & "    'վ���ַ ������/����" &vbcrlf & vbcrlf

	tempTxt=tempTxt&"strStyle=" & chr(34) & trim(request("strStyle")) & chr(34) & "        '����Ĭ�Ϸ��" & vbcrlf & vbcrlf

	tempTxt=tempTxt&"Enable_Register=" & trim(request("Enable_Register")) & "        '�����Ƿ�����ע�ᣬ0�رգ�1����" & vbcrlf & vbcrlf
	
	tempTxt=tempTxt&"Enable_Cookies=" & trim(request("Enable_Cookies")) & "        '��������cookie��ʽ�����¼��Ϣ��0�رգ�1����" & vbcrlf & vbcrlf
	
	tempTxt=tempTxt&"contextPageSize=" & trim(request("contextPageSize")) & "        '�ظ���ҳ�������ٸ��ظ�����ʼ��ҳ" & vbcrlf & vbcrlf
	
	tempTxt=tempTxt&"ListPageSize=" & trim(request("ListPageSize")) & "        '�б��ҳ���������б����������ҳ" & vbcrlf & vbcrlf

	tempTxt=tempTxt&"Hotkw=" & chr(34) & trim(request("Hotkw")) & chr(34) & "        '�����ȴʣ���|�ŷֿ�" & vbcrlf & vbcrlf& vbcrlf & vbcrlf

	tempTxt=tempTxt&"strBadWords=" & chr(34) & trim(request("strBadWords")) & chr(34) & "        '����������ˣ���|�ŷֿ�" & vbcrlf & vbcrlf& vbcrlf & vbcrlf

	tempTxt=tempTxt&"'+++++++++++++++++++�������ð�Ȩ��ʾ����Ϣ++++++++++++++++" & vbcrlf & vbcrlf
	
	tempTxt=tempTxt&"'--------����һ���趨��Ȩ��Ϣ����ʾ��ҳβ����Ҫʹ��Ӣ��˫���ţ�����" & vbcrlf
	tempTxt=tempTxt&"strCopyright=" & chr(34) & trim(Replace(Replace(Replace(Replace(request("strCopyright"),chr(13),""),chr(10),""),chr(9),""),chr(34),"")) & chr(34) & vbcrlf & vbcrlf
	
	tempTxt=tempTxt&"'--------����һ���趨��Ȩ��������Ҫʹ��Ӣ��˫���ţ�����" & vbcrlf
	tempTxt=tempTxt&"strfooterin=" & chr(34) & trim(Replace(Replace(Replace(Replace(request("strfooterin"),chr(13),""),chr(10),""),chr(9),""),chr(34),"")) & chr(34) & vbcrlf & vbcrlf

	tempTxt=tempTxt&"'--------����һ���趨������Ϣ����Ҫʹ��Ӣ��˫���ţ�����" & vbcrlf
	tempTxt=tempTxt&"strBeiAn=" & chr(34) & trim(Replace(Replace(Replace(Replace(request("strBeiAn"),chr(13),""),chr(10),""),chr(9),""),chr(34),"")) & chr(34) & vbcrlf & vbcrlf
	
	tempTxt=tempTxt&"'--------����һ���趨��վͳ�ƴ��룬��Ҫʹ��Ӣ��˫���ţ�����" & vbcrlf
	tempTxt=tempTxt&"strtongji=" & chr(34) & trim(Replace(Replace(Replace(Replace(request("strtongji"),chr(13),""),chr(10),""),chr(9),""),chr(34),"")) & chr(34) & vbcrlf & vbcrlf
	
	tempTxt=tempTxt&"'--------����һ���趨��վ��Ҫ����Ҫʹ��Ӣ��˫���ţ�����" & vbcrlf
	tempTxt=tempTxt&"strAbout=" & chr(34) & trim(request("strAbout")) & chr(34) &"     " & vbcrlf & vbcrlf

	tempTxt=tempTxt&"'--------����һ���趨�ؼ�����Ϣ������վ��������������¼" & vbcrlf
	tempTxt=tempTxt&"strKeyword=" & chr(34) & trim(Replace(Replace(Replace(Replace(request("strKeyword"),chr(13),""),chr(10),""),chr(9),""),chr(34),"")) & chr(34) & vbcrlf & vbcrlf

	tempTxt=tempTxt&"%" & ">"

	WriteToTextFile "../include/config.asp",tempTxt,"GB2312"

    '���������ȴ�hotkw.txt�ļ�
	DIM str,i,tmpStr
	str=""&trim(request("Hotkw"))&""
	str=split(str,"|")
	For i=0 To ubound(str)
	tmpStr=tmpStr&"<a href=""search.asp?kw="&str(i)&""">"&str(i)&"</a>"
	Next
	Call Save("../txt/hotkw.txt",tmpStr)

	Call ShowError("���óɹ���",1)
    Response.end
end sub
%>