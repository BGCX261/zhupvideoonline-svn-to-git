<!--#include file="bsconfig.asp"-->

<%
dim rs
dim sql
dim count
dim bigClass

set rs=server.createobject("adodb.recordset")
sql = "select * from SmallClass order by SmallClassID asc"
rs.open sql,conn,1,1

bigClass=trim(request.form("bigclass"))
if bigClass="" then
	dim sql6
	set rs6=server.createobject("adodb.recordset")
	sql6 = "select * from BigClass"
	rs6.open sql6,conn,1,1
	if rs6.eof and rs6.bof then
		bigClass = "���������Ŀ��"
	else
		bigClass=rs6("BigClassName")
	end if
	
	rs6.close
end if
%>
<script language = "JavaScript">
var onecount;
subcat = new Array();
        <%
        count = 0
        do while not rs.eof 
        %>
subcat[<%=count%>] = new Array("<%= trim(rs("SmallClassName"))%>","<%= trim(rs("BigClassName"))%>","<%= trim(rs("SmallClassName"))%>");
        <%
        count = count + 1
        rs.movenext
        loop
        rs.close
        %>
onecount=<%=count%>;

function changelocation(locationid)
    {
    document.myform.SmallClassName.length = 1; 
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.myform.SmallClassName.options[document.myform.SmallClassName.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
    }    

function CheckForm()
{
  if (editor.EditMode.checked==true)
	  document.myform.Content.value=editor.HtmlEdit.document.body.innerText;
  else
	  document.myform.Content.value=editor.HtmlEdit.document.body.innerHTML; 

  if (document.myform.Title.value=="")
  {
    alert("���ⲻ��Ϊ�գ�");
	document.myform.Title.focus();
	return false;
  }
  if (document.myform.Product_Id.value=="")
  {
    alert("��Ų���Ϊ�գ�");
	document.myform.Key.focus();
	return false;
  }
  if (document.myform.Key.value=="")
  {
    alert("�ؼ��ֲ���Ϊ�գ�");
	document.myform.Key.focus();
	return false;
  }
  if (document.myform.Content.value=="")
  {
    alert("���ݲ���Ϊ�գ�");
	editor.HtmlEdit.focus();
	return false;
  }
  return true;  
}
function loadForm()
{
  editor.HtmlEdit.document.body.innerHTML=document.myform.Content.value;
  return true
}
</script>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("50")%>
<body onLoad="javascipt:setTimeout('loadForm()',1000);">
<table align="center" class="a2">
	<tr>
		<td class="a1">�������γ�</td>
	</tr>
	<tr class="a4">
		<td>
			<form method="POST" name="myform" onSubmit="return CheckForm();" action="Bs_Article_Save.asp?action=add" target="_self">
						<table width="100%">
							<tr> 
								<td class="iptlab">�������</td>
								<td class="ipt"> 
<%
		sql = "select * from BigClass"
		rs.open sql,conn,1,1
if rs.eof and rs.bof then
	response.write "���������Ŀ��"
else
%>
<select name="BigClassName" onChange="changeName(document.myform.BigClassName.options[document.myform.BigClassName.selectedIndex].value);" size="1">
<option selected value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
<%
dim selclass
	selclass=rs("BigClassName")
		rs.movenext
	do while not rs.eof
%>
<option value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
<%
			rs.movenext
		loop
end if
	rs.close
%>
</select>
<select name="SmallClassName">
<option value="" selected>��ָ��С��</option>
<%
sql="select * from SmallClass where BigClassName='" & selclass & "'"
rs.open sql,conn,1,1
if not(rs.eof and rs.bof) then
%>
<option value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
<%
rs.movenext
do while not rs.eof
%>
<option value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
<%
rs.movenext
loop
end if
rs.close
%>
<%
ranNum=int(9*rnd)+10
iddata=month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
%>
</select>
����<input name="author" type="text" id="author" value='<%if session("realname") then response.write(session("realname")) else response.write ("������")end if%>' readonly="readonly" />								</td>
							</tr>
							<tr> 
								<td class="iptlab">��ţ�</td>
								<td class="ipt"> <input name="Product_Id" type="text" id="Product_Id2" value="<%=iddata%>" size="15" maxlength="15">
								<span class="red">*��Ų�������ͬ�����㲻��ȷ�����ظ�������Ķ�����</span></td>
							</tr>
							<tr> 
								<td class="iptlab">���ƣ�</td>
								<td class="ipt"> <input name="Title" type="text" id="Title2" size="50" maxlength="80">
								<span class="red">*</span></td>
							</tr>
<%							
sql="select * from BigClass where BigClassName='" & bigClass & "'"
rs.open sql,conn,1,1
if rs("Wrd1")<>"" then
%>
							<tr> 
							  <td class="iptlab"><%=trim(rs("Wrd1"))%>��
								</td>
							  <td class="ipt"> <input name="Zi1" type="text" id="Zi1" size="50" maxlength="80"></td>
							</tr>

<%
end if
if rs("Wrd2")<>"" then	
%>				
							<tr> 
							  <td class="iptlab"><%=trim(rs("Wrd2"))%>��</td>
							  <td class="ipt"> <input name="Zi2" type="text" id="Zi2" size="50" maxlength="80"></td>
							</tr>
<%
end if
if rs("Wrd3")<>"" then	
%>
							<tr> 
								<td class="iptlab"><%=trim(rs("Wrd3"))%>��</td>
							  <td class="ipt"> <input name="Zi3" type="text" id="Zi32" size="50" maxlength="80"></td>
							</tr>
<%
end if
if rs("Wrd4")<>"" then	
%>
							<tr> 
								<td class="iptlab"><%=trim(rs("Wrd4"))%>��</td>
							  <td class="ipt"> <input name="Zi4" type="text" id="Zi4" size="50" maxlength="80"></td>
							</tr>
<%
end if
if rs("Wrd5")<>"" then	
%>
							<tr> 
								<td class="iptlab"><%=trim(rs("Wrd5"))%>��</td>
							  <td class="ipt"> <input name="Zi5" type="text" id="Zi5" size="50" maxlength="80"></td>
							</tr>
<%
end if
if rs("Wrd6")<>"" then	
%>
							<tr> 
								<td class="iptlab"><%=trim(rs("Wrd6"))%>��</td>
							  <td class="ipt"> <input name="Zi6" type="text" id="Zi6" size="50" maxlength="80"></td>
							</tr>
<%
end if
if rs("Wrd7")<>"" then	
%>
							<tr> 
								<td class="iptlab"><%=trim(rs("Wrd7"))%>��</td>
							  <td class="ipt"> <input name="Zi7" type="text" id="Zi7" size="50" maxlength="80"></td>
							</tr>
<%
end if
rs.close
%>
							<tr> 
								<td class="iptlab">�ؼ��֣�</td>
								<td class="ipt"> <input name="Key" type="text" id="Key" size="50" maxlength="80">
								<span class="red">*<br>��������������£����������ؼ��֣��м��ÿո�ֿ������ܳ���&quot;&quot;'*?,.()���ַ���</span></td>
							</tr>
							<tr> 
								<td class="iptlab">˵����</td>
								<td class="ipt">&nbsp;</td>
							</tr>
							<tr> 
								<td colspan="2" class="ipt"> 
								<textarea name="Content" style="display:none"></textarea>
								<iframe ID="editor" src="../editor.asp" frameborder=1 scrolling=no width="700" height="405"></iframe>
								</td>
							</tr>
							<tr> 
								<td class="iptlab">��ҳͼƬ�� 
								<input name="IncludePic" type="hidden" id="IncludePic" value="yes"></td>
								<td class="ipt">
								<input name="DefaultPicUrl" type="text" id="DefaultPicUrl" value="img/nopic.jpg" size="40" maxlength="120"> 
								<br>��ҳ��ͼƬ,ֱ�Ӵ��ϴ�ͼƬ��ѡ�� 
								<select name="DefaultPicList" id="select" onChange="DefaultPicUrl.value=this.value;">
								<option selected value="img/nopic.jpg">��ָ����ҳͼƬ</option>
								</select> <input name="UploadFiles" type="hidden" id="UploadFiles">	
														</td>
							</tr>
							<tr> 
								<td class="iptlab">��ͨ����ˣ�</td>
								<td class="ipt">
                                <input name="Passed" type="checkbox" id="Passed" value="yes" <%call checkverify("53")%>>
								��<span class="red">�����ѡ�еĻ���ֱ�ӷ�����</span></td>
							</tr>
							<tr> 
								<td class="iptlab">��ҳ��ʾ��</td>
							  <td class="ipt"><input name="Elite" type="checkbox" id="Passed" value="yes">
								��<span class="red">�����ѡ�еĻ�������ҳ��ʾ��</span></td>
							</tr>
							<tr> 
								<td class="iptlab">¼��ʱ�䣺</td>
								<td class="ipt">
								<input name="UpdateTime" type="text" id="UpdateTime" value="<%=now()%>" maxlength="50">��ǰʱ��Ϊ��<%=now()%> ע�ⲻҪ�ı��ʽ��</td>
							</tr>
							<tr><td colspan="2" class="iptsubmit"><input	name="Add" type="submit"  id="Add" value=" �� �� " 
				onClick="document.myform.action='Bs_Article_Save.asp?action=add';document.myform.target='_self';">
				<input type="hidden" name='bigclass'>
				</td></tr>
						</table>

				
			</form>
		</td>
	</tr>
</table>
<%htmlend%>
<script language = "JavaScript">

function changeName(locationid){
	document.myform.bigclass.value=locationid
	document.myform.action="BS_Article_Add.asp";
	document.myform.submit();
}

if('<%=bigClass%>'!=''){
	document.myform.BigClassName.value='<%=bigClass%>';

    document.myform.SmallClassName.length = 1; 
    var locationid='<%=bigClass%>';
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.myform.SmallClassName.options[document.myform.SmallClassName.length] = 
					new Option(subcat[i][0], subcat[i][2]);
            }        
        }
}
</script>