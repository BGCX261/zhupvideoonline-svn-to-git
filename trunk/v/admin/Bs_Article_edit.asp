<!--#include file="bsconfig.asp"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->

<%
dim ArticleID,rsArticle,FoundErr,ErrMsg

ArticleID=trim(request("ArticleID"))
FoundErr=False
if ArticleID="" then 
	response.Redirect("Bs_Article.asp")
end if

if Session("adminName")="" then 
	response.Redirect("Bs_Article.asp")
end if

sql="select * from Product where ArticleID=" & ArticleID & ""
Set rsArticle= Server.CreateObject("ADODB.Recordset")
rsArticle.open sql,conn,1,1
if FoundErr=True then
	call WriteErrMsg()
else
%>
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
		bigClass = "������ӷ��ࡣ"
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
  if (document.myform.Key.value=="")
  {
    alert("�ؼ��ֲ���Ϊ�գ�");
	document.myform.Key.focus();
	return false;
  }
  if (document.myform.Uploaddown.value=="" && document.myform.UploaddownJ.value =="")
  {
    alert("���ٱ����ϴ�һ����Ƶ�ļ�������д����Ӧ�ļ���С");
	document.myform.Uploaddown.focus();
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
<body onLoad="javascipt:setTimeout('loadForm()',1000);">
<table align="center" class="a2">
	<tr>
		<td class="a1">�� �� �� Ƶ �� �� �� Ϣ</td>
	</tr>
	<tr class="a4">
		<td>
			<form method="POST" name="myform" onSubmit="return CheckForm();" action="Bs_Article_Save.asp?action=Modify">
	
						<table width="100%">
							<tr> 
								<td class="iptlab">�������ࣺ</td>
								<td class="ipt"> 
<%
if session("purview")=3 or session("purview")=4 then
	response.write rsArticle("BigClassName") & "<input name='BigClassName' type='hidden' value='" & rsArticle("BigClassName") & "'>&gt;&gt;"
else		
			sql = "select * from BigClass"
			rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "������ӷ��ࡣ"
	else
%>
<select name="BigClassName" 
onChange="changeName(document.myform.BigClassName.options[document.myform.BigClassName.selectedIndex].value);" size="1">
<%
do while not rs.eof
%>
<option <% if rs("BigClassName")=rsArticle("BigClassName") then response.Write("selected") end if%> 
value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
<%
rs.movenext
loop
end if
rs.close
%>
</select> 
<%
end if
if session("purview")=4 then
	response.write rsArticle("SmallClassName") & "<input name='SmallClassName' type='hidden' value='" & rsArticle("SmallClassName") & "'>"
else
%>
<select name="SmallClassName">
<option value="" <%if rsArticle("SmallClassName")="" then response.write "selected"%>>��ָ��С��</option>
<%
sql="select * from SmallClass where BigClassName='" & rsArticle("BigClassName") & "'"
rs.open sql,conn,1,1
if not(rs.eof and rs.bof) then
do while not rs.eof
%>
<option <% if rs("SmallClassName")=rsArticle("SmallClassName") then response.Write("selected") end if%> 
value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
<%
rs.movenext
loop
end if
rs.close
%>
</select> 
<%
end if
%>								</td>
							</tr>
							<tr> 
								<td class="iptlab">��ţ�</td>
								<td class="ipt">
								<input name="Product_Id" type="text" id="Product_Id" value="<%=rsArticle("Product_Id")%>" size="15" maxlength="15"> 
								<span class="red">*</span> </td>
							</tr>
							<tr> 
								<td class="iptlab">���ƣ�</td>
								<td class="ipt">
								<input name="Title" type="text" id="Title" value="<%=rsArticle("Title")%>" size="50" maxlength="80"> 
								<span class="red">*</span></td>
							</tr>
                            <tr> 
								<td class="iptlab">����Ժϵ��</td>
								<td class="ipt"> <select name="faculty">
                                      <option selected="selected" value="<%=rsArticle("faculty")%>"><%=rsArticle("faculty")%></option>
                                      <option value="��������ϢѧԺ">��������ϢѧԺ</option>
                                      <option value="���ù���ѧԺ">���ù���ѧԺ</option>
                                      <option value="���ʺ����뽻��ѧԺ">���ʺ����뽻��ѧԺ</option>
                                      <option value="�������ϵ">�������ϵ</option>
                                      <option value="�������ϵ">�������ϵ</option>
                                      <option value="����ϵ">����ϵ</option>
                                      <option value="��������ϢѧԺ">��������ϢѧԺ</option>
                                      <option value="˼����">˼����</option>
                                      <option value="���˽���ѧԺ">���˽���ѧԺ</option>
                                      <option value="����������">����������</option>
                                    </select>
								<span class="red">*</span></td>
							</tr>
<%							
sql="select * from BigClass where BigClassName='" & bigClass & "'"
rs.open sql,conn,1,1
if rs("Wrd1")<>"" then
%>
							<tr> 
								<td class="iptlab"><%=trim(rs("Wrd1"))%>��</td>
								<td class="ipt">
								<input name="Zi1" type="text" id="Zi1" value="<%=rsArticle("Zi1")%>" size="50" maxlength="80"></td>
							</tr>
<%
end if
if rs("Wrd2")<>"" then	
%>	
							<tr> 
								<td class="iptlab"><%=trim(rs("Wrd2"))%>��</td>
								<td class="ipt">
								<input name="Zi2" type="text" id="Zi2" value="<%=rsArticle("Zi2")%>" size="50" maxlength="80">							</td>
							</tr>
<%
end if
if rs("Wrd3")<>"" then	
%>
							<tr> 
								<td class="iptlab"><%=trim(rs("Wrd3"))%>��</td>
								<td class="ipt">
								<input name="Zi3" type="text" id="Zi3" value="<%=rsArticle("Zi3")%>" size="50" maxlength="80">							</td>
							
<%
end if
if rs("Wrd4")<>"" then	
%>
							<tr> 
								<td class="iptlab"><%=trim(rs("Wrd4"))%>��</td>
							  <td class="ipt"> <input name="Zi4" type="text" id="Zi4" value="<%=rsArticle("Zi4")%>" size="50" maxlength="80"></td>
							</tr>
<%
end if
if rs("Wrd5")<>"" then	
%>
							<tr> 
								<td class="iptlab"><%=trim(rs("Wrd5"))%>��</td>
							  <td class="ipt"> <input name="Zi5" type="text" id="Zi5" value="<%=rsArticle("Zi5")%>" size="50" maxlength="80"></td>
							</tr>
<%
end if
if rs("Wrd6")<>"" then	
%>
							<tr> 
								<td class="iptlab"><%=trim(rs("Wrd6"))%>��</td>
							  <td class="ipt"> <input name="Zi6" type="text" id="Zi6" value="<%=rsArticle("Zi6")%>" size="50" maxlength="80"></td>
							</tr>
<%
end if
if rs("Wrd7")<>"" then	
%>
							<tr> 
								<td class="iptlab"><%=trim(rs("Wrd7"))%>��</td>
							  <td class="ipt"> <input name="Zi7" type="text" id="Zi7" value="<%=rsArticle("Zi7")%>" size="50" maxlength="80"></td>
							</tr>
<%
end if
rs.close
%>
							<tr> 
								<td class="iptlab">�ؼ��֣�</td>
								<td class="ipt">
								<input name="Key" type="text" id="Key" value="<%=trim(rsArticle("Key"))%>" size="50" maxlength="255"> 
								<span class="red">*</span><br>�������������Ƶ�����������ؼ��֣��м��ÿո�ֿ������ܳ���&quot;'*?()���ַ���</td>
							</tr>
							<tr> 
								<td class="iptlab">˵����</td>
								<td class="ipt"> </td>
							</tr>
								<tr class="iptlab"> 
								<td colspan="2" class="ipt"> 
								<textarea name="Content" style="display:none"><%=rsArticle("Content")%></textarea>
								<iframe ID="editor" src="../editor.asp" frameborder=1 scrolling=no width="700" height="405"></iframe>
								</td>
							</tr>
							<tr>
								<td class="iptlab">��ҳͼƬ�� 
								<input name="IncludePic" type="hidden" id="IncludePic" value="yes"></td>
								<td class="ipt">
								<input name="DefaultPicUrl" type="text" id="DefaultPicUrl" value="<%=rsArticle("DefaultPicUrl")%>" size="50" maxlength="180">
								<br>��ҳ��ͼƬ,ֱ�Ӵ��ϴ�ͼƬ��ѡ�� 
<select name="DefaultPicList" id="DefaultPicList" onChange="DefaultPicUrl.value=this.value;">
<option value="img/nopic.jpg" <% if rsArticle("DefaultPicUrl")="" then response.write "selected" %>>��ָ����ҳͼƬ</option>
<%
if rsArticle("UploadFiles")<>"" then
	dim IsOtherUrl
	IsOtherUrl=True
	if instr(rsArticle("UploadFiles"),"|")>1 then
		dim arrUploadFiles,intTemp
		arrUploadFiles=split(rsArticle("UploadFiles"),"|")						
		for intTemp=0 to ubound(arrUploadFiles)
			if rsArticle("DefaultPicUrl")=arrUploadFiles(intTemp) then
				response.write "<option value='" & arrUploadFiles(intTemp) & "' selected>" & arrUploadFiles(intTemp) & "</option>"
				IsOtherUrl=False
			else
				response.write "<option value='" & arrUploadFiles(intTemp) & "'>" & arrUploadFiles(intTemp) & "</option>"
			end if
		next
	else
		if rsArticle("UploadFiles")=rsArticle("DefaultPicUrl") then
			response.write "<option value='" & rsArticle("UploadFiles") & "' selected>" & rsArticle("UploadFiles") & "</option>"
			IsOtherUrl=False
		else
			response.write "<option value='" & rsArticle("UploadFiles") & "'>" & rsArticle("UploadFiles") & "</option>"		
		end if
	end If
	if IsOtherUrl=True then
		response.write "<option value='" & rsArticle("DefaultPicUrl") & "' selected>" & rsArticle("DefaultPicUrl") & "</option>"
	end if
end if
%>
</select>
								<input name="UploadFiles" type="hidden" id="UploadFiles" value="<%=rsArticle("UploadFiles")%>">	
															</td>
							</tr>
							<tr>
							  <td class="iptlab">��ƵJPGԤ��ͼ�ϴ�:</td>
							  <td class="ipt"><iframe class="TBGen" style="top:2px" ID="Uploaddown" src="../Progress_up.asp" frameborder=0 scrolling=no width="380" height="60"></iframe></td>
							</tr>
							<tr>
							  <td class="iptlab">��ƵԤ��ͼ�ϴ���ַ:</td>
							  <td class="ipt"><input name="Uploaddown" type="text" id="Uploaddown" value="<%=rsArticle("Uploaddown")%>" size="60" readonly> <input name="downSize" id="downSize" type="hidden" value="<%=rsArticle("downSize")%>" size="12"></td>
							</tr>
                            <tr>
							  <td class="iptlab">��Ƶ�ļ��ϴ�[FLV��ʽ]:</td>
							  <td class="ipt"><iframe class="TBGen" style="top:2px" ID="UploaddownJ" src="../Progress_upJ.asp" frameborder=0 scrolling=no width="380" height="60"></iframe></td>
							</tr>
							<tr>
							  <td class="iptlab">��Ƶ�ļ��ϴ���ַ:</td>
							  <td class="ipt"><input name="UploaddownJ" type="text" id="UploaddownJ" value="<%=rsArticle("UploaddownJ")%>" size="60" readonly>�ļ���С<input name="downSizeJ" id="downSizeJ" type="text" value="<%=rsArticle("downSizeJ")%>" size="12">K</td>
							</tr>
                            <tr>
							  <td class="iptlab">��Ƶ����С:</td>
							  <td class="ipt">��:
							    <input name="flvWidth" type="text" id="flvWidth" size="12" value="<%=rsArticle("flvWidth")%>" />
							    px 
							    ��:
							    <input name="flvHight" id="flvHight" type="text" size="12" value="<%=rsArticle("flvHight")%>" />
						      px<span class="red">����Ƶ�ĳ��߲��ɳ���640���̱߲�����480��</span></td>
							</tr>
							<tr> 
								<td class="iptlab">��ͨ����ˣ�</td>
								<td class="ipt">
                                <input name="Passed" type="checkbox" id="Passed" value="yes" <% if instr(manage,str)<=0 then response.Write("") else if rsArticle("Passed")=true then response.Write("checked") end if%> <%call checkverify("19")%>>
								��<span class="red">�����ѡ�еĻ���ֱ�ӷ�����</span></td>
							</tr>
							<tr> 
								<td class="iptlab">��ҳ��ʾ��</td>
							  <td class="ipt">
								<input name="Elite" type="checkbox" id="Elite" value="yes" <% if rsArticle("Elite")=true then response.Write("checked") end if%>>
								��<span class="red">�����ѡ�еĻ�������ҳ��ʾ��</span> <input name="showMov" type="checkbox" id="showMov" value="yes" <% if rsArticle("showMov")=true then response.Write("checked") end if%>>
								�Ƿ�����<span class="red">�����ѡ�еĻ�ǰ̨�û��������������Ʒ��</span></td>
							</tr>
							<tr> 
								<td class="iptlab">¼��ʱ�䣺</td>
								<td class="ipt">
								<input name="UpdateTime" type="text" id="UpdateTime" value="<%=now()%>" maxlength="50">
								��ǰʱ��Ϊ��<%=now()%> ע�ⲻҪ�ı��ʽ��</td>
							</tr>
							<tr> 
								<td colspan="2" class="iptsubmit"><input name="ArticleID" type="hidden" id="ArticleID" value="<%=rsArticle("ArticleID")%>">
								<input type="hidden" name='bigclass'>
				                <input name="Save" type="submit"  id="Save" value="�����޸Ľ��">
                                <input type="text" name="Author" value="<%=rsArticle("Author")%>" />
                                  </td>
							</tr>
						</table>
			</form>
		</td>
	</tr>
</table>
<%htmlend%>
<script language = "JavaScript">

function changeName(locationid){
	document.myform.bigclass.value=locationid
	document.myform.action="BS_Article_edit.asp";
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
<%
end if
rsArticle.close
set rsArticle=nothing
call CloseConn()
%>