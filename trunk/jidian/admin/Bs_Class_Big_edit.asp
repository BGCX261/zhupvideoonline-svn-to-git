<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/Function.asp"-->

<%
dim BigClassID,Action,rs,NewBigClassName,OldBigClassName,FoundErr,ErrMsg,NewWrd1,oldWrd1,NewWrd2,oldWrd2,NewWrd3,oldWrd3,NewWrd4,oldWrd4,NewWrd5,oldWrd5,NewWrd6,oldWrd6,NewWrd7,oldWrd7,NewBrief,OldBrief
BigClassID=trim(Request("BigClassID"))
Action=trim(Request("Action"))
NewBigClassName=trim(Request("NewBigClassName"))
OldBigClassName=trim(Request("OldBigClassName"))
NewBrief=trim(Request("NewBrief"))
OldBrief=trim(Request("OldBrief"))
NewWrd1=trim(Request("NewWrd1"))
OldWrd1=trim(Request("OldWrd1"))
NewWrd2=trim(Request("NewWrd2"))
OldWrd2=trim(Request("OldWrd2"))
NewWrd3=trim(Request("NewWrd3"))
OldWrd3=trim(Request("OldWrd3"))
NewWrd4=trim(Request("NewWrd4"))
OldWrd4=trim(Request("OldWrd4"))
NewWrd5=trim(Request("NewWrd5"))
OldWrd5=trim(Request("OldWrd5"))
NewWrd6=trim(Request("NewWrd6"))
OldWrd6=trim(Request("OldWrd6"))
NewWrd7=trim(Request("NewWrd7"))
OldWrd7=trim(Request("OldWrd7"))

if BigClassID="" then
  response.Redirect("Bs_Class.asp")
end if
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open "Select * from BigClass where BigClassID=" & CLng(BigClassID),conn,1,3
if rs.Bof and rs.EOF then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>此大类不存在！</li>"
else
	if Action="Modify" then
		if NewBigClassName="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>大类名不能为空！</li>"
		end if
		if FoundErr<>True then
			rs("BigClassName")=NewBigClassName
			rs("Brief")=NewBrief
			rs("Wrd1")=NewWrd1
			rs("Wrd2")=NewWrd2
			rs("Wrd3")=NewWrd3
			rs("Wrd4")=NewWrd4
			rs("Wrd5")=NewWrd5
			rs("Wrd6")=NewWrd6
			rs("Wrd7")=NewWrd7
			
			rs("Admin")=Admin
    	 	rs.update
			rs.Close
	     	set rs=Nothing
		if NewBigClassName<>OldBigClassName or  NewWrd1<>OldWrd1 or  NewWrd2<>OldWrd2 then
				conn.execute "Update SmallClass set BigClassName='" & NewBigClassName & "' where BigClassName='" & OldBigClassName & "'"
				conn.execute "Update Product set BigClassName='" & NewBigClassName & "' where BigClassName='" & OldBigClassName & "'"
				
     		end if		
			call CloseConn()
     		Response.Redirect "Bs_Class.asp" 
		end if
	end if
	if FoundErr=True then
		call WriteErrMsg()
	else
%>
<script language="JavaScript" type="text/JavaScript">
function checkBig()
{
  if (editor.EditMode.checked==true)
	  document.myform.NewBrief.value=editor.HtmlEdit.document.body.innerText;
  else
	  document.myform.NewBrief.value=editor.HtmlEdit.document.body.innerHTML; 
	  
  if (document.myform.NewBigClassName.value=="")
  {
    alert("大类名称不能为空！");
    document.myform.NewBigClassName.focus();
    return false;
  }
   if (document.myform.NewBrief.value=="")
  {
    alert("内容不能为空！");
	editor.HtmlEdit.focus();
	return false;
  }
  return true;  
}
function loadForm()
{
  editor.HtmlEdit.document.body.innerHTML=document.myform.NewBrief.value;
  return true
}
</script>
<body onLoad="javascipt:setTimeout('loadForm()',1000);">
<!-- #include file="Inc/Head.asp" -->
<table align="center" class="a2">
	<tr>
		<td class="a1">网络课程类别设置</td>
	</tr>
	<tr class="a4">
		<td>
			<form name="myform" method="post" action="Bs_Class_Big_edit.asp" onSubmit="return checkBig()">
			<table width="100%">
				<tr> 
					<td colspan="2" class="iptsubmit">修改大类名称</td>
				</tr>
				<tr> 
					<td class="iptlab">大类ID：</td>
					<td class="ipt"> 
					<%=rs("BigClassID")%><input name="BigClassID" type="hidden" id="BigClassID" value="<%=rs("BigClassID")%>">
					<input name="OldBigClassName" type="hidden" id="OldBigClassName" value="<%=rs("BigClassName")%>">
					<input name="OldBrief" type="hidden" id="OldBrief" value="<% Brief=replace(rs("Brief"),"<br>",chr(13))
Brief=replace(Brief,"&nbsp;"," ")%>">
					<input name="OldWrd1" type="hidden" id="OldWrd1" value="<%=rs("Wrd1")%>">
					<input name="OldWrd2" type="hidden" id="OldWrd2" value="<%=rs("Wrd2")%>">
					<input name="OldWrd3" type="hidden" id="OldWrd3" value="<%=rs("Wrd3")%>">
					<input name="OldWrd4" type="hidden" id="OldWrd4" value="<%=rs("Wrd4")%>">
					<input name="OldWrd5" type="hidden" id="OldWrd5" value="<%=rs("Wrd5")%>">
					<input name="OldWrd6" type="hidden" id="OldWrd6" value="<%=rs("Wrd6")%>">
					<input name="OldWrd7" type="hidden" id="OldWrd7" value="<%=rs("Wrd7")%>">
					</td>
				</tr>
				<tr> 
					<td class="iptlab">大类名称：</td>
					<td class="ipt">
					<input name="NewBigClassName" type="text" id="NewBigClassName" value="<%=rs("BigClassName")%>" size="20" maxlength="30"></td>
				</tr>
				<tr> 
					<td class="iptlab">字段1：</td>
					<td class="ipt">
					<input name="NewWrd1" type="text" id="NewWrd1" value="<%=rs("Wrd1")%>" size="20" maxlength="30"></td>
				</tr>
				<tr> 
					<td class="iptlab">字段2：</td>
					<td class="ipt">
					<input name="NewWrd2" type="text" id="NewWrd2" value="<%=rs("Wrd2")%>" size="20" maxlength="30"></td>
				</tr>
				<tr> 
					<td class="iptlab">字段3：</td>
					<td class="ipt">
					<input name="NewWrd3" type="text" id="NewWrd3" value="<%=rs("Wrd3")%>" size="20" maxlength="30"></td>
				</tr>
				<tr> 
					<td class="iptlab">字段4：</td>
					<td class="ipt">
					<input name="NewWrd4" type="text" id="NewWrd4" value="<%=rs("Wrd4")%>" size="20" maxlength="30"></td>
				</tr>
				<tr> 
					<td class="iptlab">字段5：</td>
					<td class="ipt">
					<input name="NewWrd5" type="text" id="NewWrd5" value="<%=rs("Wrd5")%>" size="20" maxlength="30"></td>
				</tr>
				<tr> 
					<td class="iptlab">字段6：</td>
					<td class="ipt">
					<input name="NewWrd6" type="text" id="NewWrd6" value="<%=rs("Wrd6")%>" size="20" maxlength="30"></td>
				</tr>
				<tr> 
					<td class="iptlab">字段7：</td>
					<td class="ipt">
					<input name="NewWrd7" type="text" id="NewWrd7" value="<%=rs("Wrd7")%>" size="20" maxlength="30"></td>
				</tr>
				<tr class="iptlab"> 
			      <td class="iptlab">大类简介：</td>
				  <td class="ipt">
				  <textarea name="NewBrief"   id="NewBrief" style="display:none"><%=rs("Brief")%>
				  </textarea>
				  <iframe ID="editor" src="../editor1.asp" frameborder=1 scrolling=no width="700" height="205"></iframe>	</td>
				</tr>
				<tr> 
					<td colspan="2" class="iptsubmit">
					<input name="Action" type="hidden" id="Action" value="Modify">
					<input name="Save" type="submit" id="Save" value=" 修 改 ">
					</td>
				</tr>
			</table> 
			</form>	
		</td>
	</tr>
</table>
<%htmlend%>
<%
	end if
end if
rs.close
set rs=nothing
call CloseConn()
%>