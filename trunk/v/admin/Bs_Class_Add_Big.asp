<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
dim Action,BigClassName,BigClassID,rs,FoundErr,ErrMsg,Wrd1,Wrd2,Brief
Action=trim(Request("Action"))
BigClassID=trim(request("BigClassID"))
BigClassName=trim(request("BigClassName"))
Brief=trim(request("Brief"))
Wrd1=trim(request("Wrd1"))
Wrd2=trim(request("Wrd2"))
Wrd3=trim(request("Wrd3"))
Wrd4=trim(request("Wrd4"))
Wrd5=trim(request("Wrd5"))
Wrd6=trim(request("Wrd6"))
Wrd7=trim(request("Wrd7"))

if Action="Add" then
	if BigClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>大类名不能为空！</li>"
	end if
	if BigClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>大类ID不能为空！</li>"
	end if
	if FoundErr<>True then
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.open "Select * From BigClass Where BigClassName='" & BigClassName & "'",conn,1,3
		if not (rs.bof and rs.EOF) then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>大类“" & BigClassName & "”已经存在！</li>"
		else
    	 	rs.addnew
     		rs("BigClassName")=BigClassName
			rs("BigClassID")=BigClassID
			rs("Brief")=Brief
			rs("Wrd1")=Wrd1
			rs("Wrd2")=Wrd2
			rs("Wrd3")=Wrd3
			rs("Wrd4")=Wrd4
			rs("Wrd5")=Wrd5
			rs("Wrd6")=Wrd6
			rs("Wrd7")=Wrd7
			
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
function checkBig()
{
  if (editor.EditMode.checked==true)
      document.myform.Brief.value=editor.HtmlEdit.document.body.innerText;
  else
  	  document.myform.Brief.value=editor.HtmlEdit.document.body.innerHTML; 
	  
  if (document.myform.BigClassName.value=="")
  {
    alert("大类名称不能为空！");
    document.myform.BigClassName.focus();
    return false;
  }
   if (document.myform.BigClassID.value=="")
  {
    alert("大类ID不能为空！");
    document.myform.BigClassID.focus();
    return false;
  }
  if (document.myform.Brief.value=="")
  {
    alert("简介内容不能为空！");
	editor.HtmlEdit.focus();
	return false;
  }
  return true; 
}
function loadForm()
{
  editor.HtmlEdit.document.body.innerHTML=document.myform.Brief.value;
  return true
}
</script>
<script>   
setTimeout('loadForm()',1000);
</script> 

<!-- #include file="Inc/Head.asp" -->
<table align="center" class="a2">
	<tr>
		<td class="a1">视 频 类 别 设 置</td>
	</tr>
	<tr class="a4">
		<td>
				<form name="myform" method="post" action="Bs_Class_Add_Big.asp" onSubmit="return checkBig()">
        <table width="100%">
          <tr> 
            <td colspan="2" class="iptsubmit">添加大类</td>
          </tr>
		  <tr> 
            <td class="iptlab">大类ID：</td>
            <td class="ipt"><input name="BigClassID" type="text" id="BigClassID" value="<%=rs("BigClassID")%>"></td>
          </tr>
          <tr> 
            <td class="iptlab">大类名称：</td>
            <td class="ipt"> <input name="BigClassName" type="text" size="20" maxlength="30"></td>
          </tr>
		  <tr> 
            <td class="iptlab">字段1：</td>
            <td class="ipt"> <input name="Wrd1" type="text" size="20" maxlength="30"></td>
          </tr>
		  <tr> 
            <td class="iptlab">字段2：</td>
            <td class="ipt"> <input name="Wrd2" type="text" size="20" maxlength="30"></td>
          </tr>
		  <tr> 
            <td class="iptlab">字段3：</td>
            <td class="ipt"> <input name="Wrd3" type="text" size="20" maxlength="30"></td>
          </tr>
		  <tr> 
            <td class="iptlab">字段4：</td>
            <td class="ipt"> <input name="Wrd4" type="text" size="20" maxlength="30"></td>
          </tr>
		  <tr> 
            <td class="iptlab">字段5：</td>
            <td class="ipt"> <input name="Wrd5" type="text" size="20" maxlength="30"></td>
          </tr>
		  <tr> 
            <td class="iptlab">字段6：</td>
            <td class="ipt"> <input name="Wrd6" type="text" size="20" maxlength="30"></td>
          </tr>
		  <tr> 
            <td class="iptlab">字段7：</td>
            <td class="ipt"> <input name="Wrd7" type="text" size="20" maxlength="30"></td>
          </tr>
		  <tr>
		    <td class="iptlab">大类简介：</td>
			<td class="ipt">
			<textarea name="Brief" style="display:none"></textarea>
			<iframe ID="editor" src="../editor1.asp" frameborder=1 scrolling=no width="700" height="205"></iframe>								            </td>
		  </tr>
          <tr>
            <td colspan="2" class="iptsubmit"> 
                <input name="Action" type="hidden" id="Action" value="Add">
                <input name="Add" type="submit" value=" 添 加 ">
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
call CloseConn()
%>
