<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/Function.asp"-->

<%
dim SmallClassID,Action,BigClassName, SmallClassName, OldSmallClassName,rs,FoundErr,ErrMsg

SmallClassID=trim(Request("SmallClassID"))
Action=trim(Request("Action"))
BigClassName=trim(Request.form("BigClassName"))
SmallClassName=trim(Request.form("SmallClassName"))
OldSmallClassName=trim(request.form("OldSmallClassName"))

sqlBig="select * from BigClass where BigClassName='" & BigClassName & "'"
Set rsBig= Server.CreateObject("ADODB.Recordset")
rsBig.open sqlBig,conn,1,1
BigClassName=rsBig("BigClassName")
rsBig.close

if SmallClassID="" then
	response.Redirect("Bs_Class.asp")
end if
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open "Select * from SmallClass where SmallClassID="&SmallClassID&"",conn,1,3
if rs.Bof or rs.EOF then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>此小类不存在！</li>"
else
	if Action="Modify" then
		if SmallClassName="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>小类名不能为空！</li>"
		end if
		if FoundErr<>True then
			rs("SmallClassName")=SmallClassName
     		rs.update
			rs.Close
    	 	set rs=Nothing
			if SmallClassName<>OldSmallClassName then
				conn.execute "Update Product set SmallClassName='" & SmallClassName & "' where BigClassName='" & BigClassName & "' and SmallClassName='" & OldSmallClassName & "'"
	     	end if	
			call CloseConn()
    	 	Response.Redirect "Bs_Class.asp"
		end if
	end if
	if FoundErr=True then
		call WriteErrMsg()
	else
%>
<!-- #include file="Inc/Head.asp" -->
<table align="center" class="a2">
	<tr>
		<td class="a1">网络课程类别设置</td>
	</tr>
	<tr class="a4">
		<td>
			<form name="form1" method="post" action="Bs_Class_Small_edit.asp">
			<table width="100%">
				<tr> 
					<td colspan="2" class="iptsubmit">修改小类名称</td>
				</tr>
				<tr> 
					<td class="iptlab">所属大类：</td>
					<td class="ipt"><%=rs("BigClassName")%> 
					<input name="SmallClassID" type="hidden" id="SmallClassID" value="<%=rs("SmallClassID")%>"> 
					<input name="OldSmallClassName" type="hidden" id="OldSmallClassName" value="<%=rs("SmallClassName")%>">  
					<input name="BigClassName" type="hidden" id="BigClassName" value="<%=rs("BigClassName")%>"></td>
				</tr>
				<tr> 
					<td class="iptlab">小类名称：</td>
					<td class="ipt">
					<input name="SmallClassName" type="text" id="SmallClassName" value="<%=rs("SmallClassName")%>" size="20" maxlength="60"></td>
				</tr>
				<tr> 
					<td colspan="2" class="iptsubmit">
					<input name="Action" type="hidden" id="Action4" value="Modify">
					<input name="Save" type="submit" id="Save" value=" 修 改 ">					</td>
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