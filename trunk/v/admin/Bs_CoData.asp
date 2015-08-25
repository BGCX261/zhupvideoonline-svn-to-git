<!--#include file="bsconfig.asp"-->

<%if Request.QueryString("no")="eshop" then

BsCompanyName=request("BsCompanyName")
BsAddress=request("BsAddress")
BsLinkman=request("BsLinkman")
BsPhone=request("BsPhone")
qq=request("qq")
BsWeb=request("BsWeb")
BsEmail=request("BsEmail")
BsZipcode=request("BsZipcode")
BsICPcode=request("BsICPcode")



If BsCompanyName="" Then
response.write "SORRY <br>"
response.write "请填入公司名称!"
response.end
end if
If BsAddress="" Then
response.write "SORRY <br>"
response.write "请填入地址!"
response.end
end if
If BsPhone="" Then
response.write "SORRY <br>"
response.write "请填入电话!"
response.end
end if
If BsWeb="" Then
response.write "SORRY <br>"
response.write "请填入网址!"
response.end
end if
If BsEmail="" Then
response.write "SORRY <br>"
response.write "请填入Email!"
response.end
end if


Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from syswork where code='101'"
rs.open sql,conn,1,3

	rs("BsCompanyName")=BsCompanyName
	rs("BsAddress")=BsAddress
	rs("BsLinkman")=BsLinkman
	rs("BsPhone")=BsPhone
	rs("qq")=qq
	rs("BsWeb")=BsWeb
	rs("BsEmail")=BsEmail
	rs("BsZipcode")=BsZipcode
	rs("BsICPcode")=BsICPcode

rs.update
rs.close
response.redirect "Bs_CoData.asp"
end if
%>
<%
'code=request.querystring("code")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="Select * From syswork where code='101'"
rs.Open sql,conn,3,3
%>
<%
htmlname="Bs_CoData.asp"
tablename="syswork"
%>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("13")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">基 本 资 料 设 置</td>
	</tr>
	<tr class="a4">
		<td>
          <form method="post" name="myform" action="Bs_CoData.asp?no=eshop">
            <input type=hidden name=id value="101">
			<table width="100%">
								<tr> 
									<td class="iptlab"> 网站名称</td>
									<td class="ipt"><input type="text" name="BsCompanyName" maxlength="80" size="50" value="<%=rs("BsCompanyName")%>"></td>
								</tr>
								<tr> 
									<td class="iptlab"> 地址</td>
									<td class="ipt"><input type="text" name="BsAddress" maxlength="80" size="50" value="<%=rs("BsAddress")%>"></td>
								</tr>
								<tr> 
									<td class="iptlab"> 联系人</td>
									<td class="ipt"><input type="text" name="BsLinkman" maxlength="30" size="25" value="<%=rs("BsLinkman")%>"></td>
								</tr>
								<tr> 
									<td class="iptlab"> 电话</td>
									<td class="ipt"><input type="text" name="BsPhone" maxlength="30" size="25" value="<%=rs("BsPhone")%>"></td>
								</tr>
								<tr> 
									<td class="iptlab"> 网址</td>
									<td class="ipt"><input type="text" name="BsWeb" maxlength="50" size="25" value="<%=rs("BsWeb")%>"> 
									  注意: 不用填入 http://</td>
								</tr>
								<tr> 
									<td class="iptlab"> Email</td>
									<td class="ipt"><input type="text" name="BsEmail" maxlength="30" size="25" value="<%=rs("BsEmail")%>"></td>
								</tr>
								<tr> 
									<td class="iptlab"> 邮政编码</td>
									<td class="ipt"><input type="text" name="BsZipcode" maxlength="30" size="25" value="<%=rs("BsZipcode")%>"></td>
								</tr>
								<tr> 
									<td class="iptlab"> ICP备案号</td>
									<td class="ipt"><input type="text" name="BsICPcode" maxlength="30" size="25" value="<%=rs("BsICPcode")%>"></td>
								</tr>
								<tr> 
									<td colspan="2" class="iptsubmit">
											<input name="submit" type="submit" value="确认修改">
											<input name="reset" type="reset" value=" 取 消 ">										</td>
								</tr>
							</table>
          </form>
		</td>
	</tr>
</table>
<%htmlend%>
<script language=Javascript>
	
	function checkchar(str)
	{
		str=str.toLowerCase()
		if (str.search("<"+"%")>0)  
		{
			window.alert("("+str+")中有非法字符,请检查!")
			return false;
		}
		if (str.search("<scrip"+"t")>0)
		{
			window.alert("("+str+")中有非法字符,请检查!")
			return false;
		}
		return true;
	}

		
</script>
