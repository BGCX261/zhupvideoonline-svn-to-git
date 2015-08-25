<!--#include file="bsconfig.asp"-->
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("13")%>

<%
action=request("ok")
advfloat=request("advfloat")
if action="" then 
Set rs = conn.Execute("select * from adv") 
if advfloat="" then advfloat=rs("advfloat")
%>


	<table align="center" class="a2">
	<form action="Bs_advFloat.asp" method="post" name="myform">
	
	<tr><td colspan=2 class="a1">漂浮广告设置</td></tr>


    <tr>
      <td class="iptlab">是否打开漂浮广告</td>
      <td class="ipt">             
      <input type="radio" value="0" <%if advfloat=0 then %> checked<% end if%> name="advfloat"  onclick="javascript:location.href='Bs_advFloat.asp?advfloat=0';">否 &nbsp; <input type="radio" value="1" <%if advfloat=1 then %> checked<% end if%>  name="advfloat"  onclick="javascript:location.href='Bs_advFloat.asp?advfloat=1';">是</td>        
    </tr>
<%
if advfloat=1 then
%>
	

    <tr>
      <td class="iptlab">文字简介</td>
      <td class="ipt"> <input type=text value="<%=rs("titleFloat")%>" name="titleFloat" id="titleFloat" size="40" maxlength="40">
      <font color="#FF0000"> * 40字内</font></td>
    </tr>
    <tr>
      <td class="iptlab">广告链接</td>
      <td class="ipt"> <input type=text value="<%=rs("lnkFloat")%>" name="lnkFloat" id="lnkFloat" size="40" maxlength="100"></td>
    </tr>
    <tr> 
		<td class="iptlab"> 广告图片</td>
		<td class="ipt">
        <input type=text value="<%=rs("imgFloat")%>" name="honor" size="40" maxlength="100" class="inputtext">
									<font color="#FF0000"> * 图片地址(不可过大)</font></td>
	</tr>
    <tr> 
		<td colspan="2" align="center"><iframe name="ad" frameborder="0" width="100%" height="25" scrolling="auto" src="../uploadb.asp"></iframe></td>
	</tr>


<%
end if
%>

	<tr><td colspan=2 class="iptsubmit">
<input name="ok" TYPE="hidden" value="ok"><INPUT name="action" TYPE="submit" value="保存设置"></td></tr>
</form></table>
	<%
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
end if%>

<%              
if action="ok" then

	if request.form("advfloat")="" then
	  response.write "<script language='javascript'>"
	  response.write "alert('填写不完整或者不符合要求，单击“确定”返回上一页！');"
	  response.write "location.href='javascript:history.go(-1)';"
	  response.write "</script>"
	  response.end
	end if

      if request("advfloat")<>0 then


	imgFloat=request.form("honor")
	imgFloat=split(imgFloat,".")
	n=UBound(imgFloat)
	if LCase(imgFloat(n))<>"jpg" and LCase(imgFloat(n))<>"gif" and LCase(imgFloat(n))<>"png"  then
	  response.write "<script language='javascript'>"
	  response.write "alert('图片格式不正确，侧边图片仅支持JPG、GIF格式的图片。');"
	  response.write "location.href='javascript:history.go(-1)';"
	  response.write "</script>"
	  response.end
	end if

      end if

     Set rs=Server.CreateObject("ADODB.Recordset")
	 sql="select * from adv"
	 rs.open sql,conn,1,3

 	 rs("advfloat")=request.form("advfloat")
	if request("advfloat")<>0 then
 	 rs("imgFloat")=trim(request.form("honor"))
 	 rs("lnkFloat")=trim(request.form("lnkFloat"))
 	 rs("titleFloat")=trim(request.form("titleFloat"))
	end if

	rs.update
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing

	response.write "<script language='javascript'>"
	response.write "alert('首页广告图片设置成功！');"
	response.write "location.href='Bs_advFloat.asp';"
	response.write "</script>"
	response.end
end if%>

<%htmlend%>

