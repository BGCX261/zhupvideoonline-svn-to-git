<!--#include file="bsconfig.asp"-->
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("13")%>

<%
action=request("ok")
cebian=request("cebian")
if action="" then 
Set rs = conn.Execute("select * from adv") 
if cebian="" then cebian=rs("cebian")
%>


	<table align="center" class="a2">
	<form action=Bs_adv.asp method=post name=myform>
	
	<tr><td colspan=2 class="a1">首页图片广告设置</td></tr>


    <tr>
      <td class="iptlab">是否打开首页图片广告</td>
      <td class="ipt">             
      <input type="radio" value="0" <%if cebian=0 then %> checked<% end if%> name="cebian"  onclick="javascript:location.href='Bs_adv.asp?cebian=0';">否 &nbsp; <input type="radio" value="1" <%if cebian=1 then %> checked<% end if%>  name="cebian"  onclick="javascript:location.href='Bs_adv.asp?cebian=1';">是</td>        
    </tr>
<%
if cebian=1 then
%>
	

    <tr>
      <td class="iptlab">文字简介</td>
      <td class="ipt"> <input type=text value="<%=rs("advbrief")%>" name=advbrief size=40 maxlength=100>
      <font color="#FF0000"> * 100字内</font></td>
    </tr>
    <tr>
      <td class="iptlab">广告链接</td>
      <td class="ipt"> <input type=text value="<%=rs("advurl")%>" name=advurl size=40 maxlength=100></td>
    </tr>
    <tr> 
		<td class="iptlab"> 广告图片</td>
		<td class="ipt">
        <input type=text value="<%=rs("advpic")%>" name=honor size=40 maxlength=100 class="inputtext">
									<font color="#FF0000"> * 图片地址(推荐制作图片大小260px * 80px)</font></td>
	</tr>
    <tr> 
		<td colspan="2" align="center"><iframe name="ad" frameborder=0 width=100% height=25 scrolling=auto src=../uploadb.asp></iframe></td>
	</tr>


<%
end if
%>

	<tr><td colspan=2 class="iptsubmit">
<INPUT name="ok" TYPE="hidden" value="ok"><INPUT name=action TYPE="submit" value="保存设置"></td></tr>
</form></table>
	<%
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
end if%>

<%              
if action="ok" then

	if request.form("cebian")="" then
	  response.write "<script language='javascript'>"
	  response.write "alert('填写不完整或者不符合要求，单击“确定”返回上一页！');"
	  response.write "location.href='javascript:history.go(-1)';"
	  response.write "</script>"
	  response.end
	end if

      if request("cebian")<>0 then


	advpic=request.form("honor")
	advpic=split(advpic,".")
	n=UBound(advpic)
	if LCase(advpic(n))<>"jpg" and LCase(advpic(n))<>"gif"  then
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

 	 rs("cebian")=request.form("cebian")
	if request("cebian")<>0 then
 	 rs("advpic")=trim(request.form("honor"))
 	 rs("advurl")=trim(request.form("advurl"))
 	 rs("advbrief")=trim(request.form("advbrief"))
	end if

	rs.update
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing

	response.write "<script language='javascript'>"
	response.write "alert('首页广告图片设置成功！');"
	response.write "location.href='Bs_adv.asp';"
	response.write "</script>"
	response.end
end if%>

<%htmlend%>

