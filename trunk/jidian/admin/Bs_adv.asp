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
	
	<tr><td colspan=2 class="a1">��ҳͼƬ�������</td></tr>


    <tr>
      <td class="iptlab">�Ƿ����ҳͼƬ���</td>
      <td class="ipt">             
      <input type="radio" value="0" <%if cebian=0 then %> checked<% end if%> name="cebian"  onclick="javascript:location.href='Bs_adv.asp?cebian=0';">�� &nbsp; <input type="radio" value="1" <%if cebian=1 then %> checked<% end if%>  name="cebian"  onclick="javascript:location.href='Bs_adv.asp?cebian=1';">��</td>        
    </tr>
<%
if cebian=1 then
%>
	

    <tr>
      <td class="iptlab">���ּ��</td>
      <td class="ipt"> <input type=text value="<%=rs("advbrief")%>" name=advbrief size=40 maxlength=100>
      <font color="#FF0000"> * 100����</font></td>
    </tr>
    <tr>
      <td class="iptlab">�������</td>
      <td class="ipt"> <input type=text value="<%=rs("advurl")%>" name=advurl size=40 maxlength=100></td>
    </tr>
    <tr> 
		<td class="iptlab"> ���ͼƬ</td>
		<td class="ipt">
        <input type=text value="<%=rs("advpic")%>" name=honor size=40 maxlength=100 class="inputtext">
									<font color="#FF0000"> * ͼƬ��ַ(�Ƽ�����ͼƬ��С260px * 80px)</font></td>
	</tr>
    <tr> 
		<td colspan="2" align="center"><iframe name="ad" frameborder=0 width=100% height=25 scrolling=auto src=../uploadb.asp></iframe></td>
	</tr>


<%
end if
%>

	<tr><td colspan=2 class="iptsubmit">
<INPUT name="ok" TYPE="hidden" value="ok"><INPUT name=action TYPE="submit" value="��������"></td></tr>
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
	  response.write "alert('��д���������߲�����Ҫ�󣬵�����ȷ����������һҳ��');"
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
	  response.write "alert('ͼƬ��ʽ����ȷ�����ͼƬ��֧��JPG��GIF��ʽ��ͼƬ��');"
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
	response.write "alert('��ҳ���ͼƬ���óɹ���');"
	response.write "location.href='Bs_adv.asp';"
	response.write "</script>"
	response.end
end if%>

<%htmlend%>

