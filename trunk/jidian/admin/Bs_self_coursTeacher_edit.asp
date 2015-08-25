<!--#include file="bsconfig.asp"-->
<%if Request.QueryString("no")="eshop" then

'id=request("id")
teachName=request("rName")
teachSubj=request("teachSubj")
teachInfo=request("teachInfo")
teachSay=request("teachSay")
Face=request("Face")
honor=server.htmlencode(Trim(Request("honor")))



If teachName="" Then
response.write "SORRY <br>"
response.write "请输入教师姓名"
response.end
end if
If teachInfo="" Then
response.write "SORRY <br>"
response.write "教师简介不能为空"
response.end
end if

If honor="" Then 
honor="UploadFiles/nophoto.jpg"
end if


Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from Teacher where teachName= '"&teachName&"'"
rs.open sql,conn,1,3

rs("teachName")=teachName
rs("teachSubj")=teachSubj
rs("teachInfo")=teachInfo
rs("teachSay")=teachSay
rs("Face")=Face
rs("honor")=honor
rs("time")=now()
rs.update
rs.close
response.redirect "Bs_coursTeacher.asp"
end if
%>
<%
rName=request.querystring("rName")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * from Teacher where teachName='"&rName&"'",conn,3,3
%>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("57")%>
<BR>
<%=rName%>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>修改教师信息</strong></td>
	</tr>
	<tr class="a4">
		<td>
          <form method="post" name="myform" action="Bs_self_coursTeacher_edit.asp?no=eshop">
            <input type=hidden name=rName value=<%=rName%>>
							<table width="100%" border="0" cellspacing="2" cellpadding="3">
								<tr> 
									<td width="18%" class="iptlab"> *姓名</td>
									<td width="31%" class="ipt"><input type="text" name="teachName" maxlength="40" value="<%=rs("teachName")%>"></td>
									<td width="31%" rowspan="5" bgcolor="#C0C0C0"><div id=pci class="1"><img src="../<%=rs("honor")%>" width="120" height="160" id="tou" style="margin:10px;"></div></td>
								</tr>
								<tr> 
									<td width="18%" class="iptlab">主讲课</td>
									<td width="31%" class="ipt"><input name="teachSubj" type="text" value="<%=rs("teachSubj")%>" maxlength="100">最多字数：100</td>
								</tr>
								<tr> 
									<td width="18%" class="iptlab"> *简介</td>
									<td width="31%" class="ipt">
  <script language=javascript>
function gbcount(message,used,remain)
{
var max;
max=300;
if(message.value.length > max){
message.value = message.value.substring(0,max);
used.value = max;
remain.value = 0;
alert('不能超过300个字!');
}
else{
used.value = message.value.length;
remain.value = max - used.value;
}
}
</script>
									  
  <textarea name='teachInfo' cols='32' rows='4' onkeydown=gbcount(this.form.teachInfo,this.form.used,this.form.remain); onkeyup=gbcount(this.form.teachInfo,this.form.used,this.form.remain);><%=rs("teachInfo")%></textarea>
  <br>
									  最多字数：300
									  已写：
									  <INPUT disabled maxLength="4" nam="used" size="3" value="0">
									  剩余：									  <INPUT disabled maxLength="4" name="remain" size="3" value="300">
								    </td>
								</tr>
								<tr> 
									<td width="18%" class="iptlab"> 寄语</td>
									<td width="31%" class="ipt"><input name="teachSay" type="text"  value="<%=rs("teachSay")%>" maxlength="50">
								    最多字数：50</td>
								</tr>
                                <tr> 
									<td width="18%" class="iptlab"> 背景图</td>
								  <td width="31%" class="ipt">
									    <style>
#pci{ width:150px; height:200px;}
.0 {background: url(../img/wallimages/face0.png) no-repeat;}
.1 {background: url(../img/wallimages/face1.png) no-repeat}
.2 {background: url(../img/wallimages/face2.png) no-repeat}
.3 {background: url(../img/wallimages/face3.png) no-repeat }
.4 {background: url(../img/wallimages/face4.png) no-repeat}
.5 {background: url(../img/wallimages/face5.png) no-repeat }
</style>
								  <select name="Face" id="Face" onChange="document.getElementById('pci').className=options[selectedIndex].value;">
								    <option value="0">0</option>
								    <option value="1">1</option>
								    <option value="2">2</option>
								    <option value="3">3</option>
								    <option value="4">4</option>
								    <option value="5">5</option>
								    </select>
								    </td>
							    </tr>
								<tr> 
									<td class="iptlab">教师图</td>
									<td colspan="2" class="ipt"><input name="honor" type="text" id="honor" value="<%=rs("honor")%>" size="32" maxlength="100"> 
										<font color="#FF0000"> * 图片地址</font><br>（图片大小120*160像素或者保持3:4的宽高比）</td>
								</tr>
								<tr> 
									<td colspan="3" class="iptsubmit"> 
											<input name="submit" type="submit" value="确定">
											&nbsp; 
											<input name="reset" type="reset" value="从来">
										</td>
								</tr>
								<tr> 
									<td colspan="3">图片上传</td>
								</tr>
								<tr> 
									<td colspan="3" align="center">
											<iframe name="ad" frameborder=0 width=100% height=25 scrolling=auto src=../uploadb.asp></iframe>
									</td>
								</tr>
							</table>
          </form>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>

