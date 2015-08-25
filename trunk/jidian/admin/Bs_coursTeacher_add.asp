<!--#include file="bsconfig.asp"-->
<%if Request.QueryString("no")="eshop" then%>
<%
teachName=server.htmlencode(Trim(Request("teachName")))
teachSubj=server.htmlencode(Trim(Request("teachSubj")))
teachInfo=server.htmlencode(Trim(Request("teachInfo")))
teachSay=server.htmlencode(Trim(Request("teachSay")))
Face=server.htmlencode(Trim(Request("Face")))
honor=server.htmlencode(Trim(Request("honor")))
sortId=Trim(Request("sortId"))
%>
<%


If teachName="" Then
response.write "SORRY <br>"
response.write "请输入教师姓名!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if
If teachInfo="" Then
response.write "SORRY <br>"
response.write "教师简介不能为空!!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if

If honor="" Then 
honor="UploadFiles/nophoto.jpg"
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from Teacher "
rs.open sql,conn,1,3
rs.addnew
rs("teachName")=teachName
rs("teachSubj")=teachSubj
rs("teachInfo")=teachInfo
rs("teachSay")=teachSay
rs("Face")=Face
rs("honor")=honor
rs("sortId")=sortId
rs.update
rs.close
response.redirect "Bs_coursTeacher.asp"
end if
%>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("20")%>
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>添加教师信息</strong></td>
	</tr>
	<tr class="a4">
		<td>
          <form method="post" action="Bs_coursTeacher_add.asp?no=eshop" name="myform">
						<table width="100%" border="0" cellspacing="2" cellpadding="3">
							<tr> 
								<td width="18%" class="iptlab"> *姓名</td>
								<td width="31%" class="ipt"><input type="text" name="teachName" class="inputtext" maxlength="40"></td>
								<td width="31%" rowspan="5" class="ipt"><div id=pci class="1"><img src="../UploadFiles/nophoto.jpg" width="120" height="160" id="tou" style="margin:10px;"></div></td>
							</tr>
							<tr> 
								<td width="18%" class="iptlab"> 主讲课</td>
								<td width="31%" class="ipt"><input name="teachSubj" type="text" class="inputtext" maxlength="100">
							    最多字数：100</td>
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
								  
  <textarea name='teachInfo' cols='32' rows='4' onkeydown=gbcount(this.form.teachInfo,this.form.used,this.form.remain); onkeyup=gbcount(this.form.teachInfo,this.form.used,this.form.remain);></textarea>
  <br>
								  最多字数：300
								  已写：
								  <INPUT disabled maxLength="4" name="used" size="3" value="0">
								  剩余：								  <INPUT disabled maxLength="4" name=remain size="3" value="300">
							    </td>
							</tr>
							<tr> 
								<td width="18%" class="iptlab"> 寄语</td>
								<td width="31%" class="ipt"><input name="teachSay" type="text" class="inputtext" maxlength="50">
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
								    </select></td>
							</tr>
							<tr> 
								<td class="iptlab"> 教师图</td>
								<td colspan="2" class="ipt"><input name="honor" type="text" class="inputtext" id="honor"  size="32" maxlength="100">
									<font color="#FF0000"> * 图片地址</font><br>（图片大小120*160像素或者保持3:4的宽高比）</td>
							</tr>
                            <tr> 
								<td class="iptlab"> 排序序号</td>
								<td colspan="2" class="ipt">
                                <%Set rsSort = Server.CreateObject("ADODB.Recordset")
rsSort.Open "select * from Teacher where sortId =(select max(sortId) from Teacher)", conn,3,3


%>
                                <input name="sortId" type="text" class="inputtext" id="sortId" value="<%=rsSort("sortId") + 1%>"  size="6" maxlength="4" readonly="readonly">
								  <font color="#FF0000"> * 1-50之间的整数</font><br>
									（注意不要和已有序号重复,自动生成的序号默认是现有的最大值，一般新增时不要更改）
                                    <%
									rsSort.Close
	     	                        set rsSort=Nothing
									%>
                              </td>
							</tr>
							<tr> 
								<td colspan="3" class="iptsubmit"> 
										<input name="submit" type="submit" value="确定">
										&nbsp; 
										<input name="reset" type="reset" value="取消">
									</td>
							</tr>
							<tr> 
								<td colspan="3" >图片上传</td>
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

