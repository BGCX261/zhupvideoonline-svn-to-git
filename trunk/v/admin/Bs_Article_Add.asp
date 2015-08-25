<!--#include file="bsconfig.asp"-->

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
		bigClass = "请先添加分类。"
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
    alert("标题不能为空！");
	document.myform.Title.focus();
	return false;
  }
  if (document.myform.Product_Id.value=="")
  {
    alert("编号不能为空！");
	document.myform.Product_Id.focus();
	return false;
  }
  if (document.myform.faculty.value=="")
  {
    alert("请选择所属院系！");
	document.myform.faculty.focus();
	return false;
  }
  if (document.myform.Key.value=="")
  {
    alert("关键字不能为空！");
	document.myform.Key.focus();
	return false;
  }
    if (document.myform.Uploaddown.value=="" || document.myform.UploaddownJ.value =="") 
  {
    alert("必须上传视频文件和对应的预览图！");
	document.myform.Uploaddown.focus();
	return false;
  }
  if (document.myform.Content.value=="")
  {
    alert("内容不能为空！");
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
<%call checkmanage("21")%>
<body onLoad="javascipt:setTimeout('loadForm()',1000);">
<table align="center" class="a2">
	<tr>
		<td class="a1">添 加 视 频 作 品</td>
	</tr>
	<tr class="a4">
		<td>
			<form method="POST" name="myform" onSubmit="return CheckForm();" action="Bs_Article_Save.asp?action=add" target="_self">
						<table width="100%">
							<tr> 
								<td class="iptlab">所属类别：</td>
								<td class="ipt"> 
<%
		sql = "select * from BigClass"
		rs.open sql,conn,1,1
if rs.eof and rs.bof then
	response.write "请先添加分类。"
else
%>
<select name="BigClassName" onChange="changeName(document.myform.BigClassName.options[document.myform.BigClassName.selectedIndex].value);" size="1">
<option selected value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
<%
dim selclass
	selclass=rs("BigClassName")
		rs.movenext
	do while not rs.eof
%>
<option value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
<%
			rs.movenext
		loop
end if
	rs.close
%>
</select>
<select name="SmallClassName">
<option value="" selected>不指定小类</option>
<%
sql="select * from SmallClass where BigClassName='" & selclass & "'"
rs.open sql,conn,1,1
if not(rs.eof and rs.bof) then
%>
<option value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
<%
rs.movenext
do while not rs.eof
%>
<option value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
<%
rs.movenext
loop
end if
rs.close
%>
<%
ranNum=int(9*rnd)+10
iddata=month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
%>
</select>								</td>
							</tr>
							<tr> 
								<td class="iptlab">编号：</td>
								<td class="ipt"> <input name="Product_Id" type="text" id="Product_Id2" value="<%=iddata%>" size="15" maxlength="15">
								<span class="red">*编号不可以相同，如你不能确定会重复，请勿改动他！</span></td>
							</tr>
							<tr> 
								<td class="iptlab">名称：</td>
								<td class="ipt"> <input name="Title" type="text" id="Title2" size="50" maxlength="80">
								<span class="red">*</span></td>
							</tr>
                            <tr> 
								<td class="iptlab">所属院系：</td>
								<td class="ipt"> <select name="faculty">
                                      <option selected="selected" value="">请选择所属院系</option>
                                      <option value="工程与信息学院">工程与信息学院</option>
                                      <option value="经济管理学院">经济管理学院</option>
                                      <option value="国际合作与交流学院">国际合作与交流学院</option>
                                      <option value="人文社科系">人文社科系</option>
                                      <option value="艺术设计系">艺术设计系</option>
                                      <option value="旅游系">旅游系</option>
                                      <option value="工程与信息学院">工程与信息学院</option>
                                      <option value="思政部">思政部</option>
                                      <option value="成人教育学院">成人教育学院</option>
                                      <option value="行政或其它">行政或其它</option>
                                    </select>
								<span class="red">*</span></td>
							</tr>
<%							
sql="select * from BigClass where BigClassName='" & bigClass & "'"
rs.open sql,conn,1,1
if rs("Wrd1")<>"" then
%>
							<tr> 
							  <td class="iptlab"><%=trim(rs("Wrd1"))%>：
								</td>
							  <td class="ipt"> <input name="Zi1" type="text" id="Zi1" size="50" maxlength="80">
                                  <span class="red">*</span></td>
							</tr>

<%
end if
if rs("Wrd2")<>"" then	
%>				
							<tr> 
							  <td class="iptlab"><%=trim(rs("Wrd2"))%>：</td>
							  <td class="ipt"> <input name="Zi2" type="text" id="Zi2" size="50" maxlength="80"></td>
							</tr>
<%
end if
if rs("Wrd3")<>"" then	
%>
							<tr> 
								<td class="iptlab"><%=trim(rs("Wrd3"))%>：</td>
							  <td class="ipt"> <input name="Zi3" type="text" id="Zi32" size="50" maxlength="80"></td>
							</tr>
<%
end if
if rs("Wrd4")<>"" then	
%>
							<tr> 
								<td class="iptlab"><%=trim(rs("Wrd4"))%>：</td>
							  <td class="ipt"> <input name="Zi4" type="text" id="Zi4" size="50" maxlength="80"></td>
							</tr>
<%
end if
if rs("Wrd5")<>"" then	
%>
							<tr> 
								<td class="iptlab"><%=trim(rs("Wrd5"))%>：</td>
							  <td class="ipt"> <input name="Zi5" type="text" id="Zi5" size="50" maxlength="80"></td>
							</tr>
<%
end if
if rs("Wrd6")<>"" then	
%>
							<tr> 
								<td class="iptlab"><%=trim(rs("Wrd6"))%>：</td>
							  <td class="ipt"> <input name="Zi6" type="text" id="Zi6" size="50" maxlength="80"></td>
							</tr>
<%
end if
if rs("Wrd7")<>"" then	
%>
							<tr> 
								<td class="iptlab"><%=trim(rs("Wrd7"))%>：</td>
							  <td class="ipt"> <input name="Zi7" type="text" id="Zi7" size="50" maxlength="80"></td>
							</tr>
<%
end if
rs.close
%>
							<tr> 
								<td class="iptlab">关键字：</td>
								<td class="ipt"> <input name="Key" type="text" id="Key" size="50" maxlength="80">
								<span class="red">*<br>用来查找相关视频，可输入多个关键字，中间用空格分开。不能出现&quot;&quot;'*?,.()等字符。</span></td>
							</tr>
							<tr> 
								<td class="iptlab">说明：</td>
								<td class="ipt">&nbsp;</td>
							</tr>
							<tr> 
								<td colspan="2" class="ipt"> 
								<textarea name="Content" style="display:none"></textarea>
								<iframe ID="editor" src="../editor.asp" frameborder=1 scrolling=no width="700" height="405"></iframe>
								</td>
							</tr>
							<tr> 
								<td class="iptlab">首页图片： 
								<input name="IncludePic" type="hidden" id="IncludePic" value="yes"></td>
								<td class="ipt">
								<input name="DefaultPicUrl" type="text" id="DefaultPicUrl" value="img/nopic.jpg" size="40" maxlength="120"> 
								<br>首页的图片,直接从上传图片中选择： 
								<select name="DefaultPicList" id="select" onChange="DefaultPicUrl.value=this.value;">
								<option selected value="img/nopic.jpg">不指定首页图片</option>
								</select> <input name="UploadFiles" type="hidden" id="UploadFiles">	
														</td>
							</tr>
							<tr>
							  <td class="iptlab">视频JPG预览图上传:</td>
							  <td class="ipt"><iframe class="TBGen" style="top:2px" ID="Uploaddown" src="../Progress_up.asp" frameborder=0 scrolling=no width="380" height="60"></iframe></td>
							</tr>
							<tr>
							  <td class="iptlab">视频预览图上传地址:</td>
							  <td class="ipt"><input name="Uploaddown" type="text" id="Uploaddown" size="60" readonly />
                              				  <input name="downSize" id="downSize" type="hidden" size="12" /><span class="red">*</span></td>
							</tr>
                            <tr>
							  <td class="iptlab">视频文件上传[FLV格式]:</td>
							  <td class="ipt"><iframe class="TBGen" style="top:2px" ID="UploaddownJ" src="../Progress_upJ.asp" frameborder=0 scrolling=no width="380" height="60"></iframe></td>
							</tr>
							<tr>
							  <td class="iptlab">视频文件上传地址:</td>
							  <td class="ipt"><input name="UploaddownJ" type="text" id="UploaddownJ" size="60" readonly />
							  <span class="red">*</span>文件大小:
							  <input name="downSizeJ" id="downSizeJ" type="text" size="12" />K</td>
							</tr>
                            <tr>
							  <td class="iptlab">视频规格大小:</td>
							  <td class="ipt">宽:
							    <input name="flvWidth" type="text" id="flvWidth" size="12" />
							    px 
							    高:
							    <input name="flvHight" id="flvHight" type="text" size="12" />
						      px<span class="red">（视频的长边不可超过640，短边不超过480）</span></td>
							</tr>
							<tr> 
								<td class="iptlab">已通过审核：</td>
								<td class="ipt"> <input name="Passed" type="checkbox" id="Passed" value="yes" <%call checkverify("19")%>>
								是<span class="red">（如果选中的话将直接发布）</span></td>
							</tr>
							<tr> 
								<td class="iptlab">首页显示：</td>
							  <td class="ipt"><input name="Elite" type="checkbox" id="Passed" value="yes">
								是<span class="red">（如果选中的话将在首页显示）</span> <input name="showMov" type="checkbox" id="Passed" value="yes">
								是否隐藏<span class="red">（如果选中的话前台用户将看不到这个作品）</span></td>
							</tr>
							<tr> 
								<td class="iptlab">录入时间：</td>
								<td class="ipt">
								<input name="UpdateTime" type="text" id="UpdateTime" value="<%=now()%>" maxlength="50">当前时间为：<%=now()%> 注意不要改变格式。</td>
							</tr>
							<tr><td colspan="2" class="iptsubmit"><input	name="Add" type="submit"  id="Add" value=" 添 加 " 
				onClick="document.myform.action='Bs_Article_Save.asp?action=add';document.myform.target='_self';">
				<input type="hidden" name='bigclass'>
                <input type="text" readonly="readonly" name="Author" value="<%if Session("adminName")="" then response.write"未知" end if%><%=Session("adminName")%>" />
				</td></tr>
						</table>

				
			</form>
		</td>
	</tr>
</table>
<%htmlend%>
<script language = "JavaScript">

function changeName(locationid){
	document.myform.bigclass.value=locationid
	document.myform.action="BS_Article_Add.asp";
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