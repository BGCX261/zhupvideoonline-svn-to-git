<!--#include file="../include/head.asp"-->
<!--#include file="../include/ofunctions.asp"-->

<%
'** 文件说明：网站设置
Dim Form_strRegTsEmail,Form_strSubTsEmail,Form_strReTsEmail,Form_strPostTsEmail,Form_strReSsbTsEmail
dim j,id,SetStyle
id=Request.QueryString("id")

If id="kw" Then
SetStyle="style=""display:none"""
End If

If User_Group<>"admin" Then
	Call ShowError("此页只有管理员才能访问！",1)
End If

Sub ShowObjectInstalled(strObjName)
	If IsObjInstalled(strObjName) Then
		Response.Write "<font color='#0066FF'><b>√</b></font>"
	Else
		Response.Write "<font color='red'><b>×</b>(无)</font>"
	End If
End Sub

Function IsObjInstalled(strClassString)
    On Error Resume Next
    IsObjInstalled = False
    Err = 0
    Dim xTestObj
    Set xTestObj = Server.CreateObject(strClassString)
    If 0 = Err Then IsObjInstalled = True
    Set xTestObj = Nothing
    Err = 0
End Function

dim ObjInstalled,Action
ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")
Action=trim(request("Action"))
if Action="" then
	Action="ShowConfig"
end If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<%
If Action="SaveConfig" Then
	Call SaveConfig()
Else
	Call ShowConfig()
End If

Call CloseAll()

sub ShowConfig()
%>
<form method="POST" action="?Action=SaveConfig" name="form">
<div class="DB">
    <div class="L1" <%=SetStyle%>>当前位置：<a href="admin_show.asp">管理面板</a> &gt; 网站基本设置</div>
	<table cellspacing="1" cellpadding="8" border="0" class="a23" bgcolor="#FFFFFF" align="center">

         <tr class="a10" <%=SetStyle%>> 
              <td colspan="2" class="bold red PadLeft">注意：本页设置内容不允许使用英文双引号（&quot;），若必须使用到，请用单引号替换。</td>
         </tr>

         <tr class="a4" <%=SetStyle%>>
              <td width="30%" height="32" class="PadLeft"><b>网站名称：</b></td>
              <td width="69%" height="32" class="PadLeft2"> 
                   <input name="strSiteName" type="text" id="WebmasterEmail14" value="<%=strSiteName%>" size="57" maxlength="100" class="IT5">
		      </td>
    	 </tr>

		<tr class="a4" <%=SetStyle%>>
              <td width="30%" height="32" class="PadLeft"><b>网站网址：</b></td>
              <td width="69%" height="32" class="PadLeft2"> 
                   <input name="strSiteUrl" type="text" id="WebmasterEmail16" value="<%=strSiteUrl%>" size="57" maxlength="100" class="IT5"> 必须以/结束
              </td>
    	 </tr>

		 <tr class="a4" <%=SetStyle%>>
              <td width="30%" height="32" class="PadLeft"><b>设置网站风格:</b></td>
              <td width="69%" height="32" class="PadLeft2"> 
                   <input name="strStyle" type="text" id="SiteName1" value="<%=strStyle%>" size="13" maxlength="50" class="IT2"> 默认风格default.css
		      </td>
    	 </tr>

		 <tr class="a4" <%=SetStyle%>>
              <td width="30%" height="32" class="PadLeft"><b>是否开放会员注册:</b></td>
              <td width="69%" height="32" class="PadLeft2"> 
                   <input name="Enable_Register" type="text" id="LogoUrl2" value="<%=Enable_Register%>" size="4" maxlength="255" class="IT3"> 0关闭，1开放
			  </td>
    	 </tr>

		 <tr class="a4" <%=SetStyle%>>
              <td width="30%" height="32" class="PadLeft"><b>是否开启cookie登录方式：</b></td>
              <td width="69%" height="32" class="PadLeft2"> 
                   <input name="Enable_Cookies" type="text" id="LogoUrl0" value="<%=Enable_Cookies%>" size="4" maxlength="255" class="IT3"> 0关闭，1开启
			  </td>
    	 </tr>

		 <tr class="a4" <%=SetStyle%>>
              <td width="30%" height="32" class="PadLeft"><b>回复分页数：</b></td>
              <td width="69%" height="32" class="PadLeft2"> 
                   <input name="contextPageSize" type="text" id="WebmasterEmail1" value="<%=contextPageSize%>" size="4" maxlength="100" class="IT3"> 多少个回复即开始分页
			  </td>
    	 </tr>

		 <tr class="a4" <%=SetStyle%>>
              <td width="30%" height="32" class="PadLeft"><b>列表分页数：</b></td>
              <td width="69%" height="32" class="PadLeft2"> 
                   <input name="ListPageSize" type="text" id="WebmasterEmail2" value="<%=ListPageSize%>" size="4" maxlength="100" class="IT3"> 每页列表显示多少条文章后即分页
			  </td>
    	 </tr>

		 <tr class="a4">
              <td width="30%" height="32" class="PadLeft"><b>搜索热词:<br></b>请用|号分隔词语</td>
              <td width="69%" height="32" >
			       <textarea rows="4" name="Hotkw" cols="88" class="IT4"><%=Hotkw%></textarea>
			  </td>
    	 </tr>

		 <tr class="a4" <%=SetStyle%>>
              <td width="30%" height="32" class="PadLeft"><b>不良词语过滤:<br></b>请用|号分隔不良词语</td>
              <td width="69%" height="32" >
			       <textarea rows="4" name="strBadWords" cols="88" class="IT4"><%=strBadWords%></textarea>
			  </td>
    	 </tr>

         <tr class="a4" <%=SetStyle%>>
              <td width="30%" height="32" class="PadLeft"><b>版权信息：</b><br>注意替换或去除双引号</td>
              <td width="69%" height="32" >
			       <textarea rows="4" name="strCopyright" cols="88" class="IT4"><%=strCopyright%></textarea>
			  </td>
    	 </tr>
    	
         <tr class="a4" <%=SetStyle%>> 
              <td width="30%" height="32" class="PadLeft"><b>备案信息：</b><br>注意替换或去除双引号</td>
              <td width="69%" height="32" > 
                   <textarea rows="4" name="strBeiAn" cols="88" class="IT4"><%=strBeiAn%></textarea>
			  </td>
         </tr>

         <tr class="a4" <%=SetStyle%>> 
              <td width="30%" height="32" class="PadLeft"><b>页面底部信息：</b><br>注意替换或去除双引号</td>
              <td width="69%" height="32" > 
                   <textarea rows="4" name="strfooterin" cols="88" class="IT4"><%=strfooterin%></textarea>
			  </td>
         </tr>

		 <tr class="a4" <%=SetStyle%>>
              <td width="30%" height="32" class="PadLeft"><b>网站简介：</b><br>当网站被搜索时，会显示此内容。</td>
              <td width="69%" height="32" > 
                   <textarea rows="4" name="strAbout" cols="88" class="IT4"><%=strAbout%></textarea>
		      </td>
    	 </tr>

         <tr class="a4" <%=SetStyle%>> 
              <td width="30%" height="32" class="PadLeft"><b>关键字信息：</b><br>利于网站被搜索引擎收录，逗号分隔</td>
              <td width="69%" height="32" > 
                   <textarea rows="4" name="strKeyword" cols="88" class="IT4"><%=strKeyword%></textarea>
			  </td>
         </tr>

         <tr class="a4" <%=SetStyle%>> 
              <td width="30%" height="32" class="PadLeft"><b>网站统计代码：</b><br>注意替换或去除双引号</td>
              <td width="69%" height="32"> 
                   <textarea rows="4" name="strtongji" cols="88" class="IT4"><%=strtongji%></textarea>
			  </td>
         </tr>

		 <tr class="a4">
              <td height="39" colspan="2" class="center"> 
                   <input name="Action" type="hidden" id="Action" value="SaveConfig">
                   <input name="cmdSave" type="submit" id="cmdSave"  <%If ObjInstalled=false Then response.write "value=服务器不支持FSO不能保存设置 disabled" else response.write "value=保存设置"%> class="btn">
		      </td>
         </tr>
<%
If ObjInstalled=false Then
	Response.Write "<tr><td height='50' colspan='3'>"
	Response.Write "<b><font color=red>你的服务器未开启FSO(Scripting.FileSystemObject)组件支持文本读写，不能保存设置。<br>请直接修改“include/config.asp”文件中的内容。</font>"
	Response.Write "<BR><a href='http://www.baidu.com/s?ie=gb2312&wd=FSO组件开启/关闭方法'>点此查看FSO组件开启/关闭方法</a></b></td></tr>"
End If
%>

    </table>
	</div>
<%end sub%>

</form>
</body>
</html>
<%
sub SaveConfig()
	dim tempTxt
	tempTxt= "<" & "%" & vbcrlf & vbcrlf

	tempTxt=tempTxt& "'++++++++++++++++网站基本设定+++++++++++++++" & vbcrlf & vbcrlf

	tempTxt=tempTxt&"'请仅设置=号之后的部分，并严格按照说明设置" & vbcrlf & vbcrlf

	tempTxt=tempTxt&"'++++++++++++++++++++++++++++++++++++++++++" & vbcrlf & vbcrlf
		
	tempTxt=tempTxt&"strSiteName=" & chr(34) & trim(request("strSiteName")) & chr(34) &"     '站点名称" & vbcrlf & vbcrlf
	
	tempTxt=tempTxt&"strSiteUrl=" & chr(34) & trim(request("strSiteUrl")) & chr(34) & "    '站点地址 必须以/结束" &vbcrlf & vbcrlf

	tempTxt=tempTxt&"strStyle=" & chr(34) & trim(request("strStyle")) & chr(34) & "        '设置默认风格" & vbcrlf & vbcrlf

	tempTxt=tempTxt&"Enable_Register=" & trim(request("Enable_Register")) & "        '设置是否允许注册，0关闭，1开启" & vbcrlf & vbcrlf
	
	tempTxt=tempTxt&"Enable_Cookies=" & trim(request("Enable_Cookies")) & "        '设置允许cookie方式保存登录信息，0关闭，1开启" & vbcrlf & vbcrlf
	
	tempTxt=tempTxt&"contextPageSize=" & trim(request("contextPageSize")) & "        '回复分页数，多少个回复即开始分页" & vbcrlf & vbcrlf
	
	tempTxt=tempTxt&"ListPageSize=" & trim(request("ListPageSize")) & "        '列表分页数，文章列表多少条即分页" & vbcrlf & vbcrlf

	tempTxt=tempTxt&"Hotkw=" & chr(34) & trim(request("Hotkw")) & chr(34) & "        '搜索热词，用|号分开" & vbcrlf & vbcrlf& vbcrlf & vbcrlf

	tempTxt=tempTxt&"strBadWords=" & chr(34) & trim(request("strBadWords")) & chr(34) & "        '不良词语过滤，用|号分开" & vbcrlf & vbcrlf& vbcrlf & vbcrlf

	tempTxt=tempTxt&"'+++++++++++++++++++以下设置版权显示等信息++++++++++++++++" & vbcrlf & vbcrlf
	
	tempTxt=tempTxt&"'--------以下一行设定版权信息，显示在页尾，不要使用英文双引号（“）" & vbcrlf
	tempTxt=tempTxt&"strCopyright=" & chr(34) & trim(Replace(Replace(Replace(Replace(request("strCopyright"),chr(13),""),chr(10),""),chr(9),""),chr(34),"")) & chr(34) & vbcrlf & vbcrlf
	
	tempTxt=tempTxt&"'--------以下一行设定版权声明，不要使用英文双引号（“）" & vbcrlf
	tempTxt=tempTxt&"strfooterin=" & chr(34) & trim(Replace(Replace(Replace(Replace(request("strfooterin"),chr(13),""),chr(10),""),chr(9),""),chr(34),"")) & chr(34) & vbcrlf & vbcrlf

	tempTxt=tempTxt&"'--------以下一行设定备案信息，不要使用英文双引号（“）" & vbcrlf
	tempTxt=tempTxt&"strBeiAn=" & chr(34) & trim(Replace(Replace(Replace(Replace(request("strBeiAn"),chr(13),""),chr(10),""),chr(9),""),chr(34),"")) & chr(34) & vbcrlf & vbcrlf
	
	tempTxt=tempTxt&"'--------以下一行设定网站统计代码，不要使用英文双引号（“）" & vbcrlf
	tempTxt=tempTxt&"strtongji=" & chr(34) & trim(Replace(Replace(Replace(Replace(request("strtongji"),chr(13),""),chr(10),""),chr(9),""),chr(34),"")) & chr(34) & vbcrlf & vbcrlf
	
	tempTxt=tempTxt&"'--------以下一行设定网站简要，不要使用英文双引号（“）" & vbcrlf
	tempTxt=tempTxt&"strAbout=" & chr(34) & trim(request("strAbout")) & chr(34) &"     " & vbcrlf & vbcrlf

	tempTxt=tempTxt&"'--------以下一行设定关键字信息，利于站长被搜索引擎收录" & vbcrlf
	tempTxt=tempTxt&"strKeyword=" & chr(34) & trim(Replace(Replace(Replace(Replace(request("strKeyword"),chr(13),""),chr(10),""),chr(9),""),chr(34),"")) & chr(34) & vbcrlf & vbcrlf

	tempTxt=tempTxt&"%" & ">"

	WriteToTextFile "../include/config.asp",tempTxt,"GB2312"

    '生成搜索热词hotkw.txt文件
	DIM str,i,tmpStr
	str=""&trim(request("Hotkw"))&""
	str=split(str,"|")
	For i=0 To ubound(str)
	tmpStr=tmpStr&"<a href=""search.asp?kw="&str(i)&""">"&str(i)&"</a>"
	Next
	Call Save("../txt/hotkw.txt",tmpStr)

	Call ShowError("设置成功！",1)
    Response.end
end sub
%>