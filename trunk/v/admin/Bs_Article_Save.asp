<!--#include file="bsconfig.asp"-->
<!--#include file="../Inc/Config.asp"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->

<%
dim rs,sql,ErrMsg,FoundErr
dim ArticleID,Product_Id,BigClassName,SmallClassName,Title,Zi1,Zi2,Zi3,Zi4,Zi5,Zi6,Zi7,Content,key,UpdateTime,Hits
dim IncludePic,DefaultPicUrl,UploadFiles,Elite,showMov,Passed,arrUploadFiles,Uploaddown,downsize,UploaddownJ,downsizeJ,Author,flvWidth,flvHight,faculty
dim ObjInstalled

ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")
FoundErr=false
ArticleID=Trim(Request.Form("ArticleID"))
Product_Id=trim(request.form("Product_Id"))
BigClassName=trim(request.form("BigClassName"))
SmallClassName=trim(request.form("SmallClassName"))
Title=trim(request.form("Title"))
faculty=trim(request.form("faculty"))
Zi1=trim(request.form("Zi1"))
Zi2=trim(request.form("Zi2"))
Zi3=trim(request.form("Zi3"))
Zi4=trim(request.form("Zi4"))
Zi5=trim(request.form("Zi5"))
Zi6=trim(request.form("Zi6"))
Zi7=trim(request.form("Zi7"))
Key=trim(request.form("Key"))
Content=trim(request.form("Content"))
UpdateTime=trim(request.form("UpdateTime"))
IncludePic=trim(request.form("IncludePic"))
DefaultPicUrl=trim(request.form("DefaultPicUrl"))
UploadFiles=trim(request.form("UploadFiles"))
Uploaddown=trim(request.form("Uploaddown"))
downsize=trim(request.form("downsize"))
flvWidth=trim(request.form("flvWidth"))
flvHight=trim(request.form("flvHight"))

UploaddownJ=trim(request.form("UploaddownJ"))
downsizeJ=trim(request.form("downsizeJ"))

Passed=trim(request.form("Passed"))
Elite=trim(request.form("Elite"))
showMov=trim(request.form("showMov"))
Hits=trim(request.form("Hits"))
Author=trim(request.form("Author"))

sqlBig="select * from BigClass where BigClassName='" & BigClassName & "'"
Set rsBig= Server.CreateObject("ADODB.Recordset")
rsBig.open sqlBig,conn,1,1
BigClassName=rsBig("BigClassName")
rsBig.close

sqlSmall="select * from SmallClass where SmallClassName='" & SmallClassName & "' and  BigClassName='" & BigClassName & "'"
Set rsSmall= Server.CreateObject("ADODB.Recordset")
rsSmall.open sqlSmall,conn,1,1
SmallClassName=rsSmall("SmallClassName")
rsSmall.close

if BigClassName="" then
	founderr=true
	errmsg=errmsg+"<li>未指定所属大类</li>"
end if
if Title="" then
	founderr=true
	errmsg="<li>标题不能为空</li>"
end if
if Key="" then
	founderr=true
	errmsg=errmsg+"<li>请输入关键字</li>"
end if
if Content="" then
	founderr=true
	errmsg=errmsg+"<li>内容不能为空</li>"
end if

if founderr=false then
	Title=dvhtmlencode(Title)
	Key=replace(replace(replace(replace(replace(replace(replace(Key,"'",""),"*",""),"?",""),"(",""),")",""),",",""),".","")
	Key=Key & " "
	Content=ubbcode(Content)
	if UpdateTime<>"" and IsDate(UpdateTime)=true then
		UpdateTime=CDate(UpdateTime)
	else
		UpdateTime=now()
	end if
	if Hits<>"" then
		Hits=CLng(Hits)
	else
		Hits=0
	end if
	set rs=server.createobject("adodb.recordset")
	if request("action")="add" then
		sql="select top 1 * from Product" 
		rs.open sql,conn,1,3
		rs.addnew
		call SaveData()
		'rs("Editor")=Editor
		rs.update
		ArticleID=rs("ArticleID")
		rs.close
		set rs=nothing
	elseif request("action")="Modify" then
  		if ArticleID<>"" then
			sql="select * from Product where articleid=" & ArticleID
			rs.open sql,conn,1,3
			if not (rs.bof and rs.eof) then
				call SaveData()
				rs.update
				rs.close
				set rs=nothing
 			else
				founderr=true
				errmsg=errmsg+"<li>找不到，可能已经被其他人删除。</li>"
				call WriteErrMsg()
			end if
		else
			founderr=true
			errmsg=errmsg+"<li>不能确定ArticleID的值</li>"
			call WriteErrMsg()
		end if
	else
		founderr=true
		errmsg=errmsg+"<li>没有选定参数</li>"
		call WriteErrMsg()
	end if

	call CloseConn()
%>
<!-- #include file="Inc/Head.asp" -->
<table align="center" class="a2">
	<tr>
		<td class="a1">添加 修改 视频信息</td>
	</tr>
	<tr class="a4">
		<td>
      <table width="100%">
        <tr> 
          <td  colspan="2" class="iptsubmit"><b> 
            <%if request("action")="add" then%>
            添加 
            <%else%>
            修改 
            <%end if%>
            成功</b></td>
        </tr>
        <tr> 
          <td class="iptlab">序号：</td>
          <td class="ipt"><%=ArticleID%></td>
        </tr>
        <tr> 
          <td class="iptlab">编号：</td>
          <td class="ipt"><%=Product_Id%></td>
        </tr>
        <tr> 
          <td class="iptlab">名称：</td>
          <td class="ipt"><%=Title%></td>
        </tr>
        <tr> 
          <td class="iptlab">所属类别：</td>
          <td class="ipt"> 
<%
response.write BigClassName
if SmallClassName<>"" then response.write " &gt;&gt; " & SmallClassName
if SpecialName<>"" then response.write "<br>所属专题：" & SpecialName
%>
          </td>
        </tr>
        <tr> 
          <td class="iptlab">关  键 字：</td>
          <td class="ipt"><%=Key%></td>
        </tr>
        <tr> 
          <td colspan="2" class="iptsubmit"> 【<a href="Bs_Article_edit.asp?ArticleID=<%=ArticleID%>">修改</a>】&nbsp;【<a href="Bs_Article_Add.asp">继续添加</a>】&nbsp;【<a href="Bs_Article.asp">管理</a>】&nbsp;【<a href="../ArticleShow.asp?ArticleID=<%=ArticleID%>" target="_blank">预览内容</a>】</td>
        </tr>
      </table>
		</td>
	</tr>
</table>
<%htmlend%>
<%
else
	WriteErrMsg
end if

sub SaveData()
	rs("Product_Id")=Product_Id
	rs("BigClassName")=BigClassName
	rs("SmallClassName")=SmallClassName
	'rs("SpecialName")=SpecialName
	rs("Title")=Title
	rs("Zi1")=Zi1
	rs("Zi2")=Zi2
	rs("Zi3")=Zi3
	rs("Zi4")=Zi4
	rs("Zi5")=Zi5
	rs("Zi6")=Zi6
	rs("Zi7")=Zi7
	rs("Content")=Content
	rs("Key")=Key
	
	rs("flvWidth")=flvWidth
	rs("flvHight")=flvHight
	rs("faculty")=faculty
	rs("Hits")=Hits
	rs("Author")=Author
	'rs("CopyFrom")=CopyFrom
	if IncludePic="yes" then
		rs("IncludePic")=True
	else
		rs("IncludePic")=False
	end if
	if Passed="yes" then
		rs("Passed")=True
	else
		if EnableArticleCheck="No" then
			rs("Passed")=True
		else
			rs("Passed")=False
		end if
	end if
	'if OnTop="yes" then
		'rs("OnTop")=True
	'else
		'rs("OnTop")=False
	'end if
	'if Hot="yes" then
		'rs("Hot")=True
	'else
		'rs("Hot")=False
	'end if
	if Elite="yes" then
		rs("Elite")=True
	else
		rs("Elite")=False
	end if
	if showMov="yes" then
		rs("showMov")=True
	else
		rs("showMov")=False
	end if
	rs("UpdateTime")=UpdateTime

	'***************************************
	'删除无用的上传文件
	if ObjInstalled=True and UploadFiles<>"" then
		dim fso,strRubbishFile
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		if instr(UploadFiles,"|")>1 then
			dim arrUploadFiles,intTemp
			arrUploadFiles=split(UploadFiles,"|")
			UploadFiles=""
			for intTemp=0 to ubound(arrUploadFiles)
				if instr(Content,arrUploadFiles(intTemp))<=0 and arrUploadFiles(intTemp)<>DefaultPicUrl then
					strRubbishFile=server.MapPath("../" & arrUploadFiles(intTemp))
					if fso.FileExists(strRubbishFile) then
						fso.DeleteFile(strRubbishFile)
						response.write "<br><li>" & arrUploadFiles(intTemp) & "在文章中没有用到，也没有被设为首页图片，所以已经被删除！</li>"
					end if
				else
					if intTemp=0 then
						UploadFiles=arrUploadFiles(intTemp)
					else
						UploadFiles=UploadFiles & "|" & arrUploadFiles(intTemp)
					end if
				end if
			next
		else
			if instr(Content,UploadFiles)<=0 and UploadFiles<>DefaultPicUrl then
				strRubbishFile=server.MapPath("../" & UploadFiles)
				if fso.FileExists(strRubbishFile) then
					fso.DeleteFile(strRubbishFile)
					response.write "<br><li>" & UploadFiles & "在文章中没有用到，也没有被设为首页图片，所以已经被删除！</li>"
				end if
				UploadFiles=""
			end if
		end if
		set fso=nothing
	end If
	'结束
	'***************************************
	rs("DefaultPicUrl")=DefaultPicUrl
	rs("UploadFiles")=UploadFiles
	
	'***************************************
	'删除无用的声音下载文件
	if ObjInstalled=True and Uploaddown<>"" then
		'dim fso,strRubbishFile
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		if instr(Uploaddown,"|")>1 then
			'dim arrUploadFiles,intTemp
			arrUploadFiles=split(Uploaddown,"|")
			Uploaddown=""
			for intTemp=0 to ubound(arrUploadFiles)
				if instr(Uploaddown,arrUploadFiles(intTemp))<=0 then
					strRubbishFile=server.MapPath("../" & arrUploadFiles(intTemp))
					if fso.FileExists(strRubbishFile) then
						fso.DeleteFile(strRubbishFile)
						response.write "<br><li>" & arrUploadFiles(intTemp) & "在文章中没有用到,所以已经被删除!</li>"
					end if
				else
					if intTemp=0 then
						Uploaddown=arrUploadFiles(intTemp)
					else
						Uploaddown=Uploaddown & "|" & arrUploadFiles(intTemp)
					end if
				end if
			next
		else
			if instr(Uploaddown,Uploaddown)<=0 then
				strRubbishFile=server.MapPath("../" & Uploaddown)
				if fso.FileExists(strRubbishFile) then
					fso.DeleteFile(strRubbishFile)
					response.write "<br><li>" & Uploaddown & "在文章中没有用到，所以已经被删除！</li>"
				end if
				Uploaddown=""
			end if
		end if
		set fso=nothing
	end If
	'结束
	'***************************************
	rs("Uploaddown")=Uploaddown
	rs("downsize")=downsize
	
	rs("UploaddownJ")=UploaddownJ
	rs("downsizeJ")=downsizeJ
	

	
	
end sub




	
%>