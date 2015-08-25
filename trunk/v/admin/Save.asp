<!--#include file="bsconfig.asp"-->
<!--#include file="../Inc/Config.asp"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->

<%
dim rs,sql,ErrMsg,FoundErr
dim ArticleID,Product_Id,BigClassName,SmallClassName,Title,Zi1,Zi2,Zi3,Content,key,UpdateTime,Hits
dim IncludePic,DefaultPicUrl,UploadFiles,Elite,Passed,arrUploadFiles
dim ObjInstalled

ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")
FoundErr=false

	if request("action")="Modify" then
			sql="update Product set Content='UploadFiles/'+Content"
			conn.execute sql
			rs.open sql,conn,1,3
			if not (rs.bof and rs.eof) then
				rs.update
				rs.close
				set rs=nothing
 			else
				founderr=true
				errmsg=errmsg+"<li>找不到此文章，可能已经被其他人删除。</li>"
				call WriteErrMsg()
			end if
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
            视频信息成功</b></td>
        </tr>
        <tr> 
          <td class="iptlab">产品序号：</td>
          <td class="ipt"><%=DefaultPicUrl%></td>
        </tr>
        <tr> 
          <td class="iptlab">产品编号：</td>
          <td class="ipt"><%=UploadFiles%></td>
        </tr>
        <tr> 
          <td class="iptlab">产品名称：</td>
          <td class="ipt"><%=Content%></td>
        </tr>
      </table>
		</td>
	</tr>
</table>
<%htmlend%>