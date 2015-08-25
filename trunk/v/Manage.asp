<!--#include file="Inc/syscode.asp" -->
<!--#include file="inc/md5.asp"-->
<%
dim Action,UserName
dim rsUser,sqlUser
Action=trim(request("Action"))
UserName=trim(request("UserName"))
if UserName="" then
	UserName=session("UserName")
end if
if  UserName="" then
	if Action="" then
		response.redirect "Server.asp"
	else
		FoundErr=True
		ErrMsg=ErrMsg & "<br>参数不足！"
	end if
end if
if FoundErr=true then
	call WriteErrMsg()
else
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from [User] where UserName='" & UserName & "'"
	rsUser.Open sqlUser,conn,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br>找不到指定的用户！"
		call writeErrMsg()
	else
		if Action="Modify" then
			dim Sex,Email,Homepage,MSN
			Sex=trim(Request("Sex"))
			Email=trim(request("Email"))
			Homepage=trim(request("Homepage"))
			Comane=trim(request("Comane"))
			Add=trim(request("Add"))
			Somane=trim(request("Somane"))
			Zip=trim(request("Zip"))
			Phone=trim(request("Phone"))
			Fox=trim(request("Fox"))
			if Sex="" then
				founderr=true
				errmsg=errmsg & "<br>性别不能为空"
			else
				sex=cint(sex)
				if Sex<>0 and Sex<>1 then
					Sex=1
				end if
			end if
			if Email="" then
				founderr=true
				errmsg=errmsg & "<br>Email不能为空"
			else
				if IsValidEmail(Email)=false then
					errmsg=errmsg & "<br>您的Email有错误"
			   		founderr=true
				end if
			end if
			if Comane="" then
				founderr=true
				errmsg=errmsg & "<br>真实名称不能为空"
			end if
			if Add="" then
				founderr=true
				errmsg=errmsg & "<br>居住地址不能为空"
			end if
			if Somane="" then
				founderr=true
				errmsg=errmsg & "<br>联系人不能为空"
			end if
			if Phone="" then
				founderr=true
				errmsg=errmsg & "<br>联系电话不能为空"
			end if
			if FoundErr<>true then
				rsUser("Sex")=Sex
				rsUser("Email")=Email
				rsUser("HomePage")=HomePage
				rsUser("Comane")=Comane
				rsUser("Add")=Add
				rsUser("Somane")=Somane
				rsUser("Zip")=Zip
				rsUser("Phone")=Phone
				rsUser("Fox")=Fox
				rsUser.update
				response.write"<SCRIPT language=JavaScript>alert('会员资料修改成功！');"
                response.write"javascript:history.go(-1)</SCRIPT>"
				'call WriteSuccessMsg("会员资料修改成功！")
			else
				call WriteErrMsg()
			end if
		else

%>
<!-- #include file="Inc/Head.asp" -->
<div id="body">
  <table  class="bodyTable">
    <tr>
      <td class="bodyLeft"><!-- 会员中心 -->
        <!-- #include file="L_vip.asp" -->
      </td>
      <td class="bodyLine"></td>
      <td class="bodyRight">
          <!-- 修改会员资料 -->
          <div class="navPath">修改资料</div>
          <div>
            <form action="Manage.asp" method="post" name="Form1" id="Form1">
             <table width="500" border="0" align="center" cellpadding="2" cellspacing="2" class='border'>
                  <tr align="center" class='title'>
                    <td height="30" colspan="2"><font color="#FF6600" size="2" class="en"><b>修改注册用户信息</b></font></td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="120" align="right"><b>用 户 名：</b></td>
                    <td><input name="UserName" value="<%=rsUser("UserName")%>" size="30"   maxlength="14" /></td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="120" align="right"><strong>性别：</strong></td>
                    <td><input type="radio" value="1" name="sex" <%if rsUser("Sex")=1 then response.write "CHECKED"%> />
                      男 &nbsp;&nbsp;
                      <input type="radio" value="0" name="sex" <%if rsUser("Sex")=0 then response.write "CHECKED"%> />
                      女</td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="120" align="right"><strong>Email地址：</strong></td>
                    <td><input name="Email" value="<%=rsUser("Email")%>" size="30"   maxlength="50" /></td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="120" align="right"><strong>有效证件号：</strong></td>
                    <td><input   maxlength="100" size="30" name="homepage" value="<%=rsUser("HomePage")%>" /></td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="120" align="right"><strong>真实姓名：</strong></td>
                    <td><input name="Comane" value="<%=rsUser("Comane")%>" size="30" maxlength="20" /></td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="120" align="right"><strong>居住地址：</strong></td>
                    <td><input name="Add" value="<%=rsUser("Add")%>" size="30" maxlength="50" /></td>
                  </tr>
                 <tr class="tdbg" >
                    <td align="right"><strong>联系人：</strong></td>
                    <td><input name="Somane" value="<%=rsUser("Somane")%>" size="30" maxlength="50" /></td>
                  </tr>
                  <tr class="tdbg" >
                    <td align="right"><strong>邮政编码：</strong></td>
                    <td><input name="Zip" value="<%=rsUser("Zip")%>" size="30" maxlength="50" /></td>
                  </tr>
                  <tr class="tdbg" >
                    <td align="right"><strong>联系电话：</strong><br /></td>
                    <td><input name="Phone" value="<%=rsUser("Phone")%>" size="30" maxlength="50" /></td>
                  </tr>
                  <tr class="tdbg" >
                    <td align="right"><strong>传 真：</strong></td>
                    <td><input name="Fox" value="<%=rsUser("Fox")%>" size="30" maxlength="50" /></td>
                  </tr>
                  <tr align="center" class="tdbg" >
                    <td height="40" colspan="2"><input name="Action" type="hidden" id="Action" value="Modify" />
                      <input name="Submit" type="submit" id="Submit" value="保存修改结果" /></td>
                  </tr>
                </table>
            </form>
          </div>
       </td>
    </tr>
  </table>
</div>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>
<%
end if
end if
rsUser.close
set rsUser=nothing
end if
call CloseConn()
%>