<!--#include file="Inc/syscode.asp" -->
<!-- #include file="Inc/Head.asp" -->
<div id="body">
  <table  class="bodyTable">
    <tr>
      <td class="bodyLeft"><!-- #include file="L_fellow.asp" --></td>
      <td class="bodyLine"></td>
      <td class="bodyRight">
          <div class="rtitle"><img src="img/title_ico.gif" /> 评 论 用 户 注 册</div>
          <form action='UserRegPost.asp' method='post' name='UserReg' id="UserReg">
                <table width="95%" border="0" align="center" cellpadding="4">
                  <tr>
                    <td colspan="2">您现在注册的是评论用户，用于对视频作品发表您的意见。如需注册上传视频用户请联系管理员添加。</td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="37%"><b>用户名：</b><br />
                      不能超过50个字符（25个汉字）</td>
                    <td width="63%"><input maxlength="50" size="30" name="UserName" />
                      <font color="#FF0000">*</font></td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="37%"><strong>真实姓名：</strong><br />
                      您的真实姓名</td>
                    <td width="63%"><input name="Comane" id="Comane" size="30"   maxlength="50" /> 
                      <font color="#FF0000">*</font>
                    </td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="37%"><b>密码(至少6位)：</b><br />
                      请输入密码，区分大小写。 不要使用类似 '*'、' '的特殊字符</td>
                    <td width="63%"><input type="password" maxlength="12" size="30" name="Password" />
                      <font color="#FF0000">*</font></td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="37%"><strong>确认密码(至少6位)：</strong><br />
                      请再输一遍确认</td>
                    <td width="63%"><input type="password" maxlength="12" size="30" name="PwdConfirm" />
                      <font color="#FF0000">*</font></td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="37%"><strong>密码问题：</strong><br />
                      忘记密码的提示问题</td>
                    <td width="63%"><input   type="text" maxlength="50" size="30" name="Question" />
                      <font color="#FF0000">*</font></td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="37%"><strong>问题答案：</strong><br />
                      忘记密码的提示问题答案，用于取回密码</td>
                    <td width="63%"><input type="text" maxlength="50" size="30" name="Answer" />
                      <font color="#FF0000">*</font></td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="37%"><strong>性别：</strong><br />
                      请选择您的性别</td>
                    <td width="63%"><input type="radio" checked="checked" value="1" name="sex" />
                      男 &nbsp;&nbsp;
                      <input type="radio" value="0" name="sex" />
                      女</td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="37%"><strong>Email地址：</strong><br />
                      请输入有效的邮件地址</td>
                    <td width="63%"><input maxlength="50" size="30" name="Email" />
                      <font color="#FF0000">*</font></td>
                  </tr>
                  <tr class="tdbg" >
                    <td><strong>入学年份：</strong><br />
                    只填4位年份数字即可，如：2010</td>
                    <td width="63%"><input name="homepage" id="homepage" size="30"   maxlength="50" />
                    <font color="#FF0000">*</font></td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="37%"><p><strong>专业班级：<br />
                      </strong>如：10机电1班</p></td>
                    <td width="63%"><input name="Add" id="Add" size="30" maxlength="100" />
                      <font color="#FF0000">*</font></td>
                  </tr>
                  <tr class="tdbg" >
                    <td><strong>QQ：</strong></td>
                    <td width="63%"><input name="Zip" id="Zip" size="30" maxlength="20" /></td>
                  </tr>
                  <tr class="tdbg" >
                    <td><strong>联系电话：<br />
                      </strong>能联系到你的电话</td>
                    <td width="63%"><input name="Phone" id="Phone"  size="30" maxlength="20" />
                      <font color="#FF0000">*</font></td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="37%"><strong>个人简介：</strong></td>
                    <td width="63%"><input name="Fox" id="Fox" size="30" /></td>
                  </tr>
                  <tr align="center">
                    <td colspan="2"><input type="submit" value=" 注 册 " name="Submit" />
                  <input name="Reset" type="reset" id="Reset" value=" 清 除 " /></td>
                  </tr>
          </table>
                  
        </form>
        </td>
    </tr>
  </table>
</div>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>
