<!--#include file="Inc/syscode.asp" -->
<!-- #include file="Inc/Head.asp" -->

<div id="body">
  <table class="bodyTable">
    <tr>
      <td class="bodyLeft">
      <div id="leftbg">
        <!-- #include file="L_vip.asp" -->
      </div>
      </td>
      <td class="bodyLine"></td>
      <td class="bodyRight">
          <!-- 会员中心首页 -->
            <div class="navPath">会员中心</div>
            <div><%If Session("UserName")="" Then %>
              <br />
              <br />
              <table width="100%" border="0" cellpadding="0" cellspacing="3">
                <tr>
                  <td width="19%" align="right"></td>
                  <td width="81%"><p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#006699">对不起，您还没有登录，无法进入会员中心，请先登录！</font></p>
                    <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#006699">如果您还不是我们的会员，请立即<a href="UserReg.asp">注册</a>。</font></p></td>
                </tr>
                <tr>
                  <td width="19%" align="right"></td>
                  <td width="81%"><div align="center">
                      <table class="main" cellspacing="0" cellpadding="2" width="100%" border="0" height="1">
                        <tbody>
                          <tr>
                            <td width="100%" height="1"><% call ShowUserLogina() %></td>
                          </tr>
                        </tbody>
                      </table>
                    </div></td>
                </tr>
                <tr>
                  <td width="19%" align="right"></td>
                  <td width="81%"><font color="#990000"><b><font color="#006699">会员中心有以下功能：</font></b></font></td>
                </tr>
                <tr>
                  <td width="19%" align="right"><font color="#666666">1、</font></td>
                  <td width="81%"><font color="#666666">修改您的会员注册资料</font></td>
                </tr>
                <tr>
                  <td width="19%" align="right"><font color="#666666">2、</font></td>
                  <td width="81%"><font color="#666666">您个人留言信息中心，您可在此查看你的留言信息的审核状态</font></td>
                </tr>
              </table>
              <%else%>
              <table width="85%" height="136" align="center" cellpadding="0" cellspacing="0">
                <tr valign="top" >
                  <td  width="100%" height="18">
                    <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"> <b><font color="#F584B1">会员中心首页：</font></b> <br />
                      <br />
                      <br />
                    </p>
                    <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><%=Session("UserName")%> 您好：<br />
                      <br />
                    </p>
                    <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">您现在已经进入会员服务中心，这里只有注册会员才能访问。您可在这里修改您的注册信
                      <br />
                      息、查看您的留言的审核状态。
                  </p></td>
                </tr>
                <tr valign="top" >
                  <td  width="100%" height="15"><a href="NetBook.asp">查看自己评论的审核情况>>></a>
                    <p>&nbsp;</p>
                    <p>　</p></td>
                </tr>
              </table>
              <%end if%>
          </div>
      </td>
    </tr>
  </table>
</div>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>
