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
          <!-- ��Ա������ҳ -->
            <div class="navPath">��Ա����</div>
            <div><%If Session("UserName")="" Then %>
              <br />
              <br />
              <table width="100%" border="0" cellpadding="0" cellspacing="3">
                <tr>
                  <td width="19%" align="right"></td>
                  <td width="81%"><p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#006699">�Բ�������û�е�¼���޷������Ա���ģ����ȵ�¼��</font></p>
                    <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#006699">��������������ǵĻ�Ա��������<a href="UserReg.asp">ע��</a>��</font></p></td>
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
                  <td width="81%"><font color="#990000"><b><font color="#006699">��Ա���������¹��ܣ�</font></b></font></td>
                </tr>
                <tr>
                  <td width="19%" align="right"><font color="#666666">1��</font></td>
                  <td width="81%"><font color="#666666">�޸����Ļ�Աע������</font></td>
                </tr>
                <tr>
                  <td width="19%" align="right"><font color="#666666">2��</font></td>
                  <td width="81%"><font color="#666666">������������Ϣ���ģ������ڴ˲鿴���������Ϣ�����״̬</font></td>
                </tr>
              </table>
              <%else%>
              <table width="85%" height="136" align="center" cellpadding="0" cellspacing="0">
                <tr valign="top" >
                  <td  width="100%" height="18">
                    <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"> <b><font color="#F584B1">��Ա������ҳ��</font></b> <br />
                      <br />
                      <br />
                    </p>
                    <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><%=Session("UserName")%> ���ã�<br />
                      <br />
                    </p>
                    <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">�������Ѿ������Ա�������ģ�����ֻ��ע���Ա���ܷ��ʡ������������޸�����ע����
                      <br />
                      Ϣ���鿴�������Ե����״̬��
                  </p></td>
                </tr>
                <tr valign="top" >
                  <td  width="100%" height="15"><a href="NetBook.asp">�鿴�Լ����۵�������>>></a>
                    <p>&nbsp;</p>
                    <p>��</p></td>
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
