<!--#include file="Inc/syscode.asp" -->
<!-- #include file="Inc/Head.asp" -->
<div id="body">
  <table  class="bodyTable">
    <tr>
      <td class="bodyLeft"><!-- #include file="L_fellow.asp" --></td>
      <td class="bodyLine"></td>
      <td class="bodyRight">
          <div class="rtitle"><img src="img/title_ico.gif" /> �� �� �� �� ע ��</div>
          <form action='UserRegPost.asp' method='post' name='UserReg' id="UserReg">
                <table width="95%" border="0" align="center" cellpadding="4">
                  <tr>
                    <td colspan="2">������ע����������û������ڶ���Ƶ��Ʒ�����������������ע���ϴ���Ƶ�û�����ϵ����Ա��ӡ�</td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="37%"><b>�û�����</b><br />
                      ���ܳ���50���ַ���25�����֣�</td>
                    <td width="63%"><input maxlength="50" size="30" name="UserName" />
                      <font color="#FF0000">*</font></td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="37%"><strong>��ʵ������</strong><br />
                      ������ʵ����</td>
                    <td width="63%"><input name="Comane" id="Comane" size="30"   maxlength="50" /> 
                      <font color="#FF0000">*</font>
                    </td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="37%"><b>����(����6λ)��</b><br />
                      ���������룬���ִ�Сд�� ��Ҫʹ������ '*'��' '�������ַ�</td>
                    <td width="63%"><input type="password" maxlength="12" size="30" name="Password" />
                      <font color="#FF0000">*</font></td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="37%"><strong>ȷ������(����6λ)��</strong><br />
                      ������һ��ȷ��</td>
                    <td width="63%"><input type="password" maxlength="12" size="30" name="PwdConfirm" />
                      <font color="#FF0000">*</font></td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="37%"><strong>�������⣺</strong><br />
                      �����������ʾ����</td>
                    <td width="63%"><input   type="text" maxlength="50" size="30" name="Question" />
                      <font color="#FF0000">*</font></td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="37%"><strong>����𰸣�</strong><br />
                      �����������ʾ����𰸣�����ȡ������</td>
                    <td width="63%"><input type="text" maxlength="50" size="30" name="Answer" />
                      <font color="#FF0000">*</font></td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="37%"><strong>�Ա�</strong><br />
                      ��ѡ�������Ա�</td>
                    <td width="63%"><input type="radio" checked="checked" value="1" name="sex" />
                      �� &nbsp;&nbsp;
                      <input type="radio" value="0" name="sex" />
                      Ů</td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="37%"><strong>Email��ַ��</strong><br />
                      ��������Ч���ʼ���ַ</td>
                    <td width="63%"><input maxlength="50" size="30" name="Email" />
                      <font color="#FF0000">*</font></td>
                  </tr>
                  <tr class="tdbg" >
                    <td><strong>��ѧ��ݣ�</strong><br />
                    ֻ��4λ������ּ��ɣ��磺2010</td>
                    <td width="63%"><input name="homepage" id="homepage" size="30"   maxlength="50" />
                    <font color="#FF0000">*</font></td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="37%"><p><strong>רҵ�༶��<br />
                      </strong>�磺10����1��</p></td>
                    <td width="63%"><input name="Add" id="Add" size="30" maxlength="100" />
                      <font color="#FF0000">*</font></td>
                  </tr>
                  <tr class="tdbg" >
                    <td><strong>QQ��</strong></td>
                    <td width="63%"><input name="Zip" id="Zip" size="30" maxlength="20" /></td>
                  </tr>
                  <tr class="tdbg" >
                    <td><strong>��ϵ�绰��<br />
                      </strong>����ϵ����ĵ绰</td>
                    <td width="63%"><input name="Phone" id="Phone"  size="30" maxlength="20" />
                      <font color="#FF0000">*</font></td>
                  </tr>
                  <tr class="tdbg" >
                    <td width="37%"><strong>���˼�飺</strong></td>
                    <td width="63%"><input name="Fox" id="Fox" size="30" /></td>
                  </tr>
                  <tr align="center">
                    <td colspan="2"><input type="submit" value=" ע �� " name="Submit" />
                  <input name="Reset" type="reset" id="Reset" value=" �� �� " /></td>
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
