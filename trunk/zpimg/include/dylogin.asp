<!--#include file="head.asp"-->
<%If User_ID="" Then%> 
document.write("<form action='include/login.asp?act=login' method='post' name='form2'>�û���<input type='text' name='username' size='8' tabindex='1' class='dlinput' />&nbsp;&nbsp;����<input type='password' name='userpass' size='8' tabindex='2' class='dlinput' />&nbsp;&nbsp;<input type='submit' value='��¼' name='B1' tabindex='3' class='dlbutton' />&nbsp;&nbsp;<a href='javascript:openreg();'>ע���Ա</a></form>");
<%Else%>
<%If User_VIP=True Then%>
document.write("<span class='red'>vip*</span>");
<%End If%>
document.write("[<span class='blue'><%=User_ID%></span>] ");
document.write("<a href=\"javascript:memberinfo('<%=User_ID%>')\">�ҵ�����</a> | <a href='javascript:openreg();' class='gray'>�޸�</a>");
<%If User_Group="admin" Then%>
document.write(" - <a href='admin/admin_main.asp'>��վ����</a>");
<%End If%>
document.write(" - <a href='include/login.asp?act=logout'>�˳�</a>");
<%End If%>
<%Call CloseAll()%>
