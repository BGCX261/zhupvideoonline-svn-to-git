<!--#include file="Inc/syscode.asp"-->
<!--#include file="Inc/Top.asp" -->
<%
'����Ķ����������д���
ShowSmallClassType=ShowSmallClassType_Default
MaxPerPage=MaxPerPage_Default
strFileName="Product.asp?BigClassName=" & BigClassName & "&SmallClassName=" & SmallClassName 
%>
<!-- #include file="Inc/Head.asp" -->
<div id="body">
  <table class="bodytable">
    <tr>
      <td class="bodyleft"><div id="leftbg">
          <!-- #include file="L_product.asp" -->
        </div><img src="img/1.gif" /></td>
      <td class="bodyline"></td>
      <td class="bodyright"><table class="mbox">
        <tr>
          <td><div class="rtitle"><img src="img/title_ico.gif" /> �� Ƶ �� �� �� ��</div></td>
        </tr>
        <tr>
          <td class="cttd"><% call ShowAllClass() %></td>
        </tr>
      </table></td>
    </tr>
  </table>
</div>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>
