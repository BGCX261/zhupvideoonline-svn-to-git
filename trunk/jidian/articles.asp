<!--#include file="Inc/syscode.asp"-->
<!--#include file="Inc/Top.asp" -->
<%
'请勿改动下面这三行代码
ShowSmallClassType=ShowSmallClassType_Default
MaxPerPage=MaxPerPage_Default
strFileName="article.asp?BigClassName=" & BigClassName & "&SmallClassName=" & SmallClassName 
%>
<!-- #include file="Inc/Head.asp" -->
<div id="body">
  <table class="bodytable">
    <tr>
      <td class="bodyleft"><!-- #include file="L_article.asp" --></td>
      <td class="bodyline"></td>
      <td class="bodyright"><table class="mbox">
        <tr>
          <td><div class="positiontitle"><img src="img/title_ico.gif" /> 网络课程</div></td>
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
