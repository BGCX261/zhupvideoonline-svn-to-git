<!--#include file="Inc/syscode.asp"-->
<!--#include file="Inc/Top.asp" -->
<%
'����Ķ����������д���
ShowSmallClassType=ShowSmallClassType_Default
MaxPerPage=MaxPerPage_Default
strFileName="Product.asp?BigClassName=" & Server.URLEncode(BigClassName) & "&SmallClassName=" & Server.URLEncode(SmallClassName) 
%>
<!-- #include file="Inc/Head.asp" -->
<div id="body">
  <table class="bodytable">
    <tr>
      <td class="bodyleft"><div id="leftbg">
          <!-- #include file="L_product.asp" -->
        </div></td>
      <td class="bodyline"></td>
      <td class="bodyright"><table class="mbox">
          <tr>
            <td><div class="rtitle"><img src="img/title_ico.gif" /> �� Ƶ �� Ʒ չ ʾ</div></td>
          </tr>
          <tr>
            <td class="cttd"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="28"><% call ShowClassGuide() %>                  </td>
                  <td width="100"><% call ShowArticleTotal() %>                  </td>
                </tr>
              </table></td>
          </tr>
          <tr>
            <td><% call ShowArticle(32) %></td>
          </tr>
          <tr>
            <td id='tpage'><%
if totalput>0 then
call showpage(strFileName,totalput,MaxPerPage,false,true,"����Ʒ")
end if
%>            </td>
          </tr>
        </table></td>
    </tr>
  </table>
</div>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>
