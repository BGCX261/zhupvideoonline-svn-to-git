<%response.end%>
<div id="header">
  <div class="logo"><a href="{$S_ROOT}"><img src="{$S_ROOT}images/logo.jpg" /></a></div>
  <div class="search">
    <form method="post" action="{$S_ROOT}index.asp?search.{$suffix}">
      <input class="text" name="search" type="text" maxlength="30" />
      <input class="button" name="submit" type="submit" value=""/>
    </form>
  </div>
  <div id="nav">
    <div class="l"></div>
    <ul>
      {for:$i,0,$nav_u}
      <li><a href="{$nav[$i]['url']}" {if:$nav[$i]['target'],eq,1}target="_blank"{/if}>{$nav[$i]['word']}</a></li>
      {/for}
      <div class="clear"></div>
    </ul>
    <div class="r"></div>
  </div>
</div>
<!-- 新秀 -->
