<%response.end%>
<div id="header">
  <div class="main">
    <div class="logo"><a href="{$S_ROOT}"><img src="{$S_ROOT}images/logo.png" /></a></div>
    <div class="search">
      <form name="form_search" method="post" action="{$S_ROOT}index.asp?search.{$suffix}">
        <input class="text" name="search" type="text" maxlength="30" />
        {if:$S_LANG,neq,"en-us"}
        <a class="button" onclick="document.form_search.submit();">{$L_SEARCH}</a>
        {else}
        <a class="button" onclick="document.form_search.submit();" style="font-size:10px;font-weight:normal;">{$L_SEARCH}</a>
        {/if}
      </form>
    </div>
  </div>
  <div id="nav">
    <ul>
      {for:$i,0,$nav_u}
      <li {if:$i,eq,0}class="first"{/if}><a href="{$nav[$i]['url']}" {if:$nav[$i]['target'],eq,1}target="_blank"{/if}>{$nav[$i]['word']}</a></li>
      {/for}
      <div class="clear"></div>
    </ul>
  </div>
</div>
<!-- 新秀 -->
