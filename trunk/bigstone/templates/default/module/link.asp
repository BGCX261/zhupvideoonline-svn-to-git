<%response.end%>
<div class="box" id="link">
  <div class="head"><span>{$L_LINK}</span></div>
  <div class="main">
    {if:$img_link_u,egt,0}
    <div class="img">
      {for:$i,0,$img_link_u}
      <a href="{$img_link[$i]['url']}" title="{$img_link[$i]['title']}" target="_blank"><img src="{$img_link[$i]['img']}"/></a>
      {/for}
      <div class="clear"></div>
    </div>
    {/if}
    {if:$word_link_u,egt,0}
    <div class="word">
      {for:$i,0,$word_link_u}
      <a href="{$word_link[$i]['url']}" title="{$word_link[$i]['title']}" target="_blank">{$word_link[$i]['word']}</a>
      {/for}
    </div>
    {/if}
  </div>
</div>
<!-- 新秀 -->
