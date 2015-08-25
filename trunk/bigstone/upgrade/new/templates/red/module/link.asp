<%response.end%>
<div id="link">
  <div class="main">
  {if:$word_link_u,egt,0}
    <span>{$L_LINK}：</span>
    {for:$i,0,$word_link_u}
    <a href="{$word_link[$i]['url']}" title="{$word_link[$i]['title']}" target="_blank">{$word_link[$i]['word']}</a>
    {/for}
  {/if}
  </div>
</div>
<!-- 新秀 -->
