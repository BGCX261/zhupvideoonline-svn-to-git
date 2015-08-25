<%response.end%>
<div class="box tree">
  <div class="head"><span>{$L_ART_TREE}</span></div>
  <div class="main">
  {for:$i,0,$tree_u}
    {if:$tree[$i]['grade'],eq,1}{$grade=1}{elseif:$tree[$i]['grade'],eq,2}{$grade=2}{else}{$grade=3}{/if}
    <div class="cat{$grade}"><a href="{$S_ROOT}?article_{$tree[$i]['id']}_0_0.{$suffix}">{$tree[$i]['name']}</a></div>
  {/for}
  </div>
</div>
<!-- 新秀 -->
