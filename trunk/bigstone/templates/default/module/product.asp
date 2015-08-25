<%response.end%>
<div class="box2 list" id="product">
  <div class="head"><span>{$L_NEW_PRODUCT}</span><a class="more" href="{$S_ROOT}?product.{$suffix}">{$L_MORE}</a></div>
  <ul class="main">
    {for:$i,0,$product_u}
	<li><a href="{$S_ROOT}?product_{$product[$i]['art_id']}.{$suffix}" target="_blank">{$cut_str($product[$i]['art_title'],20)}</a><span>{$datevalue($product[$i]['art_add_time'])}</span><div class="clear"></div></li>
    {/for}
  </ul>
</div>
<!-- 新秀 -->
