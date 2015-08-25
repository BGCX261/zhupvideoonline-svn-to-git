<%response.end%>
<div class="box" id="research">
  <div class="head"><div class="l"></div><span>{$L_RESEARCH}</span><div class="r"></div></div>
  <div class="main">
  {if:$research_u,egt,0}
    <form action="{$S_ROOT}index/submit.asp" method="post">
    <input name="cmd" type="hidden" value="research"/>
    {$sign=chr(123)&"v"&chr(125)}
    {$k=0}
    {for:$i,0,$research_u}
      {$arr=$split($research[$i]['text'],$sign)}
      <div class="question">{$arr[0]}</div>
      <div>
      {for:$j,2,ubound($arr)}
        {select:$arr[1]}
          {case:radio}<span><input name="res_{$k}" type="radio" value="{$arr[$j]}" />{$arr[$j]}</span>
          {case:checkbox}<span><input name="res_{$k}" type="checkbox" value="{$arr[$j]}" />{$arr[$j]}</span>{$k=$k+1}
          {case:text}<span class="text">{$arr[$j]}<input name="res_{$k}" type="text" /></span>{$k=$k+1}
        {/select}
      {/for}
      </div>
      {if:$arr[1],eq,"radio"}{$k=$k+1}{/if}
    {/for}
    <div class="bt"><input class="button" type="submit" value="{$L_SUBMIT}" /></div>
    </form>
  {/if}
  </div>
</div>
<!-- 新秀 -->