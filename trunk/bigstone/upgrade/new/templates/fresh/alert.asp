<%response.end%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<meta name ="keywords" content="{$keywords}" />
<meta name ="description" content="{$describe}" />
<title>{$site_title}</title>
<link href="{$S_TPL_PATH}css.css" rel="stylesheet" type="text/css">
</head>
<body>
{select case $alert_type}
  {case:leave_word}{$alert=$L_SUBMIT_LEAVE_WORD}
  {case:research}{$alert=$L_SUBMIT_RESEARCH}
  {case:empty}{$alert=$L_FORBID_EMPTY}
  {default}{$alert=$L_FORBID_EMPTY}
{/select}
<script language="javascript">
  alert("{$alert}");
  document.location.href = "{$S_ROOT}{$href}";
</script>
</body>
</html>