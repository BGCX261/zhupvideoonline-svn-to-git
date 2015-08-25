<!-- #include file="../../../include/config.asp" -->
<!-- #include file="../../../include/function.asp" -->
<!-- #include file="../../../languages/en-us.asp" -->
<%
url_page = "message"
site_title = "新秀免费企业网站系统 sinsiu 1.2 beta1"
channel_title = "留言板"
cat_id = 0
id = 0
url = "message.html"
suffix = "html"
keywords = "网站关键字"
describe = "网站描述"
dim nav(7,2)
nav(0,0) = "首页"
nav(0,1) = "/test/likai/sinsiu_v1.2_beta1/?index.html"
nav(0,2) = "0"
nav(1,0) = "关于我们"
nav(1,1) = "/test/likai/sinsiu_v1.2_beta1/?about.html"
nav(1,2) = "0"
nav(2,0) = "产品展示"
nav(2,1) = "/test/likai/sinsiu_v1.2_beta1/?product.html"
nav(2,2) = "0"
nav(3,0) = "资讯中心"
nav(3,1) = "/test/likai/sinsiu_v1.2_beta1/?article.html"
nav(3,2) = "0"
nav(4,0) = "人才招聘"
nav(4,1) = "/test/likai/sinsiu_v1.2_beta1/?recruit.html"
nav(4,2) = "0"
nav(5,0) = "下载中心"
nav(5,1) = "/test/likai/sinsiu_v1.2_beta1/?download.html"
nav(5,2) = "0"
nav(6,0) = "留言板"
nav(6,1) = "/test/likai/sinsiu_v1.2_beta1/?message.html"
nav(6,2) = "0"
nav(7,0) = "新秀官网"
nav(7,1) = "http://www.sinsiu.com"
nav(7,2) = "1"
nav_u = 7
dim nav_name(2)
nav_name(0) = "word"
nav_name(1) = "url"
nav_name(2) = "target"
dim focus(1,3)
focus(0,0) = "/test/likai/sinsiu_v1.2_beta1/images/focus_1.jpg"
focus(0,1) = "index.asp"
focus(0,2) = "焦点图片"
focus(0,3) = "default"
focus(1,0) = "/test/likai/sinsiu_v1.2_beta1/images/focus_2.jpg"
focus(1,1) = "index.asp"
focus(1,2) = "焦点图片"
focus(1,3) = "default"
focus_u = 1
dim focus_name(2)
focus_name(0) = "path"
focus_name(1) = "url"
focus_name(2) = "title"
dim research(1,1)
research(0,0) = "29"
research(0,1) = "您是如何知道本网站的？{v}radio{v}广告{v}名片{v}搜索引擎"
research(1,0) = "30"
research(1,1) = "您认为本公司产品质量如何？{v}radio{v}很好{v}一般{v}很差"
research_u = 1
dim research_name(1)
research_name(0) = "id"
research_name(1) = "text"
message_u = -1
dim message_name(3)
message_name(0) = "mes_user"
message_name(1) = "mes_time"
message_name(2) = "mes_text"
message_name(3) = "mes_reply"
page_sum = 1
record_sum = 1
cat = 0
page = 1
dim footer_nav(4,2)
footer_nav(0,0) = "首页"
footer_nav(0,1) = "/test/likai/sinsiu_v1.2_beta1/?index.html"
footer_nav(0,2) = "0"
footer_nav(1,0) = "关于我们"
footer_nav(1,1) = "/test/likai/sinsiu_v1.2_beta1/?about.html"
footer_nav(1,2) = "0"
footer_nav(2,0) = "产品展示"
footer_nav(2,1) = "/test/likai/sinsiu_v1.2_beta1/?product.html"
footer_nav(2,2) = "0"
footer_nav(3,0) = "资讯中心"
footer_nav(3,1) = "/test/likai/sinsiu_v1.2_beta1/?article.html"
footer_nav(3,2) = "0"
footer_nav(4,0) = "人才招聘"
footer_nav(4,1) = "/test/likai/sinsiu_v1.2_beta1/?recruit.html"
footer_nav(4,2) = "0"
footer_nav_u = 4
dim footer_nav_name(2)
footer_nav_name(0) = "word"
footer_nav_name(1) = "url"
footer_nav_name(2) = "target"
sit_code = "粤ICP备12345678号"
sit_code_url = "http://www.sinsiu.com"
sit_tec = "新秀工作室"
sit_tec_url = "http://www.sinsiu.com"
sit_count_code = "统计代码"
service = "在线客服"
%>
<!-- #include file="../compile/message.asp" -->
