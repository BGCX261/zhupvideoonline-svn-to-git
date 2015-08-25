<!-- #include file="../../../include/config.asp" -->
<!-- #include file="../../../include/function.asp" -->
<!-- #include file="../../../languages/en-us.asp" -->
<%
url_page = "recruit"
site_title = "新秀免费企业网站系统 sinsiu 1.2 beta1"
channel_title = "人才招聘"
cat_id = 0
id = 0
url = "recruit.html"
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
dim contact(8,1)
contact(0,0) = "联系人"
contact(0,1) = "陈湘国"
contact(1,0) = "电话"
contact(1,1) = "15919510148"
contact(2,0) = "手机"
contact(2,1) = "15919510148"
contact(3,0) = "传真"
contact(3,1) = "15919510148"
contact(4,0) = "邮编"
contact(4,1) = "521011"
contact(5,0) = "地址"
contact(5,1) = "广东省潮州市湘桥区景山村"
contact(6,0) = "网址"
contact(6,1) = "www.sinsiu.com"
contact(7,0) = "QQ/MSN"
contact(7,1) = "627780354"
contact(8,0) = "电子邮件"
contact(8,1) = "627780354@qq.com"
contact_u = 8
dim contact_name(1)
contact_name(0) = "cue"
contact_name(1) = "content"
dim recruit(0,2)
recruit(0,0) = "JOB"
recruit(0,1) = "<p>JOBS</p>"
recruit(0,2) = "2013/8/23 21:04:00"
recruit_u = 0
dim recruit_name(2)
recruit_name(0) = "art_title"
recruit_name(1) = "art_text"
recruit_name(2) = "art_add_time"
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
<!-- #include file="../compile/recruit.asp" -->
