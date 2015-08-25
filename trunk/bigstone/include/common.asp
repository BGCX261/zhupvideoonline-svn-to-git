<%'option explicit%>
<!-- #include file="config.asp" -->
<!-- #include file="db.class.asp" -->
<!-- #include file="function.asp" -->
<!-- #include file="interface.asp" -->

<%
call interface("reset_config",get_plu_name())

dim db:set db = new db_class
call db.set_db(S_ROOT,S_DB_TYPE,S_DB_PATH,S_DB_NAME,S_DB_USER,S_DB_PWD,S_DB_PREFIX)
call db.sql_connect()

'新秀
%>