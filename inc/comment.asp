<!-- #include file="AntiAttack.asp" -->
<!-- #include file="conn.asp" -->
<!-- #include file="html_clear.asp" -->
<!-- #include file="Create.asp" -->
<!-- #include file="md5.asp" -->
<!-- #include file="web_AdvancedSettings.asp" -->
<!-- #include file="x_to_html/article_to_html.asp" -->
<!-- #include file="x_to_html/post_index_to_html.asp" -->

<%'�ж�
if request("act")="add" then

'�����ļ��л�ȡ
set rs_1=server.createobject("adodb.recordset")
sql="select FolderName from web_Models_type where [id]=9"
rs_1.open(sql),cn,1,1
if not rs_1.eof then
if rs_1("FolderName")<>"" then
Article_FolderName1="/"&rs_1("FolderName")
end if
end if
rs_1.close
set rs_1=nothing

article_id=request("id")
name1=trim(request.form("name"))
email1=trim(request.form("email"))
qq1=trim(request.form("qq"))
comment=trim(request.form("content"))
input_code=trim(request.form("verycode"))


if comment="" then
response.Write "<script language='javascript'>alert('�����������������ݣ�');history.go(-1)</script>"
else

    if request("verycode")="" then
    response.write "<script language=javascript>alert('���������֤������^_^');history.go(-1);</script>"
  	Response.End 
	elseif session("getcode")="9999" then
    session("getcode")=""
	elseif session("getcode")="" then
    response.write "<script language=javascript>alert('���������֤������^_^');history.go(-1);</script>"
 	Response.End 
	elseif cstr(session("getcode"))<>cstr(trim(request("verycode"))) then
    response.write "<script language=javascript>alert('���������֤������^_^');history.go(-1);</script>"
	Response.End 
	end if

' ��������
set rs=server.createobject("adodb.recordset")
sql="select * from web_article_comment where [content]='"&nohtml(comment)&"'"
rs.open(sql),cn,1,3
if not rs.eof then  
response.Write "<script language='javascript'>alert('�벻Ҫ�ظ��������ۣ�');history.go(-1)</script>"
else
rs.addnew
if article_id<>"" then
rs("article_id")=article_id
if web_FeedComment=1 then
rs("view_yes")=0
end if
else
if web_FeedAdvice=1 then
rs("view_yes")=0
end if

end if

rs("name")=name1
rs("email")=email1
rs("qq")=qq1
rs("content")=nohtml(comment)
rs("ip")=Request.serverVariables("REMOTE_ADDR")
rs("time")=now()
rs.update
rs.close
set rs=nothing


'������������1
if article_id<>"" then
	Pre_url=request.servervariables("HTTP_REFERER")
set rs=server.createobject("adodb.recordset")
sql="select [comment],[id],[title],file_path from [article] where [id]="&article_id&""
rs.open(sql),cn,1,3
if not rs.eof then
rs("comment")=rs("comment")+1
rs.update
a_id=rs("id")
a_title=rs("title")
a_link=Article_FolderName1&"/"&rs("file_path")
end if
rs.close
set rs=nothing
call article_to_html(a_id)
	response.Write "<script language='javascript'>alert('���������Ѿ������ɹ�^_^');location.href='"&Pre_url&"';</script>"
	else
call post_index_to_html()
			response.write"<SCRIPT language=JavaScript>alert('���������Ѿ������ɹ�^_^');"
  response.write"javascript:history.go(-1)</SCRIPT>"
end if


end if
end if
end if 
%>