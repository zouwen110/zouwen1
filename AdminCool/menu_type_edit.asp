<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/x_to_html/index_to_html.asp" -->
<% '�������ݵ����ݱ�
page=request.querystring("page")
act=request.querystring("act")
keywords=request.querystring("keywords")
article_id=cint(request.querystring("id"))


act1=Request("act1")
If act1="save" Then 
l_id=trim(request.form("l_id"))
l_name=trim(request.form("name"))
l_memo=trim(request.form("memo"))
l_number=trim(request.form("number"))
l_order=trim(request.form("order"))
l_time=now()

set rs=server.createobject("adodb.recordset")
sql="select * from web_menu_type where id="&l_id&""
rs.open(sql),cn,1,3
rs("name")=l_name
rs("memo")=l_memo
rs("number")=l_number
if l_order<>"" then
rs("order")=l_order
end if
rs.update
rs.close
set rs=nothing
response.Write "<script language='javascript'>alert('�޸ĳɹ���');location.href='menu_type_list.asp?page="&page&"&act="&act&"&keywords="&keywords&"';</script>"
end if 
 %>
 

	<%
Call header()

%>
<% set rs=server.createobject("adodb.recordset")
sql="select * from web_menu_type where id="&article_id&""
rs.open(sql),cn,1,1
if not rs.eof and not rs.bof then%>
  <form id="form1" name="form1" method="post" action="?act1=save&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>">
         <script language='javascript'>
function checksignup1() {
if ( document.form1.name.value == '' ) {
window.alert('�������������^_^');
document.form1.name.focus();
return false;}

if ( document.form1.number.value == '' ) {
window.alert('�����뵼������^_^');
document.form1.number.focus();
return false;}

if(document.form1.number.value.search(/^([0-9]*)([.]?)([0-9]*)$/)   ==   -1)   
      {   
  window.alert("��������ֻ��������^_^");   
document.form1.number.focus();
return false;}

if(document.form1.order.value.search(/^([0-9]*)([.]?)([0-9]*)$/)   ==   -1)   
      {   
  window.alert("����ֻ��������^_^");   
document.form1.order.focus();
return false;}

return true;}
</script>
	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th class='tableHeaderText' colspan=2 height=25>�޸ĵ�������</th>
	<tr>
	<td width="15%" height=23 class='forumRow'>�������� (����) </td>
	<td width="85%" class='forumRow'><input name='name' type='text' id='name'  value="<%=rs("name")%>" size='70'>
	 <input name='l_id' type='hidden' id='l_id' value="<%=rs("id")%>" size='70'>
	  &nbsp;</td>
	</tr>
		<tr>
	  <td class='forumRow' height=11>�������� (����) </td>
	  <td class='forumRow'><input name='number' type='text' id='number' size='20' maxlength="2" value="<%=rs("number")%>"> 
      ֻ�������� </td>
	  </tr>
	<tr>
	  <td class='forumRowHighLight' height=11>����</td>
	  <td class='forumRowHighLight'><input name='order' type='text' id='order' size='20' maxlength="2" value="<%=rs("order")%>">
ֻ�������֣�����ԽС����Խ��ǰ</td>
	  </tr>
	<tr>

	<tr>
	  <td class='forumRow' height=11>��ע</td>
	  <td class='forumRow'><textarea name='memo'  cols="100" rows="6" id="memo" ><%=rs("memo")%></textarea></td>
	</tr>
	<tr><td height="50" colspan=2  class='forumRow'><div align="center">
	  <INPUT type=submit value='�ύ' onClick='javascript:return checksignup1()' name=Submit>
	  </div></td></tr>
	</table>
</form>
<%
else
response.write"δ�ҵ�����"
end if%>
<%
Call DbconnEnd()
 %>