<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<%
act=Request("act")
If act="save" Then 
FileName=trim(request.form("FileName"))
FileSize=trim(request.form("FileSize"))
FileMemo=trim(request.form("FileMemo"))
FileTime=now()

set rs=server.createobject("adodb.recordset")
sql="select * from web_Files"
rs.open(sql),cn,1,3
rs.addnew
rs("FileName")=FileName
rs("FileSize")=FileSize
rs("FileMemo")=FileMemo
rs("FileTime")=FileTime
rs.update
rs.close
set rs=nothing

response.Write "<script language='javascript'>alert('��ӳɹ���');location.href='File_List.asp';</script>"

end if
 %>

	<%
Call header()

%>
<%
set rs2=server.createobject("adodb.recordset")
sql="select * from web_FileSetting "
rs2.open(sql),cn,1,1
if not rs2.eof  then
FileFolder=rs2("FileFolder")
FileType=rs2("FileType")
FileSize=rs2("FileSize")
end if
rs2.close
set rs2=nothing
%>
  <form id="form1" name="form1" method="post" action="?act=save">
         <script language='javascript'>
function checksignup1() {
if ( document.form1.FileName.value == '' ) {
window.alert('��ѡ���ļ��ϴ�^_^');
document.form1.FileName.focus();
return false;}

if ( document.form1.FileSize.value == '' ) {
window.alert('�������ļ���С^_^');
document.form1.FileSize.focus();
return false;}

if(document.form1.FileSize.value.search(/^([0-9]*)([.]?)([0-9]*)$/)   ==   -1)   
      {   
  window.alert("�ļ���Сֻ��������^_^");   
document.form1.FileSize.focus();
return false;}

return true;}
</script>
	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>

	<tr>
	  <th class='tableHeaderText' colspan=2 height=25>�ϴ��ļ�</th>
	<tr>
	<tr>
	  <td height=23 colspan="2" class='forumRow'><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="20" class='TipTitle'>&nbsp;�� ������ʾ</td>
        </tr>
        <tr>
          <td height="30" valign="top" class="TipWords"><p>1��Ŀǰϵͳ�����ϴ��ļ���С���Ϊ<%=FileSize%>KB��</p>
            <p>2��Ŀǰϵͳ�����ϴ� <%=replace(FileType,"/"," ��")%>����չ�����ļ���</p>
			 <p>3������ļ��Ƚϴ��ϴ����ܻ��ʱ�����������ĵȴ�����Ҫ���������</p>
			<p>4���ļ�����޷��ϴ����������¼���ԭ��(1)��Ŀռ䲻֧��FSO�����(2)��Ŀռ�д��Ȩ��δ��;(3)�ϴ��ļ����Ͳ�֧��;(4)�ϴ��ļ�������С;(5)�ļ�����ļ��в����ڣ�(6)��Ŀռ�������(7)��Ŀռ��ٶȹ��ͣ�(8)�ڿ������ˡ�</p>
			<p>5�������ȷ�����������û�г��ֵĻ�����ô������ϵ���������ˡ�</p></td>
        </tr>
        <tr>
          <td height="10">&nbsp;</td>
        </tr>
      </table></td>
	  </tr>
	<td width="15%" height=23 class='forumRowHighLight'>�ϴ��ļ�</td>
	<td class='forumRowHighLight'><input name='FileName' type='text' id='FileName' size='30'  value="" >
	  &nbsp;<iframe frameborder="0" width="330" height="23" scrolling="No" src="Upload_File.asp?Action=upFile&Field=FileName&FieldSize=FileSize&FF=<%=FileFolder%>&FS=<%=FileSize%>&FT=<%=FileType%>"></iframe></td>
	</tr>
	  <tr>
	    <td class='forumRow' height=23>�ļ���С</td>
	    <td class='forumRow'><input name='FileSize' type='text' id='FileSize'  size='20'>KB��ϵͳ�Զ�����ļ���С�������޸ġ�</td>
      </tr>	  	
<tr>
	  <td class='forumRow' height=11>��ע</td>
	  <td class='forumRow'><textarea name='FileMemo'  cols="100" rows="6" id="FileMemo" ></textarea></td>
	</tr>	  
	<tr><td height="50" colspan=2  class='forumRow'><div align="center">
	  <INPUT type=submit value='�ύ' onClick='javascript:return checksignup1()' name=Submit>
	  </div></td></tr>
	</table>
</form>
<%
Call DbconnEnd()
 %>