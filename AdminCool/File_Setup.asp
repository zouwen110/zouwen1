<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<%
act=Request("act")
If act="save" Then 
FileFolder=trim(request.form("FileFolder"))
FileType=trim(request.form("FileType"))
FileSize=trim(request.form("FileSize"))
'FileNameType=trim(request.form("FileNameType"))
FileTime=now()


set rs=server.createobject("adodb.recordset")
sql="select * from web_FileSetting"
rs.open(sql),cn,1,3
OldFolderDir=rs("FileFolder")
rs("FileFolder")=FileFolder
rs("FileType")=FileType
rs("FileSize")=FileSize
'rs("FileNameType")=FileNameType
rs("FileTime")=FileTime
rs.update
rs.close
set rs=nothing

'���ԭ�ļ����Ƿ���ڣ����򴴽�
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath("/"&OldFolderDir))=false Then
NewFolderDir="/"&OldFolderDir
call CreateFolderB(NewFolderDir)
end if
'������ļ����Ƿ���ԭ�ļ��в�ͬ�����������
if FileFolder<>OldFolderDir  then
NewFolderDir="/"&FileFolder
call renamefolder("/"&OldFolderDir,NewFolderDir) 
end if


response.Write "<script language='javascript'>alert('�޸ĳɹ���')</script>"

end if
 %>

	<%
Call header()

%>
<%
set rs2=server.createobject("adodb.recordset")
sql="select * from web_FileSetting "
rs2.open(sql),cn,1,3
if not rs2.eof and not rs2.bof then
%>
  <form id="form1" name="form1" method="post" action="?act=save">
         <script language='javascript'>
function checksignup1() {
if ( document.form1.FileFolder.value == '' ) {
window.alert('�������ļ���λ��^_^');
document.form1.FileFolder.focus();
return false;}

if ( document.form1.FileType.value == '' ) {
window.alert('�����������ϴ��ļ�����^_^');
document.form1.FileType.focus();
return false;}

if ( document.form1.FileSize.value == '' ) {
window.alert('�����������ϴ��ļ���С^_^');
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
	  <th class='tableHeaderText' colspan=2 height=25>�ϴ�����</th>
	<tr>
	<tr>
	  <td height=23 colspan="2" class='forumRow'><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="20" class='TipTitle'>&nbsp;�� ������ʾ</td>
        </tr>
        <tr>
          <td height="30" valign="top" class="TipWords"><p>1�����ļ����λ�á���ָ���ϴ����ļ����õĵط���һ��λ����Ŀռ�ĸ�Ŀ¼�µ�ĳ���ļ��С�</p>
            <p>2���������ϴ��ļ����͡���������Щ���͵��ļ��ǿ����ϴ��ģ����鲻Ҫ���ù���ĺ�İ�����ļ����ͣ���ȷ��ϵͳ��ȫ��</p>
			<p>3���������ϴ��ļ���С�����鲻Ҫ���ù����ϴ�̫����ļ����ܵ��³�ʱ���޷��ϴ��������ϴ��ļ��Ĵ�С�����ܻ��ܵ��ռ�����ơ�</p>
			<p>4�����ڳ����İ칫�ĵ���WORD,EXCEL,POWERPOINT���ļ�һ�㲻��̫�󣬲���ϵͳĬ�ϵ�2M�����Ѿ����á�</p>
            </td>
        </tr>
        <tr>
          <td height="10">&nbsp;</td>
        </tr>
      </table></td>
	  </tr>
	<td width="15%" height=23 class='forumRowHighLight'>�ļ����λ��</td>
	<td class='forumRowHighLight'><input name='FileFolder' type='text' id='FileFolder' size='40'  value="<%=rs2("FileFolder")%>" >
	  &nbsp;����ϵͳ��Ŀ¼�½����ļ��д�����ϴ����ļ�</td>
	</tr>
	  <tr>
	    <td class='forumRow' height=23>�����ϴ��ļ�����</td>
	    <td class='forumRow'><input name='FileType' type='text' id='FileType' value="<%=rs2("FileType")%>" size='80'>
        &nbsp;����ļ���չ���� / �ֿ���</td>
      </tr>
	  <tr>
	    <td class='forumRowHighLight' height=23>�����ϴ��ļ���С</td>
	    <td class='forumRowHighLight'><input name='FileSize' type='text' id='FileSize' value="<%=rs2("FileSize")%>" size='20'>KB��ϵͳĬ�����ƴ�СΪ2MB��1MB=1024KB��</td>
      </tr>	  

	<tr><td height="50" colspan=2  class='forumRow'><div align="center">
	  <INPUT type=submit value='�ύ' onClick='javascript:return checksignup1()' name=Submit>
	  </div></td></tr>
	</table>
</form>
<%
end if
rs2.close
set rs2=nothing
%>
<%
Call DbconnEnd()
 %>