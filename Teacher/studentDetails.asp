<!--#include file="../Database/conn.asp"-->
<!--#include file="../System/md5.asp"--> 
<!--#include file="../System/function.asp"-->
<%
  '�������Ա��û�е�½
  if session("teacherNumber")="" and  session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
  '�õ�ĳ��ѧ������ϸ��Ϣ
  set studentDetailRs = Server.CreateObject("ADODB.RecordSet")
  sqlString = "select * from [studentInfo] where studentNumber='" & Request("studentNumber") & "'"
  studentDetailRs.Open sqlString,conn,1,1
%>
<HTML>
<HEAD>
	<Title>ѧ����ϸ��Ϣ</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/edit.gif" width=14px height=14px>ѧ����Ϣ����--&gt;ѧ����ϸ��Ϣ
			 </td>
	   </tr><br>
		 <%
		   '�����ѧ��������ͼƬ����ʾ��ѧ����ͷ��
		   if studentDetailRs("studentPhoto") <> "" then
			   Response.Write "<tr><td>ѧ��ͷ��:</td><td><img src='../admin/" & studentDetailRs("studentPhoto") & "' border=0 height=100 width=100></td></tr>"
			 end if 
		 %>
		 <tr>
			 <td>���ڰ༶:</td>
			 <td>
			   <%=GetClassNameByNumber(studentDetailRs("studentClassNumber"))%>
			 </td>
		 </tr>
	   <tr>
	     <td style="height: 26px">
		     ѧ��:</td><td><%=studentDetailRs("studentNumber")%></td>
			 </td>
		 </tr>
		 <tr>
		  <td>ѧ������:</td><td><%=studentDetailRs("studentName")%></td>
		 </tr>
		 <tr>
		   <td>�Ա�:</td>
			 <td>
			   <%=studentDetailRs("studentSex")%>
			 </td>
		 </tr>
		 <tr>
		   <td>ѧ������:</td>
			 <td>
			   <%=studentDetailRs("studentBirthday")%>
			</td>
		 </tr>
		 <tr>
		   <td>������ò:</td>
			 <td><%=studentDetailRs("studentState")%>
			 </td>
		 </tr>
		  <tr>
		    <td>��ͥ��ַ:</td>
			  <td><%=studentDetailRs("studentAddress")%></td>
		  </tr>
		  <tr>
		    <td width=100 align="right">������Ϣ:</td>
		    <td><textarea cols=40 rows=5 name=studentMemo><%=studentDetailRs("studentMemo")%></textarea></td>
		  </tr>
      <tr bgcolor="#ffffff">
        <td height="30" colspan="2" align="center">
		      <input type="button" value=" ����" onClick="javascript:location.href='studentInfoQuery.asp?studentNumber=<%=Request("studentQueryNumber")%>&studentName=<%=Request("studentQueryName")%>&studentClass=<%=Request("studentQueryClass")%>';">
		    </td>
      </tr>
	  </table>
  </form>
</BODY>
</HTML>
