<!--#include virtual="/DataBase/conn.asp"-->
<!--#include virtual="/System/function.asp"-->
<%
  dim sqlString,teachId,teachInfoRs,courseName,termName,teacherName
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'ȡ���Ͽ���Ϣ���
	teachId = CInt(Request("teachId"))
	sqlString = "select * from [publicCourseTeach] where teachId=" & teachId
	set teachInfoRs = Server.CreateObject("ADODB.RecordSet")
	teachInfoRs.Open sqlString,conn,1,1
	if not teachInfoRs.EOF then
	  'ȡ���ϿεĿγ̵�����
	  courseName = GetClassCourseNameByNumber(teachInfoRs("courseNumber"))
	  'ȡ�ø��Ͽ���Ϣ���ڵ�ѧ����Ϣ
	  termName = GetTermnameById(teachInfoRs("termId"))
	  'ȡ�ø��Ͽ���Ϣ���ڿν�ʦ������
	  teacherName = GetTeacherNameByNumber(teachInfoRs("teacherNumber"))
	end if
%>
<HTML>
<HEAD>
	<Title>�γ��ſ���ϸ��Ϣ�鿴</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/list.gif" width=14px height=14px>�ſ���Ϣ����--&gt;ѡ�޿γ���ϸ�ſ���Ϣ�鿴
			 </td>
	   </tr>
		 <tr>
		   <td width=100 align="right">�γ�����:</td>
			 <td><%=courseName%></td>
		 </tr>
		 <tr>
		   <td width=100px align="right">�Ͽ�ѧ��:</td>
			 <td>
			   <%=termName%>
			 </td>
		 </tr>
		 <tr>
		   <td width=100px align="right">�Ͽεص�:</td>
			 <td>
			   <%=teachInfoRs("teachClassRoom")%>
			 </td>
		 </tr>
		  <tr>
		   <td width=100px align="right">�Ͽ�ʱ��:</td>
			 <td>����
			  <%=teachInfoRs("teachDay")%>
				 </select>
			 </td>
		 </tr>
		 <tr>
		   <td width=100px align="right">��ϸ�Ͽ���Ϣ</td>
			 <td>
			   &nbsp;<input type="checkbox" name=MorningOne value="1" <%if teachInfoRs("MorningOne") = True then Response.Write "checked" end if%>>�����һ��
				 &nbsp;<input type="checkbox" name=MorningTwo value="1" <% if CInt(teachInfoRs("MorningTwo")) = True then Response.Write "checked" end if%>>����ڶ���<br>
				 &nbsp;<input type="checkbox" name=MorningThree value="1" <% if CInt(teachInfoRs("MorningThree")) = True then Response.Write "checked" end if%>>���������
				 &nbsp;<input type="checkbox" name=MorningFour value="1" <% if CInt(teachInfoRs("MorningFour")) = True then Response.Write "checked" end if%>>������Ľ�<br>
				 &nbsp;<input type="checkbox" name=MorningFive value="1" <% if CInt(teachInfoRs("MorningFive")) = True then Response.Write "checked" end if%>>��������
				 &nbsp;<input type="checkbox" name=AfternoonOne value="1" <% if CInt(teachInfoRs("AfternoonOne")) = True then Response.Write "checked" end if%>>�����һ��<br>
				 &nbsp;<input type="checkbox" name=AfternoonTwo value="1" <% if CInt(teachInfoRs("AfternoonTwo")) = True then Response.Write "checked" end if%>>����ڶ���
				 &nbsp;<input type="checkbox" name=AfternoonThree value="1" <% if CInt(teachInfoRs("AfternoonThree")) = True then Response.Write "checked" end if%>>���������<br>
				 &nbsp;<input type="checkbox" name=AfternoonFour value="1" <% if CInt(teachInfoRs("AfternoonFour")) = True then Response.Write "checked" end if%>>������Ľ�
				 &nbsp;<input type="checkbox" name=EveningOne value="1" <% if CInt(teachInfoRs("EveningOne")) = True then Response.Write "checked" end if%>>���ϵ�һ��<br>
				 &nbsp;<input type="checkbox" name=EveningTwo value="1" <% if CInt(teachInfoRs("EveningTwo")) = True then Response.Write "checked" end if%>>���ϵڶ���
				 &nbsp;<input type="checkbox" name=EveningThree value="1" <% if CInt(teachInfoRs("EveningThree")) = True then Response.Write "checked" end if%>>���ϵ�����
			 </td>
		 </tr>
		 <tr bgcolor="#ffffff">
        <td height="30" colspan="2" align="center">
				 <input type="button" value="����" onClick="javascript:location.href='publicCourseTeachMakeSecond.asp?courseNumber=<%=teachInfoRs("courseNumber")%>'">
		    </td>
      </tr>
	 </table>
 </form>
</body>
</html>