<!--#include virtual="/Database/conn.asp"-->
<%
  'errMessage���������Ϣ
  dim errMessage,sqlString
  errMessage = ""
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'�������Ա������µİ༶��Ϣ���ύ
	if Request("submit") <> "" then
	  '���û��ѡ�����ڵ�רҵ
	  if Request("classSpecialFieldNumber") = "" then
	    errMessage = "��ѡ������רҵ"
	  end if 
	  if Request("classBeginTime") = "" then
	    errMessage = "��ѡ��༶��ʼʱ��"
	  end if
	  '���û������༶����
	  if Request("className") = "" then
	    errMessage = "������༶����"
	  end if
	  '���ð༶�����Ƿ����
	  sqlString = "select * from [classInfo] where className='" & Request("className") & "'"
	  set classInfoRs = Server.CreateObject("ADODB.RecordSet")
	  classInfoRs.Open sqlString,conn,1,1
	  if not classInfoRs.EOF then
	    errMessage = "�ð༶������Ϣ�Ѿ�����"
		end if
		classInfoRs.Close
		
		'����errMessage�ж��Ƿ�Ҫ�����°༶��Ϣ���������
		if errMessage = "" then
		  '���濪ʼ�����°༶��Ϣ��ӳ�������
		  sqlString = "select * from [classInfo]"
		  classInfoRs.Open sqlString,conn,1,3
		  classInfoRs.AddNew
		  classInfoRs("classNumber") = Request("classNumber")
		  classInfoRs("className") = Request("className")
		  classInfoRs("classSpecialFieldNumber") = Request("classSpecialFieldNumber")
		  classInfoRs("classBeginTime") = CDate(Request("classBeginTime"))
		  classInfoRs("classYearsTime") = CInt(Request("classYearsTime"))
		  classInfoRs("classTeacherCharge") = Trim(Request("classTeacherCharge"))
		  classInfoRs("classMemo") = Trim(Request("classMemo"))
		  classInfoRs.Update
		  classInfoRs.Close
		  Response.Write "<script>alert('�༶��Ϣ��ӳɹ�!');</script>"
	  else
	    Response.Write "<script>alert('" & errMessage & "');</script>"
	  end if
	end if
%>
<HTML>
<HEAD>
	<Title>�༶��Ϣ���</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
		<script language=javascript>
	function seltime(inputName)
	{
	  window.open('seltime.asp?InputName='+inputName+'','','width=250,height=220,left=360,top=250,scrollbars=yes');  
	}
	</script>
</HEAD>
<BODY>
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height:14px" colspan=2 align="center">
		  <img src="../images/ADD.gif" width=14px height=14px>�༶��Ϣ����--&gt;�༶��Ϣ���<br>
			 </td>
	   </tr>
		 <tr>
		   <td width=100 align="right">ѡ��רҵ:
			 <td>
			   <select name=classSpecialFieldNumber>
				   <option value="">��ѡ��</option>
				  <%
					  sqlString = "select specialFieldNumber,specialFieldName from [specialFieldInfo]"
					  set specialFieldInfoRs = Server.CreateObject("ADODB.RecordSet")
					  specialFieldInfoRs.Open sqlString,conn,1,1
					  while not specialFieldInfoRs.EOF
					    Response.Write "<option value='" & specialFieldInfoRs("specialFieldNumber") & "'>" & specialFieldInfoRs("specialFieldName") & "</option>"
						  specialFieldInfoRs.MoveNext
					  wend
				  %>
				 </select>
			 </td>
		 </tr>
		 <tr>
		    <td width=100 align="right">�༶���:</td>
		    <td><input type=text name=classNumber size=20></td>
		  </tr>
		  <tr>
		    <td width=100 align="right">�༶����:</td>
		    <td><input type=text name=className size=20></td>
		  </tr>
		  <tr>
		    <td width=100 align="right">�༶��ʼʱ��:</td>
		    <td><input type=text name=classBeginTime width=25px>
			    <input class="submit" name="Button" onclick="seltime('classBeginTime');" style="width:30px" type="button" value="ѡ��">
				</td>
		  </tr>
		   <tr>
		    <td width=100 align="right">�༶ʱ��:</td>
		    <td><select name=classYearsTime>
			        <option value="3">3</option>
					    <option value="4">4</option>
					 </select>����</td>
		  </tr>
			<tr>
		    <td width=100 align="right">�༶������:</td>
		    <td><input type=text name=classTeacherCharge size=20></td>
		  </tr>
		  <tr>
		    <td width=100 align="right">������Ϣ:</td>
		    <td><textarea cols=40 rows=5 name=classMemo></textarea></td>
		  </tr>
      <tr bgcolor="#ffffff">
        <td height="30" colspan="4" align="center">
		      <input name="submit"  type="submit" value=" ȷ����� ">
		      <input type="reset" value=" ������д ">
		    </td>
      </tr>
	  </table>
  </form>
</BODY>

</HTML>
