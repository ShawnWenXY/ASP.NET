<!--#include virtual="/DataBase/conn.asp"-->
<!--#include virtual="/System/function.asp"-->
<%
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'���Ҫ��黹ĳ���豸
	if Request("action") = "return" then
	  useInfoId = Request("useInfoId")
	  '���ȸ��¸���������Ϣ(�޸Ĺ黹ʱ��͹黹��־����)
	  sqlString = "select * from [deviceUseInfo] where useInfoId=" & useInfoId
	  set rs = server.CreateObject("ADODB.RecordSet")
	  rs.Open sqlString,conn,1,3
	  rs("useEndTime") = Now
	  rs("isReturn") = 1
	  rs.Update
	  'Ȼ�����Ӷ�Ӧ�豸�Ŀ��
	  sqlString = "update [deviceStockInfo] set deviceStock = deviceStock + 1 where deviceName='" & rs("deviceName") & "'"
    conn.Execute(sqlString)
	  Response.Write "<script>alert('�豸�黹�ɹ�!');</script>"
  end if
  
	'ȡ�ø�����ѯ�Ĳ�����Ϣ
	studentNumber = Request("studentNumber")
	startTime = Request("startTime")
	endTime = Request("endTime")
	sqlString = "select * from [deviceUseInfo] where isReturn=0"
	'���ݸ�����������Ϣ�����ѯsql
	if studentNumber <> "" then
	  sqlString = sqlString & " and studentNumber like '%" & studentNumber & "%'"
	end if
	if startTime <> "" then
	  sqlString = sqlString & " and useBeginTime > '" & CDate(startTime) & "'"
	end if
	if endTime <> "" then
	  sqlString = sqlString & " and useBeginTime < '" & CDate(endTime) & "'"
	end if
	set deviceUseInfoRs = Server.CreateObject("ADODB.RecordSet")
	deviceUseInfoRs.Open sqlString,conn,1,1
	
%>
<html>
<HEAD>
	<Title>�豸�黹</Title>
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
<br>
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=650 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=6 align="center">
		      <img src="../images/ADD.gif" width=14px height=14px>�豸��Ϣ����--&gt;�豸�黹
			 </td>
	   </tr>
		 <tr>
		   <td colspan=6 height=10>
			   ѧ��:<input type="text" name=studentNumber size=15 value=<%=Request("studentNumber")%>>
				 ��ʼʱ��:<input type=text name=startTime size=10 value='<%=Request("startTime")%>'>
				 <input class="submit" name="Button" onclick="seltime('startTime');" style="width:30px" type="button" value="ѡ��">
			   ����ʱ��:
				 <input type=text name=endTime size=10 value='<%=Request("endTime")%>'>
				 <input class="submit" name="Button" onclick="seltime('endTime');" style="width:30px" type="button" value="ѡ��">
			   &nbsp;&nbsp
				<input type="submit" value=" ���� " class="button1">
		   </td>
		 </tr>
		 <tr>
		   <td>�豸����</td>
			 <td>�豸����</td>
			 <td>ѧ��</td>
			 <td>����</td>
			 <td>����ʱ��</td>
			 <td>�黹</td>
		 </tr>
		 <%
			 while not deviceUseInfoRs.EOF
			   Response.Write "<tr><td>" & deviceUseInfoRs("deviceName") & "</td><td>" & GetDeviceTypeNameById(deviceUseInfoRs("deviceTypeId")) & "</td><td>" & deviceUseInfoRs("studentNumber") & "</td><td>" & GetStudentNameByNumber(deviceUseInfoRs("studentNumber")) & "</td><td>" & deviceUseInfoRs("useBeginTime") & "</td><td><a href='DeviceReturn.asp?action=return&useInfoId=" & deviceUseInfoRs("useInfoId") & "' onclick='return confirm(" & """" & "ȷ�Ϲ黹��?" & """" & ");'>�黹</a></td></tr>"
				 deviceUseInfoRs.MoveNext
			wend
		 %>
	 </table>
 </form>
</body>
</html>