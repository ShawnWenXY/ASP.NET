<!--#include virtual="/DataBase/conn.asp"-->
<!--#include virtual="/System/function.asp"-->
<%
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	sqlString = "select * from [deviceRepairInfo] where 1=1"
	deviceName = Request("deviceName")
	deviceTypeId = Request("deviceTypeId")
	startTime = Request("startTime")
	endTime = Request("endTime")
	if deviceName <> "" then
	  sqlString = sqlString & " and deviceName like '%" & deviceName & "%'"
	end if
	if deviceTypeId <> "" then
	  sqlString = sqlString & " and deviceTypeId=" & deviceTypeId
	end if
	if startTime <> "" then
	  sqlString = sqlString & " and repairDate > '" & startTime & "'"
	end if
	if endTime <> "" then
	  sqlString = sqlString & " and repairDate < '" & endTime & "'"
	end if
	set deviceRepairInfoRs = Server.CreateObject("ADODB.RecordSet")
	deviceRepairInfoRs.Open sqlString,conn,1,1
%>
<html>
<HEAD>
	<Title>�豸ά��</Title>
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
 
	 <table width='95%' border=0 cellpadding=0 cellspacing=0 align="center">
	 <tr>
		   <td colspan=6 height=10>
			 <a href="DeviceRepairInfoAdd.asp"><font color=red>�Ǽ�ά����Ϣ</font></a>
			  </td>
	</tr>
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=6 align="center">
		      <img src="../images/ADD.gif" width=14px height=14px>�豸��Ϣ����--&gt;�豸ά����Ϣ�б�
			 </td>
	   </tr>
		
		 <tr>
		   <td colspan=6 height=10>
			   �豸����:<input type="text" name=deviceName size=10 value=<%=Request("deviceName")%>>
				 �豸���:
				 <select name=deviceTypeId>
				   <option value="">ѡ�����</option>
				    <%
					   sqlString = "select * from [deviceTypeInfo]"
						 set deviceInfoRs = Server.CreateObject("ADODB.RecordSet")
						 deviceInfoRs.Open sqlString,conn,1,1
						 while not deviceInfoRs.EOF
						   Response.Write "<option value='" & deviceInfoRs("deviceTypeId") & "'>" & deviceInfoRs("deviceTypeName") & "</option>"
							deviceInfoRs.MoveNext
						 wend
						 deviceInfoRs.Close
					 %>
				 </select><br>
				 ��ʼʱ��:<input type=text name=startTime size=10 value='<%=Request("startTime")%>'>
				 <input class="submit" name="Button" onclick="seltime('startTime');" style="width:30px" type="button" value="ѡ��">
			   ����ʱ��:
				 <input type=text name=endTime size=10 value='<%=Request("endTime")%>'>
				 <input class="submit" name="Button" onclick="seltime('endTime');" style="width:30px" type="button" value="ѡ��">
				<input type="submit" value=" ���� " class="button1">
		   </td>
		 </tr>
		 <tr><td colspan=7 height=20></td></tr>
		 <tr>
		   <td>�豸����</td>
			 <td>�豸����</td>
			 <td>ά����</td>
			 <td>ά�޵ص�</td>
			 <td>ά�޽��</td>
			 <td>����ԭ��</td>
			 <td>ά������</td>
		 </tr>
		 <%
			 while not deviceRepairInfoRs.EOF
			   Response.Write "<tr><td>" & deviceRepairInfoRs("deviceName") & "</td><td>" & GetDeviceTypeNameById(deviceRepairInfoRs("deviceTypeId")) & "</td>"
				 Response.Write "<td>" & deviceRepairInfoRs("repairMan") & "</td>"
				 Response.Write "<td>" & deviceRepairInfoRs("repairPlace") & "</td>"
				 Response.Write "<td>" & deviceRepairInfoRs("repairMoney") & "</td>"
				 Response.Write "<td>" & deviceRepairInfoRs("errorReason") & "</td>"
				 Response.Write "<td>" & deviceRepairInfoRs("repairDate") & "</td>"
				 Response.Write "</tr>"
			   deviceRepairInfoRs.MoveNext
			 wend
		 %>
	 </table>
 </form>
</body>
</html>