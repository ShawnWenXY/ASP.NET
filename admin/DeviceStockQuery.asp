<!--#include virtual="/DataBase/conn.asp"-->
<%
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'��������:�����豸���͵õ��豸����
	Function GetDeviceTypeNameById(deviceTypeId)
	  dim sqlString,deviceTypeName
	  sqlString = "select deviceTypeName from [deviceTypeInfo] where deviceTypeId=" & deviceTypeId
	  set deviceTypeInfoRs = Server.CreateObject("ADODB.RecordSet")
	  deviceTypeInfoRs.Open sqlString,conn,1,1
	  if not deviceTypeInfoRs.EOF then
	    deviceTypeName = deviceTypeInfoRs("deviceTypeName")
		else
		  deviceTypeName = ""
		end if
		GetDeviceTypeNameById = deviceTypeName
	End Function
%>
<html>
<HEAD>
	<Title>�豸�����Ϣ��ѯ</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
<br>
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=400 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=3 align="center">
		      <img src="../images/ADD.gif" width=14px height=14px>�豸��Ϣ����--&gt;�豸����ѯ
			 </td>
	   </tr>
		 <tr><td colspan=3 height=10></td></tr>
		 <tr>
		   <td>�豸����</td>
			 <td>�豸����</td>
			 <td>�豸���</td>
		 </tr>
		 <%
		   sqlString = "select * from [deviceStockInfo]"
			 set deviceStockInfoRs = Server.CreateObject("ADODB.RecordSet")
			 deviceStockInfoRs.Open sqlString,conn,1,1
			 while not deviceStockInfoRs.EOF
			   Response.Write "<tr><td>" & deviceStockInfoRs("deviceName") & "</td><td>" & GetDeviceTypeNameById(deviceStockInfoRs("deviceTypeId")) & "</td><td>" & deviceStockInfoRs("deviceStock") & "</td></tr>"
			   deviceStockInfoRs.MoveNext
			 wend
			 deviceStockInfoRs.Close
		 %>
	 </table>
 </form>
</body>
</html>