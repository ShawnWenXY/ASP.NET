<!--#include virtual="/DataBase/conn.asp"-->
<%
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'函数功能:根据设备类型得到设备名称
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
	<Title>设备库存信息查询</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
<br>
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=400 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=3 align="center">
		      <img src="../images/ADD.gif" width=14px height=14px>设备信息管理--&gt;设备库存查询
			 </td>
	   </tr>
		 <tr><td colspan=3 height=10></td></tr>
		 <tr>
		   <td>设备名称</td>
			 <td>设备类型</td>
			 <td>设备库存</td>
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