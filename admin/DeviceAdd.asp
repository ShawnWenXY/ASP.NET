<!--#include virtual="/DataBase/conn.asp"-->
<%
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'���Ҫ�����豸�ĵǼ�
	if Request("submit") <> "" then
	  deviceName = Request("deviceName")
	  deviceTypeId = CInt(Request("deviceTypeId"))
	  deviceModel = Request("deviceModel")
	  deviceMadePlace = Request("deviceMadePlace")
	  if Request("devicePurchaseTime") <> "" then
	    devicePurchaseTime = CDate(Request("devicePurchaseTime"))
		else
		  devicePurchaseTime = CDate("1900-1-1")
		end if
		if Request("deviceCount") <> "" then
		  deviceCount = CInt(Request("deviceCount"))
		else
		  deviceCount = 0
		end if
		deviceMemo = Request("deviceMemo")
		if deviceName = "" then
		  Response.Write "<script>alert('�������豸������!');</script>"
		else
		  '���豸������Ϣ���뵽�豸������Ϣ����
		  sqlString = "select * from [deviceBuyInfo]"
		  set deviceBuyInfoRs = Server.CreateObject("ADODB.RecordSet")
		  deviceBuyInfoRs.Open sqlString,conn,1,3
		  deviceBuyInfoRs.AddNew
		  deviceBuyInfoRs("deviceName") = deviceName
		  deviceBuyInfoRs("deviceTypeId") = deviceTypeId
		  deviceBuyInfoRs("deviceModel") = deviceModel
		  deviceBuyInfoRs("deviceMadePlace") = deviceMadePlace
		  deviceBuyInfoRs("devicePurchaseTime") = devicePurchaseTime
		  deviceBuyInfoRs("deviceCount") = deviceCount
		  deviceBuyInfoRs("deviceMemo") = deviceMemo
		  deviceBuyInfoRs.Update
		  deviceBuyInfoRs.Close
		  '�ڿ��������Ӷ�Ӧ�豸�Ŀ����,�����û�и��豸����Ϣ����������Ϣ
		  sqlString = "select * from [deviceStockInfo] where deviceName='" & deviceName & "'"
		  set deviceStockInfoRs = Server.CreateObject("ADODB.RecordSet")
		  deviceStockInfoRs.Open sqlString,conn,1,3
		  if deviceStockInfoRs.EOF then
		    sqlString = "insert into [deviceStockInfo] (deviceName,deviceTypeId,deviceStock) values ('" & deviceName & "'," & deviceTypeId & "," & deviceCount & ")"
			  conn.Execute(sqlString)
		  else
		    deviceStockInfoRs("deviceStock") = deviceStockInfoRs("deviceStock") + deviceCount
			  deviceStockInfoRs.Update
		  end if
		  Response.Write "<script>alert('���豸�Ǽǳɹ�!');</script>"
		end if
	end if
%>
<HEAD>
	<Title>�豸��Ϣ���</Title>
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
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/ADD.gif" width=14px height=14px>�豸��Ϣ����--&gt;���豸��Ϣ���
			 </td>
	   </tr>
		 <tr>
		   <td width=100 align="right">�豸����:</td>
			 <td><input type=text  name=deviceName size=20></td>
		 </tr>
		 <tr>
		   <td width=100px align="right">�豸����:</td>
			 <td>
			   <select name=deviceTypeId>
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
				 </select>
			 </td>
		 </tr>
		 <tr>
		   <td width=100px align="right">�豸�ͺ�:</td>
			 <td>
			   <input type=text name=deviceModel size=20>
			 </td>
		 </tr>
		 <tr>
		   <td width=100px align="right">��������:</td>
			 <td>
			   <input type=text name=deviceMadePlace size=40>
			 </td>
		 </tr>
		 <tr>
		   <td width=100px align="right">��������:</td>
			 <td>
			   <input type=text name=devicePurchaseTime width=20px>
				 <input class="submit" name="Button" onclick="seltime('devicePurchaseTime');" style="width:30px" type="button" value="ѡ��">
			 </td>
		 </tr>
		 <tr>
		   <td width=100px align="right">��������:</td>
			 <td><input type=text name=deviceCount size=5>��</td>
		 </tr>
		 <tr>
		    <td width=100 align="right">������Ϣ:</td>
		    <td><textarea cols=40 rows=5 name=deviceMemo></textarea></td>
		  </tr>
		 <tr bgcolor="#ffffff">
        <td height="30" colspan="2" align="center">
		      <input name="submit"  type="submit" value="�豸�Ǽ�"> &nbsp;
				  <input type="reset" value="ȡ��"></td>
      </tr>
	 </table>
 </form>
</body>
</html>