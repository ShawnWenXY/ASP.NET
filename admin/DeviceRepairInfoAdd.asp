<!--#include virtual="/DataBase/conn.asp"-->
<!--#include virtual="/System/function.asp"-->
<%
  '���������Ϣ
  dim errMessage
  errMessage = ""
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'���Ҫ�����豸ά����Ϣ�ĵǼ�
	if Request("submit") <> "" then
	  if Request("deviceName") = "" then
	    errMessage = "��ѡ���豸����"
	  end if
	  sqlString = "select * from [deviceRepairInfo]"
	  set deviceRepairInfoRs = Server.CreateObject("ADODB.RecordSet")
	  deviceRepairInfoRs.Open sqlString,conn,1,3
	  deviceRepairInfoRs.AddNew
	  deviceRepairInfoRs("deviceName") = Request("deviceName")
	  deviceRepairInfoRs("deviceTypeId") = Request("deviceTypeId")
	  deviceRepairInfoRs("repairPlace") = Request("repairPlace")
	  deviceRepairInfoRs("repairMan") = Request("repairMan")
	  if Request("repairMoney") <> "" then
	    deviceRepairInfoRs("repairMoney") = CSng(Request("repairMoney"))
		else
		  deviceRepairInfoRs("repairMoney") = 0.0
		end if
		deviceRepairInfoRs("errorReason") = Request("errorReason")
		if Request("repairDate") <> "" then
		  deviceRepairInfoRs("repairDate") = CDate(Request("repairDate"))
		else
		  deviceRepairInfoRs("repairDate") = CDate("1900-1-1")
		end if
		deviceRepairInfoRs("repairMemo") = Request("repairMemo")
		deviceRepairInfoRs.Update
		Response.Write "<script>alert('�豸ά����Ϣ�Ǽǳɹ�!');location.href='DeviceRepairInfoList.asp';</script>"
	end if
%>
<HEAD>
	<Title>�豸ά����Ϣ���</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
	<script language=javascript>
	function seltime(inputName)
	{
	  window.open('seltime.asp?InputName='+inputName+'','','width=250,height=220,left=360,top=250,scrollbars=yes');  
	}
	var deviceType_deviceName = new Array();
	//��ʼ������(�豸���ͣ��豸������Ϣ)��Ϣ��¼����
  function initArray() {
  <%
    sql = "select * from deviceStockInfo" 
	  set deviceStockInfoRs = conn.Execute(sql)
	  i = 0
	  while not deviceStockInfoRs.eof
	    Response.Write "deviceType_deviceName[" & i & "]='" & deviceStockInfoRs("deviceTypeId") & ":" & deviceStockInfoRs("deviceName") & "';" & vbCrLf
		  i = i + 1
		  deviceStockInfoRs.MoveNext
	  wend
  %>
 }
 //��ѡ��ͬ���豸���ʱ��ʾ������µ��豸��Ϣ
function changeDeviceType() {
  var searchDeviceTypeId; //Ҫ�������豸����
  var eachDeviceTypeId; //ÿ����¼���豸����
  var eachDeviceName; //��¼ÿ���豸������
  var indexOfSplit; // :�ŷָ���ŵ�λ��
  var innerHtmlText;
  var oOption; 
  var index;
  innerHtmlText = "";
  searchDeviceTypeId = document.all.deviceTypeId.value;
  initArray(); //��ʼ����Ϣ����
  index = document.all.deviceName.length
  for(;index>0;index--) {
    document.all.deviceName.remove(index);
  }
  for(var i=0;i<deviceType_deviceName.length;i++) {
    indexOfSplit = deviceType_deviceName[i].indexOf(":"); //�õ�:�ŷָ���ŵ�λ��
	  eachDeviceTypeId = deviceType_deviceName[i].substr(0,indexOfSplit); //ȡ�õ�ǰ��¼���豸���ͱ��
	  if(searchDeviceTypeId == eachDeviceTypeId) { //����豸����һ���Ͱ��豸��Ϣȡ���������豸��Ϣ��������
	    eachDeviceName = deviceType_deviceName[i].substr(indexOfSplit+1);
		   oOption = document.createElement("OPTION");
		   document.all.deviceName.options.add(oOption);
			  oOption.innerText = eachDeviceName;
       oOption.value = eachDeviceName;
	  }
  }
}
	</script>
</HEAD>
<BODY>
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/ADD.gif" width=14px height=14px>�豸��Ϣ����--&gt;�豸ά����Ϣ���
			 </td>
	   </tr>
		 <tr>
		   <td width=100px align="right">�豸����:</td>
			 <td>
			   <select name=deviceTypeId onchange="changeDeviceType();">
				  <option value=''>��ѡ���豸����</option>
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
		   <td width=100 align="right">�豸����:</td>
			 <td>
			   <select name=deviceName>
				   <option value="">��ѡ���豸</option>
				 </select>
			 </td>
		 </tr>
		 <tr>
		   <td width=100px align="right">ά����:</td>
			 <td>
			   <input type=text name=repairMan size=20>
			 </td>
		 </tr>
		 <tr>
		   <td width=100px align="right">ά�޵ص�:</td>
			 <td>
			   <input type=text name=repairPlace size=40>
			 </td>
		 </tr>
		 <tr>
		   <td width=100px align="right">ά�޽��:</td>
			 <td>
			   <input type=text name=repairMoney size=5>Ԫ
			 </td>
		 </tr>
		 <tr>
		   <td width=100px align="right">����ԭ��:</td>
			 <td><input type=text name=errReason size=40></td>
		 </tr>
		 <tr>
		   <td width=100px align="right">ά������:</td>
			 <td>
			   <input type=text name=repairDate width=20px>
				 <input class="submit" name="Button" onclick="seltime('repairDate');" style="width:30px" type="button" value="ѡ��">
			 </td>
		 </tr>
		 <tr>
		    <td width=100 align="right">������Ϣ:</td>
		    <td><textarea cols=40 rows=5 name=repairMemo></textarea></td>
		  </tr>
		 <tr bgcolor="#ffffff">
        <td height="30" colspan="2" align="center">
		      <input name="submit"  type="submit" value="ά�޵Ǽ�"> &nbsp;
				  <input type="reset" value="ȡ��"></td>
      </tr>
	 </table>
 </form>
</body>
</html>