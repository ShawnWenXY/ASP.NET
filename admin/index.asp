<!--#include virtual="/DataBase/conn.asp"-->
<!--#include virtual="/System/config.asp"-->
<%
  '�жϹ���Ա�Ƿ��Ѿ���½ϵͳ
  if session("adminUsername") = "" then
    Response.Redirect "../login.asp"
	end if
%>
<HTML>
	<HEAD>
		<title><%=systemName%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<LINK href="../css/style.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body onmousemove="hiddenMenu();">
		<form id="Form1" method="post" runat="server">
			<table align="center" cellSpacing="0" cellPadding="0" width="760" border="0" ID="Table5">
				<TBODY>
					<tr>
						<td>
							<!--   /�˵���   -->
							<table class="tbl" id="Table2" cellSpacing="0" cellPadding="0" width="760" border="0">
								<tr>
									<td colspan=3>
										<!--   ����������   -->
											<!--#include file="head.inc"-->
										<!--   /�������˵�   --></td>
								</tr>
								<tr width=100%>
									<td width=45%>&nbsp;����Ա<font color=blue><%=session("adminUsername")%></font>
</td><td width=20%>&nbsp;</td>
								</tr>
							</table>
							<!-- END PAGE HEADER MODULE -->
							<!--   ���ݲ�   -->
							<table class="lrb" align="center" cellSpacing="0" cellPadding="0" width="760" border="0"
								ID="Table3">
								<tr>
									<td bgcolor="#d6ebff" style="height: 530px">
									<iframe style="height: 530px; width: 760px;" frameborder="0" name="ContentFrame"  scrolling="yes" src="../System/systemInfo.asp" width="760"></IFRAME>
									</td>
								</tr>
							</table>
							<!--   /���ݲ�   -->
						</td>
					</tr>
					<tr><td>
					  <table align=center width=100% cellspacing=0 cellpadding=0>
	            <tr>
					     <td align=center bgcolor="#00ff33" width=100%></td>
	            </tr>
            </table>
				  </td></tr>
				</TBODY>
			</table>
		</form>
	</body>
</HTML>



