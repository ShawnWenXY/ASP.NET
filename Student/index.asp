<!--#include file="../DataBase/conn.asp"-->
<!--#include file="../System/config.asp"-->
<%
  '�ж�ѧ���Ƿ��Ѿ���½ϵͳ
  if session("studentNumber") = "" then
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
									<td width=45%>&nbsp;ѧ��<font color=blue><%=session("studentNumber")%></font>,��ã���ĵ�½ʱ����:<%=Now%></td><td width=35%></td><td width=20%>&nbsp;</td>
								</tr>
							</table>
							<!-- END PAGE HEADER MODULE -->
							<!--   ���ݲ�   -->
							<table class="lrb" align="center" cellSpacing="0" cellPadding="0" width="760" border="0"
								ID="Table3">
								<tr>
									<td bgcolor="#d6ebff" style="height: 400px">
									<iframe style="height: 500px; width: 760px;" frameborder="0"  name="ContentFrame" scrolling="auto" src="../System/systemInfo.asp" width="760"></IFRAME>
									</td>
								</tr>
							</table>
							<!--   /���ݲ�   -->
						</td>
					</tr>
					<tr><td>
					  <table align=center width=100% cellspacing=0 cellpadding=0>
	            <tr>
					     <td align=center bgcolor="#00ff33" width=100%>������� by <%=author%>;   QQ: <%=qq%>  �ֻ�:<%=phone%> Email:<a href="mailto:<%=email%>"><%=email%></a></td>
	            </tr>
            </table>
				  </td></tr>
				</TBODY>
			</table>
		</form>
	</body>
</HTML>



