<%@ Page Language="vb" AutoEventWireup="false" Codebehind="LaunchUnumSession.aspx.vb" Inherits="UnumBridge.LaunchUnumSession"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
  <HEAD>
		<title>LaunchUnumSession</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK title="BVIStyle" href="BVI.css" type="text/css" rel="stylesheet">
  </HEAD>
	<body>
		<form id="form1" method="post" runat="server">
			<asp:PlaceHolder id="phDisplayRegID" runat="server">
<TABLE class=ReturnTbl style="LEFT: 140px; POSITION: absolute; TOP: 14px" 
cellSpacing=0 cellPadding=0 width=500 border=0>
  <TR>
    <TD style="FONT-SIZE: 16pt; FONT-FAMILY: arial">UnumID Registration</TD></TR>
  <TR>
    <TD class=CellSeparator></TD></TR>
  <TR>
    <TD>If you have not yet registered as an enroller with Unum, please contact your supervisor and do so now. Then you must enter your Unum UserID (NOT your BVI UserID) here. Once you do so, this page will not display.</TD></TR>
  <TR>
    <TD>&nbsp;</TD></TR>
  <TR>
    <TD>Unum UserID:&nbsp;&nbsp;&nbsp;&nbsp;<INPUT type=text 
      name=txtUnumUserID></TD></TR>
  <TR>
    <TD>&nbsp;</TD></TR>
  <TR>
    <TD><INPUT onclick=fnSubmit() type=button value=Submit <></TD></TR></TABLE>
			</asp:PlaceHolder>	
			<asp:PlaceHolder id="phDisplayErrorMsg" runat="server">
<TABLE class=ReturnTbl style="LEFT: 140px; POSITION: absolute; TOP: 14px" 
cellSpacing=0 cellPadding=0 width=500 border=0>
  <TR>
    <TD style="FONT-SIZE: 16pt; FONT-FAMILY: arial">An error has 
occurred</TD></TR>
  <TR>
    <TD class=CellSeparator></TD></TR>
  <TR>
    <TD>Unable to obtain UnumID</TD></TR></TABLE>
			</asp:PlaceHolder>				
		</form>
		<script language="javascript">
			function fnClose()  {				
				window.opener= "x"
				window.close()
			}
			
			function fnSubmit() {
				if (form1.txtUnumUserID.value.length > 0)  {
					form1.submit()
				}
			}
		</script>
	</body>
</HTML>
