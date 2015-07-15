Imports UnumBridge.biz.plane.www

Public Class LaunchUnumSessionConnect
    Inherits System.Web.UI.Page

    Private Common As New Common
    Private DBase As New DBase
    Private UnumSession As UnumSession
    Private ACESSave As New ACESSave

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents litExchangeSessionID As System.Web.UI.WebControls.Literal
    Protected WithEvents litExchangeURL As System.Web.UI.WebControls.Literal
    Protected WithEvents litGoToUnum As System.Web.UI.WebControls.Literal
    Protected WithEvents lblMsg As System.Web.UI.WebControls.Label

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

#Region " Page_Load "
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            ' ___ BVI objects
            UnumSession = Session("UnumSession")
            Dim BVIEmployee As New BVIEmployee
            Dim BVIDependent As New BVIDependent

            ' ___ Unum objects
            Dim Employee As Employee = BVIEmployee.SetEmployeeValues(UnumSession)
            Dim LogonEnroller As LogonEnroller = SetLogonEnroller()
            Dim Dependents As DependentsDependent() = BVIDependent.SetDependentValues(UnumSession)
            Dim LogonEmployee As LogonEmployee = SetLogonEmployee()
            Dim BenefitInfo As BenefitInfo = SetBenefitInfo()
            Dim HostURL As HostURL = SetHostURL()
            Dim HTMLLayout As New HTMLLayout

            ' ___ Web service objects
            Dim InitReturn As New InitReturn
            Dim ec As New ExchangeConnect
            ec.Url = UnumSession.ecURL
            ec.HeaderInfoValue = New HeaderInfo
            ec.HeaderInfoValue.EmployerID = New Guid(UnumSession.ecGUID)

            ' ___ Call the web service
            InitReturn = ec.GetExchangeStatus(Employee, BenefitInfo, HTMLLayout, LogonEmployee, LogonEnroller, Dependents, HostURL)

            ' ___ Pass SessionID and URL to the page
            litExchangeSessionID.Text = "<input type=""hidden"" name=""ExchangeSessionID"" value=""" & InitReturn.SessionID & """ > "
            litExchangeURL.Text = "<input type=""hidden"" name=""hdExchangeURL"" value=""" & InitReturn.ExchangeURL & """>"

            ' ___ Save session value
            UnumSession.SessionID = InitReturn.SessionID

            ' ___ Save employee and dependent data sent
            If System.Environment.MachineName.ToUpper <> "LT-5ZFYRC1" Then
                Try
                    ACESSave.RecordEmpAndDepValuesSent(Employee, Dependents)
                    litGoToUnum.Text = "<input type='hidden' name='hdGoToUnum' value='1'>"
                Catch ex As Exception
                    litGoToUnum.Text = "<input type='hidden' name='hdGoToUnum' value='0'>"
                    lblMsg.Text = "Error: " & ex.Message
                End Try
            Else
                litGoToUnum.Text = "<input type='hidden' name='hdGoToUnum' value='1'>"
            End If


        Catch se As System.Web.Services.Protocols.SoapException
            If (se.Detail.InnerText.Length > 0) Then
                lblMsg.Text = se.Detail.InnerText
            Else
                lblMsg.Text = se.Detail.InnerText
            End If

        Catch ex As Exception
            lblMsg.Text = "Error: " & ex.Message
        End Try

    End Sub
#End Region

#Region " Logon Employee "
    Private Function SetLogonEmployee() As LogonEmployee
        Dim le As New LogonEmployee
        le.AuthenticationType = EmployeeAuthenticationTypeList.VoiceOther
        le.AuthenticationTypeSpecified = True
        le.IDKey = LogonEmployeeIDKey.SSN
        Return le
    End Function
#End Region

#Region " Logon Enroller "
    Private Function SetLogonEnroller() As LogonEnroller
        Dim le As New LogonEnroller
        Dim dt As DataTable

        dt = DBase.GetConvertDTSql("SELECT * FROM BVI..BVIUser WHERE UserID = '" & UnumSession.EnrollerID & "'", UnumSession.ServerIPAddress, "BVI")

        le.AuthenticationType = EnrollerAuthenticationTypeList.SignaturePad
        le.AuthenticationTypeSpecified = True
        le.EnrollmentCity = ""
        le.EnrollmentState = EnrollerStateList.CA
        le.EnrollmentStateSpecified = False
        le.FirstName = Common.DoTrim(dt.Rows(0)("FirstName").ToText, 25)
        le.LastName = Common.DoTrim(dt.Rows(0)("LastName").ToText, 25)
        le.LogonName = Common.DoTrim(UnumSession.UnumEnrollerID, 25)
        le.LogonType = LogonTypeList.CallCenterOther
        le.LogonTypeSpecified = True
        UnumSession.LastLoginIP = Common.DoTrim(dt.Rows(0)("LastLoginIP").ToText, 50)
        le.MiddleName = String.Empty
        Return le
    End Function
#End Region

#Region " Benefit Info "
    Private Function SetBenefitInfo() As BenefitInfo
        Dim bi As New BenefitInfo
        ' bi.BenefitType = BenefitTypeList.VWBAccident
        bi.BenefitType = UnumSession.BenefitType
        bi.BenefitTypeSpecified = True
        Return bi
    End Function
#End Region

#Region " HostURL Info "
    Private Function SetHostURL() As HostURL
        Dim hu As New HostURL
        hu.ReturnURL = "http://netserver.benefitvision.com/UnumBridge/Return.aspx"
        Return hu
    End Function
#End Region

End Class

