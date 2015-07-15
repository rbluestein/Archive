Imports UnumBridge.biz.plane.www
Imports System.Data

Partial Class ReturnFromUnum
    Inherits System.Web.UI.Page

    Private Common As New Common
    Private DBase As New DBase
    Private ACESSave As New ACESSave
    Private UnumSession As UnumSession

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub DoTemp()
        ' UnumSession.EmpSSN = "011-46-3915"
        UnumSession.EmpSSN = "999-99-9999"
        UnumSession.DBName = "HT"
        UnumSession.ServerIPAddress = "192.168.1.15"
        UnumSession.SessionID = "a8f0aed6-dfed-4a46-a958-dac9d0d51709"
        UnumSession.EmpID = "101058"
        UnumSession.ActivityID = "0a673588-f728-488f-9d0b-01bb40fd4164"
        UnumSession.EnrollerID = "ccwright"
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim BenefitObj As BenefitObj

        Try

            UnumSession = Session("UnumSession")

            If System.Environment.MachineName.ToUpper = "LT-5ZFYRC1" Then
                DoTemp()
            End If

            UnumSession.GetClientVariables()
            Dim ec As New ExchangeConnect
            ec.Url = UnumSession.ecURL
            ec.HeaderInfoValue = New HeaderInfo
            ec.HeaderInfoValue.EmployerID = New Guid(UnumSession.ecGUID)
            Dim ed As New EnrollmentData
            Dim esp As New EnrollmentSummaryParameters

            'esp.BenefitType = ESBenefitTypeList.VWBAccident
            esp.BenefitType = UnumSession.ESBenefitType

            esp.SessionID = UnumSession.SessionID

            ed = ec.GetExchangeEnrollmentSummary(esp)

            DeleteOldValues(ed)
            BenefitObj = GetBenefitObj(ed)

            ' Test for opt out
            If ed.TotalPremiumEmployee.Annual > 0 Then
                RecordNewValues(ed, BenefitObj)
            End If

            ACESSave.RecordEmpAndDepValuesReceived(ed, UnumSession)

            GetPrompt(True, Nothing)

        Catch se As System.Web.Services.Protocols.SoapException
            If (se.Detail.InnerText.Length > 0) Then
                GetPrompt(False, "Service Error 1: " & se.Detail.InnerText)
            Else
                GetPrompt(False, "Service Error 2: " & se.Message)
            End If

        Catch ex As Exception
            GetPrompt(False, "Service Error 3: " & ex.Message)
        End Try

    End Sub

    Private Function GetPrompt(ByVal Success As Boolean, ByVal ErrorMsg As String) As String
        Dim sb As New System.Text.StringBuilder

        If Success Then
            sb.Append("<td><p style='font-family:arial;font-size:16pt;'>")
            sb.Append("ENROLLMENT DATA BRIDGE TO ACES:<br>Data Captured!</p>")
            sb.Append("<p style='font-family:arial;font-size:10pt;'>Please close this window now.</p><br></td>")
        Else
            sb.Append("<td><p style='font-family:arial;font-size:16pt;'>")
            sb.Append("ENROLLMENT DATA BRIDGE TO ACES:<br>Error transferring data!</p>")
            sb.Append("<p style='color:black;'><b>ENROLLER:</b> Although any applications you just took were saved in the UNUM system, an error has occurred trying to pass the information back to ACES.")
            sb.Append("Please continue the enrollment but you must do two things:<br>")
            sb.Append("&nbsp;&nbsp;1. Manually add in the per-pay cost of the policy in the post-tax total when communicating deduction totals.<br>")
            sb.Append("&nbsp;&nbsp;2. Notify your supervisor and submit a helpdesk report so that the confirmation statement may be recreated.</p>")
            sb.Append(ErrorMsg & "<br><br></td>")
        End If

        litPrompt.Text = sb.ToString
    End Function

    Private Function GetBenefitObj(ByRef ed As EnrollmentData) As BenefitObj
        Dim i, p, b As Integer
        Dim ProductID As String
        Dim PolicyScope As PolicyScopeEnum
        Dim BenefitObj As New BenefitObj

        ' ___ Cycle through the Persons array
        If ed.Persons.GetUpperBound(0) > -1 Then
            For p = 0 To ed.Persons.GetUpperBound(0)

                ' ___ Cycle through the Benefits array
                If ed.Persons(p).Benefits.GetUpperBound(0) > -1 Then
                    For b = 0 To ed.Persons(p).Benefits.GetUpperBound(0)

                        ProductID = ed.Persons(p).Benefits(b).ProductID
                        If Not BenefitObj.ProductIDExists(ProductID) Then
                            PolicyScope = GetPolicyScope(ed, ProductID)
                            BenefitObj.Add(ProductID, PolicyScope)
                        End If

                    Next  ' Benefits
                End If

            Next   ' Persons
        End If
        Return BenefitObj
    End Function

    Private Function GetPolicyScope(ByRef ed As EnrollmentData, ByVal ProductID As String) As PolicyScopeEnum
        Dim p, b, c As Integer

        ' ___ Cycle through the Persons array
        If ed.Persons.GetUpperBound(0) > -1 Then
            For p = 0 To ed.Persons.GetUpperBound(0)

                ' ___ Cycle through the Benefits array
                If ed.Persons(p).Benefits.GetUpperBound(0) > -1 Then
                    For b = 0 To ed.Persons(p).Benefits.GetUpperBound(0)

                        If ed.Persons(p).Benefits(b).ProductID = ProductID Then

                            ' __ Cycle through the Coverages array
                            If ed.Persons(p).Benefits(b).Coverages.GetUpperBound(0) > -1 Then
                                For c = 0 To ed.Persons(p).Benefits(b).Coverages.GetUpperBound(0)

                                    If ed.Persons(p).Benefits(b).Coverages(c).Plan.ToLower = "individual" Then
                                        Return PolicyScopeEnum.Individual
                                    Else
                                        Return PolicyScopeEnum.Multiple
                                    End If

                                Next
                            End If

                        End If

                    Next
                End If
            Next
        End If
    End Function

    Private Sub DeleteOldValues(ByRef ed As EnrollmentData)
        Dim Sql As New System.Text.StringBuilder

        Sql.Append("DELETE IAMS..AppsAndPolsSummary WHERE EmpID = '" & UnumSession.EmpID & "' ")
        Sql.Append("AND ClientID='" & UnumSession.ClientID & "' ")
        Sql.Append(" AND ProductID = '" & UnumSession.BVIProductID & "'")
        Sql.Append(" AND BVIAppStatus = 'NEW' ")
        Sql.Append(" AND Cast(Year(AddDate) as varchar(4)) = '" & DateTime.Now.ToString("yyyy") & "' ")
        Sql.Append(" AND Cast(Month(AddDate) as varchar(4)) = '" & DateTime.Now.ToString("MM") & "'")

        Dim CmdAsst As New CmdAsst(DBase, CommandType.Text, Sql.ToString, UnumSession.ServerIPAddress, "IAMS")
        Dim QueryPack As CmdAsst.QueryPack
        QueryPack = CmdAsst.Execute

        '        If Not QueryPack.Success Then
        '        Throw New Exception("Unable to delete old records.")
        '        End If
    End Sub

    Private Sub RecordNewValues(ByRef ed As EnrollmentData, ByRef BenefitObj As BenefitObj)
        Dim i As Integer
        Dim PolicyScope As PolicyScopeEnum

        For i = 1 To BenefitObj.Coll.Count
            Select Case BenefitObj.Coll(i).PolicyScope
                Case PolicyScopeEnum.Individual
                    HandleIndividual(ed, BenefitObj.Coll(i).ProductID)
                Case PolicyScopeEnum.Multiple
                    HandleMultiple(ed, BenefitObj.Coll(i).ProductID)
            End Select
        Next
    End Sub

    Private Sub HandleIndividual(ByRef ed As EnrollmentData, ByVal ProductID As String)
        Dim p, b, c As Integer

        ' Cycle through the Persons array
        If ed.Persons.GetUpperBound(0) > -1 Then
            For p = 0 To ed.Persons.GetUpperBound(0)

                ' Cycle through the Benefits array
                If ed.Persons(p).Benefits.GetUpperBound(0) > -1 Then
                    For b = 0 To ed.Persons(p).Benefits.GetUpperBound(0)

                        If ed.Persons(p).Benefits(b).ProductID = ProductID Then

                            ' Insert into AppPolsSummary one record per individual per product
                            Dim AppPolSummary As New AppPolSummary
                            AppPolSummary.FirstName = ed.Persons(p).FirstName
                            AppPolSummary.MiddleName = ed.Persons(p).MiddleName
                            AppPolSummary.LastName = ed.Persons(p).LastName
                            AppPolSummary.ProductID = ed.Persons(p).Benefits(b).ProductID
                            AppPolSummary.CoverageBeginDate = ed.Persons(p).Benefits(b).CoverageBeginDate
                            AppPolSummary.MonthlyPremium = ed.Persons(p).Benefits(b).TotalBenefitPremiumEmployee.Annual / 12
                            AppPolSummary.PayPeriodPremium = ed.Persons(p).Benefits(b).TotalBenefitPremiumEmployee.Annual / UnumSession.PayFrequencyFactor

                            ' Cycle through the Coverages array
                            If ed.Persons(p).Benefits(b).Coverages.GetUpperBound(0) > -1 Then
                                For c = 0 To ed.Persons(p).Benefits(b).Coverages.GetUpperBound(0)

                                    ' Prepare new AppPolData record
                                    'Dim AppPolData As New AppPolData
                                    'AppPolData.Description = AppPolSummary.ProductID & "|" & ed.Persons(p).Benefits(b).Coverages(c).Description
                                    'AppPolSummary.CoverageBeginDate = ed.Persons(p).Benefits(b).CoverageBeginDate
                                    'AppPolData.MonthlyPremium = ed.Persons(p).Benefits(b).Coverages(c).Premium.Annual / 12
                                    'AppPolData.PayPeriodPremium = ed.Persons(p).Benefits(b).Coverages(c).Premium.Annual / UnumSession.PayFrequencyFactor
                                    'AppPolData.CoverageAmount = ed.Persons(p).Benefits(b).Coverages(c).CoverageAmount
                                    'AppPolData.CoverageAmountSpecified = ed.Persons(p).Benefits(b).Coverages(c).CoverageAmountSpecified
                                    'acessave.AppPolDataSave(ed, AppPolSummary, DBNull.Value)

                                    ' Accumulating coverage amount for AppPolSummary
                                    AppPolSummary.CoverageAmount += ed.Persons(p).Benefits(b).Coverages(c).CoverageAmount

                                Next  ' Coverages
                            End If

                            If AppPolSummary.CoverageAmount > 0 Then
                                AppPolSummary.CoverageAmountSpecified = True
                            End If


                            ACESSave.AppPolSummarySave(ed, AppPolSummary, DBNull.Value, UnumSession)

                        End If

                    Next  ' Benefits
                End If

            Next   ' Persons
        End If
    End Sub

    Private Sub HandleMultiple(ByRef ed As EnrollmentData, ByVal ProductID As String)
        Dim p, b, c As Integer

        ' ___ Insert into AppPolsSummary one record per product
        Dim AppPolSummary As New AppPolSummary
        AppPolSummary.FirstName = ed.Persons(0).FirstName
        AppPolSummary.MiddleName = ed.Persons(0).MiddleName
        AppPolSummary.LastName = ed.Persons(0).LastName

        AppPolSummary.ProductID = ed.Persons(0).Benefits(0).ProductID
        AppPolSummary.CoverageBeginDate = ed.Persons(0).Benefits(0).CoverageBeginDate
        AppPolSummary.MonthlyPremium = ed.Persons(0).Benefits(0).TotalBenefitPremiumEmployee.Annual / 12
        AppPolSummary.PayPeriodPremium = ed.Persons(0).Benefits(0).TotalBenefitPremiumEmployee.Annual / UnumSession.PayFrequencyFactor

        For c = 0 To ed.Persons(0).Benefits(0).Coverages.GetUpperBound(0)
            If ed.Persons(0).Benefits(0).ProductID = ProductID Then
                AppPolSummary.CoverageAmount += ed.Persons(0).Benefits(0).Coverages(c).CoverageAmount
            End If
        Next

        If AppPolSummary.CoverageAmount > 0 Then
            AppPolSummary.CoverageAmountSpecified = True
        End If
        ACESSave.AppPolSummarySave(ed, AppPolSummary, DBNull.Value, UnumSession)

    End Sub

    'Private Sub HandleMultiple(ByRef ed As EnrollmentData, ByVal ProductID As String)
    '    Dim p, b, c As Integer

    '    ' Insert into AppPolsSummary one record per product
    '    Dim AppPolSummary As New AppPolSummary

    '    ' Cycle through the Persons array
    '    If ed.Persons.GetUpperBound(0) > -1 Then
    '        For p = 0 To ed.Persons.GetUpperBound(0)

    '            If ed.Persons(p).SSN = UnumSession.EmpSSN Then
    '                AppPolSummary.FirstName = ed.Persons(p).FirstName
    '                AppPolSummary.MiddleName = ed.Persons(p).MiddleName
    '                AppPolSummary.LastName = ed.Persons(p).LastName
    '            End If

    '            ' Cycle through the Benefits array
    '            If ed.Persons(p).Benefits.GetUpperBound(0) > -1 Then
    '                For b = 0 To ed.Persons(p).Benefits.GetUpperBound(0)

    '                    If ed.Persons(p).Benefits(b).ProductID = ProductID AndAlso ed.Persons(p).SSN = UnumSession.EmpSSN Then
    '                        AppPolSummary.ProductID = ed.Persons(p).Benefits(b).ProductID
    '                        AppPolSummary.CoverageBeginDate = ed.Persons(p).Benefits(b).CoverageBeginDate
    '                        AppPolSummary.MonthlyPremium = ed.Persons(p).Benefits(b).TotalBenefitPremiumEmployee.Annual / 12
    '                        AppPolSummary.PayPeriodPremium = ed.Persons(p).Benefits(b).TotalBenefitPremiumEmployee.Annual / UnumSession.PayFrequencyFactor
    '                    End If

    '                    ' __ Cycle through the Coverages array
    '                    If ed.Persons(p).Benefits(b).Coverages.GetUpperBound(0) > -1 Then
    '                        For c = 0 To ed.Persons(p).Benefits(b).Coverages.GetUpperBound(0)

    '                            'Dim AppPolData As New AppPolData
    '                            'AppPolData.Description = AppPolSummary.ProductID & "|" & ed.Persons(p).Benefits(b).Coverages(c).Description
    '                            'AppPolSummary.CoverageBeginDate = ed.Persons(p).Benefits(b).CoverageBeginDate
    '                            'AppPolData.MonthlyPremium = ed.Persons(p).Benefits(b).Coverages(c).Premium.Annual / 12
    '                            'AppPolData.PayPeriodPremium = ed.Persons(p).Benefits(b).Coverages(c).Premium.Annual / UnumSession.PayFrequencyFactor
    '                            'AppPolData.CoverageAmount = ed.Persons(p).Benefits(b).Coverages(c).CoverageAmount
    '                            'AppPolData.CoverageAmountSpecified = ed.Persons(p).Benefits(b).Coverages(c).CoverageAmountSpecified


    '                            If ed.Persons(p).Benefits(b).ProductID = ProductID AndAlso ed.Persons(p).SSN = UnumSession.EmpSSN Then
    '                                AppPolSummary.CoverageAmount += ed.Persons(p).Benefits(b).Coverages(c).CoverageAmount
    '                            End If

    '                            'acessave.AppPolDataSave(ed, AppPolSummary, DBNull.Value)

    '                        Next  ' Coverages
    '                    End If

    '                    If AppPolSummary.CoverageAmount > 0 Then
    '                        AppPolSummary.CoverageAmountSpecified = True
    '                    End If

    '                    If ed.Persons(p).SSN = UnumSession.EmpSSN Then
    '                        ACESSave.AppPolSummarySave(ed, AppPolSummary, DBNull.Value, UnumSession)
    '                    End If

    '                Next  ' Benefits
    '            End If

    '        Next   ' Persons

    '    End If
    'End Sub

End Class