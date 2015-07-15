Imports UnumBridge.biz.plane.www
Imports System.Data

Public Class ACESSave
    Private UnumSession As UnumSession
    Private DBase As New DBase
    Private Common As New Common

    Public Sub RecordEmpAndDepValuesReceived(ByRef ed As EnrollmentData, ByRef UnumSession As UnumSession)
        Dim p, b, c, ec As Integer
        Dim sb As New System.Text.StringBuilder
        Dim FullPath As String

        sb.Append("<?xml version=""1.0"" encoding=""UTF-8"" ?> ")
        sb.Append(" <EnrollmentData>")

        sb.Append("<TotalPremium")
        sb.Append(" TotalPremiumEmployee_Annual=""" & ed.TotalPremiumEmployee.Annual & """")
        sb.Append(" TotalPremiumEmployee_PayPeriod=""" & ed.TotalPremiumEmployee.PayPeriod & """")
        sb.Append(" TotalPremiumEmployer_Annual=""" & ed.TotalPremiumEmployer.Annual & """")
        sb.Append(" TotalPremiumEmployer_PayPeriod=""" & ed.TotalPremiumEmployer.PayPeriod & """")
        sb.Append(" />")

        ' ___ Existing coverage
        If ed.ExistingCoverage.GetUpperBound(0) > -1 Then
            sb.Append(" <ExistingCoverage>")
            For ec = 0 To ed.ExistingCoverage.GetUpperBound(0)
                sb.Append("<ExistingCoverage_" & ec.ToString)
                sb.Append(" InsuredName=""" & ed.ExistingCoverage(ec).InsuredName & """")
                sb.Append(" Premium_Annual=""" & ed.ExistingCoverage(ec).Premium.Annual & """")
                sb.Append(" Premium_PayPeriod=""" & ed.ExistingCoverage(ec).Premium.PayPeriod & """")
                ' sb.Append(" ProductDescription=""" & ed.ExistingCoverage(ec).ProductDescription & """")

                ' ___ Convert illegal xml character
                sb.Append(" ProductDescription=""" & Replace(ed.ExistingCoverage(ec).ProductDescription, "&", " and") & """")

                sb.Append(" ProductID=""" & ed.ExistingCoverage(ec).ProductID & """")
                sb.Append(" />")
            Next
            sb.Append("</ExistingCoverage>")
        End If

        '  ___Persons
        If ed.Persons.GetUpperBound(0) > -1 Then
            sb.Append("<Persons>")
            For p = 0 To ed.Persons.GetUpperBound(0)
                'For p = 0 To 0
                sb.Append("<Person_" & p.ToString & ">")
                sb.Append(" FirstName=""" & ed.Persons(p).FirstName & """")
                sb.Append(" MiddleName=""" & ed.Persons(p).MiddleName & """")
                sb.Append(" LastName=""" & ed.Persons(p).LastName & """")
                sb.Append(" Gender=""" & ed.Persons(p).Gender.ToString & """")

                sb.Append(" SSN=""" & ed.Persons(p).SSN & """")
                sb.Append(" MaritalStatus=""" & ed.Persons(p).MaritalStatus.ToString & """")
                sb.Append(" Relationship=""" & ed.Persons(p).Relationship.ToString & """")
                sb.Append(" Occupation=""" & ed.Persons(p).Occupation & """")

                sb.Append(" State=""" & ed.Persons(p).State & """")
                sb.Append(" Zip=""" & ed.Persons(p).Zip & """")
                sb.Append(" Title=""" & ed.Persons(p).Title & """")
                sb.Append(" HomePhone=""" & ed.Persons(p).HomePhone & """")
                sb.Append(" WorkPhone=""" & ed.Persons(p).WorkPhone & """")


                ' ___ Benefits
                If ed.Persons(p).Benefits.GetUpperBound(0) > -1 Then
                    sb.Append("<Benefits>")
                    For b = 0 To ed.Persons(p).Benefits.GetUpperBound(0)
                        sb.Append("<Benefit_" & b.ToString)
                        sb.Append(" Description=""" & ed.Persons(p).Benefits(b).Description & """")
                        sb.Append(" FirstDeductionDate=""" & ed.Persons(p).Benefits(b).FirstDeductionDate.ToString & """")
                        sb.Append(" ProductID=""" & ed.Persons(p).Benefits(b).ProductID & """")
                        sb.Append(" Status=""" & ed.Persons(p).Benefits(b).Status.ToString & """")
                        sb.Append(" >")

                        sb.Append("<Dates")
                        sb.Append(" CoverageBeginDate=""" & ed.Persons(p).Benefits(b).CoverageBeginDate.ToString & """")
                        sb.Append(" CoverageChangeDate=""" & ed.Persons(p).Benefits(b).CoverageChangeDate.ToString & """")
                        sb.Append(" CoverageEndDate=""" & ed.Persons(p).Benefits(b).CoverageEndDate.ToString & """")
                        sb.Append(" />")

                        sb.Append("<Premiums")
                        sb.Append(" AfterTaxBenefitPremiumEmployee_Annual=""" & ed.Persons(p).Benefits(b).AfterTaxBenefitPremiumEmployee.Annual & """")
                        sb.Append(" AfterTaxBenefitPremiumEmployee_PayPeriod=""" & ed.Persons(p).Benefits(b).AfterTaxBenefitPremiumEmployee.PayPeriod & """")
                        sb.Append(" PreTaxBenefitPremiumEmployee_Annual=""" & ed.Persons(p).Benefits(b).PreTaxBenefitPremiumEmployee.Annual & """")
                        sb.Append(" PreTaxBenefitPremiumEmployee_PayPeriod=""" & ed.Persons(p).Benefits(b).PreTaxBenefitPremiumEmployee.PayPeriod & """")
                        sb.Append(" TotalBenefitPremiumEmployee_Annual=""" & ed.Persons(p).Benefits(b).TotalBenefitPremiumEmployee.Annual & """")
                        sb.Append(" TotalBenefitPremiumEmployee_PayPeriod=""" & ed.Persons(p).Benefits(0).TotalBenefitPremiumEmployee.PayPeriod & """")
                        sb.Append(" TotalBenefitPremiumEmployer_Annual=""" & ed.Persons(p).Benefits(b).TotalBenefitPremiumEmployer.Annual & """")
                        sb.Append(" TotalBenefitPremiumEmployer_PayPeriod=""" & ed.Persons(p).Benefits(b).TotalBenefitPremiumEmployer.PayPeriod & """")
                        sb.Append("></Premiums>")

                        ' ___ Coverages
                        If ed.Persons(p).Benefits(b).Coverages.GetUpperBound(0) > -1 Then
                            sb.Append("<Coverages>")
                            For c = 0 To ed.Persons(p).Benefits(b).Coverages.GetUpperBound(0)
                                sb.Append("<Coverage_" & c.ToString)
                                sb.Append(" CoverageAmount=""" & ed.Persons(p).Benefits(b).Coverages(c).CoverageAmount & """")
                                sb.Append(" CoverageAmountSpecified=""" & ed.Persons(p).Benefits(b).Coverages(c).CoverageAmountSpecified & """")

                                ' ___ Convert illegal xml character
                                sb.Append(" Description=""" & Replace(ed.Persons(p).Benefits(b).Coverages(c).Description, "&", " and ") & """")

                                sb.Append(" Premium_Annual=""" & ed.Persons(p).Benefits(b).Coverages(c).Premium.Annual & """")
                                sb.Append(" Premium_PayPeriod=""" & ed.Persons(p).Benefits(b).Coverages(c).Premium.PayPeriod & """")
                                sb.Append(" Plan=""" & ed.Persons(p).Benefits(b).Coverages(c).Plan & """")
                                sb.Append(" Units=""" & ed.Persons(p).Benefits(b).Coverages(c).Units & """")
                                sb.Append(" UnitsSpecified=""" & ed.Persons(p).Benefits(b).Coverages(c).UnitsSpecified & """")
                                sb.Append(" >")
                                sb.Append("<EliminationBenefitPeriod ")
                                sb.Append(" AccidentEliminationPeriod=""" & ed.Persons(p).Benefits(b).Coverages(c).EliminationBenefitPeriod.AccidentEliminationPeriod & """")
                                sb.Append(" AccidentEliminationPeriodSpecified=""" & ed.Persons(p).Benefits(b).Coverages(c).EliminationBenefitPeriod.AccidentEliminationPeriodSpecified.ToString & """")
                                sb.Append(" BenefitPeriod=""" & ed.Persons(p).Benefits(b).Coverages(c).EliminationBenefitPeriod.BenefitPeriod & """")
                                sb.Append(" BenefitPeriodSpecified=""" & ed.Persons(p).Benefits(b).Coverages(c).EliminationBenefitPeriod.BenefitPeriodSpecified.ToString & """")
                                sb.Append(" SicknessEliminationPeriod=""" & ed.Persons(p).Benefits(b).Coverages(c).EliminationBenefitPeriod.SicknessEliminationPeriod & """")
                                sb.Append(" SicknessEliminationPeriodSpecified=""" & ed.Persons(p).Benefits(b).Coverages(c).EliminationBenefitPeriod.SicknessEliminationPeriodSpecified & """")
                                sb.Append(" />")
                                sb.Append("</Coverage_" & c.ToString & ">")
                            Next
                            sb.Append("</Coverages>")
                        End If

                        sb.Append("</Benefit_" & b.ToString & ">")

                    Next
                    sb.Append("</Benefits>")
                End If

                sb.Append(" </Person_" & p.ToString & ">")

            Next
            sb.Append("</Persons>")
        End If

        sb.Append("</EnrollmentData>")

        Try

            ' FullPath = "V:\prints\HarrisTeeter\NH\HarrisTeeterUnum" & ee.EmployeeNumber & DateTime.Now.ToString("yyyyMMddHHmmss") & ".txt"
            'FullPath = "\\hbg-web\eprog\prints\HarrisTeeter\NH\HarrisTeeterUnum" & ee.EmployeeNumber & DateTime.Now.ToString("yyyyMMddHHmmss") & ".txt"

            If System.Environment.MachineName.ToUpper = "LT-5ZFYRC1" Then
                ' FullPath = "C:\Inetpub\wwwroot\UnumBridge\prints\UnumReturn_" & Common.GetServerDateTime.ToString("yyyyMMdd_HHmmssfff") & ".xml"
            Else
                ' FullPath = "\\hbg-web\eprog\prints\HarrisTeeter\NH\HarrisTeeterUnum" & ee.EmployeeNumber & DateTime.Now.ToString("yyyyMMddHHmmss") & ".txt"
                '  FullPath = "\\hbg-web\eprog\prints\HarrisTeeter\NH\UnumDataReturned_" & UnumSession.EmpID & Common.GetServerDateTime.ToString("yyyyMMddHHmmss") & ".xml"
                ' FullPath = "\\hbg-web\eprog\prints\HarrisTeeter\NH\recdtest.xml"
                FullPath = "\\hbg-web\eprog\prints\HarrisTeeter\NH\UnumDataReturned_" & UnumSession.EmpID & Common.GetServerDateTime.ToString("yyyyMMddHHmmss") & ".xml"
            End If

            Dim StreamWriter As System.IO.StreamWriter
            StreamWriter = New System.IO.StreamWriter(FullPath)
            StreamWriter.Write(sb.ToString)
            StreamWriter.Close()

        Catch ex As Exception
            '   Throw New Exception("Unable to save employee/dependent data. " & ex.Message)
        End Try

    End Sub

    Public Sub RecordEmpAndDepValuesSent(ByRef ee As Employee, ByRef depdep As DependentsDependent())
        Dim FullPath As String
        Dim sb As New System.Text.StringBuilder

        Try

            ' FullPath = "V:\prints\HarrisTeeter\NH\HarrisTeeterUnum" & ee.EmployeeNumber & DateTime.Now.ToString("yyyyMMddHHmmss") & ".txt"
            FullPath = "\\hbg-web\eprog\prints\HarrisTeeter\NH\UnumDataSent_" & ee.EmployeeNumber & Common.GetServerDateTime.ToString("yyyyMMddHHmmss") & ".xml"
            ' EmpAndDepData = GetEmployeeString(ee) & GetDependentString(depdep)

            sb.Append("<?xml version=""1.0"" encoding=""UTF-8"" ?> ")
            sb.Append("<BVIDataPassedToUnum>")
            sb.Append(GetEmployeeString(ee))
            sb.Append(GetDependentString(depdep))
            sb.Append("</BVIDataPassedToUnum>")


            Dim StreamWriter As System.IO.StreamWriter
            StreamWriter = New System.IO.StreamWriter(FullPath)
            StreamWriter.Write(sb.ToString)
            StreamWriter.Close()

        Catch ex As Exception
            '   Throw New Exception("Unable to save employee/dependent data. " & ex.Message)
        End Try
    End Sub

    Private Function GetEmployeeString(ByRef ee As Employee) As String
        Dim sb As New System.Text.StringBuilder

        sb.Append(" <Employee ")
        sb.Append(" Address1=""" & ee.Address1 & """")
        sb.Append(" Address2=""" & ee.Address2 & """")
        sb.Append(" AnnWorkDays=""" & ee.AnnWorkDays & """")
        sb.Append(" AnnWorkDaysSpecified=""" & ee.AnnWorkDaysSpecified.ToString & """")
        sb.Append(" BirthCountry=""" & ee.BirthCountry & """")
        sb.Append(" BirthState=""" & ee.BirthState.ToString & """")
        sb.Append(" BirthStateSpecified=""" & ee.BirthStateSpecified.ToString & """")
        sb.Append(" Bonus=""" & ee.Bonus & """")
        sb.Append(" BonusSpecified=""" & ee.BonusSpecified.ToString & """")
        sb.Append(" City=""" & ee.City & """")
        sb.Append(" Class=""" & ee.Class & """")
        sb.Append(" ClassSpecified=""" & ee.ClassSpecified.ToString & """")

        sb.Append(" Country=""" & ee.Country & """")
        sb.Append(" County=""" & ee.County & """")
        sb.Append(" DataChangeDate=""" & ee.DataChangeDate.ToString & """")
        sb.Append(" BirthDate=""" & ee.DateOfBirth.ToString & """")
        sb.Append(" HireDate=""" & ee.DateOfHire.ToString & """")
        sb.Append(" Department=""" & ee.Department & """")
        sb.Append(" Duties=""" & ee.Duties & """")

        sb.Append(" Email=""" & ee.Email & """")
        sb.Append(" EmployeeNumber=""" & ee.EmployeeNumber & """")
        sb.Append(" Fax=""" & ee.Fax & """")
        sb.Append(" FirstName=""" & ee.FirstName & """")
        sb.Append(" Gender=""" & ee.Gender.ToString & """")
        sb.Append(" Height_FT=""" & ee.Height_FT & """")
        sb.Append(" Height_FTSpecified=""" & ee.Height_FTSpecified.ToString & """")
        sb.Append(" Height_IN=""" & ee.Height_IN & """")
        sb.Append(" Height_INSpecified=""" & ee.Height_INSpecified.ToString & """")

        sb.Append(" HomePhone=""" & ee.HomePhone & """")
        sb.Append(" HoursWorkedPerWeek=""" & ee.HoursWorkedPerWeek.ToString & """")
        sb.Append(" LastName=""" & ee.LastName & """")
        sb.Append(" LengthUSResidence=""" & ee.LengthUSResidence & """")
        sb.Append(" Location=""" & ee.Location & """")
        sb.Append(" LocationSpecified=""" & ee.LocationSpecified.ToString & """")

        sb.Append(" LocationInfo=""" & ee.LocationInfo & """")
        sb.Append(" MaritalStatus=""" & ee.MaritalStatus.ToString & """")
        sb.Append(" MaritalStatusSpecified=""" & ee.MaritalStatusSpecified & """")
        sb.Append(" MiddleName=""" & ee.MiddleName & """")
        sb.Append(" NameSuffix=""" & ee.NameSuffix & """")
        sb.Append(" Occupation=""" & ee.Occupation & """")
        sb.Append(" PayType=""" & ee.PayType & """")
        sb.Append(" PayTypeSpecified=""" & ee.PayTypeSpecified.ToString & """")
        sb.Append(" PinCode=""" & ee.PinCode & """")

        sb.Append(" Salary=""" & ee.Salary & """")
        sb.Append(" SalarySpecified=""" & ee.SalarySpecified.ToString & """")
        sb.Append(" Shift=""" & ee.Shift & """")
        sb.Append(" SSN=""" & ee.SSN & """")

        sb.Append(" State=""" & ee.State & """")
        sb.Append(" StateSpecified=""" & ee.StateSpecified.ToString & """")
        sb.Append(" TaxRate=""" & ee.TaxRate & """")
        sb.Append(" TaxRateSpecified=""" & ee.TaxRateSpecified.ToString & """")
        sb.Append(" Title=""" & ee.Title & """")

        sb.Append(" W4Exempt=""" & ee.W4Expemt & """")
        sb.Append(" W4ExemptSpecified=""" & ee.WeightSpecified.ToString & """")
        sb.Append(" Weight=""" & ee.Weight & """")
        sb.Append(" WeightSpecified=""" & ee.WeightSpecified.ToString & """")
        sb.Append(" WorkPhone=""" & ee.WorkPhone & """")
        sb.Append(" Zip=""" & ee.Zip & """")

        sb.Append(" />")

        Return sb.ToString
    End Function

    Public Function GetDependentString(ByRef depdep As DependentsDependent()) As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder

        sb.Append("<Dependents>")
        For i = 0 To depdep.GetUpperBound(0)
            sb.Append("<Dependent_" & i.ToString)
            sb.Append(" Address1=""" & depdep(i).Address1 & """")
            sb.Append(" Address2=""" & depdep(i).Address2 & """")
            sb.Append(" City=""" & depdep(i).City & """")
            sb.Append(" DateOfBirth=""" & depdep(i).DateOfBirth & """")
            sb.Append(" DateOfBirthSpecified=""" & depdep(i).DateOfBirthSpecified.ToString & """")

            sb.Append(" DependentStatus=""" & depdep(i).DependentStatus.ToString & """")
            sb.Append(" DependentStatusSpecified=""" & depdep(i).DependentStatusSpecified.ToString & """")
            sb.Append(" Email=""" & depdep(i).EMail & """")
            sb.Append(" FirstName=""" & depdep(i).FirstName & """")
            sb.Append(" Gender=""" & depdep(i).Gender.ToString & """")
            sb.Append(" GenderSpecified=""" & depdep(i).GenderSpecified.ToString & """")
            sb.Append(" HomePhone=""" & depdep(i).HomePhone & """")

            sb.Append(" LastName=""" & depdep(i).LastName & """")
            sb.Append(" MaritalStatus=""" & depdep(i).MaritalStatus.ToString & """")
            sb.Append(" MaritalStatusSpecified=""" & depdep(i).MaritalStatusSpecified.ToString & """")
            sb.Append(" MiddleName=""" & depdep(i).MiddleName & """")
            sb.Append(" Occupation=""" & depdep(i).Occupation & """")
            sb.Append(" RelationType=""" & depdep(i).RelationType.ToString & """")
            sb.Append(" RelationTypeSpecified=""" & depdep(i).RelationTypeSpecified.ToString & """")

            sb.Append(" SchoolName=""" & depdep(i).SchoolName & """")
            sb.Append(" SSN=""" & depdep(i).SSN & """")
            sb.Append(" State=""" & depdep(i).State & """")
            sb.Append(" StateSpecified=""" & depdep(i).StateSpecified.ToString & """")
            sb.Append(" Title=""" & depdep(i).Title & """")
            sb.Append(" WorkPhone=""" & depdep(i).WorkPhone & """")
            sb.Append(" Zip=""" & depdep(i).Zip & """")
            sb.Append(" />")
        Next
        sb.Append("</Dependents>")

        Return sb.ToString
    End Function

    Public Sub AppPolSummarySave(ByRef ed As EnrollmentData, ByRef AppPolSummary As AppPolSummary, ByVal AppID As Object, ByVal UnumSession As UnumSession)
        Dim sb As New System.Text.StringBuilder
        Dim Now As DateTime

        Now = Common.GetServerDateTime
        sb.Append("exec Usp_AppsAndPolsSummary ")
        sb.Append("@BVIAppStatus='NEW', ")
        sb.Append("@CompletedInd=1, ")
        sb.Append("@CompletedDate='" & Now.ToString & "', ")
        sb.Append("@TransmittedInd=1, ")
        sb.Append("@TransmittedDate='" & Now.ToString & "', ")
        sb.Append("@ClientID='" & UnumSession.ClientID & "', ")
        '  sb.Append("@ProductID='" & Common.DoTrim(AppPolSummary.ProductID, 20) & "', ")
        sb.Append("@ProductID='" & Common.DoTrim(UnumSession.BVIProductID, 20) & "', ")
        sb.Append("@AppDate='" & Now.ToString & "', ")
        sb.Append("@EffectiveDate='" & AppPolSummary.CoverageBeginDate & "', ")
        sb.Append("@EmpID='" & UnumSession.EmpID & "', ")
        sb.Append("@AppSSN='" & Common.DoTrim(Replace(UnumSession.EmpSSN, "-", ""), 9) & "', ")
        sb.Append("@InsuredFirstName='" & Common.DoTrim(AppPolSummary.FirstName, 20) & "', ")
        sb.Append("@InsuredLastName='" & Common.DoTrim(AppPolSummary.LastName, 30) & "', ")
        sb.Append("@MonthlyPremium=" & AppPolSummary.MonthlyPremium & ", ")
        sb.Append("@PerPayPremium=" & AppPolSummary.PayPeriodPremium & ", ")
        sb.Append("@PayFrequencyCode='" & UnumSession.PayFrequencyCode & "', ")
        sb.Append("@PlanBenefitAmt=" & AppPolSummary.CoverageAmount & ", ")
        sb.Append("@PlanBenefitAmtFreqCode='M', ")
        sb.Append("@CarrierAppNum=null, ")
        sb.Append("@LicensedEnroller='" & Common.DoTrim(UnumSession.EnrollerID, 50) & "', ")
        If UnumSession.LastLoginIP = Nothing Then
            UnumSession.LastLoginIP = String.Empty
        End If
        sb.Append("@EnrollerComputerName='" & Common.DoTrim(UnumSession.LastLoginIP, 50) & "', ")
        sb.Append("@ActivityID='" & UnumSession.ActivityID & "'")

        Dim CmdAsst As New CmdAsst(DBase, CommandType.Text, sb.ToString, UnumSession.ServerIPAddress, "IAMS")
        Dim QueryPack As CmdAsst.QueryPack
        QueryPack = CmdAsst.Execute

        If Not QueryPack.dt.rows(0)(0) = 0 Then
            Throw New Exception("Error #" & QueryPack.dt.rows(0)(0) & ": " & QueryPack.dt.rows(0)(1))
        End If

        If Not QueryPack.Success Then
            Throw New Exception(QueryPack.TechErrMsg)
        End If
    End Sub

End Class
