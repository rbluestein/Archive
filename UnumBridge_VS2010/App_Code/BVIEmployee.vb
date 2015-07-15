Imports UnumBridge.biz.plane.www
Imports System.Data

Public Class BVIEmployee
    Private Common As New Common
    Private DBase As New DBase

    Public Function SetEmployeeValues(ByRef UnumSession As UnumSession) As Employee
        Dim dt As DataTable

        dt = DBase.GetConvertDTSql("SELECT empl.*, cpf.PayFrequencyDesc FROM Employee empl  LEFT JOIN " & UnumSession.DBName & "..Codes_PayFrequency cpf ON empl.PayFrequencyCode = cpf.PayFrequencyCode WHERE empl.EmpID = '" & UnumSession.EmpID & "'", UnumSession.ServerIPAddress, UnumSession.DBName)

        UnumSession.PayFrequencyDesc = dt.Rows(0)("PayFrequencyDesc").ToText

        Dim ee As New Employee                                                                  'Comments                                    '
        ee.Address1 = Common.DoTrim(dt.Rows(0)("AddrLine1").ToText, 50)
        ee.Address2 = Common.DoTrim(dt.Rows(0)("AddrLine2").ToText, 50)
        ee.AnnWorkDays = -1
        ee.AnnWorkDaysSpecified = False
        ee.BirthCountry = ""                                                                      'Birth country not sent
        ee.BirthState = EmployeeStateList.CA
        ee.BirthStateSpecified = False
        ee.Bonus = -1
        ee.BonusSpecified = False
        ee.City = Common.DoTrim(dt.Rows(0)("City").ToText, 50)
        ee.Class = Common.GetUnumClass(dt.Rows(0)("PayFrequencyCode").ToText, UnumSession)
        ' ee.Class = 1                                                                                   'Unum: 1=Fulltime. BVI: All entries are benefits eligible. Confirm this.     
        ' Rose: Generally 20+ hours / week qualifies for voluntary benefits. Check with Unum.
        ' Employee.StatusCode: If employee on leave, are they eligible.

        ee.ClassSpecified = True
        ee.Country = ""                                                                              'Country not sent
        ee.County = ""                                                                               'County not sent
        ee.DataChangeDate = Common.GetServerDateTime.ToString("yyyy-MM-ddTHH:mm:ss")                        'Send DateTime.Now??? Time zone??
        Common.SetDateValues(dt.Rows(0)("BirthDate").ToText, ee.DateOfBirth, ee.DateOfBirthSpecified)

        ' Common.SetDateValues(dt.Rows(0)("HireDate").ToText, ee.DateOfHire, ee.DateOfHireSpecified)
        If IsDate(dt.Rows(0)("HireDate").ToText) AndAlso IsDate(dt.Rows(0)("PromotionDate").ToText) Then
            If DateTime.Compare(dt.Rows(0)("PromotionDate").ToText, dt.Rows(0)("HireDate").ToText) >= 0 Then
                Common.SetDateValues(dt.Rows(0)("PromotionDate").ToText, ee.DateOfHire, ee.DateOfHireSpecified)
            Else
                Common.SetDateValues(dt.Rows(0)("HireDate").ToText, ee.DateOfHire, ee.DateOfHireSpecified)
            End If
        Else
            Common.SetDateValues(dt.Rows(0)("HireDate").ToText, ee.DateOfHire, ee.DateOfHireSpecified)
        End If

        ee.Department = Common.DoTrim(dt.Rows(0)("DeptNumber").ToText, 30)
        ee.Duties = ""
        ee.Email = Common.DoTrim(dt.Rows(0)("Email").ToText, 50)
        ee.EmployeeNumber = Common.DoTrim(dt.Rows(0)("EmpID").ToText, 25)
        ee.Fax = ""
        ee.FirstName = Common.DoTrim(dt.Rows(0)("FirstName").ToText, 25)
        Common.SetGenderValues(dt.Rows(0)("SexCode").ToText, ee.Gender, ee.GenderSpecified)
        ee.Height_FT = 0
        ee.Height_FTSpecified = False
        ee.Height_IN = 0
        ee.Height_INSpecified = False
        ee.HomePhone = Common.DoTrim(dt.Rows(0)("HomePhone").ToText, 25)
        Common.SetHoursWorkedPerWeekValues(dt.Rows(0)("TypeCode").ToText, ee.HoursWorkedPerWeek, ee.HoursWorkedPerWeekSpecified)
        ee.LastName = Common.DoTrim(dt.Rows(0)("LastName").ToText, 25)
        ee.LengthUSResidence = ""                                                            'LengthUSResidence not sent
        ee.Location = 0
        ee.LocationSpecified = True

        ee.LocationInfo = Common.DoTrim(dt.Rows(0)("Location").ToText, 50)            'LocationInfo not sent
        Common.SetMaritalStatusValues(dt.Rows(0)("MaritalStatusCode").ToText, ee.MaritalStatus, ee.MaritalStatusSpecified)
        ee.MiddleName = Common.DoTrim(dt.Rows(0)("MI").ToText, 25)
        ee.NameSuffix = Common.DoTrim(dt.Rows(0)("NameSuffix").ToText, 5)
        ee.Occupation = ""                                                                        'Occupation not sent
        Common.SetPayTypeValues(dt.Rows(0)("TypeCode").ToText, ee.PayType, ee.PayTypeSpecified)
        ee.PinCode = ""                                                                               'PinCode not sent"

        'ee.SalarySpecified = True
        If IsDBNull(dt.Rows(0)("BenefitBaseSalary").Value) Then
            ee.Salary = 0
            ee.SalarySpecified = False
        Else
            ee.Salary = dt.Rows(0)("BenefitBaseSalary").ToText
            ee.SalarySpecified = True
        End If

        ' SetSalaryValues(dt.Rows(0)("BenefitBaseSalary").ToText, dt.Rows(0)("PayFrequencyCode").ToText, ee.Salary, ee.SalarySpecified)
        ee.Shift = ""          'Shift not sent

        ee.SSN = Common.GetFormattedSSN(dt.Rows(0)("SSN").ToText)
        UnumSession.EmpSSN = ee.SSN

        ''If UnumSession.UseRealSSN Then
        'ee.SSN = Common.GetFormattedSSN(dt.Rows(0)("SSN").ToText)
        ''Else
        ''    ee.SSN = "999-99-2345"
        ''End If
        ''  UnumSession.EmpSSN = Common.GetFormattedSSN(dt.Rows(0)("SSN").ToText)
        'UnumSession.EmpSSN = ee.SSN


        Common.SetStateValues(dt.Rows(0)("State").ToText, ee.State, ee.StateSpecified)
        ee.TaxRate = -1
        ee.TaxRateSpecified = False
        ee.Title = Common.DoTrim(dt.Rows(0)("JobTitle").ToText, 10)

        If IsDBNull(dt.Rows(0)("FedTaxExemptions").Value) Then
            ee.W4Expemt = -1
            ee.W4ExpemtSpecified = False
        Else
            ee.W4Expemt = dt.Rows(0)("FedTaxExemptions").ToText
            ee.W4ExpemtSpecified = True
        End If

        ee.Weight = -1
        ee.WeightSpecified = False
        ee.WorkPhone = Common.DoTrim(dt.Rows(0)("WorkPhone").ToText, 25)
        ee.Zip = Common.GetFullZipCode(dt.Rows(0)("Zip5").ToText, dt.Rows(0)("Zip4").ToText)

        Return ee
    End Function



    Public Function SetTestEmployeeValues() As Employee
        Dim ee As New Employee
        ee.Address1 = "1492 Maple St."
        ee.Address2 = ""
        ee.AnnWorkDays = -1
        ee.AnnWorkDaysSpecified = False
        ee.AnnWorkDaysSpecified = False
        ee.BirthCountry = "United States"
        ee.BirthState = EmployeeStateList.CA
        ee.BirthStateSpecified = False
        ee.Bonus = -1
        ee.BonusSpecified = False
        ee.City = "Oakland"
        ee.Class = 1
        ee.ClassSpecified = True
        ee.Country = "USA"
        ee.County = "Alameda"
        ee.DataChangeDate = Common.GetServerDateTime.ToString("yyyy-MM-ddTHH:mm:ss")
        ee.DateOfBirth = DateTime.Parse("12/1/1970")
        ee.DateOfBirthSpecified = True
        ee.DateOfHire = DateTime.Parse("12/1/2000")
        ee.DateOfHireSpecified = True
        ee.Department = "Cust Service"
        ee.Duties = ""
        ee.Email = "r@b.com"
        ee.EmployeeNumber = ""
        ee.Fax = ""
        ee.FirstName = "Benjamin"
        ee.Gender = EmployeeGender.Male
        ee.GenderSpecified = True
        ee.Height_FT = 0
        ee.Height_FTSpecified = False
        ee.Height_IN = 0
        ee.Height_INSpecified = False
        ee.HomePhone = "510.482.8989"
        ee.HoursWorkedPerWeek = 40
        ee.HoursWorkedPerWeekSpecified = True
        ee.LastName = "Vega"
        ee.LengthUSResidence = "20"
        ee.Location = 0
        ee.LocationSpecified = True
        ee.LocationInfo = ""
        ee.MaritalStatus = EmployeeMaritalStatus.Single
        ee.MaritalStatusSpecified = True
        ee.MiddleName = ""
        ee.NameSuffix = ""
        ee.Occupation = "Administrator"
        ee.PayType = PayTypeList.Unknown
        ee.PayTypeSpecified = False
        ee.PinCode = ""
        ee.Salary = -1
        ee.SalarySpecified = False
        ee.Shift = ""
        ee.SSN = "999-99-5001"
        ee.State = EmployeeStateList.CA
        ee.StateSpecified = True
        ee.TaxRate = -1
        ee.TaxRateSpecified = False
        ee.Title = ""
        ee.W4Expemt = -1
        ee.W4ExpemtSpecified = False
        ee.Weight = -1
        ee.WeightSpecified = False
        ee.WorkPhone = "510.482.8888"
        ee.Zip = "94602"
        Return ee
    End Function
End Class
