Imports System.Data.SqlClient
Imports UnumBridge.biz.plane.www
Imports System.Data

Public Class Results
    Public Success As Boolean
    Public Msg As String
    Public ObtainConfirm As Boolean
End Class

Public Class Common
    Public Function GetDTSql(ByRef dt As DataTable) As DataTable
        Dim Row As Integer
        Dim Col As Integer
        Dim dt2 As New DataTable
        Dim dr2 As DataRow
        For Col = 0 To dt.Columns.Count - 1
            dt2.Columns.Add(dt.Columns(Col).ColumnName, GetType(DataTypeSQL))
        Next

        For Row = 0 To dt.Rows.Count - 1
            dr2 = dt2.NewRow
            For Col = 0 To dt.Columns.Count - 1
                Select Case dt.Columns(Col).DataType.ToString.ToLower
                    Case "system.guid"
                        Dim VarcharSql As New VarcharSQL(dt.Rows(Row)(Col).ToString)
                        Dim DataTypeSql As DataTypeSQL = VarcharSql
                        dr2(Col) = DataTypeSql
                    Case "system.string"
                        Dim VarcharSql As New VarcharSQL(dt.Rows(Row)(Col))
                        Dim DataTypeSql As DataTypeSQL = VarcharSql
                        dr2(Col) = DataTypeSql
                    Case "system.int64", "system.int32", "system.int16"
                        Dim IntSql As New IntSQL(dt.Rows(Row)(Col))
                        Dim DataTypeSql As DataTypeSQL = IntSql
                        dr2(Col) = DataTypeSql
                    Case "system.boolean"
                        Dim BitSql As New BitSQL(dt.Rows(Row)(Col))
                        Dim DataTypeSql As DataTypeSQL = BitSql
                        dr2(Col) = DataTypeSql
                    Case "system.datetime"
                        Dim DateTimeSql As New DateTimeSQL(dt.Rows(Row)(Col))
                        Dim DataTypeSql As DataTypeSQL = DateTimeSql
                        dr2(Col) = DataTypeSql
                    Case "system.double", "system.single"
                        Dim FloatSql As New FloatSQL(dt.Rows(Row)(Col))
                        Dim DataTypeSql As DataTypeSQL = FloatSql
                        dr2(Col) = DataTypeSql
                    Case "system.decimal"
                        Dim MoneySql As New MoneySQL(dt.Rows(Row)(Col))
                        Dim DataTypeSql As DataTypeSQL = MoneySql
                        dr2(Col) = DataTypeSql
                    Case Else
                        Stop
                End Select
            Next
            dt2.Rows.Add(dr2)
        Next
        Return dt2
    End Function

    Public Function GetServerDateTime() As DateTime
        Return Date.Now.ToUniversalTime.AddHours(-5)
        'Dim CmdAsst As New CmdAsst(CommandType.StoredProcedure, "ServerDateTime")
        'Dim QueryPack As CmdAsst.QueryPack
        'Try
        '    QueryPack = CmdAsst.Execute
        'Catch
        '    QueryPack.Success = False
        'End Try
        'Return QueryPack.dt.rows(0)("ServerDateTime")
    End Function

#Region " Helpers "
    Public Function xGetUnumClass() As Integer
        '   case  
        'when payfrequencycode in ('9') then '1' --warehouse weekly 
        'when payfrequencycode in ('8' then '2' --hourly weekly 
        'when payfrequencycode in ('11','12') then '3' --Salaried semimonthly
        'when payfrequencycode in ('4') then '4' --Monthly Ruddick
        'else '3' --default value just in case
        '   End
    End Function
    Public Function GetUnumClass(ByVal PayFrequencyCode As String, ByVal UnumSession As UnumSession) As Integer
        If IsNumeric(PayFrequencyCode) Then
            Select Case PayFrequencyCode
                Case 9
                    UnumSession.PayFrequencyCode = "W"
                    UnumSession.PayFrequencyFactor = 52
                    Return 1
                Case 8
                    UnumSession.PayFrequencyCode = "W"
                    UnumSession.PayFrequencyFactor = 52
                    Return 2
                Case 11, 12
                    UnumSession.PayFrequencyCode = "S"
                    UnumSession.PayFrequencyFactor = 24
                    Return 3
                Case 4
                    UnumSession.PayFrequencyCode = "M"
                    UnumSession.PayFrequencyFactor = 12
                    Return 4
                Case Else
                    UnumSession.PayFrequencyCode = "S"
                    UnumSession.PayFrequencyFactor = 24
                    Return 3
            End Select
        Else
            UnumSession.PayFrequencyCode = "S"
            UnumSession.PayFrequencyFactor = 24
            Return 3
        End If
    End Function

    Public Function DoTrim(ByVal Input As String, ByVal Length As Integer) As String
        If Input.Length <= Length Then
            Return Input
        Else
            Return Input.Substring(0, Length)
        End If
    End Function
    Public Function GetFullZipCode(ByVal Zip5 As String, ByVal Zip4 As String) As String
        Zip5 = Trim(Zip5)
        Zip4 = Trim(Zip4)
        If Zip4.Length > 0 Then
            Return DoTrim(Zip5, 5) & "-" & DoTrim(Zip4, 4)
        Else
            Return DoTrim(Zip5, 5)
        End If
    End Function
    Public Sub SetDateValues(ByVal Input As String, ByRef DateFld As DateTime, ByRef SpecifiedFld As Boolean)
        If Input = String.Empty Then
            DateFld = GetServerDateTime.ToString("yyyy-MM-ddTHH:mm:ss")
            SpecifiedFld = False
        ElseIf IsDate(Input) Then
            DateFld = DateTime.Parse(Input)
            SpecifiedFld = True
        Else
            DateFld = GetServerDateTime.ToString("yyyy-MM-ddTHH:mm:ss")
            SpecifiedFld = False
        End If
    End Sub

    Public Sub SetGenderValues(ByVal Input As String, ByRef GenderFld As EmployeeGender, ByRef SpecifiedFld As Boolean)
        Select Case Input.ToLower
            Case "m"
                GenderFld = EmployeeGender.Male
                SpecifiedFld = True
            Case "f"
                GenderFld = EmployeeGender.Female
                SpecifiedFld = True
            Case "u"
                GenderFld = EmployeeGender.Female
                SpecifiedFld = False
            Case Else
                GenderFld = EmployeeGender.Female
                SpecifiedFld = False
        End Select
    End Sub

    Public Sub SetMaritalStatusValues(ByVal Input As String, ByRef MaritalStatusFld As EmployeeMaritalStatus, ByRef SpecifiedFld As Boolean)
        Select Case Input.ToLower
            Case String.Empty, "h", "u"
                MaritalStatusFld = EmployeeMaritalStatus.Unknown
                SpecifiedFld = False
            Case "d"
                MaritalStatusFld = EmployeeMaritalStatus.Divorced
                SpecifiedFld = True
            Case "m"
                MaritalStatusFld = EmployeeMaritalStatus.Married
                SpecifiedFld = True
            Case "s"
                MaritalStatusFld = EmployeeMaritalStatus.Single
                SpecifiedFld = True
            Case "w"
                MaritalStatusFld = EmployeeMaritalStatus.Widow
                SpecifiedFld = True
            Case Else
                MaritalStatusFld = EmployeeMaritalStatus.Unknown
                SpecifiedFld = False
        End Select
    End Sub

    Public Sub SetPayTypeValues(ByVal Input As String, ByRef PayTypeFld As PayTypeList, ByRef SpecifiedFld As Boolean)
        Select Case Input.ToLower
            Case String.Empty
                PayTypeFld = PayTypeList.Unknown
                SpecifiedFld = False
            Case "f"
                PayTypeFld = PayTypeList.Salary
                SpecifiedFld = True
            Case "h"
                PayTypeFld = PayTypeList.Hourly
                SpecifiedFld = True
            Case "p"
                PayTypeFld = PayTypeList.Unknown
                SpecifiedFld = False
            Case "s"
                PayTypeFld = PayTypeList.Salary
                SpecifiedFld = True
            Case Else
                PayTypeFld = PayTypeList.Unknown
                SpecifiedFld = False
        End Select
    End Sub

    Public Sub SetHoursWorkedPerWeekValues(ByVal Input As String, ByRef HoursFld As Integer, ByRef SpecifiedFld As Boolean)
        Select Case Input.ToLower
            Case String.Empty
                HoursFld = 0
                SpecifiedFld = False
            Case "f"
                HoursFld = 40
                SpecifiedFld = True
            Case "h"
                HoursFld = 40
                SpecifiedFld = True
            Case "p"
                HoursFld = 20
                SpecifiedFld = True
            Case "s"
                HoursFld = 40
                SpecifiedFld = True
            Case Else
                HoursFld = 0
                SpecifiedFld = False
        End Select
    End Sub

    'Private Sub SetSalaryValues(ByVal BenefitBaseSalaryInput As Decimal, ByVal PayFrequencyCode As String, ByRef SalaryFld As Decimal, ByRef SpecifiedFld As Boolean)
    '    Dim PayFrequencyCodeInt As Integer
    '    Try
    '        PayFrequencyCodeInt = CType(PayFrequencyCode, System.Int64)
    '        Select Case PayFrequencyCodeInt
    '            Case 6          'daily
    '                SalaryFld = BenefitBaseSalaryInput * 2080
    '                SpecifiedFld = True
    '            Case 1, 8, 9    'weekly
    '                SalaryFld = BenefitBaseSalaryInput * 52.14286
    '                SpecifiedFld = True
    '            Case 2          'biweekly
    '                SalaryFld = BenefitBaseSalaryInput * 26.07143
    '                SpecifiedFld = True
    '            Case 10, 11, 3  'semimonthly
    '                SalaryFld = BenefitBaseSalaryInput * 24
    '                SpecifiedFld = True
    '            Case 4          'monthly
    '                SalaryFld = BenefitBaseSalaryInput * 12
    '                SpecifiedFld = True
    '            Case 7          'annual
    '                SalaryFld = BenefitBaseSalaryInput
    '                SpecifiedFld = True
    '            Case Else
    '                SalaryFld = 0
    '                SpecifiedFld = False
    '        End Select
    '    Catch ex As Exception
    '        Select Case PayFrequencyCode.ToLower
    '            Case "Q"        'quarterly
    '                SalaryFld = BenefitBaseSalaryInput * 4
    '                SpecifiedFld = True
    '            Case "S"        'semiannual
    '                SalaryFld = BenefitBaseSalaryInput * 2
    '                SpecifiedFld = True
    '            Case Else
    '                SalaryFld = BenefitBaseSalaryInput
    '                SpecifiedFld = False
    '        End Select
    '    End Try
    'End Sub

    Public Function GetFormattedSSN(ByVal Value As String) As String
        Dim Results As String = String.Empty
        If Value.Length > 0 Then
            Results = Value.Substring(0, 3) & "-" & Value.Substring(3, 2) & "-" & Value.Substring(5, 4)
        End If
        Return Results
    End Function

    Public Sub SetStateValues(ByVal Input As String, ByRef StateFld As EmployeeStateList, ByRef SpecifiedFld As Boolean)
        SpecifiedFld = True
        Select Case Input.ToUpper
            Case "AK"
                StateFld = EmployeeStateList.AK
            Case "AL"
                StateFld = EmployeeStateList.AL
            Case "AR"
                StateFld = EmployeeStateList.AR
            Case "AZ"
                StateFld = EmployeeStateList.AZ
            Case "BC"
                StateFld = EmployeeStateList.BC
            Case "CA"
                StateFld = EmployeeStateList.CA
            Case "CO"
                StateFld = EmployeeStateList.CO
            Case "CT"
                StateFld = EmployeeStateList.CT
            Case "DC"
                StateFld = EmployeeStateList.DC
            Case "DE"
                StateFld = EmployeeStateList.DE
            Case "FL"
                StateFld = EmployeeStateList.FL

            Case "GA"
                StateFld = EmployeeStateList.GA
            Case "GU"
                StateFld = EmployeeStateList.GU
            Case "HI"
                StateFld = EmployeeStateList.HI
            Case "IA"
                StateFld = EmployeeStateList.IA
            Case "ID"
                StateFld = EmployeeStateList.ID
            Case "IL"
                StateFld = EmployeeStateList.IL
            Case "IN"
                StateFld = EmployeeStateList.IN
            Case "KS"
                StateFld = EmployeeStateList.KS
            Case "KY"
                StateFld = EmployeeStateList.KY
            Case "LA"
                StateFld = EmployeeStateList.LA

            Case "MA"
                StateFld = EmployeeStateList.MA
            Case "MD"
                StateFld = EmployeeStateList.MD
            Case "ME"
                StateFld = EmployeeStateList.ME
            Case "MI"
                StateFld = EmployeeStateList.MI
            Case "MN"
                StateFld = EmployeeStateList.MN
            Case "MO"
                StateFld = EmployeeStateList.MO
            Case "MS"
                StateFld = EmployeeStateList.MS
            Case "MT"
                StateFld = EmployeeStateList.MT
            Case "NA"
                StateFld = EmployeeStateList.NA
            Case "NC"
                StateFld = EmployeeStateList.NC

            Case "ND"
                StateFld = EmployeeStateList.ND
            Case "NE"
                StateFld = EmployeeStateList.NE
            Case "NH"
                StateFld = EmployeeStateList.NH
            Case "NJ"
                StateFld = EmployeeStateList.NJ
            Case "NM"
                StateFld = EmployeeStateList.NM
            Case "NV"
                StateFld = EmployeeStateList.NV
            Case "NY"
                StateFld = EmployeeStateList.NY
            Case "OH"
                StateFld = EmployeeStateList.OH
            Case "OK"
                StateFld = EmployeeStateList.OK
            Case "ON"
                StateFld = EmployeeStateList.ON


            Case "OR"
                StateFld = EmployeeStateList.OR
            Case "PA"
                StateFld = EmployeeStateList.PA
            Case "PR"
                StateFld = EmployeeStateList.PR
            Case "RI"
                StateFld = EmployeeStateList.RI
            Case "SC"
                StateFld = EmployeeStateList.SC
            Case "SD"
                StateFld = EmployeeStateList.SD
            Case "TN"
                StateFld = EmployeeStateList.TN
            Case "TX"
                StateFld = EmployeeStateList.TX
            Case "UT"
                StateFld = EmployeeStateList.UT
            Case "VA"
                StateFld = EmployeeStateList.VA

            Case "VI"
                StateFld = EmployeeStateList.VI
            Case "VT"
                StateFld = EmployeeStateList.VT
            Case "WA"
                StateFld = EmployeeStateList.WA
            Case "WI"
                StateFld = EmployeeStateList.WI
            Case "WV"
                StateFld = EmployeeStateList.WV
            Case "WY"
                StateFld = EmployeeStateList.WY
            Case Else
                StateFld = EmployeeStateList.AK
                SpecifiedFld = False
        End Select
    End Sub

    Public Sub SetDependentStatus(ByVal HandicappedInd As String, ByVal StudentInd As String, ByRef DependentStatus As DependentStatusList, ByRef SpecifiedFld As Boolean)
        DependentStatus = DependentStatusList.NA

        If HandicappedInd.ToLower = "y" Then
            If StudentInd.ToLower = "y" Then
                DependentStatus = DependentStatusList.HandicapStudent
            Else
                DependentStatus = DependentStatus.Handicap
            End If
        Else
            If StudentInd.ToLower = "y" Then
                DependentStatus = DependentStatusList.Student
            Else
                DependentStatus = DependentStatusList.NA
            End If
        End If

        'If HandicappedInd.ToLower = "n" Then
        '    If StudentInd = 0 Then
        '        DependentStatus = DependentStatusList.NA
        '    Else
        '        DependentStatus = DependentStatusList.Student
        '    End If
        'Else
        '    If StudentInd = 0 Then
        '        DependentStatus = DependentStatus.Handicap
        '    Else
        '        DependentStatus = DependentStatusList.HandicapStudent
        '    End If
        'End If
    End Sub



    ' Rose: Add UnumTranslation column for relationcodes.
    ' Rose: Look into doing this with the other code tables.
    Public Sub SetRelationTypeValues(ByVal BRelationType As String, ByVal SexCode As String, ByRef RelationType As DependentRelationType, ByRef SpecifiedFld As Boolean)
        Dim BRelationTypeInt As Integer
        If IsNumeric(BRelationType) Then
            BRelationTypeInt = CType(BRelationType, System.Int64)
        Else
            RelationType = DependentRelationType.Others
            SpecifiedFld = False
            Return
        End If

        SpecifiedFld = True
        Select Case BRelationTypeInt
            Case 1
                RelationType = DependentRelationType.Spouse
            Case 3
                RelationType = DependentRelationType.Parent
            Case 4
                Select Case SexCode
                    Case "F"
                        RelationType = DependentRelationType.GrandMother
                    Case "M"
                        RelationType = DependentRelationType.GrandFather
                    Case "U"
                        RelationType = DependentRelationType.Others
                End Select
            Case 5
                RelationType = DependentRelationType.GrandChild
            Case 6
                Select Case SexCode
                    Case "F"
                        RelationType = DependentRelationType.Aunt
                    Case "M"
                        RelationType = DependentRelationType.Uncle
                    Case "U"
                        RelationType = DependentRelationType.Others
                End Select
            Case 7
                Select Case SexCode
                    Case "F"
                        RelationType = DependentRelationType.Niece
                    Case "M"
                        RelationType = DependentRelationType.Nephew
                    Case "U"
                        RelationType = DependentRelationType.Others
                End Select
            Case 8
                RelationType = DependentRelationType.Cousin
            Case 9
                RelationType = DependentRelationType.AdoptedChild
            Case 10
                RelationType = DependentRelationType.FosterChild
            Case 11
                RelationType = DependentRelationType.Others
            Case 12
                Select Case SexCode
                    Case "F"
                        RelationType = DependentRelationType.Soninlaw
                    Case "M"
                        RelationType = DependentRelationType.Brotherinlaw
                    Case "U"
                        RelationType = DependentRelationType.Others
                End Select
            Case 13
                RelationType = DependentRelationType.Others
                SpecifiedFld = False
            Case 14
                Select Case SexCode
                    Case "F"
                        RelationType = DependentRelationType.Sister
                    Case "M"
                        RelationType = DependentRelationType.Brother
                    Case "U"
                        RelationType = DependentRelationType.Others
                End Select
            Case 15
                RelationType = DependentRelationType.Ward
            Case 17
                RelationType = DependentRelationType.StepChild
            Case 18
                RelationType = DependentRelationType.Self
            Case 19
                RelationType = DependentRelationType.Child
            Case 38
                RelationType = DependentRelationType.CollateralDependent
            Case 53
                RelationType = DependentRelationType.DomesticPartner
            Case 94
                RelationType = DependentRelationType.Others
                SpecifiedFld = False
            Case 95
                RelationType = DependentRelationType.Friends
            Case 96
                RelationType = DependentRelationType.Guardian
            Case 97
                RelationType = DependentRelationType.Others
                SpecifiedFld = False
            Case 98
                RelationType = DependentRelationType.Others
                SpecifiedFld = False
            Case 99
                RelationType = DependentRelationType.Others
            Case Else
                SpecifiedFld = False
        End Select
    End Sub
#End Region
End Class

Public Class KeyValueObj
    Private cKey As String
    Private cValue As String

    Public Property Key()
        Get
            Return cKey
        End Get
        Set(ByVal Value)
            cKey = Value
        End Set
    End Property
    Public Property Value()
        Get
            Return cValue
        End Get
        Set(ByVal Value)
            cValue = Value
        End Set
    End Property
    Public Sub New(ByVal Key As String, ByVal Value As String)
        cKey = Key
        cValue = Value
    End Sub
End Class
