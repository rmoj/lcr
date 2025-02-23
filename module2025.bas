Attribute VB_Name = "Module2007"
Option Compare Database
Option Explicit
Public Function GetYear(ByVal code As Long) As Integer
    GetYear = code \ 100  '200701/100 = 2007
End Function

Public Function GetMonth(ByVal code As Long) As Byte
    GetMonth = CByte(code Mod 100)  '200701 mod 100 = 1
End Function

Function GetLastOfMonth(ByVal dteDate As Date) As Date
    ' This function calculates the last day of a month, given a date.
    ' Find the first day of the next month, then subtract one day.
    GetLastOfMonth = DateSerial(Year(dteDate), month(dteDate) + 1, 1) - 1
End Function

Public Function IsWorkday(ByVal dat As Date) As Boolean
    'Based on six-workday week, Sunday is not a workday
    If Weekday(dat) = vbSunday Then
        IsWorkday = False
    Else
        IsWorkday = True
    End If
End Function

Public Function GetWorkDaysInPayperiod(ByVal Year As Integer, ByVal month As Byte, ByVal half As Byte) As Byte
    Dim start, finish As Byte
    Dim d, w As Byte
    Dim dat As Date

    If half = 1 Then
        start = 1
        finish = 15
    Else
        start = 16
        finish = Day(GetLastOfMonth(DateSerial(Year, month, 1)))
    End If

    For d = start To finish
        dat = DateSerial(Year, month, d)
        If IsWorkday(dat) = True Then
            w = w + 1
        End If
    Next
    GetWorkDaysInPayperiod = w
End Function

Public Function GetHourlyRate(ByVal IsMonthly As Boolean, ByVal Salary As Currency) As Currency
    Dim hr As Currency  'currency type substitutes for decimal type

    Select Case IsMonthly
        Case True
            hr = Salary * 12 / 313 / 8 'hr = basic*12/313/8
        Case False
            hr = Salary / 8
    End Select

    GetHourlyRate = hr

End Function

Public Function GetBasic(ByVal Salary As Currency, ByVal workdaysinpayperiod As Byte, ByVal monthly As Boolean)
    If monthly = True Then
        GetBasic = Salary / 2
    Else
        GetBasic = Salary * workdaysinpayperiod
    End If
End Function


Public Function GetAllowances(ByVal daysWithAllowance As Byte, ByVal dailyallowance As Currency) As Currency
    GetAllowances = daysWithAllowance * dailyallowance
End Function

Public Function GetOT(ByVal hourlyrate As Currency, ByVal hrs11 As Currency, ByVal hrs125 As Currency, ByVal hrs1375 As Currency, ByVal hrs13 As Currency, ByVal hrs15 As Currency, ByVal hrs20 As Currency) As Currency
    GetOT = hourlyrate * (hrs11 * 1.1 + hrs125 * 1.25 + hrs1375 * 1.375 + hrs13 * 1.3 + hrs15 * 1.5 + hrs20 * 2)
End Function

Public Function GetLUA(ByVal hourlyrate As Currency, ByVal hrsLateUndertime As Currency, ByVal daysAbsent As Byte) As Currency
    GetLUA = hourlyrate * (hrsLateUndertime + daysAbsent * 8)
End Function

Public Function GetGross(ByVal Basic As Currency, ByVal Allowances As Currency, ByVal additions As Currency, ByVal OT As Currency, ByVal lua As Currency) As Currency
    GetGross = Basic + Allowances + additions + OT - lua
End Function

Public Function GetNontaxable(ByVal hourlyrate As Currency, ByVal hrs11 As Currency, ByVal hrs125 As Currency, ByVal hrs1375 As Currency, ByVal hrs13 As Currency, ByVal hrs15 As Currency, ByVal hrs20 As Currency, ByVal ntadditions As Currency) As Currency
    GetNontaxable = hourlyrate * (hrs11 * 1.1 + hrs125 * 1.25 + hrs1375 * 1.375 + hrs13 * 1.3 + hrs15 * 1.5 + hrs20 * 2) + ntadditions
End Function

Public Function GetDeductions(ByVal deduction1 As Currency, ByVal deduction2 As Currency, ByVal deduction3 As Currency, ByVal deduction4 As Currency, ByVal deduction5 As Currency, ByVal PhilhealthCont As Currency, ByVal PagibigCont As Currency, ByVal PagibigLoan As Currency, ByVal SSSCont As Currency, ByVal SSSLoan As Currency, ByVal Wtax As Currency, ByVal half As Byte) As Currency
If half = 1 Then
    GetDeductions = PhilhealthCont + PagibigCont + SSSLoan + deduction1 + deduction2 + deduction3 + deduction4 + deduction5
Else
    GetDeductions = SSSCont + Wtax + PagibigLoan + deduction1 + deduction2 + deduction3 + deduction4 + deduction5
End If
End Function

Public Function GetWTax(ByVal grossformonth As Currency, ByVal ExemptionStatus As String) As Currency
        'tax table 2009
        Dim g As Currency
        Dim w As Currency

        g = grossformonth

        Select Case ExemptionStatus
            Case "S"
                If g < 20833 Then
                    w = 0
                ElseIf g < 33332 Then
                    w = (g - 20833) * 0.15
                ElseIf g < 66666 Then
                    w = (g - 33333) * 0.2 + 1875
                ElseIf g < 166666 Then
                    w = (g - 66667) * 0.25 + 8541.8
                ElseIf g < 666666 Then
                    w = (g - 166667) * 0.3 + 33541.8

                Else
                    w = (g - 666667) * 0.35 + 200833.33
                End If



            Case "ME"
                If g < 20833 Then
                    w = 0
                ElseIf g < 33332 Then
                    w = (g - 20833) * 0.15
                ElseIf g < 66666 Then
                    w = (g - 33333) * 0.2 + 1875
                ElseIf g < 166666 Then
                    w = (g - 66667) * 0.25 + 8541.8
                ElseIf g < 666666 Then
                    w = (g - 166667) * 0.3 + 33541.8

                Else
                    w = (g - 666667) * 0.35 + 200833.33
                End If


            Case "S1"
                If g < 20833 Then
                    w = 0
                ElseIf g < 33332 Then
                    w = (g - 20833) * 0.15
                ElseIf g < 66666 Then
                    w = (g - 33333) * 0.2 + 1875
                ElseIf g < 166666 Then
                    w = (g - 66667) * 0.25 + 8541.8
                ElseIf g < 666666 Then
                    w = (g - 166667) * 0.3 + 33541.8

                Else
                    w = (g - 666667) * 0.35 + 200833.33
                End If


            Case "S2"
                If g < 20833 Then
                    w = 0
                ElseIf g < 33332 Then
                    w = (g - 20833) * 0.15
                ElseIf g < 66666 Then
                    w = (g - 33333) * 0.2 + 1875
                ElseIf g < 166666 Then
                    w = (g - 66667) * 0.25 + 8541.8
                ElseIf g < 666666 Then
                    w = (g - 166667) * 0.3 + 33541.8

                Else
                    w = (g - 666667) * 0.35 + 200833.33
                End If


            Case "S3"
                If g < 20833 Then
                    w = 0
                ElseIf g < 33332 Then
                    w = (g - 20833) * 0.15
                ElseIf g < 66666 Then
                    w = (g - 33333) * 0.2 + 1875
                ElseIf g < 166666 Then
                    w = (g - 66667) * 0.25 + 8541.8
                ElseIf g < 666666 Then
                    w = (g - 166667) * 0.3 + 33541.8

                Else
                    w = (g - 666667) * 0.35 + 200833.33
                End If

            Case "S4"
                If g < 20833 Then
                    w = 0
                ElseIf g < 33332 Then
                    w = (g - 20833) * 0.15
                ElseIf g < 66666 Then
                    w = (g - 33333) * 0.2 + 1875
                ElseIf g < 166666 Then
                    w = (g - 66667) * 0.25 + 8541.8
                ElseIf g < 666666 Then
                    w = (g - 166667) * 0.3 + 33541.8

                Else
                    w = (g - 666667) * 0.35 + 200833.33
                End If


            Case "ME1"
                If g < 20833 Then
                    w = 0
                ElseIf g < 33332 Then
                    w = (g - 20833) * 0.15
                ElseIf g < 66666 Then
                    w = (g - 33333) * 0.2 + 1875
                ElseIf g < 166666 Then
                    w = (g - 66667) * 0.25 + 8541.8
                ElseIf g < 666666 Then
                    w = (g - 166667) * 0.3 + 33541.8

                Else
                    w = (g - 666667) * 0.35 + 200833.33
                End If

            Case "ME2"
                If g < 20833 Then
                    w = 0
                ElseIf g < 33332 Then
                    w = (g - 20833) * 0.15
                ElseIf g < 66666 Then
                    w = (g - 33333) * 0.2 + 1875
                ElseIf g < 166666 Then
                    w = (g - 66667) * 0.25 + 8541.8
                ElseIf g < 666666 Then
                    w = (g - 166667) * 0.3 + 33541.8

                Else
                    w = (g - 666667) * 0.35 + 200833.33
                End If


            Case "ME3"
                If g < 20833 Then
                    w = 0
                ElseIf g < 33332 Then
                    w = (g - 20833) * 0.15
                ElseIf g < 66666 Then
                    w = (g - 33333) * 0.2 + 1875
                ElseIf g < 166666 Then
                    w = (g - 66667) * 0.25 + 8541.8
                ElseIf g < 666666 Then
                    w = (g - 166667) * 0.3 + 33541.8

                Else
                    w = (g - 666667) * 0.35 + 200833.33
                End If


            Case "ME4"
                If g < 20833 Then
                    w = 0
                ElseIf g < 33332 Then
                    w = (g - 20833) * 0.15
                ElseIf g < 66666 Then
                    w = (g - 33333) * 0.2 + 1875
                ElseIf g < 166666 Then
                    w = (g - 66667) * 0.25 + 8541.8
                ElseIf g < 666666 Then
                    w = (g - 166667) * 0.3 + 33541.8

                Else
                    w = (g - 666667) * 0.35 + 200833.33
                End If


                Case "MWE"

                    w = 0


        End Select

        GetWTax = w

End Function

Public Function GetExemptions(ByVal exemptionstat As String) As Currency
        Dim exemptions As Currency


        Select Case exemptionstat
            Case "MWE"
                exemptions = 200000

            Case "S"
               exemptions = 50000


            Case "S1"
                exemptions = 75000

            Case "S2"
                exemptions = 100000

            Case "S3"
                exemptions = 125000

            Case "S4"
                exemptions = 150000

            Case "ME"
                exemptions = 50000

            Case "ME1"
                exemptions = 75000

            Case "ME2"
                exemptions = 100000

            Case "ME3"
                exemptions = 125000

            Case "ME4"
                exemptions = 150000

        End Select

        GetExemptions = exemptions

End Function

Public Function GetSSSCont(ByVal gross As Currency) As Currency
       'Computes SSS Employee Contributions
        'gross = monthly gross + nontaxable
        Dim sss As Currency

        If gross < 4250 Then
            sss = 180
        ElseIf gross < 4750 Then
            sss = 202.5
        ElseIf gross < 5250 Then
            sss = 180
        ElseIf gross < 4750 Then
            sss = 225
        ElseIf gross < 5750 Then
            sss = 247.5
        ElseIf gross < 6250 Then
            sss = 270

        ElseIf gross < 6750 Then
            sss = 292.5
        ElseIf gross < 7250 Then
            sss = 315
        ElseIf gross < 7750 Then
            sss = 337.5
        ElseIf gross < 8250 Then
            sss = 360
        ElseIf gross < 8750 Then
            sss = 382.5
        ElseIf gross < 9250 Then
            sss = 405

        ElseIf gross < 9750 Then
            sss = 427.5
        ElseIf gross < 10250 Then
            sss = 450
        ElseIf gross < 10750 Then
            sss = 472.5
        ElseIf gross < 11250 Then
            sss = 495
        ElseIf gross < 11750 Then
            sss = 517.5
        ElseIf gross < 12250 Then
            sss = 540

        ElseIf gross < 12750 Then
            sss = 562.5
        ElseIf gross < 13250 Then
            sss = 585
        ElseIf gross < 13750 Then
            sss = 607.5
        ElseIf gross < 14250 Then
            sss = 630
        ElseIf gross < 14750 Then
            sss = 652.5
        ElseIf gross < 15250 Then
            sss = 675

        ElseIf gross < 15750 Then
            sss = 697.5
        ElseIf gross < 16250 Then
            sss = 720
        ElseIf gross < 16750 Then
            sss = 742.5
        ElseIf gross < 17250 Then
            sss = 765
        ElseIf gross < 17750 Then
            sss = 787.5
        ElseIf gross < 18250 Then
            sss = 810
        ElseIf gross < 18750 Then
            sss = 832.5
        ElseIf gross < 19250 Then
            sss = 855
        ElseIf gross < 19750 Then
            sss = 877.5
        ElseIf gross < 20250 Then
            sss = 900
        ElseIf gross < 20750 Then
            sss = 922.5
        ElseIf gross < 21250 Then
            sss = 945
        ElseIf gross < 21750 Then
            sss = 967.5
        ElseIf gross < 22250 Then
            sss = 990
        ElseIf gross < 22750 Then
            sss = 1012.5
        ElseIf gross < 23250 Then
            sss = 1035
        ElseIf gross < 23750 Then
            sss = 1057.5
        ElseIf gross < 24250 Then
            sss = 1080
        ElseIf gross < 24750 Then
            sss = 1102.5
        ElseIf gross < 25250 Then
            sss = 1125
         ElseIf gross < 25750 Then
            sss = 1147.5
        ElseIf gross < 26250 Then
            sss = 1170
        ElseIf gross < 26750 Then
            sss = 1192.5
        ElseIf gross < 27250 Then
            sss = 1215
        ElseIf gross < 27750 Then
            sss = 1237.5
        ElseIf gross < 28250 Then
            sss = 1260
        ElseIf gross < 28750 Then
            sss = 1282.5
        ElseIf gross < 29250 Then
            sss = 1305
        ElseIf gross < 29750 Then
            sss = 1327.5
        Else
            sss = 1350
        End If

        GetSSSCont = sss

End Function

Public Function GetER(ByVal ee As Currency) As Currency
        'Computes SSS Employer Contributions
        'from SSS Schedule 2007

        Dim er As Currency

           Select Case ee
            Case 0
               er = 0
            Case 180
               er = 380
            Case 202.5
               er = 427.5
            Case 225
               er = 475
            Case 247.5
               er = 522.5
            Case 270
               er = 570
            Case 292.5
               er = 617.5
            Case 315
               er = 665
            Case 337.5
               er = 712.5
            Case 360
               er = 760
            Case 382.5
               er = 807.5
            Case 405
               er = 855
            Case 427.5
               er = 902.5
            Case 450
               er = 950
            Case 472.5
               er = 997.5
            Case 495
               er = 1045
            Case 517.5
               er = 1092.5
            Case 540
               er = 1140
            Case 562.5
               er = 1187.5
            Case 585
               er = 1235
            Case 607.5
               er = 1282.5
            Case 630
               er = 1330
            Case 652.5
               er = 1377.5
            Case 675
               er = 1425
            Case 697.5
               er = 1472.5
            Case 720
               er = 1520
            Case 742.5
               er = 1567.5
            Case 765
               er = 1615
            Case 787.5
               er = 1662.5
            Case 810
               er = 1710
            Case 832.5
               er = 1757.5
            Case 855
               er = 1805
            Case 877.5
               er = 1852.5
            Case 900
               er = 1900
            Case 922.5
               er = 1947.5
            Case 945
               er = 1995
            Case 967.5
               er = 2042.5
            Case 990
               er = 2090
            Case 1012.5
               er = 2137.5
            Case 1035
               er = 2185
            Case 1057.5
               er = 2232.5
            Case 1080
               er = 2280
            Case 1102.5
               er = 2327.5
            Case 1125
               er = 2375
            Case 1147.5
               er = 2422.5
            Case 1170
               er = 2470
            Case 1192.5
               er = 2517.5
            Case 1215
               er = 2565
            Case 1237.5
               er = 2612.5
            Case 1260
               er = 2660
            Case 1282.5
               er = 2707.5
            Case 1305
               er = 2755
            Case 1327.5
               er = 2802.5
            Case 1350
               er = 2850

        End Select

        GetER = er

End Function

Public Function GetEC(ByVal ee As Currency) As Currency
'from SSS Contributions Table 2025

Dim ec As Currency
If ee < 750 Then
    ec = 10
Else
    ec = 30
End If
GetEC = ec
End Function

Public Function GetAnnualWTax(ByVal Taxable As Currency) As Currency
        'from tax table 2003
        Dim taxdue As Currency

        If Taxable < 10000 Then
            taxdue = Taxable * 0.05
        ElseIf Taxable < 30000 Then
            taxdue = 500 + 0.1 * (Taxable - 10000)
        ElseIf Taxable < 70000 Then
            taxdue = 2500 + 0.15 * (Taxable - 30000)
        ElseIf Taxable < 140000 Then
            taxdue = 8500 + 0.2 * (Taxable - 70000)
        ElseIf Taxable < 250000 Then
            taxdue = 22500 + 0.25 * (Taxable - 140000)
        ElseIf Taxable < 500000 Then
            taxdue = 50000 + 0.3 * (Taxable - 250000)
        Else
            taxdue = 125000 + 0.32 * (Taxable - 500000)
        End If
        GetAnnualWTax = taxdue
End Function
