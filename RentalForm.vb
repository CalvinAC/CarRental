'Calvin Coxen
'RCET 6025
'Fall 2020
'Car Rental
'https://github.com/CalvinAC/CarRental

Option Explicit On
Option Strict On
Option Compare Binary
Public Class RentalForm
    Dim totalClients As Integer
    Dim totalCharges, totalMiles, totalMilesDrove, charges As Double

    Function Verification() As Boolean
        Dim data As Boolean
        Dim zip, odmEnd, odmStart, days As Integer
        Dim name As Integer

        If NameTextBox.Text = "" Then
            data = False
            MsgBox("Fill out your name sir")
            Return IsNumeric(NameTextBox.Text)
            Exit Function
        Else
            data = True
        End If

        If AddressTextBox.Text = "" Then
            data = False
            MsgBox("Please type your address")
            Exit Function
        Else
            data = True
        End If

        If CityTextBox.Text = "" Then
            data = False
            MsgBox("Please type your city")
            Exit Function
        Else
            data = True
        End If

        If StateTextBox.Text = "" Then
            data = False
            MsgBox("Please type your state")
            Exit Function
        Else
            data = True
        End If

        If ZipCodeTextBox.Text = "" Then
            data = False
            MsgBox("Please type your zipcode")
            Exit Function
        Else
            data = True
        End If

        Try
            zip = CInt(ZipCodeTextBox.Text)
        Catch ex As Exception
            MsgBox("Zipcode should be a numeric value")
            ZipCodeTextBox.Clear()
            Exit Function
        End Try

        If BeginOdometerTextBox.Text = "" Then
            data = False
            MsgBox("Please enter beginning odometer reading")
            Exit Function
        Else
            data = True
        End If

        Try
            odmStart = CInt(BeginOdometerTextBox.Text)
        Catch ex As Exception
            MsgBox("Odometer reading should be a numeric value")
            BeginOdometerTextBox.Clear()
            Exit Function
        End Try

        If EndOdometerTextBox.Text = "" Then
            data = False
            MsgBox("Please enter ending odometer reading")
            Exit Function
        Else
            data = True
        End If

        Try
            odmEnd = CInt(EndOdometerTextBox.Text)
        Catch ex As Exception
            MsgBox("Odometer reading should be a numeric value")
            EndOdometerTextBox.Clear()
            Exit Function
        End Try

        If BeginOdometerTextBox.Text > EndOdometerTextBox.Text Then
            MsgBox("The ending milage should be higher than beginning mileage")
            BeginOdometerTextBox.Clear()
            EndOdometerTextBox.Clear()
            Exit Function
        Else
            data = True
        End If

        If DaysTextBox.Text = "" Then
            data = False
            MsgBox("Please enter the amount of days")
            Exit Function
        Else
            data = True
        End If

        Try
            days = CInt(DaysTextBox.Text)
            If days = 0 Then
                DaysTextBox.Clear()
                data = False
            ElseIf days >= 45 Then
                DaysTextBox.Clear()
                data = False
                MsgBox("Car cannnot be rented for longer than 45 days")
                Exit Function
            Else
                data = True
            End If
        Catch ex As Exception
            MsgBox("Days should be a numeric value")
            DaysTextBox.Clear()
            Exit Function
        End Try

        Return data
    End Function

    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        Dim startOdm, endOdm, milesCost, seniorDisc, aaaDisc As Double
        Dim dayCharge, dayChargeCalc As Integer

        If Verification() Then
            SummaryButton.Enabled = True
        Else

        End If

        Try
            startOdm = CDbl(BeginOdometerTextBox.Text)
            endOdm = CDbl(EndOdometerTextBox.Text)

            If KilometersradioButton.Checked = True Then
                totalMiles = ((endOdm * 0.621) - (startOdm * 0.621))
            Else
                totalMiles = (endOdm - startOdm)
            End If

            TotalMilesTextBox.Text = (Str(totalMiles) & "mi")

            If totalMiles <= 200 Then
                milesCost = 0
            ElseIf totalMiles <= 500 Then
                milesCost = ((totalMiles - 200) * 0.12)
            ElseIf totalMiles > 500 Then
                milesCost = ((300 * 0.12) + ((totalMiles - 500) * 0.1))
            End If

            MileageChargeTextBox.Text = ("$" & CStr(milesCost))
            dayCharge = CInt(DaysTextBox.Text)
            dayChargeCalc = dayCharge * 15
            DayChargeTextBox.Text = CStr(dayChargeCalc)

            aaaDisc = milesCost * 0.05
            seniorDisc = milesCost * 0.03

            If Seniorcheckbox.Checked = True And AAAcheckbox.Checked = True Then
                TotalDiscountTextBox.Text = ("$" & CStr(aaaDisc + seniorDisc))
                TotalChargeTextBox.Text = ("$" & CStr(milesCost - aaaDisc - aaaDisc))
            ElseIf AAAcheckbox.Checked = True Then
                TotalDiscountTextBox.Text = ("$" & CStr(aaaDisc))
                TotalChargeTextBox.Text = ("$" & CStr(milesCost - aaaDisc))
            ElseIf Seniorcheckbox.Checked = True Then
                TotalDiscountTextBox.Text = ("$" & CStr(seniorDisc))
                TotalChargeTextBox.Text = ("$" & CStr(milesCost - seniorDisc))
            Else
                TotalChargeTextBox.Text = ("$" & CStr(milesCost))
                TotalDiscountTextBox.Text = "$0"
            End If

            charges = CDbl(TotalChargeTextBox.Text)
        Catch ex As Exception

        End Try


    End Sub
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        Dim result As MsgBoxResult

        result = MsgBox("Would you like to exit the program?", MsgBoxStyle.YesNo)

        If result = 6 Then
            Me.Close()
        End If

    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click

        Delete()
    End Sub

    Private Sub RentalForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SummaryButton.Enabled = False
    End Sub

    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click
        totalClients += 1
        totalMilesDrove += totalMiles
        totalCharges += charges

        Delete()
        MsgBox("# of customer: " & totalClients & vbNewLine &
               "Total distance (miles): " & totalMilesDrove & vbNewLine &
               "Total Charges: " & totalCharges)
    End Sub

    Sub Delete()
        NameTextBox.Clear()
        AddressTextBox.Clear()
        CityTextBox.Clear()
        StateTextBox.Clear()
        ZipCodeTextBox.Clear()
        BeginOdometerTextBox.Clear()
        EndOdometerTextBox.Clear()
        DaysTextBox.Clear()
        TotalMilesTextBox.Clear()
        MileageChargeTextBox.Clear()
        DayChargeTextBox.Clear()
        TotalDiscountTextBox.Clear()
        TotalChargeTextBox.Clear()
        MilesradioButton.Checked = True
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False
    End Sub


End Class
