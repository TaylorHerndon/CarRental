Option Explicit On
Option Strict On
Option Compare Binary

'Taylor Herndon
'RCET0265
'Spring 2021
'CarRentalForum
'https://github.com/TaylorHerndon/CarRental

Public Class CarRentalForum

    Sub Startup() Handles Me.Load

        'Disable the summary button on startup
        SummaryButton.Enabled = False

    End Sub

    Sub CalculateButtonPress() Handles CalculateButton.Click, CalculateToolStripMenuItem.Click

        'Set all text colors to black
        NameTextBox.ForeColor = Color.Black
        AddressTextBox.ForeColor = Color.Black
        CityTextBox.ForeColor = Color.Black
        StateTextBox.ForeColor = Color.Black
        ZipCodeTextBox.ForeColor = Color.Black
        BeginOdometerTextBox.ForeColor = Color.Black
        EndOdometerTextBox.ForeColor = Color.Black
        DaysTextBox.ForeColor = Color.Black

        'Check Inputs
        Dim inputCheck(7) As Boolean
        inputCheck(7) = CheckString(NameTextBox.Text) 'Check if has number or empty
        inputCheck(6) = AddressTextBox.Text = "" 'Check if empty
        inputCheck(5) = CheckString(CityTextBox.Text) 'Check if has number or empty
        inputCheck(4) = CheckString(StateTextBox.Text) 'Check if has number or empty

        inputCheck(3) = CheckStringToInteger(ZipCodeTextBox.Text) 'Check if only numbers or empty
        inputCheck(3) = Len(ZipCodeTextBox.Text) <> 5 'Check if zip code is 8 characters long

        inputCheck(2) = CheckStringToInteger(BeginOdometerTextBox.Text) 'Check if only numbers or empty
        inputCheck(1) = CheckStringToInteger(EndOdometerTextBox.Text) 'Check if only numbers or empty

        inputCheck(0) = CheckStringToInteger(DaysTextBox.Text) 'Check if only numbers or empty

        'Go through the marked problems and set the focus to the top most text box with a problem
        Dim exitSub As Boolean = False

        For i = 0 To 7

            If inputCheck(i) = True Then

                Select Case i

                    Case 7
                        NameTextBox.Focus() 'Set the focus
                        NameTextBox.ForeColor = Color.Red 'Change text color to red to indicate the problem
                    Case 6
                        AddressTextBox.Focus()
                        AddressTextBox.ForeColor = Color.Red
                    Case 5
                        CityTextBox.Focus()
                        CityTextBox.ForeColor = Color.Red
                    Case 4
                        StateTextBox.Focus()
                        StateTextBox.ForeColor = Color.Red
                    Case 3
                        ZipCodeTextBox.Focus()
                        ZipCodeTextBox.ForeColor = Color.Red
                    Case 2
                        BeginOdometerTextBox.Focus()
                        BeginOdometerTextBox.ForeColor = Color.Red
                    Case 1
                        EndOdometerTextBox.Focus()
                        EndOdometerTextBox.ForeColor = Color.Red
                    Case 0
                        DaysTextBox.Focus()
                        DaysTextBox.ForeColor = Color.Red

                End Select

                exitSub = True

            End If

        Next

        If exitSub Then

            Exit Sub

        End If

        Dim distanceDriven As Double = CInt(EndOdometerTextBox.Text) - CInt(BeginOdometerTextBox.Text) 'Get distance driven

        'The following two statements catch errors that can be difficult to understand so they require msg boxes 

        'If distance driven is less than 0mi then notify the user and exit the sub
        If distanceDriven < 0 Then

            MsgBox("Error: Distance driven is less than 0." & vbNewLine & "Please check the begining and ending odometer reading and try again.", MsgBoxStyle.SystemModal, "Odometer Reading Error")
            Exit Sub

        End If

        'If number if days is less than 0 or greater than 45 prompt user and exit sub
        If CInt(DaysTextBox.Text) > 45 Or CInt(DaysTextBox.Text) < 0 Then

            MsgBox("Error: Days cannot be less than 0 or greater than 45.", MsgBoxStyle.ApplicationModal, "Error: Invalid Days")
            Exit Sub

        End If

        'If kilometers is selected then convert kilometers to miles
        If KilometersradioButton.Checked Then

            distanceDriven *= 0.62

        End If

        'Calculate the mileage charge based off of miles driven
        Dim mileageCharge As Double = 0

        If distanceDriven > 200 And distanceDriven <= 500 Then

            'If distance driven is between 200 and 500 miles then calculate as follows
            mileageCharge = (distanceDriven - 200) * 0.12

        Else

            If distanceDriven > 500 Then

                'If distance driven is greater than 500 then calculate as follows
                mileageCharge = ((distanceDriven - 500) * 0.1) + 32

            Else

                'If no other option is selected than distanceDriven is less than 200mi

            End If

        End If

        'Calculate day charge
        Dim dayCharge As Integer = CInt(DaysTextBox.Text) * 15

        'Assign discounts
        Dim minusDiscount As Double = 1

        If AAAcheckbox.Checked Then

            minusDiscount += 0.05

        End If

        If Seniorcheckbox.Checked Then

            minusDiscount += 0.03

        End If

        'Calculate total cost
        Dim totalCost As Double = (dayCharge + mileageCharge) / minusDiscount

        'Write all variables to the output text boxes
        TotalMilesTextBox.Text = distanceDriven & "mi"
        MileageChargeTextBox.Text = FormatCurrency(mileageCharge, 2)
        DayChargeTextBox.Text = FormatCurrency(dayCharge, 2)
        TotalDiscountTextBox.Text = CInt((minusDiscount - 1) * 100) & "%"
        TotalChargeTextBox.Text = FormatCurrency(totalCost, 2)

        'Store the distance drive and total cost
        StoreSummary(distanceDriven, totalCost, False)

        'Enable the summary button (At this point a calculation has been made)
        SummaryButton.Enabled = True

    End Sub

    Sub ClearButtonPress() Handles ClearButton.Click, ClearToolStripMenuItem.Click, ClearToolStripMenuItem1.Click

        'Clear all input text boxes
        NameTextBox.Clear()
        AddressTextBox.Clear()
        CityTextBox.Clear()
        StateTextBox.Clear()
        ZipCodeTextBox.Clear()
        BeginOdometerTextBox.Clear()
        EndOdometerTextBox.Clear()
        DaysTextBox.Clear()

        'Clear output text boxes
        TotalMilesTextBox.Clear()
        MileageChargeTextBox.Clear()
        DayChargeTextBox.Clear()
        TotalDiscountTextBox.Clear()
        TotalChargeTextBox.Clear()

        'Set units to miles
        MilesradioButton.Checked = True

        'Disable all discounts
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False

    End Sub

    Sub SummaryButtonPress() Handles SummaryButton.Click, SummaryToolStripMenuItem.Click, SummaryToolStripMenuItem1.Click

        'Write stored summary to a message box
        MsgBox(StoreSummary(-1, -1, False), MsgBoxStyle.SystemModal, "Detailed Summary")

    End Sub

    Sub ExitButtonPress() Handles ExitButton.Click, ExitToolStripMenuItem.Click, ExitToolStripMenuItem1.Click

        'Promt the user and ensure they want to quit
        If MsgBox("Are you sure you want to quit?", MsgBoxStyle.YesNo) = 6 Then

            End

        End If

    End Sub

    Function StoreSummary(distanceDriven As Double, charge As Double, Clear As Boolean) As String

        Static numberOfCustomers As Integer
        Static totalDistanceDriven As Double
        Static totalCharges As Double
        Static messageString As String

        'If clear is true then clear all history and reset the message string
        If Clear Then

            numberOfCustomers = 0
            totalDistanceDriven = 0
            totalCharges = 0
            messageString = "Number of Customers: 0" & vbNewLine & "Total Miles = 0mi" & vbNewLine & "Total Charge = $0"
            Return messageString

        End If

        'If either distanceDriven or totalCharge is less than 1 then abort the program
        If distanceDriven < 1 Or charge < 1 Then

            Return messageString

        End If

        'Add new numbers to static variables
        numberOfCustomers += 1
        totalDistanceDriven += distanceDriven
        totalCharges += charge

        'Recreate the message string with new numbers
        messageString = "Number of Customers: " & numberOfCustomers & vbNewLine &
                        "Total Miles = " & CInt(totalDistanceDriven) & vbNewLine &
                        "Total Charge = " & FormatCurrency(totalCharges)

        Return messageString

    End Function

    Function CheckString(checkThisString As String) As Boolean

        'If the string is empty then return "Empty"
        If checkThisString = "" Then

            Return True

        Else

            'If the string is not empty test if each character is a number
            For i = 0 To Len(checkThisString) - 1

                Try

                    Convert.ToInt32(checkThisString.Substring(i, 1)) 'Test the character
                    'If the code continues then the tested character is a number
                    Return True

                Catch ex As Exception

                End Try

            Next

        End If

        Return False

    End Function

    Function CheckStringToInteger(TestThisString As String) As Boolean

        'Try to convert the given string to an integer
        Try

            Dim p = CDbl(TestThisString)
            Return False 'If the string can be converted to an integer return false

        Catch ex As Exception

            'If the string cannot be converted to an integer return true
            Return True

        End Try

    End Function

    Function StoreMessage(Message As String, Clear As Boolean) As String

        Static storedMessages As String

        'If clear is true reset stored messages
        If Clear Then

            storedMessages = ""
            Return storedMessages

        End If

        'If message is empty then return the stored messages and continue
        If Message = "" Then

            Return storedMessages

        End If

        'Add the new message to the StoredMessages String
        storedMessages = storedMessages & vbNewLine & Message

        Return storedMessages

    End Function

End Class