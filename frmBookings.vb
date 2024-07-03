Imports System.Text
Imports BookingApp.Common
Imports BookingApp.Models
Imports BookingAppBL
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Net.Mime.MediaTypeNames
Imports System.Net.Security

Public Class frmBookings

#Region "Declaration"
    Dim BookingID As Int32
    Dim objBookingsBal As New clsBookingsBAL
    Dim objClientsBal As New clsClientsBAL
    Dim objRentDetailsBal As New clsRentDetailsBAL
    Dim chosenData As New clsBookings
    Dim booking As New clsBookings
    Dim Data As New List(Of clsBookingsSearch)
    Dim RentID As Int32
    Dim rentData As New List(Of clsRentDetails)
    Dim chosenRent As New clsRentDetails
    Dim SearchQuery As String
#End Region

#Region "Events"

    Private Sub frmBookings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitializeForm()
    End Sub

    Private Sub InitializeForm()
        Try
            SearchQuery = "SELECT Bookings.BookingId, Bookings.BookingDate, Bookings.LeadAgent, Bookings.DivisionType, Bookings.TransactionType, Bookings.PropertyName, " &
                          "ListClients.ClientName AS ListClientName, ProcClients.ClientName AS ProcClientName, ListFirms.ClientName AS ListFirmName, ProcFirms.ClientName AS ProcFirmName " &
                          "FROM (((Bookings " &
                          "Left JOIN Clients As ListClients On ListClients.ClientID = Bookings.ListClientID) " &
                          "LEFT JOIN Clients AS ProcClients ON ProcClients.ClientID = Bookings.ProcClientID) " &
                          "LEFT JOIN Clients AS ListFirms ON ListFirms.ClientID = Bookings.ListFirmID) " &
                          "LEFT JOIN Clients AS ProcFirms ON ProcFirms.ClientID = Bookings.ProcFirmID " &
                          "WHERE 1 = 1 " &
                          "ORDER BY Bookings.BookingDate DESC"
            DropDownsPopulate()
            IniGrid()
            LoadAllData(SearchQuery)
            IniRentGrid()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub textBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtPropInfoZip.KeyPress, TxtListingManual.KeyPress, TxtListingCommRate.KeyPress, TxtLFOutBroker1.KeyPress, TxtLFOutBroker2.KeyPress, TxtTotConsidation.KeyPress, TxtTotSf.KeyPress, TxtProcuringManual.KeyPress, TxtProcuringCommRate.KeyPress, TxtPFOutBroker1.KeyPress, TxtPFOutBroker2.KeyPress,
TxtLFGrossAmt1.KeyPress, TxtLFLEEAgt1.KeyPress, TxtLFSplitPct1.KeyPress, TxtPFGrossAmt1.KeyPress, TxtPFLEEAgt1.KeyPress, TxtPFSplitPct1.KeyPress,
TxtLFGrossAmt2.KeyPress, TxtLFLEEAgt2.KeyPress, TxtLFSplitPct2.KeyPress, TxtPFGrossAmt2.KeyPress, TxtPFLEEAgt2.KeyPress, TxtPFSplitPct2.KeyPress,
TxtLFGrossAmt3.KeyPress, TxtLFLEEAgt3.KeyPress, TxtLFSplitPct3.KeyPress, TxtPFGrossAmt3.KeyPress, TxtPFLEEAgt3.KeyPress, TxtPFSplitPct3.KeyPress,
TxtLFGrossAmt4.KeyPress, TxtLFLEEAgt4.KeyPress, TxtLFSplitPct4.KeyPress, TxtPFGrossAmt4.KeyPress, TxtPFLEEAgt4.KeyPress, TxtPFSplitPct4.KeyPress,
TxtLFGrossAmt5.KeyPress, TxtLFLEEAgt5.KeyPress, TxtLFSplitPct5.KeyPress, TxtPFGrossAmt5.KeyPress, TxtPFLEEAgt5.KeyPress, TxtPFSplitPct5.KeyPress,
TxtLFGrossAmt6.KeyPress, TxtLFLEEAgt6.KeyPress, TxtLFSplitPct6.KeyPress, TxtPFGrossAmt6.KeyPress, TxtPFLEEAgt6.KeyPress, TxtPFSplitPct6.KeyPress
        Try
            If Char.IsLetter(e.KeyChar) Or e.KeyChar = "-"c Then
                e.Handled = True
            End If
        Catch ex As Exception
            'logger.Log(ex.ToString)
            Throw ex
        End Try
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If ValidateBookingFields() Then
            If BookingID <> 0 Then
                UpdateBooking()
                'Call EmailSend()    ' SN 20240701
            Else
                InsertBooking()
            End If
            'ClearFieldValues()
        End If
    End Sub

    Private Sub BtnFetchData_Click(sender As Object, e As EventArgs) Handles BtnFetchData.Click
        FilterBookings()
    End Sub

    Private Sub BtnResetData_Click(sender As Object, e As EventArgs) Handles BtnResetData.Click
        ClearSearchFields()
        LoadAllData(SearchQuery)
    End Sub

    Private Sub ClearSearchFields()
        If comboBoxLeadAgent.Items.Count > 0 Then
            comboBoxLeadAgent.SelectedIndex = 0
        End If
        If CmbTranTypeSearch.Items.Count > 0 Then
            CmbTranTypeSearch.SelectedIndex = 0
        End If
        If CmbBrokerageFirm.Items.Count > 0 Then
            CmbBrokerageFirm.SelectedIndex = 0
        End If
        DTFromDateSearch.Value = DateTime.Now : DTFromDateSearch.Checked = False
        DTToDateSearch.Value = DateTime.Now : DTToDateSearch.Checked = False
        TxtPropNameSearch.Clear()
        TxtCompanyNameSearch.Clear()
    End Sub

    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        Try
            'ToolsEnable()
            ButtonEnableDisable(False, False, True, True, False, False, True)
            ClearFieldValues()
            BookingID = 0

            ' Add 12 rows by default to the Rent grid
            If GrdRentDetails.Rows.Count < 12 Then
                Dim Counter As Integer = 0

                For Counter = (GrdRentDetails.Rows.Count + 1) To 12
                    Dim values As Object() = {
                    0,
                    DBNull.Value,
                    DBNull.Value,
                    DBNull.Value,
                    DBNull.Value,
                    DBNull.Value,
                    DBNull.Value,
                    DBNull.Value,
                    DBNull.Value,
                    DBNull.Value
                }
                    GrdRentDetails.Rows.Add(values)
                Next
            End If
            GrdRentDetails.Refresh()

            'NavigationButtonAllDisable()
            'enabling the group box rdobtn
        Catch ex As Exception
            'logger.Log(ex.ToString)
            'Throw ex
        End Try
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        Try
            ToolsEnableDisable(False)
            ButtonEnableDisable(False, False, True, True, False, False, False)
            'NavigationButtonAllDisable()
        Catch ex As Exception
            'logger.Log(ex.ToString)
            Throw ex
        End Try
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to delete this Booking?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            ' Delete the current booking
            objBookingsBal.Delete(BookingID)
            MessageBox.Show("booking deleted successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            LoadAllData(SearchQuery)
            TabBookings.SelectedTab = TabPageBL
            'chosenData = Nothing
            'ClearFieldValues()
            ' Update button states
            ButtonEnableDisable(True, True, False, False, True, True)
            'NavigationButtonEnableDisable()
            'RecPosLabelChange()
        End If
    End Sub

    Private Sub grd_KeyDown(sender As Object, e As KeyEventArgs) Handles grdBookings.KeyDown
        Try
            If grdBookings.RowCount <= 0 Then
                Return
            End If
            If e.KeyCode = Keys.Enter Then
                e.SuppressKeyPress = True
                If grdBookings.CurrentRow.Selected Then
                    'If grdBookings.Rows.Count > 0 AndAlso grdBookings.CurrentCell.RowIndex < grdBookings.RowCount - 1 Then
                    Dim BookingID As Integer = CInt(grdBookings.CurrentRow.Cells("BookingID").Value)
                    booking = objBookingsBal.GetByID(BookingID)
                    TabBookings.SelectedTab = tabPageBD
                    SetFieldValues(booking)
                    ButtonEnableDisable(True, True, False, False, True, True)
                    'NavigationButtonEnableDisable()
                    'End If
                End If
            End If
        Catch ex As Exception
            'logger.Log(ex.ToString)
            Throw ex
        End Try
    End Sub

    Private Sub grd_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdBookings.CellDoubleClick
        Try
            If e.RowIndex >= 0 AndAlso e.RowIndex < grdBookings.Rows.Count Then
                Dim BookingID As Integer = CInt(grdBookings.Rows(e.RowIndex).Cells(BookingID).Value)
                booking = objBookingsBal.GetByID(BookingID)
                TabBookings.SelectedTab = tabPageBD
                SetFieldValues(booking)
                'ButtonEnableDisable(True, True, False, False, True, True)
                'NavigationButtonEnableDisable()
            End If
        Catch ex As Exception
            'logger.Log(ex.ToString)
            Throw ex
        End Try
    End Sub

    Private Sub grdBookings_SelectionChanged(sender As Object, e As EventArgs) Handles grdBookings.SelectionChanged
        Try
            If grdBookings.SelectedRows.Count > 0 Then
                Dim selectedRow As DataGridViewRow = grdBookings.SelectedRows(0)
                If selectedRow.Cells("BookingID").Value IsNot Nothing Then
                    BookingID = CInt(selectedRow.Cells("BookingID").Value)
                    Dim booking As clsBookings = objBookingsBal.GetByID(BookingID)
                    If booking IsNot Nothing AndAlso booking.BookingID <> 0 Then
                        SetFieldValues(booking)
                    End If
                    ButtonEnableDisable(True, True, False, False, True, True)
                    NavigationButtonEnableDisable()
                End If
            End If
        Catch ex As Exception
            'logger.Log(ex.ToString)
            Throw ex
        End Try
    End Sub

    Private Sub btnCancelChanges_Click(sender As Object, e As EventArgs) Handles btnCancelChanges.Click
        Dim dr As DialogResult
        Try
            dr = MessageBox.Show("Are you sure that you want to cancel the changes?", "Booking Form", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If dr = DialogResult.No Then
                Return
            Else
                SetFieldValues(chosenData)
                ToolsEnableDisable(True)
            End If
        Catch

        End Try
    End Sub

    Private Sub TabPageBL_Leave(sender As Object, e As EventArgs) Handles TabPageBL.Leave
        'RecPosLabelChange()
        ToolsEnableDisable(True)
        NavigationButtonEnableDisable()
        AssignComboBox()
        If grdBookings.SelectedRows.Count > 0 Then
            Dim bookingID As Integer = Convert.ToInt32(grdBookings.SelectedRows(0).Cells("bookingID").Value)
            Dim booking As clsBookings = objBookingsBal.GetByID(bookingID)
            SetFieldValues(booking)
            LoadAllRentData()
        Else
            If grdBookings.RowCount = 0 Then
                ButtonEnableDisable(True, False, False, False, False, False)
            End If
        End If
    End Sub

    Private Sub btnFirst_Click(sender As Object, e As EventArgs) Handles btnFirst.Click
        Try
            If grdBookings.Rows.Count > 0 Then
                ' Navigate to the first row
                grdBookings.ClearSelection()
                grdBookings.Rows(0).Selected = True
                grdBookings.CurrentCell = grdBookings.Rows(0).Cells("bookingID")
                ' Retrieve the booking object by ID
                Dim currentbookingID As Integer = CInt(grdBookings.CurrentRow.Cells("bookingID").Value)
                Dim booking As clsBookings = objBookingsBal.GetByID(currentbookingID)

                If booking IsNot Nothing Then
                    ' Pass the booking object to SetFieldValues
                    SetFieldValues(booking)
                    'RecPosLabelChange()
                    ButtonEnableDisable(True, True, False, False, True, True)
                    NavigationButtonEnableDisable()
                End If
            End If
        Catch ex As Exception
            ' Log and handle the exception
            ' logger.Log(ex.ToString)
            MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnPrev_Click(sender As Object, e As EventArgs) Handles btnPrev.Click
        Try
            Dim currentRowIndex As Integer = If(grdBookings.CurrentCell?.RowIndex, -1)
            Dim rowCount As Integer = grdBookings.RowCount

            If currentRowIndex > 0 AndAlso currentRowIndex < rowCount Then
                ' Move to the previous row
                grdBookings.CurrentCell = grdBookings.Rows(currentRowIndex - 1).Cells(grdBookings.CurrentCell.ColumnIndex)
                grdBookings.Rows(currentRowIndex).Selected = False
                grdBookings.Rows(currentRowIndex - 1).Selected = True

                ' Get the new bookingID from the selected row
                Dim selectedbookingID As Integer = CInt(grdBookings.Rows(currentRowIndex - 1).Cells("bookingID").Value)
                booking = objBookingsBal.GetByID(selectedbookingID)

                If booking IsNot Nothing AndAlso booking.BookingID <> 0 Then
                    SetFieldValues(booking)
                    'RecPosLabelChange()
                    NavigationButtonEnableDisable()
                End If
            End If
        Catch ex As Exception
            ' Handle any exceptions here (e.g., logging)
            MessageBox.Show("An error occurred: " & ex.Message)
        End Try
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        Try
            Dim currentRowIndex As Integer = If(grdBookings.CurrentCell?.RowIndex, -1)
            Dim rowCount As Integer = grdBookings.RowCount

            If currentRowIndex >= 0 AndAlso currentRowIndex < rowCount - 1 Then
                ' Move to the next row
                grdBookings.CurrentCell = grdBookings.Rows(currentRowIndex + 1).Cells(grdBookings.CurrentCell.ColumnIndex)
                grdBookings.Rows(currentRowIndex).Selected = False
                grdBookings.Rows(currentRowIndex + 1).Selected = True

                ' Get the new bookingID from the selected row
                Dim selectedbookingID As Integer = CInt(grdBookings.Rows(currentRowIndex + 1).Cells("bookingID").Value)
                booking = objBookingsBal.GetByID(selectedbookingID)

                If booking IsNot Nothing AndAlso booking.BookingID <> 0 Then
                    SetFieldValues(booking)
                    'RecPosLabelChange()
                    NavigationButtonEnableDisable()
                End If
            End If
        Catch ex As Exception
            ' Handle any exceptions here (e.g., logging)
            MessageBox.Show("An error occurred: " & ex.Message)
        End Try
    End Sub
    Private Sub CmbDivisionType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbDivisionType.SelectedIndexChanged
        GetDropDownValues("submarkets", ComboBoxSubMarkets)
    End Sub

    ' SN 20240701
    Private Function EmailSend() As Boolean
        Try
            Return clsGlobals.EmailSend(Me.BookingID)
        Catch ex As Exception
            MessageBox.Show("An error occurred in the event EmailSend()" & Environment.NewLine() & ex.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Private Sub btnLast_Click(sender As Object, e As EventArgs) Handles btnLast.Click
        Try
            If grdBookings.Rows.Count > 0 Then
                ' Navigate to the last row
                Dim lastRowIndex As Integer = grdBookings.Rows.Count - 1
                grdBookings.ClearSelection()
                grdBookings.Rows(lastRowIndex).Selected = True
                grdBookings.CurrentCell = grdBookings.Rows(lastRowIndex).Cells("bookingID")

                ' Retrieve the booking object by ID
                Dim currentbookingID As Integer = CInt(grdBookings.CurrentRow.Cells("bookingID").Value)
                Dim booking As clsBookings = objBookingsBal.GetByID(currentbookingID)

                If booking IsNot Nothing Then
                    ' Pass the booking object to SetFieldValues
                    SetFieldValues(booking)
                    'RecPosLabelChange()
                    ButtonEnableDisable(True, True, False, False, True, True)
                    NavigationButtonEnableDisable()
                End If
            End If
        Catch ex As Exception
            ' Log and handle the exception
            ' logger.Log(ex.ToString)
            MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'assigning info based on dropdown change
    Private Sub ComboBoxClients_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxClients.SelectedIndexChanged, ComboBoxClients2.SelectedIndexChanged, ComboBoxLF1.SelectedIndexChanged, ComboBoxLF2.SelectedIndexChanged, ComboBoxPF1.SelectedIndexChanged, ComboBoxPF2.SelectedIndexChanged
        ComboBox_SelectedIndexChanged(sender, e)
    End Sub

    Private Sub checkBoxLEELF_CheckedChanged(sender As Object, e As EventArgs) Handles checkBoxLEELF.CheckedChanged
        If checkBoxLEELF.Checked Then
            Dim ownCompanyLF As clsClients = objClientsBal.GetOwnCompany()
            If ownCompanyLF IsNot Nothing Then
                PopulateListingComboBoxWithOwnCompany(ComboBoxLF1, ownCompanyLF)
                'PopulateListingComboBoxWithOwnCompany(ComboBoxLF2, ownCompanyLF)
            Else
                MessageBox.Show("Own company data not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Else
            GetDropDownValues("listingFirms", ComboBoxLF1)
            checkBoxLEELF.Checked = False
            'RestoreFullListingFirms(ComboBoxLF1)
            'RestoreFullListingFirms(ComboBoxLF2)
        End If
    End Sub

    Private Sub CheckBoxLEEPF_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxLEEPF.CheckedChanged
        If CheckBoxLEEPF.Checked Then
            Dim ownCompanyLF As clsClients = objClientsBal.GetOwnCompany()
            If ownCompanyLF IsNot Nothing Then
                PopulateProcuringComboBoxWithOwnCompany(ComboBoxPF1, ownCompanyLF)
                'PopulateListingComboBoxWithOwnCompany(ComboBoxLF2, ownCompanyLF)
            Else
                MessageBox.Show("Own company data not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Else
            GetDropDownValues("procuringfirms", ComboBoxPF1)
            CheckBoxLEEPF.Checked = False
            'RestoreFullListingFirms(ComboBoxLF1)
            'RestoreFullListingFirms(ComboBoxLF2)
        End If
    End Sub

    Private Sub TabPageLease_Enter(sender As Object, e As EventArgs) Handles TabPageLease.Enter
        'IniRentGrid()
        'LoadAllRentData()
    End Sub

    'Private Sub grdRent_KeyDown(sender As Object, e As KeyEventArgs) Handles GrdRentDetails.KeyDown
    '    Try
    '        If GrdRentDetails.RowCount <= 0 Then
    '            Return
    '        End If
    '        If e.KeyCode = Keys.Enter Then
    '            e.SuppressKeyPress = True
    '            If GrdRentDetails.CurrentRow.Selected Then
    '                RentID = CInt(GrdRentDetails.CurrentRow.Cells("RentID").Value)
    '                Dim rentDetails As clsRentDetails = objRentDetailsBal.GetByID(RentID)
    '                'TabRentDetails.SelectedTab = tabPageRD
    '                SetRentFieldValues(rentDetails)
    '                'ButtonEnableDisable(True, True, False, False, True, True)
    '            End If
    '        End If

    '    Catch ex As Exception
    '        'logger.Log(ex.ToString)
    '        Throw ex
    '    End Try
    'End Sub

    Private Sub grdRent_SelectionChanged(sender As Object, e As EventArgs) Handles GrdRentDetails.SelectionChanged
        Try

            If GrdRentDetails.SelectedRows.Count > 0 Then
                Dim selectedRow As DataGridViewRow = GrdRentDetails.SelectedRows(0)
                If selectedRow.Cells("RentID").Value IsNot Nothing Then
                    RentID = CInt(selectedRow.Cells("RentID").Value)
                    Dim rent As clsRentDetails = objRentDetailsBal.GetByID(RentID)
                    If rent IsNot Nothing AndAlso rent.RentID <> 0 Then
                        SetRentFieldValues(rent)
                    End If
                    'ButtonEnableDisable(True, True, False, False, True, True)
                    'NavigationButtonEnableDisable()
                End If
            End If
        Catch ex As Exception
            'logger.Log(ex.ToString)
            Throw ex
        End Try
    End Sub

    Private Sub btnRentSave_Click(sender As Object, e As EventArgs) Handles btnRentSave.Click
        Try
            ' Populate the chosenRent object with current form values
            PopulateRentFromFields()

            If chosenRent.RentID = 0 Then
                objRentDetailsBal.Insert(chosenRent)
                MessageBox.Show($"Inserted Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                chosenRent.lastModifiedBy = "" 'change
                objRentDetailsBal.Update(chosenRent)
                MessageBox.Show($"Updated Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
            LoadAllRentData()
        Catch ex As Exception
            MessageBox.Show($"Error saving rent details: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnRentDel_Click(sender As Object, e As EventArgs) Handles BtnRentDel.Click
        Try
            If chosenRent.RentID = 0 Then
                MessageBox.Show("No rent details selected for deletion.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' Ask for confirmation before deletion
            Dim result As DialogResult = MessageBox.Show("Are you sure you want to delete this rent detail?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If result = DialogResult.Yes Then
                ' Proceed with deletion
                objRentDetailsBal.Delete(chosenRent.RentID)
                LoadAllRentData()
                'ClearRentFormFields()
                MessageBox.Show("Rent detail deleted successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show($"Error deleting rent detail: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TxtRentSFOverride_Leave(sender As Object, e As EventArgs) Handles TxtRentSFOverride.Leave
        If (Not String.IsNullOrEmpty(TxtRentSFOverride.Text)) Then
            TxtRentSfCal.Text = TxtRentSFOverride.Text
        ElseIf (Not String.IsNullOrEmpty(TxtLeasedSF.Text)) Then
            TxtRentSfCal.Text = TxtLeasedSF.Text
        End If
    End Sub

    Private Sub TxtRentMonthlyOverride_Leave(sender As Object, e As EventArgs) Handles TxtRentMonthlyOverride.Leave
        TxtRentMonthlyCal.Text = TxtRentMonthlyOverride.Text
    End Sub

    Private Sub TxtRentTotOverride_Leave(sender As Object, e As EventArgs) Handles TxtRentTotOverride.Leave
        TxtRentTotCal.Text = TxtRentTotOverride.Text
    End Sub

    Private Sub TxtRentMonthlyCal_Leave(sender As Object, e As EventArgs) Handles TxtRentMonthlyCal.Leave, TxtRentMonthlyCal.TextChanged, TxtRentNoOfMonth.TextChanged
        If String.IsNullOrEmpty(TxtRentMonthlyOverride.Text) AndAlso String.IsNullOrEmpty(TxtRentTotOverride.Text) Then
            If Not String.IsNullOrEmpty(TxtRentMonthlyCal.Text) AndAlso Not String.IsNullOrEmpty(TxtRentNoOfMonth.Text) Then
                TxtRentTotCal.Text = TxtRentMonthlyCal.Text * TxtRentNoOfMonth.Text
            End If
        End If
    End Sub

    Private Sub TxtRentToMonth_Leave(sender As Object, e As EventArgs) Handles TxtRentToMonth.Leave, TxtRentFromMonth.Leave
        Dim fromMonth As Int32
        Dim toMonth As Int32

        If Not String.IsNullOrEmpty(TxtRentFromMonth.Text) AndAlso Not String.IsNullOrEmpty(TxtRentToMonth.Text) Then
            fromMonth = Convert.ToInt32(TxtRentFromMonth.Text)
            toMonth = CInt(TxtRentToMonth.Text)
            TxtRentNoOfMonth.Text = ((toMonth - fromMonth) + 1).ToString()
        End If
    End Sub

    Private Sub TxtRentPerSF_Leave(sender As Object, e As EventArgs) Handles TxtRentPerSF.Leave
        If String.IsNullOrEmpty(TxtRentPerSF.Text) Then
            TxtRentMonthlyCal.Text = ""
            TxtRentTotCal.Text = ""
            TxtRentSfCal.Text = ""
            ' If TxtRentTotOverride or TxtRentMonthlyOverride has value, set it in the respective Cal textboxes
            If Not String.IsNullOrEmpty(TxtRentTotOverride.Text) Then
                TxtRentTotCal.Text = TxtRentTotOverride.Text
            Else
                TxtRentTotCal.Text = ""
            End If

            If Not String.IsNullOrEmpty(TxtRentMonthlyOverride.Text) Then
                TxtRentMonthlyCal.Text = TxtRentMonthlyOverride.Text
            Else
                TxtRentMonthlyCal.Text = ""
            End If
        End If
    End Sub

    Private Sub TxtRentSfCal_TextChanged(sender As Object, e As EventArgs) Handles TxtRentSfCal.TextChanged, cmbRSFCalculator.SelectedIndexChanged
        If Not cmbRSFCalculator.SelectedItem = "Select" Then

            If Not String.IsNullOrEmpty(TxtRentSfCal.Text) And Not String.IsNullOrEmpty(TxtRentPerSF.Text) Then
                If (cmbRSFCalculator.SelectedItem = "Monthly") Then
                    TxtRentMonthlyCal.Text = TxtRentPerSF.Text * TxtRentSfCal.Text
                ElseIf (cmbRSFCalculator.SelectedItem = "Annually") Then
                    TxtRentMonthlyCal.Text = (TxtRentPerSF.Text * TxtRentSfCal.Text) / 12
                End If
            End If
        End If
    End Sub
#End Region

#Region "Methods"
    Private Function DropDownsPopulate() As Boolean
        Try
            Dim leaseTypes As String() = {"Select", "NewLease", "SubLease", "Renewal", "Expansion"}
            Dim transactionTypes As String() = {"Select", "Lease", "Sale", "Consulting"}
            Dim leaseRateTypes As String() = {"Select", "FSG", "GRS", "IND GRS", "MOD GRS", "NNN"}
            Dim rsfCalculators As String() = {"Select", "Monthly", "Annually"}
            Dim divisionTypes As String() = {"Industrial", "Investment", "Multi Family", "Office", "Retail"}
#Region "Search Fields"
            LeadAgentsDropdown()
            PopulateComboBox(CmbTranTypeSearch, transactionTypes)
#End Region
            'ComboBoxLeaseType
            PopulateComboBox(ComboBoxLeaseType, leaseTypes)
            'CmbTransactionType
            PopulateComboBox(CmbTransactionType, transactionTypes)
            'CmbLeaseRateType
            PopulateComboBox(CmbLeaseRateType, leaseRateTypes)
            'cmbRSFCalculator
            PopulateComboBox(cmbRSFCalculator, rsfCalculators)
            'CmbDivisionType
            PopulateComboBox(CmbDivisionType, divisionTypes)

            Return True
        Catch ex As Exception
            ' Log the exception if necessary
            ' logger.Log(ex.ToString())
            Throw ex
            ' If an exception occurs, return False
            ' Return False
        End Try
    End Function

    Private Sub PopulateComboBox(comboBox As ComboBox, items As String())
        comboBox.Items.Clear()
        comboBox.Items.AddRange(items)
        comboBox.SelectedIndex = 0
    End Sub

    Private Sub BrokerageListFirmDropdown()
        Try
            CmbBrokerageFirm.Items.Clear()
            Dim listBrokerageFirms As List(Of String) = objBookingsBal.GetBrokerageFirmsAgents()
            CmbBrokerageFirm.Items.Add("Select")
            For Each listBrokerageFirm As String In listBrokerageFirms
                CmbBrokerageFirm.Items.Add(listBrokerageFirm)
            Next
            If CmbBrokerageFirm.Items.Count > 0 Then
                CmbBrokerageFirm.SelectedIndex = 0
            End If
        Catch ex As Exception
            MessageBox.Show("An error occurred while populating the Lead Agent dropdown: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub LeadAgentsDropdown()
        Try
            comboBoxLeadAgent.Items.Clear()
            comboBoxLeadAgent.Items.Add("Select")
            Dim leadAgents As List(Of String) = objBookingsBal.GetLeadAgents()
            For Each leadAgent As String In leadAgents
                comboBoxLeadAgent.Items.Add(leadAgent)
            Next
            If comboBoxLeadAgent.Items.Count > 0 Then
                comboBoxLeadAgent.SelectedIndex = 0
            End If
            BrokerageListFirmDropdown()
        Catch ex As Exception
            MessageBox.Show("An error occurred while populating the Lead Agent dropdown: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub FilterBookings()
        Dim CriteriaStr As String = ""
        Dim filteredData As New List(Of clsBookingsSearch)

        ' Filter by LeadAgent
        If comboBoxLeadAgent.SelectedItem IsNot Nothing AndAlso comboBoxLeadAgent.SelectedItem.ToString() <> "Select" Then
            CriteriaStr = "LeadAgent = '" & comboBoxLeadAgent.SelectedItem.ToString() & "' "
        End If

        ' Filter by TransactionType
        If CmbTranTypeSearch.SelectedItem IsNot Nothing AndAlso CmbTranTypeSearch.SelectedItem.ToString() <> "Select" Then
            CriteriaStr = CriteriaStr & IIf(CriteriaStr = "", "", " AND ") & "TransactionType = '" & CmbTranTypeSearch.SelectedItem.ToString() & "' "
        End If

        ' Filter by BrokerageFirm (assuming ListFirmID is being used to represent brokerage firm)
        If CmbBrokerageFirm.SelectedItem IsNot Nothing AndAlso CmbBrokerageFirm.SelectedItem.ToString() <> "Select" Then
            CriteriaStr = CriteriaStr & IIf(CriteriaStr = "", "", " AND ") & "(ListFirms.ClientName = '" & CmbBrokerageFirm.SelectedItem.ToString() & "' OR ProcFirms.ClientName = '" & CmbBrokerageFirm.SelectedItem.ToString() & "') "
        End If

        ' Filter by From Date
        If DTFromDateSearch.Checked Then
            If DTFromDateSearch.Value <> Nothing Then
                CriteriaStr = CriteriaStr & IIf(CriteriaStr = "", "", " AND ") & " CAST(BookingDate AS DATE) >= '" & DTFromDateSearch.Value.Date & "' "
            End If
        End If

        ' Filter by To Date
        If DTToDateSearch.Checked Then
            If DTToDateSearch.Value <> Nothing Then
                CriteriaStr = CriteriaStr & IIf(CriteriaStr = "", "", " AND ") & " CAST(BookingDate AS DATE) <= '" & DTToDateSearch.Value.Date & "' "
            End If
        End If

        ' Filter by PropertyName
        If Not String.IsNullOrWhiteSpace(TxtPropNameSearch.Text) Then
            CriteriaStr = CriteriaStr & IIf(CriteriaStr = "", "", " AND ") & "PropertyName LIKE '%" & TxtPropNameSearch.Text & "%' "
        End If

        ' Filter by CompanyName (assuming BillCompany is being used to represent company name)
        If Not String.IsNullOrWhiteSpace(TxtCompanyNameSearch.Text) Then
            CriteriaStr = CriteriaStr & IIf(CriteriaStr = "", "", " AND ") & "(ListClients.ClientName LIKE '%" & TxtCompanyNameSearch.Text & "%' OR ProcClients.ClientName LIKE '%" & TxtCompanyNameSearch.Text & "%') "
        End If

        If CriteriaStr = "" Then
            Call LoadAllData(SearchQuery)
        Else
            Call LoadAllData(Replace(SearchQuery, "1 = 1", CriteriaStr))
        End If
    End Sub

    Private Sub LoadAllData(DataSource As String)
        Try
            Data = Nothing
            Me.Cursor = Cursors.WaitCursor
            Data = objBookingsBal.ExecuteQuery(DataSource)
            ShowData(Data)
            'txtFilter.Focus()
            ButtonEnableDisable(True, True, False, False, True, True, True)
            Me.Cursor = Cursors.Default
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            Throw ex
        End Try
    End Sub

    Private Sub ShowData(data As List(Of clsBookingsSearch))
        grdBookings.Rows.Clear()
        If (data.Count = 0) Then
            grdBookings.Rows.Clear()
        Else
            'GetCityOrState("state")
            'GetCityOrState("city")
            For Each item As clsBookingsSearch In data
                Dim values As Object() = {
             item.BookingID,
             item.BookingDate,
             item.LeadAgent,
             item.DivisionType,
             item.TransactionType,
             item.PropertyName,
             item.ListClientName,
             item.ProcClientName,
             item.ListFirmName,
             item.ProcFirmName
         }
                grdBookings.Rows.Add(values)
            Next
        End If
        grdBookings.Refresh()
    End Sub

    Private Sub IniGrid()
        Try
            'DesignGridView(grdBookings, 10, False, False, False, True, DataGridViewAutoSizeColumnsMode.Fill)
            'grdBookings.Columns.Clear()
            grdBookings.Columns.Add("BookingID", "Booking ID") : grdBookings.Columns(0).DataPropertyName = "BookingID" : grdBookings.Columns(0).Visible = False
            grdBookings.Columns.Add("BookingDate", "Booking Date") : grdBookings.Columns(1).DataPropertyName = "BookingDate" : grdBookings.Columns(1).DefaultCellStyle.Format = "MM/dd/yyyy"
            grdBookings.Columns.Add("LeadAgent", "Lead Agent") : grdBookings.Columns(2).DataPropertyName = "LeadAgent"
            grdBookings.Columns.Add("DivisionType", "Division Type") : grdBookings.Columns(3).DataPropertyName = "DivisionType"
            grdBookings.Columns.Add("TransactionType", "Transaction Type") : grdBookings.Columns(4).DataPropertyName = "TransactionType"
            grdBookings.Columns.Add("PropertyName", "Property Name") : grdBookings.Columns(5).DataPropertyName = "PropertyName"
            grdBookings.Columns.Add("ListClientName", "Listing Company Name") : grdBookings.Columns(6).DataPropertyName = "ListClientName"
            grdBookings.Columns.Add("ProcClientName", "Procuring Company Name") : grdBookings.Columns(7).DataPropertyName = "ProcClientName"
            grdBookings.Columns.Add("ListFirmName", "Brokerage Listing Firm Name") : grdBookings.Columns(8).DataPropertyName = "ListFirmName"
            grdBookings.Columns.Add("ProcFirmName", "Brokerage Procuring Firm Name") : grdBookings.Columns(9).DataPropertyName = "ProcFirmName"

            grdBookings.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub SetFieldValues(booking As clsBookings)
        Try
            ' Check if booking is not null and necessary properties have non-null, non-empty values
            If booking IsNot Nothing Then
                ' Set form fields with booking values
#Region "TextBoxes"
                BookingID = booking.BookingID
                txtLeadAgent.Text = booking.LeadAgent
                txtBookingDate.Value = booking.BookingDate
                CmbDivisionType.Text = booking.DivisionType
                If (Not String.IsNullOrEmpty(booking.TransactionType)) Then
                    CmbTransactionType.SelectedItem = booking.TransactionType
                Else
                    CmbTransactionType.SelectedIndex = 0
                End If

                TxtPropInfoName.Text = booking.PropertyName
                ComboBoxSubMarkets.SelectedValue = booking.SubMarketID
                TxtPropInfoAddress.Text = booking.PropertyAddress
                TxtPropInfoCity.Text = booking.PropertyCity
                TxtPropInfoState.Text = booking.PropertyState
                TxtPropInfoZip.Text = booking.PropertyZIP
	TxtTotConsidation.Text = booking.TotalConsidation
                TxtTotSf.Text = booking.TotalSf
                ComboBoxClients.SelectedValue = booking.ListClientID
                ComboBoxLF1.SelectedValue = booking.ListFirmID
                CmbBFIAttention1.SelectedValue = booking.ListFirmAddressID
                CmbCIAttention1.SelectedValue = booking.ListClientAddressID
                checkBoxLEELF.Checked = booking.ListIsLeeArizona
                ComboBoxLF2.SelectedValue = booking.AddRefListFirmID
                CmbBFIAttention3.SelectedValue = booking.AddRefListFirmAddressID
                TxtLFTotalFee.Text = booking.ListTotalFees.ToString()
                TxtListingManual.Text = booking.ListTotalFeesManual.ToString()
                TxtListingCommRate.Text = booking.ListCommRate.ToString()
                TxtLFOutBroker1.Text = booking.ListOutsideBroker1.ToString()
                TxtLFOutBroker2.Text = booking.ListOutsideBroker2.ToString()
                TxtLFLeeGrossComm.Text = booking.ListLeeGrossComm.ToString()
                TxtLFGrossAmt1.Text = booking.ListGrossAmt1.ToString()
                TxtLFGrossAmt2.Text = booking.ListGrossAmt2.ToString()
                TxtLFGrossAmt3.Text = booking.ListGrossAmt3.ToString()
                TxtLFGrossAmt4.Text = booking.ListGrossAmt4.ToString()
                TxtLFGrossAmt5.Text = booking.ListGrossAmt5.ToString()
                TxtLFGrossAmt6.Text = booking.ListGrossAmt6.ToString()
                TxtLFLEEAgt1.Text = booking.ListLeeAgent1
                TxtLFLEEAgt2.Text = booking.ListLeeAgent2
                TxtLFLEEAgt3.Text = booking.ListLeeAgent3
                TxtLFLEEAgt4.Text = booking.ListLeeAgent4
                TxtLFLEEAgt5.Text = booking.ListLeeAgent5
                TxtLFLEEAgt6.Text = booking.ListLeeAgent6
                TxtLFSplitPct1.Text = booking.ListSplitPct1.ToString()
                TxtLFSplitPct2.Text = booking.ListSplitPct2.ToString()
                TxtLFSplitPct3.Text = booking.ListSplitPct3.ToString()
                TxtLFSplitPct4.Text = booking.ListSplitPct4.ToString()
                TxtLFSplitPct5.Text = booking.ListSplitPct5.ToString()
                TxtLFSplitPct6.Text = booking.ListSplitPct6.ToString()
                ComboBoxClients2.SelectedValue = booking.ProcClientID
                checkBoxPropAddress.Checked = booking.ProcIsPropertyAddress
                ComboBoxPF1.SelectedValue = booking.ProcFirmID
                CmbCIAttention2.SelectedValue = booking.ProcClientAddressID
                CmbBFIAttention2.SelectedValue = booking.ProcFirmAddressID
                CheckBoxLEEPF.Checked = booking.ProcIsLeeArizona
                ComboBoxPF2.SelectedValue = booking.AddRefProcFirmID
                CmbBFIAttention4.SelectedValue = booking.AddRefProcFirmAddressID
                TxtPFTotalFee.Text = booking.ProcTotalFees.ToString()
                TxtProcuringManual.Text = booking.ProcTotalFeesManual.ToString()
                TxtProcuringCommRate.Text = booking.ProcCommRate.ToString()
                TxtPFOutBroker1.Text = booking.ProcOutsideBroker1.ToString()
                TxtPFOutBroker2.Text = booking.ProcOutsideBroker2.ToString()
                TxtPFLeeGrossComm.Text = booking.ProcLeeGrossComm.ToString()
                TxtPFGrossAmt1.Text = booking.ProcGrossAmt1.ToString()
                TxtPFGrossAmt2.Text = booking.ProcGrossAmt2.ToString()
                TxtPFGrossAmt3.Text = booking.ProcGrossAmt3.ToString()
                TxtPFGrossAmt4.Text = booking.ProcGrossAmt4.ToString()
                TxtPFGrossAmt5.Text = booking.ProcGrossAmt5.ToString()
                TxtPFGrossAmt6.Text = booking.ProcGrossAmt6.ToString()
                TxtPFLEEAgt1.Text = booking.ProcLeeAgent1
                TxtPFLEEAgt2.Text = booking.ProcLeeAgent2
                TxtPFLEEAgt3.Text = booking.ProcLeeAgent3
                TxtPFLEEAgt4.Text = booking.ProcLeeAgent4
                TxtPFLEEAgt5.Text = booking.ProcLeeAgent5
                TxtPFLEEAgt6.Text = booking.ProcLeeAgent6
                TxtPFSplitPct1.Text = booking.ProcSplitPct1.ToString()
                TxtPFSplitPct2.Text = booking.ProcSplitPct2.ToString()
                TxtPFSplitPct3.Text = booking.ProcSplitPct3.ToString()
                TxtPFSplitPct4.Text = booking.ProcSplitPct4.ToString()
                TxtPFSplitPct5.Text = booking.ProcSplitPct5.ToString()
                TxtPFSplitPct6.Text = booking.ProcSplitPct6.ToString()
                ChkBillListing.Checked = booking.BillListBroker
                ChkBillClient.Checked = booking.BillClient
                ChkBillOther.Checked = booking.BillOther
                TxtBillCompany.Text = booking.BillCompany
                TxtBillInfoAttention.Text = booking.BillAttention
                TxtBillInfoAddress.Text = booking.BillAddress
                TxtBillInfoCity.Text = booking.BillCity
                TxtBillInfoState.Text = booking.BillState
                TxtBillInfoZip.Text = booking.BillZIP
                TxtBillInfoPhnNo.Text = booking.BillPhoneNo
                CheckBoxSendWireInst.Checked = booking.WireInstructionsSend
                TxtEscrowNo.Text = booking.EscrowNumber
                CheckBoxDontSndInv.Checked = booking.DoNotSendInvoice
                CheckBoxSndMailInv.Checked = booking.InvoiceMail
                CheckBoxSendEmailInv.Checked = booking.InvoiceEmail
                dtNewDate1.Value = booking.DueDate1
                TxtDuePct1.Text = booking.DuePercent1.ToString()
                dtNewDate2.Value = booking.DueDate2
                TxtDuePct2.Text = booking.DuePercent2.ToString()
                dtNewDate3.Value = booking.DueDate3
                TxtDuePct3.Text = booking.DuePercent3.ToString()
                dtNewDate4.Value = booking.DueDate4
                TxtDuePct4.Text = booking.DuePercent4.ToString()
                TxtBillEmail1.Text = booking.BillEmailAddress1
                TxtBillEmail2.Text = booking.BillEmailAddress2
                TxtBillEmail3.Text = booking.BillEmailAddress3
                If (Not String.IsNullOrEmpty(booking.LeaseType)) Then
                    ComboBoxLeaseType.SelectedItem = booking.LeaseType
                Else
                    ComboBoxLeaseType.SelectedIndex = 0
                End If
                If (Not String.IsNullOrEmpty(booking.LeaseRateType)) Then
                    CmbLeaseRateType.SelectedItem = booking.LeaseRateType
                Else
                    CmbLeaseRateType.SelectedIndex = 0
                End If
                TxtLeaseTerm.Text = booking.LeaseTermMonths.ToString()
                TxtLeasedSF.Text = booking.LeasedSF
                If (Not String.IsNullOrEmpty(booking.RSFCalculator)) Then
                    cmbRSFCalculator.SelectedItem = booking.RSFCalculator
                Else
                    cmbRSFCalculator.SelectedIndex = 0
                End If
                TxtLeaseConsid.Text = booking.LeaseConsideration.ToString()
                DTLeaseCommence.Value = booking.CommencementDate
                DTLeaseExpiry.Value = booking.ExpDate
                DTLeaseOccupancy.Value = booking.OccupancyDate
                TxtLeaseTIAllow.Text = booking.TIAllowance.ToString()
                TxtParkingRatio.Text = booking.ParkingRatio.ToString()
                TxtReservedAmt.Text = booking.ReservedAmt.ToString()
                TxtUnReservedAmt.Text = booking.UnreservedAmt.ToString()
                TxtRoofTopAmt.Text = booking.RooftopAmt.ToString()
                TxtLeaseNotes.Text = booking.LeaseNotes
                TxtLFCommAdd.Text = booking.ListCommAddition.ToString()
                TxtLFCommDed.Text = booking.ListCommDeduction.ToString()
                TxtPFCommAdd.Text = booking.ProcCommAddition.ToString()
                TxtPFCommDed.Text = booking.ProcCommDeduction.ToString()

                'TxtCreatedBy.Text = booking.createdBy.ToString()
                'TxtCreatedOn.Text = booking.createdOn.ToString()
                'TxtModBy.Text = booking.lastModifiedBy.ToString()
                'TxtModOn.Text = booking.lastModifiedOn.ToString()
                CalculateTotalSf()
                CalculatePercetageTotal()
                CalculateEffectiveRate()
                'CalculateAgentLeave()
#End Region
                ' Set focus to a specific control if necessary
                txtLeadAgent.Focus()
                chosenData = booking
                ToolsEnableDisable(True)
                ButtonEnableDisable(True, True, False, False, True, True)
            Else
                ButtonEnableDisable(True, False, False, False, True, True)
                ' Clear the form fields if the booking object is null or contains invalid data
                ClearFieldValues()
                MessageBox.Show("Selected booking has invalid data or is null.", "Invalid Data", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
            'NavigationButtonEnableDisable()

        Catch ex As Exception
            ' Log and rethrow the exception
            ' logger.Log(ex.ToString())
            'Throw ex
        End Try
    End Sub

    Private Sub ButtonEnableDisable(IsEnableNew As Boolean,
                                    IsEnableEdit As Boolean,
                                    IsEnableSave As Boolean,
                                    IsEnableCancelChanges As Boolean,
                                    IsEnableDelete As Boolean,
                                    IsEnableExit As Boolean, Optional IsClickPreviousTabButtonEnable As Boolean = True)
        btnNew.Enabled = IsEnableNew
        btnEdit.Enabled = IsEnableEdit
        btnSave.Enabled = IsEnableSave
        btnCancelChanges.Enabled = IsEnableCancelChanges
        btnDelete.Enabled = IsEnableDelete
        'btnExit.Enabled = IsEnableExit
    End Sub

    Private Sub ClearFieldValues()
        Try
            ' Clear text fields
            txtLeadAgent.Text = String.Empty
            txtBookingDate.Value = DateTime.Now
            CmbDivisionType.SelectedIndex = -1
            CmbTransactionType.SelectedIndex = 0
            TxtPropInfoName.Text = String.Empty
            TxtPropInfoAddress.Text = String.Empty
            TxtPropInfoCity.Text = String.Empty
            TxtPropInfoState.Text = String.Empty
            TxtPropInfoZip.Text = String.Empty
            TxtLFTotalFee.Text = String.Empty
            TxtLFOutBroker1.Text = String.Empty
            TxtLFOutBroker2.Text = String.Empty
            TxtLFLeeGrossComm.Text = String.Empty
            TxtPFTotalFee.Text = String.Empty
            TxtPFOutBroker1.Text = String.Empty
            TxtPFOutBroker2.Text = String.Empty
            TxtPFLeeGrossComm.Text = String.Empty
            TxtBillInfoAttention.Text = String.Empty
            TxtBillInfoAddress.Text = String.Empty
            TxtBillInfoCity.Text = String.Empty
            TxtBillInfoState.Text = String.Empty
            TxtBillInfoZip.Text = String.Empty
            TxtBillInfoPhnNo.Text = String.Empty
            ComboBoxLeaseType.SelectedIndex = 0
            CmbLeaseRateType.SelectedIndex = 0
            TxtLeaseTerm.Text = String.Empty
            TxtLeasedSF.Text = String.Empty
            ComboBoxLeaseType.SelectedIndex = 0
            TxtLeaseConsid.Text = String.Empty
            DTLeaseCommence.Value = DateTime.Now
            DTLeaseExpiry.Value = DateTime.Now
            DTLeaseOccupancy.Value = DateTime.Now
            TxtLeaseTIAllow.Text = String.Empty
            TxtParkingRatio.Text = String.Empty
            TxtReservedAmt.Text = String.Empty
            TxtUnReservedAmt.Text = String.Empty
            TxtRoofTopAmt.Text = String.Empty
            TxtLeaseNotes.Text = String.Empty
            TxtLFCommAdd.Text = String.Empty
            TxtLFCommDed.Text = String.Empty
            TxtPFCommAdd.Text = String.Empty
            TxtPFCommDed.Text = String.Empty

            ' Uncheck checkboxes
            checkBoxLEELF.Checked = False
            checkBoxPropAddress.Checked = False
            CheckBoxLEEPF.Checked = False
            CheckBoxSendWireInst.Checked = False
            CheckBoxDontSndInv.Checked = False
            CheckBoxSndMailInv.Checked = False
            CheckBoxSendEmailInv.Checked = False

            ' Set comboboxes selected value to Nothing
            ComboBoxSubMarkets.SelectedValue = -1
            ComboBoxClients.SelectedValue = -1
            ComboBoxLF1.SelectedValue = -1
            ComboBoxLF2.SelectedValue = -1
            ComboBoxClients2.SelectedValue = -1
            ComboBoxPF1.SelectedValue = -1
            ComboBoxPF2.SelectedValue = -1

            TxtCreatedBy.Text = ""
            TxtCreatedOn.Text = ""
            TxtModBy.Text = ""
            TxtModOn.Text = ""

#Region "Change by HB"
            TxtTotConsidation.Text = String.Empty
            TxtTotSf.Text = String.Empty
            TxtListingManual.Text = String.Empty
            TxtProcuringManual.Text = String.Empty
            TxtPFLeeGrossComm.Text = String.Empty
            TxtPFNetAgt1.Text = String.Empty
            txtPFGrossComm1.Text = String.Empty
            TxtPFSplitPct1.Text = String.Empty
            TxtPFGrossAmt1.Text = String.Empty
            TxtPFLEEAgt1.Text = String.Empty
            TxtPFNetAgt2.Text = String.Empty
            txtPFGrossComm2.Text = String.Empty
            TxtPFSplitPct2.Text = String.Empty
            TxtPFGrossAmt2.Text = String.Empty
            TxtPFLEEAgt2.Text = String.Empty
            TxtPFNetAgt3.Text = String.Empty
            txtPFGrossComm3.Text = String.Empty
            TxtPFSplitPct3.Text = String.Empty
            TxtPFGrossAmt3.Text = String.Empty
            TxtPFLEEAgt3.Text = String.Empty
            TxtPFNetAgt4.Text = String.Empty
            txtPFGrossComm4.Text = String.Empty
            TxtPFSplitPct4.Text = String.Empty
            TxtPFGrossAmt4.Text = String.Empty
            TxtPFLEEAgt4.Text = String.Empty
            TxtPFNetAgt5.Text = String.Empty
            txtPFGrossComm5.Text = String.Empty
            TxtPFSplitPct5.Text = String.Empty
            TxtPFGrossAmt5.Text = String.Empty
            TxtPFLEEAgt5.Text = String.Empty
            TxtPFNetAgt6.Text = String.Empty
            txtPFGrossComm6.Text = String.Empty
            TxtPFSplitPct6.Text = String.Empty
            TxtPFGrossAmt6.Text = String.Empty
            TxtPFLEEAgt6.Text = String.Empty


            TxtLFLeeGrossComm.Text = String.Empty
            TxtLFNetAgt1.Text = String.Empty
            txtLFGrossComm1.Text = String.Empty
            TxtLFSplitPct1.Text = String.Empty
            TxtLFGrossAmt1.Text = String.Empty
            TxtLFLEEAgt1.Text = String.Empty
            TxtLFNetAgt2.Text = String.Empty
            txtLFGrossComm2.Text = String.Empty
            TxtLFSplitPct2.Text = String.Empty
            TxtLFGrossAmt2.Text = String.Empty
            TxtLFLEEAgt2.Text = String.Empty
            TxtLFNetAgt3.Text = String.Empty
            txtLFGrossComm3.Text = String.Empty
            TxtLFSplitPct3.Text = String.Empty
            TxtLFGrossAmt3.Text = String.Empty
            TxtLFLEEAgt3.Text = String.Empty
            TxtLFNetAgt4.Text = String.Empty
            txtLFGrossComm4.Text = String.Empty
            TxtLFSplitPct4.Text = String.Empty
            TxtLFGrossAmt4.Text = String.Empty
            TxtLFLEEAgt4.Text = String.Empty
            TxtLFNetAgt5.Text = String.Empty
            txtLFGrossComm5.Text = String.Empty
            TxtLFSplitPct5.Text = String.Empty
            TxtLFGrossAmt5.Text = String.Empty
            TxtLFLEEAgt5.Text = String.Empty
            TxtLFNetAgt6.Text = String.Empty
            txtLFGrossComm6.Text = String.Empty
            TxtLFSplitPct6.Text = String.Empty
            TxtLFGrossAmt6.Text = String.Empty
            TxtLFLEEAgt6.Text = String.Empty
#End Region

            ' Set focus to the first text field
            txtLeadAgent.Focus()

            ToolsEnableDisable(False)
        Catch ex As Exception
            ' Handle any potential exceptions here
            ' logger.Log(ex.ToString())
            Throw ex
        End Try
    End Sub

    Private Sub ToolsEnableDisable(enable As Boolean)
        Try
            ' Enable or disable the controls based on the boolean parameter
            txtLeadAgent.ReadOnly = enable
            txtBookingDate.Enabled = Not enable
            CmbDivisionType.Enabled = Not enable
            CmbTransactionType.Enabled = Not enable
            TxtPropInfoName.ReadOnly = enable
            ComboBoxSubMarkets.Enabled = Not enable
            TxtPropInfoAddress.ReadOnly = enable
            TxtPropInfoCity.ReadOnly = enable
            TxtPropInfoState.ReadOnly = enable
            TxtPropInfoZip.ReadOnly = enable
            ComboBoxClients.Enabled = Not enable
            ComboBoxLF1.Enabled = Not enable
            checkBoxLEELF.Enabled = Not enable
            CmbCIAttention1.Enabled = Not enable
            CmbCIAttention2.Enabled = Not enable
            ComboBoxLF2.Enabled = Not enable
            TxtListingCommRate.ReadOnly = enable
            TxtLFOutBroker1.ReadOnly = enable
            TxtLFOutBroker2.ReadOnly = enable
            TxtLFLeeGrossComm.ReadOnly = enable
            TxtLFGrossAmt1.ReadOnly = enable
            TxtLFGrossAmt2.ReadOnly = enable
            TxtLFGrossAmt3.ReadOnly = enable
            TxtLFGrossAmt4.ReadOnly = enable
            TxtLFGrossAmt5.ReadOnly = enable
            TxtLFGrossAmt6.ReadOnly = enable
            TxtLFLEEAgt1.ReadOnly = enable
            TxtLFLEEAgt2.ReadOnly = enable
            TxtLFLEEAgt3.ReadOnly = enable
            TxtLFLEEAgt4.ReadOnly = enable
            TxtLFLEEAgt5.ReadOnly = enable
            TxtLFLEEAgt6.ReadOnly = enable
            TxtLFSplitPct1.ReadOnly = enable
            TxtLFSplitPct2.ReadOnly = enable
            TxtLFSplitPct3.ReadOnly = enable
            TxtLFSplitPct4.ReadOnly = enable
            TxtLFSplitPct5.ReadOnly = enable
            TxtLFSplitPct6.ReadOnly = enable
            ComboBoxClients2.Enabled = Not enable
            checkBoxPropAddress.Enabled = Not enable
            ComboBoxPF1.Enabled = Not enable
            CheckBoxLEEPF.Enabled = Not enable
            ComboBoxPF2.Enabled = Not enable
            TxtProcuringCommRate.ReadOnly = enable
            TxtPFOutBroker1.ReadOnly = enable
            TxtPFOutBroker2.ReadOnly = enable
            TxtPFLeeGrossComm.ReadOnly = enable
            TxtPFGrossAmt1.ReadOnly = enable
            TxtPFGrossAmt2.ReadOnly = enable
            TxtPFGrossAmt3.ReadOnly = enable
            TxtPFGrossAmt4.ReadOnly = enable
            TxtPFGrossAmt5.ReadOnly = enable
            TxtPFGrossAmt6.ReadOnly = enable
            TxtPFLEEAgt1.ReadOnly = enable
            TxtPFLEEAgt2.ReadOnly = enable
            TxtPFLEEAgt3.ReadOnly = enable
            TxtPFLEEAgt4.ReadOnly = enable
            TxtPFLEEAgt5.ReadOnly = enable
            TxtPFLEEAgt6.ReadOnly = enable
            TxtPFSplitPct1.ReadOnly = enable
            TxtPFSplitPct2.ReadOnly = enable
            TxtPFSplitPct3.ReadOnly = enable
            TxtPFSplitPct4.ReadOnly = enable
            TxtPFSplitPct5.ReadOnly = enable
            TxtPFSplitPct6.ReadOnly = enable
            ChkBillListing.Enabled = Not enable
            ChkBillClient.Enabled = Not enable
            ChkBillOther.Enabled = Not enable
            TxtBillCompany.ReadOnly = enable
            TxtBillInfoAttention.ReadOnly = enable
            TxtBillInfoAddress.ReadOnly = enable
            TxtBillInfoCity.ReadOnly = enable
            TxtBillInfoState.ReadOnly = enable
            TxtBillInfoZip.ReadOnly = enable
            TxtBillInfoPhnNo.ReadOnly = enable
            CheckBoxSendWireInst.Enabled = Not enable
            TxtEscrowNo.ReadOnly = enable
            CheckBoxDontSndInv.Enabled = Not enable
            CheckBoxSndMailInv.Enabled = Not enable
            CheckBoxSendEmailInv.Enabled = Not enable
            dtNewDate1.Enabled = Not enable
            TxtDuePct1.ReadOnly = enable
            dtNewDate2.Enabled = Not enable
            TxtDuePct2.ReadOnly = enable
            dtNewDate3.Enabled = Not enable
            TxtDuePct3.ReadOnly = enable
            dtNewDate4.Enabled = Not enable
            TxtDuePct4.ReadOnly = enable
            TxtBillEmail1.ReadOnly = enable
            TxtBillEmail2.ReadOnly = enable
            TxtBillEmail3.ReadOnly = enable
            ComboBoxLeaseType.Enabled = Not enable
            CmbLeaseRateType.Enabled = Not enable
            TxtLeaseTerm.ReadOnly = enable
            TxtLeasedSF.ReadOnly = enable
            cmbRSFCalculator.Enabled = Not enable
            TxtLeaseConsid.ReadOnly = enable
            DTLeaseCommence.Enabled = Not enable
            DTLeaseExpiry.Enabled = Not enable
            DTLeaseOccupancy.Enabled = Not enable
            TxtLeaseTIAllow.ReadOnly = enable
            TxtParkingRatio.ReadOnly = enable
            TxtReservedAmt.ReadOnly = enable
            TxtUnReservedAmt.ReadOnly = enable
            TxtRoofTopAmt.ReadOnly = enable
            TxtLeaseNotes.ReadOnly = enable
            TxtLFCommAdd.ReadOnly = enable
            TxtLFCommDed.ReadOnly = enable
            TxtPFCommAdd.ReadOnly = enable
            TxtPFCommDed.ReadOnly = enable
            TxtListingManual.ReadOnly = enable
            TxtProcuringManual.ReadOnly = enable
            ' Optional: Enable/disable created/modified info fields if they exist and are used
            ' txtCreatedBy.ReadOnly = enable
            ' txtCreatedOn.Enabled = Not enable
            ' txtLastModifiedBy.ReadOnly = enable
            ' txtLastModifiedOn.Enabled = Not enable

            ' Set focus to the first input field if enabling
            If Not enable Then
                txtLeadAgent.Focus()
            End If
        Catch ex As Exception
            ' Handle any potential exceptions here
            ' logger.Log(ex.ToString())
            Throw ex
        End Try
    End Sub



    Private Function ValidateBookingFields() As Boolean
        Dim isValid As Boolean = True
        Dim ValidationMessage As New StringBuilder()

        'If String.IsNullOrWhiteSpace(txtLeadAgent.Text) Then
        '    ValidationMessage.AppendLine("- Lead Agent must be entered.")
        '    isValid = False
        'End If

        'If String.IsNullOrWhiteSpace(txtBookingDate.Value) Then
        '    ValidationMessage.AppendLine("- Booking Date must be entered.")
        '    isValid = False
        'End If

        'If String.IsNullOrWhiteSpace(TxtTranType.Text) Then
        '    ValidationMessage.AppendLine("- Transaction Type must be entered.")
        '    isValid = False
        'End If

        'If Not isValid Then
        '    MessageBox.Show("Mandatory fields needing to be filled. " & Environment.NewLine() & ValidationMessage.ToString(), "Bookings Validation - Validation", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'End If

        Return isValid
    End Function

    Private Sub InsertBooking()
        Try
            Dim newBooking As New clsBookings() With {
        .LeadAgent = txtLeadAgent.Text,
        .BookingDate = txtBookingDate.Value,
        .DivisionType = CmbDivisionType.Text,
        .TransactionType = If(Not String.IsNullOrWhiteSpace(CmbTransactionType.SelectedItem), CmbTransactionType.SelectedItem, ""),
    .TotalConsidation = If(Decimal.TryParse(TxtTotConsidation.Text, Nothing), Decimal.Parse(TxtTotConsidation.Text), 0),
        .TotalSf = If(Decimal.TryParse(TxtTotSf.Text, Nothing), Decimal.Parse(TxtTotSf.Text), 0),
        .PropertyName = If(Not String.IsNullOrWhiteSpace(TxtPropInfoName.Text), TxtPropInfoName.Text, ""),
        .SubMarketID = Convert.ToInt32(ComboBoxSubMarkets.SelectedValue),
        .PropertyAddress = If(Not String.IsNullOrWhiteSpace(TxtPropInfoAddress.Text), TxtPropInfoAddress.Text, ""),
        .PropertyCity = If(Not String.IsNullOrWhiteSpace(TxtPropInfoCity.Text), TxtPropInfoCity.Text, ""),
        .PropertyState = If(Not String.IsNullOrWhiteSpace(TxtPropInfoState.Text), TxtPropInfoState.Text, ""),
        .PropertyZIP = If(Not String.IsNullOrWhiteSpace(TxtPropInfoZip.Text), TxtPropInfoZip.Text, ""),
        .ListClientID = Convert.ToInt32(ComboBoxClients.SelectedValue),
        .ListClientAddressID = Convert.ToInt32(CmbCIAttention1.SelectedValue),
        .ListFirmID = ComboBoxLF1.SelectedValue,
        .ListFirmAddressID = CmbBFIAttention1.SelectedValue,
        .ListIsLeeArizona = checkBoxLEELF.Checked,
        .AddRefListFirmID = ComboBoxLF2.SelectedValue,
        .AddRefListFirmAddressID = CmbBFIAttention3.SelectedValue,
        .ListTotalFees = If(Decimal.TryParse(TxtLFTotalFee.Text, Nothing), Decimal.Parse(TxtLFTotalFee.Text), 0),
        .ListTotalFeesManual = If(String.IsNullOrWhiteSpace(TxtListingManual.Text), Nothing, If(Decimal.TryParse(TxtListingManual.Text, Nothing), Decimal.Parse(TxtListingManual.Text), 0)),
        .ListCommRate = If(Decimal.TryParse(TxtListingCommRate.Text, Nothing), Decimal.Parse(TxtListingCommRate.Text), 0),
        .ListOutsideBroker1 = If(Decimal.TryParse(TxtLFOutBroker1.Text, Nothing), Decimal.Parse(TxtLFOutBroker1.Text), 0),
        .ListOutsideBroker2 = If(Decimal.TryParse(TxtLFOutBroker2.Text, Nothing), Decimal.Parse(TxtLFOutBroker2.Text), 0),
        .ListLeeGrossComm = If(Decimal.TryParse(TxtLFLeeGrossComm.Text, Nothing), Decimal.Parse(TxtLFLeeGrossComm.Text), 0),
        .ListGrossAmt1 = If(Decimal.TryParse(TxtLFGrossAmt1.Text, Nothing), Decimal.Parse(TxtLFGrossAmt1.Text), 0),
        .ListGrossAmt2 = If(Decimal.TryParse(TxtLFGrossAmt2.Text, Nothing), Decimal.Parse(TxtLFGrossAmt2.Text), 0),
        .ListGrossAmt3 = If(Decimal.TryParse(TxtLFGrossAmt3.Text, Nothing), Decimal.Parse(TxtLFGrossAmt3.Text), 0),
        .ListGrossAmt4 = If(Decimal.TryParse(TxtLFGrossAmt4.Text, Nothing), Decimal.Parse(TxtLFGrossAmt4.Text), 0),
        .ListGrossAmt5 = If(Decimal.TryParse(TxtLFGrossAmt5.Text, Nothing), Decimal.Parse(TxtLFGrossAmt5.Text), 0),
        .ListGrossAmt6 = If(Decimal.TryParse(TxtLFGrossAmt6.Text, Nothing), Decimal.Parse(TxtLFGrossAmt6.Text), 0),
        .ListLeeAgent1 = If(Not String.IsNullOrWhiteSpace(TxtLFLEEAgt1.Text), TxtLFLEEAgt1.Text, ""),
        .ListLeeAgent2 = If(Not String.IsNullOrWhiteSpace(TxtLFLEEAgt2.Text), TxtLFLEEAgt2.Text, ""),
        .ListLeeAgent3 = If(Not String.IsNullOrWhiteSpace(TxtLFLEEAgt3.Text), TxtLFLEEAgt3.Text, ""),
        .ListLeeAgent4 = If(Not String.IsNullOrWhiteSpace(TxtLFLEEAgt4.Text), TxtLFLEEAgt4.Text, ""),
        .ListLeeAgent5 = If(Not String.IsNullOrWhiteSpace(TxtLFLEEAgt5.Text), TxtLFLEEAgt5.Text, ""),
        .ListLeeAgent6 = If(Not String.IsNullOrWhiteSpace(TxtLFLEEAgt6.Text), TxtLFLEEAgt6.Text, ""),
        .ListSplitPct1 = If(Decimal.TryParse(TxtLFSplitPct1.Text, Nothing), Decimal.Parse(TxtLFSplitPct1.Text), 0),
        .ListSplitPct2 = If(Decimal.TryParse(TxtLFSplitPct2.Text, Nothing), Decimal.Parse(TxtLFSplitPct2.Text), 0),
        .ListSplitPct3 = If(Decimal.TryParse(TxtLFSplitPct3.Text, Nothing), Decimal.Parse(TxtLFSplitPct3.Text), 0),
        .ListSplitPct4 = If(Decimal.TryParse(TxtLFSplitPct4.Text, Nothing), Decimal.Parse(TxtLFSplitPct4.Text), 0),
        .ListSplitPct5 = If(Decimal.TryParse(TxtLFSplitPct5.Text, Nothing), Decimal.Parse(TxtLFSplitPct5.Text), 0),
        .ListSplitPct6 = If(Decimal.TryParse(TxtLFSplitPct6.Text, Nothing), Decimal.Parse(TxtLFSplitPct6.Text), 0),
        .ProcClientID = ComboBoxClients2.SelectedValue,
        .ProcClientAddressID = Convert.ToInt32(CmbCIAttention2.SelectedValue), ' Assuming this is a fixed value or should be retrieved from UI
        .ProcIsPropertyAddress = checkBoxPropAddress.Checked,
        .ProcFirmID = ComboBoxPF1.SelectedValue,
        .ProcFirmAddressID = CmbBFIAttention2.SelectedValue,
        .ProcIsLeeArizona = CheckBoxLEEPF.Checked,
        .AddRefProcFirmID = ComboBoxPF2.SelectedValue,
        .AddRefProcFirmAddressID = CmbBFIAttention4.SelectedValue,
        .ProcTotalFees = If(Decimal.TryParse(TxtPFTotalFee.Text, Nothing), Decimal.Parse(TxtPFTotalFee.Text), 0),
        .ProcTotalFeesManual = If(String.IsNullOrWhiteSpace(TxtProcuringManual.Text), Nothing, If(Decimal.TryParse(TxtProcuringManual.Text, Nothing), Decimal.Parse(TxtProcuringManual.Text), 0)),
        .ProcCommRate = If(Decimal.TryParse(TxtProcuringCommRate.Text, Nothing), Decimal.Parse(TxtProcuringCommRate.Text), 0),
        .ProcOutsideBroker1 = If(Decimal.TryParse(TxtPFOutBroker1.Text, Nothing), Decimal.Parse(TxtPFOutBroker1.Text), 0),
        .ProcOutsideBroker2 = If(Decimal.TryParse(TxtPFOutBroker2.Text, Nothing), Decimal.Parse(TxtPFOutBroker2.Text), 0),
        .ProcLeeGrossComm = If(Decimal.TryParse(TxtPFLeeGrossComm.Text, Nothing), Decimal.Parse(TxtPFLeeGrossComm.Text), 0),
        .ProcGrossAmt1 = If(Decimal.TryParse(TxtPFGrossAmt1.Text, Nothing), Decimal.Parse(TxtPFGrossAmt1.Text), 0),
        .ProcGrossAmt2 = If(Decimal.TryParse(TxtPFGrossAmt2.Text, Nothing), Decimal.Parse(TxtPFGrossAmt2.Text), 0),
        .ProcGrossAmt3 = If(Decimal.TryParse(TxtPFGrossAmt3.Text, Nothing), Decimal.Parse(TxtPFGrossAmt3.Text), 0),
        .ProcGrossAmt4 = If(Decimal.TryParse(TxtPFGrossAmt4.Text, Nothing), Decimal.Parse(TxtPFGrossAmt4.Text), 0),
        .ProcGrossAmt5 = If(Decimal.TryParse(TxtPFGrossAmt5.Text, Nothing), Decimal.Parse(TxtPFGrossAmt5.Text), 0),
        .ProcGrossAmt6 = If(Decimal.TryParse(TxtPFGrossAmt6.Text, Nothing), Decimal.Parse(TxtPFGrossAmt6.Text), 0),
        .ProcLeeAgent1 = If(Not String.IsNullOrWhiteSpace(TxtPFLEEAgt1.Text), TxtPFLEEAgt1.Text, ""),
        .ProcLeeAgent2 = If(Not String.IsNullOrWhiteSpace(TxtPFLEEAgt2.Text), TxtPFLEEAgt2.Text, ""),
        .ProcLeeAgent3 = If(Not String.IsNullOrWhiteSpace(TxtPFLEEAgt3.Text), TxtPFLEEAgt3.Text, ""),
        .ProcLeeAgent4 = If(Not String.IsNullOrWhiteSpace(TxtPFLEEAgt4.Text), TxtPFLEEAgt4.Text, ""),
        .ProcLeeAgent5 = If(Not String.IsNullOrWhiteSpace(TxtPFLEEAgt5.Text), TxtPFLEEAgt5.Text, ""),
        .ProcLeeAgent6 = If(Not String.IsNullOrWhiteSpace(TxtPFLEEAgt6.Text), TxtPFLEEAgt6.Text, ""),
        .ProcSplitPct1 = If(Decimal.TryParse(TxtPFSplitPct1.Text, Nothing), Decimal.Parse(TxtPFSplitPct1.Text), 0),
        .ProcSplitPct2 = If(Decimal.TryParse(TxtPFSplitPct2.Text, Nothing), Decimal.Parse(TxtPFSplitPct2.Text), 0),
        .ProcSplitPct3 = If(Decimal.TryParse(TxtPFSplitPct3.Text, Nothing), Decimal.Parse(TxtPFSplitPct3.Text), 0),
        .ProcSplitPct4 = If(Decimal.TryParse(TxtPFSplitPct4.Text, Nothing), Decimal.Parse(TxtPFSplitPct4.Text), 0),
        .ProcSplitPct5 = If(Decimal.TryParse(TxtPFSplitPct5.Text, Nothing), Decimal.Parse(TxtPFSplitPct5.Text), 0),
        .ProcSplitPct6 = If(Decimal.TryParse(TxtPFSplitPct6.Text, Nothing), Decimal.Parse(TxtPFSplitPct6.Text), 0),
        .BillListBroker = ChkBillListing.Checked,
        .BillClient = ChkBillClient.Checked,
        .BillOther = ChkBillOther.Checked,''CheckBoxBillOther.Checked
         .BillCompany = TxtBillCompany.Text,
         .BillAttention = TxtBillInfoAttention.Text,
         .BillAddress = TxtBillInfoAddress.Text,
         .BillCity = TxtBillInfoCity.Text,
         .BillState = TxtBillInfoState.Text,
         .BillZIP = TxtBillInfoZip.Text,
         .BillPhoneNo = TxtBillInfoPhnNo.Text,
         .WireInstructionsSend = CheckBoxSendWireInst.Checked,
         .EscrowNumber = TxtEscrowNo.Text, ' Assuming this is a fixed value or should be retrieved from UI
         .DoNotSendInvoice = CheckBoxDontSndInv.Checked,
         .InvoiceMail = CheckBoxSndMailInv.Checked,
         .InvoiceEmail = CheckBoxSendEmailInv.Checked,
         .DueDate1 = dtNewDate1.Value,
         .DuePercent1 = If(Decimal.TryParse(TxtDuePct1.Text, Nothing), Decimal.Parse(TxtDuePct1.Text), 0),
         .DueDate2 = dtNewDate2.Value,
         .DuePercent2 = If(Decimal.TryParse(TxtDuePct2.Text, Nothing), Decimal.Parse(TxtDuePct2.Text), 0),
         .DueDate3 = dtNewDate3.Value,
         .DuePercent3 = If(Decimal.TryParse(TxtDuePct3.Text, Nothing), Decimal.Parse(TxtDuePct3.Text), 0),
         .DueDate4 = dtNewDate4.Value,
         .DuePercent4 = If(Decimal.TryParse(TxtDuePct4.Text, Nothing), Decimal.Parse(TxtDuePct4.Text), 0),
         .BillEmailAddress1 = TxtBillEmail1.Text,
         .BillEmailAddress2 = TxtBillEmail2.Text,
         .BillEmailAddress3 = TxtBillEmail3.Text,
         .LeaseType = If(Not String.IsNullOrWhiteSpace(ComboBoxLeaseType.SelectedItem), ComboBoxLeaseType.SelectedItem, ""),
         .LeaseRateType = If(Not String.IsNullOrWhiteSpace(CmbLeaseRateType.SelectedItem), CmbLeaseRateType.SelectedItem, ""),
         .LeaseTermMonths = If(Decimal.TryParse(TxtLeaseTerm.Text, Nothing), Decimal.Parse(TxtLeaseTerm.Text), 0),
         .LeasedSF = If(Not String.IsNullOrWhiteSpace(TxtLeasedSF.Text), TxtLeasedSF.Text, ""),
         .RSFCalculator = If(Not String.IsNullOrWhiteSpace(cmbRSFCalculator.SelectedItem), cmbRSFCalculator.SelectedItem, ""),
         .LeaseConsideration = If(Decimal.TryParse(TxtLeaseConsid.Text, Nothing), Decimal.Parse(TxtLeaseConsid.Text), 0),
         .CommencementDate = DTLeaseCommence.Value,
         .ExpDate = DTLeaseExpiry.Value,
         .OccupancyDate = DTLeaseOccupancy.Value,
         .TIAllowance = If(Decimal.TryParse(TxtLeaseTIAllow.Text, Nothing), Decimal.Parse(TxtLeaseTIAllow.Text), 0),
         .ParkingRatio = If(Decimal.TryParse(TxtParkingRatio.Text, Nothing), Decimal.Parse(TxtParkingRatio.Text), 0),
         .ReservedAmt = If(Decimal.TryParse(TxtReservedAmt.Text, Nothing), Decimal.Parse(TxtReservedAmt.Text), 0),
         .UnreservedAmt = If(Decimal.TryParse(TxtUnReservedAmt.Text, Nothing), Decimal.Parse(TxtUnReservedAmt.Text), 0),
         .RooftopAmt = If(Decimal.TryParse(TxtRoofTopAmt.Text, Nothing), Decimal.Parse(TxtRoofTopAmt.Text), 0),
         .LeaseNotes = If(Not String.IsNullOrWhiteSpace(TxtLeaseNotes.Text), TxtLeaseNotes.Text, ""),
         .ListCommAddition = If(Decimal.TryParse(TxtLFCommAdd.Text, Nothing), Decimal.Parse(TxtLFCommAdd.Text), 0),
         .ListCommDeduction = If(Decimal.TryParse(TxtLFCommDed.Text, Nothing), Decimal.Parse(TxtLFCommDed.Text), 0),
         .ProcCommAddition = If(Decimal.TryParse(TxtPFCommAdd.Text, Nothing), Decimal.Parse(TxtPFCommAdd.Text), 0),
         .ProcCommDeduction = If(Decimal.TryParse(TxtPFCommDed.Text, Nothing), Decimal.Parse(TxtPFCommDed.Text), 0),
         .createdBy = clsGlobals.gblLoginName,
         .createdOn = DateTime.Now
    }
            ' .lastModifiedBy = "",
            ' .lastModifiedOn = DateTime.Now

            ' Insert the new booking object into the collection or database
            objBookingsBal.Insert(newBooking)

            MsgBox("New Booking has been successfully added!", MsgBoxStyle.Information, "Success")
            LoadAllData(SearchQuery)
            'ClearFieldValues()
            TabBookings.SelectedTab = TabPageBL
        Catch ex As Exception
            MsgBox($"An error occurred: {ex.Message}", MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateBooking()
        Dim _ProcTotalFeesManual As Decimal? = Nothing
        Dim _ListTotalFeesManual As Decimal? = Nothing

        ' If(Not String.IsNullOrWhiteSpace(TxtPropInfoCity.Text), TxtPropInfoCity.Text, "")
        '  If(String.IsNullOrWhiteSpace(TxtListingManual.Text), Nothing, Decimal.Parse(TxtListingManual.Text))
        If String.IsNullOrWhiteSpace(TxtListingManual.Text) Then
            _ListTotalFeesManual = Nothing
        Else
            _ListTotalFeesManual = Decimal.Parse(TxtListingManual.Text)
        End If

        ' If(String.IsNullOrWhiteSpace(TxtProcuringManual.Text), Nothing, Decimal.Parse(TxtProcuringManual.Text))
        If String.IsNullOrWhiteSpace(TxtProcuringManual.Text) Then
            _ProcTotalFeesManual = Nothing
        Else
            _ProcTotalFeesManual = Decimal.Parse(TxtProcuringManual.Text)
        End If
        Dim _TotalConsidation As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtTotConsidation.Text) Then
            _TotalConsidation = Nothing
        Else
            _TotalConsidation = Decimal.Parse(TxtTotConsidation.Text)
        End If

        Dim _TotalSf As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtTotSf.Text) Then
            _TotalSf = Nothing
        Else
            _TotalSf = Decimal.Parse(TxtTotSf.Text)
        End If

        Dim _SubMarketID As Decimal? = Nothing
        If ComboBoxSubMarkets.SelectedValue Is Nothing Then
            _SubMarketID = Nothing
        Else
            _SubMarketID = Convert.ToInt32(ComboBoxSubMarkets.SelectedValue)
        End If

        Dim _ListClientID As Decimal? = Nothing
        If ComboBoxClients.SelectedValue Is Nothing Then
            _ListClientID = Nothing
        Else
            _ListClientID = Convert.ToInt32(ComboBoxClients.SelectedValue)
        End If

        Dim _ListClientAddressID As Decimal? = Nothing
        If CmbCIAttention1.SelectedValue Is Nothing Then
            _ListClientAddressID = Nothing
        Else
            _ListClientAddressID = Convert.ToInt32(CmbCIAttention1.SelectedValue)
        End If

        Dim _ListTotalFees As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtLFTotalFee.Text) Then
            _ListTotalFees = Nothing
        Else
            _ListTotalFees = Decimal.Parse(TxtLFTotalFee.Text)
        End If

        Dim _ListCommRate As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtListingCommRate.Text) Then
            _ListCommRate = Nothing
        Else
            _ListCommRate = Decimal.Parse(TxtListingCommRate.Text)
        End If

        Dim _ListOutsideBroker1 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtLFOutBroker1.Text) Then
            _ListOutsideBroker1 = Nothing
        Else
            _ListOutsideBroker1 = Decimal.Parse(TxtLFOutBroker1.Text)
        End If

        Dim _ListOutsideBroker2 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtLFOutBroker2.Text) Then
            _ListOutsideBroker2 = Nothing
        Else
            _ListOutsideBroker2 = Decimal.Parse(TxtLFOutBroker2.Text)
        End If

        Dim _ListLeeGrossComm As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtLFLeeGrossComm.Text) Then
            _ListLeeGrossComm = Nothing
        Else
            _ListLeeGrossComm = Decimal.Parse(TxtLFLeeGrossComm.Text)
        End If

        Dim _ListGrossAmt1 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtLFGrossAmt1.Text) Then
            _ListGrossAmt1 = Nothing
        Else
            _ListGrossAmt1 = Decimal.Parse(TxtLFGrossAmt1.Text)
        End If

        Dim _ListGrossAmt2 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtLFGrossAmt2.Text) Then
            _ListGrossAmt2 = Nothing
        Else
            _ListGrossAmt2 = Decimal.Parse(TxtLFGrossAmt2.Text)
        End If

        Dim _ListGrossAmt3 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtLFGrossAmt3.Text) Then
            _ListGrossAmt3 = Nothing
        Else
            _ListGrossAmt3 = Decimal.Parse(TxtLFGrossAmt3.Text)
        End If

        Dim _ListGrossAmt4 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtLFGrossAmt4.Text) Then
            _ListGrossAmt4 = Nothing
        Else
            _ListGrossAmt4 = Decimal.Parse(TxtLFGrossAmt4.Text)
        End If

        Dim _ListGrossAmt5 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtLFGrossAmt5.Text) Then
            _ListGrossAmt5 = Nothing
        Else
            _ListGrossAmt5 = Decimal.Parse(TxtLFGrossAmt5.Text)
        End If

        Dim _ListGrossAmt6 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtLFGrossAmt6.Text) Then
            _ListGrossAmt6 = Nothing
        Else
            _ListGrossAmt6 = Decimal.Parse(TxtLFGrossAmt6.Text)
        End If

        Dim _ListSplitPct1 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtLFSplitPct1.Text) Then
            _ListSplitPct1 = Nothing
        Else
            _ListSplitPct1 = Decimal.Parse(TxtLFSplitPct1.Text)
        End If

        Dim _ListSplitPct2 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtLFSplitPct2.Text) Then
            _ListSplitPct2 = Nothing
        Else
            _ListSplitPct2 = Decimal.Parse(TxtLFSplitPct2.Text)
        End If

        Dim _ListSplitPct3 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtLFSplitPct3.Text) Then
            _ListSplitPct3 = Nothing
        Else
            _ListSplitPct3 = Decimal.Parse(TxtLFSplitPct3.Text)
        End If

        Dim _ListSplitPct4 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtLFSplitPct4.Text) Then
            _ListSplitPct4 = Nothing
        Else
            _ListSplitPct4 = Decimal.Parse(TxtLFSplitPct4.Text)
        End If

        Dim _ListSplitPct5 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtLFSplitPct5.Text) Then
            _ListSplitPct5 = Nothing
        Else
            _ListSplitPct5 = Decimal.Parse(TxtLFSplitPct5.Text)
        End If

        Dim _ListSplitPct6 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtLFSplitPct6.Text) Then
            _ListSplitPct6 = Nothing
        Else
            _ListSplitPct6 = Decimal.Parse(TxtLFSplitPct6.Text)
        End If

        Dim _ProcClientAddressID As Decimal? = Nothing
        If CmbCIAttention2.SelectedValue Is Nothing Then
            _ProcClientAddressID = Nothing
        Else
            _ProcClientAddressID = Convert.ToInt32(CmbCIAttention2.SelectedValue)
        End If

        Dim _ProcTotalFees As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtPFTotalFee.Text) Then
            _ProcTotalFees = Nothing
        Else
            _ProcTotalFees = Decimal.Parse(TxtPFTotalFee.Text)
        End If

        Dim _ProcCommRate As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtProcuringCommRate.Text) Then
            _ProcCommRate = Nothing
        Else
            _ProcCommRate = Decimal.Parse(TxtProcuringCommRate.Text)
        End If

        Dim _ProcOutsideBroker1 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtPFOutBroker1.Text) Then
            _ProcOutsideBroker1 = Nothing
        Else
            _ProcOutsideBroker1 = Decimal.Parse(TxtPFOutBroker1.Text)
        End If

        Dim _ProcOutsideBroker2 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtPFOutBroker2.Text) Then
            _ProcOutsideBroker2 = Nothing
        Else
            _ProcOutsideBroker2 = Decimal.Parse(TxtPFOutBroker2.Text)
        End If

        Dim _ProcLeeGrossComm As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtPFLeeGrossComm.Text) Then
            _ProcLeeGrossComm = Nothing
        Else
            _ProcLeeGrossComm = Decimal.Parse(TxtPFLeeGrossComm.Text)
        End If

        Dim _ProcGrossAmt1 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtPFGrossAmt1.Text) Then
            _ProcGrossAmt1 = Nothing
        Else
            _ProcGrossAmt1 = Decimal.Parse(TxtPFGrossAmt1.Text)
        End If

        Dim _ProcGrossAmt2 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtPFGrossAmt2.Text) Then
            _ProcGrossAmt2 = Nothing
        Else
            _ProcGrossAmt2 = Decimal.Parse(TxtPFGrossAmt2.Text)
        End If

        Dim _ProcGrossAmt3 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtPFGrossAmt3.Text) Then
            _ProcGrossAmt3 = Nothing
        Else
            _ProcGrossAmt3 = Decimal.Parse(TxtPFGrossAmt3.Text)
        End If

        Dim _ProcGrossAmt4 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtPFGrossAmt4.Text) Then
            _ProcGrossAmt4 = Nothing
        Else
            _ProcGrossAmt4 = Decimal.Parse(TxtPFGrossAmt4.Text)
        End If

        Dim _ProcGrossAmt5 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtPFGrossAmt5.Text) Then
            _ProcGrossAmt5 = Nothing
        Else
            _ProcGrossAmt5 = Decimal.Parse(TxtPFGrossAmt5.Text)
        End If

        Dim _ProcGrossAmt6 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtPFGrossAmt6.Text) Then
            _ProcGrossAmt6 = Nothing
        Else
            _ProcGrossAmt6 = Decimal.Parse(TxtPFGrossAmt6.Text)
        End If

        Dim _ProcSplitPct1 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtPFSplitPct1.Text) Then
            _ProcSplitPct1 = Nothing
        Else
            _ProcSplitPct1 = Decimal.Parse(TxtPFSplitPct1.Text)
        End If

        Dim _ProcSplitPct2 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtPFSplitPct2.Text) Then
            _ProcSplitPct2 = Nothing
        Else
            _ProcSplitPct2 = Decimal.Parse(TxtPFSplitPct2.Text)
        End If

        Dim _ProcSplitPct3 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtPFSplitPct3.Text) Then
            _ProcSplitPct3 = Nothing
        Else
            _ProcSplitPct3 = Decimal.Parse(TxtPFSplitPct3.Text)
        End If

        Dim _ProcSplitPct4 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtPFSplitPct4.Text) Then
            _ProcSplitPct4 = Nothing
        Else
            _ProcSplitPct4 = Decimal.Parse(TxtPFSplitPct4.Text)
        End If

        Dim _ProcSplitPct5 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtPFSplitPct5.Text) Then
            _ProcSplitPct5 = Nothing
        Else
            _ProcSplitPct5 = Decimal.Parse(TxtPFSplitPct5.Text)
        End If

        Dim _ProcSplitPct6 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtPFSplitPct6.Text) Then
            _ProcSplitPct6 = Nothing
        Else
            _ProcSplitPct6 = Decimal.Parse(TxtPFSplitPct6.Text)
        End If

        Dim _DuePercent1 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtDuePct1.Text) Then
            _DuePercent1 = Nothing
        Else
            _DuePercent1 = Decimal.Parse(TxtDuePct1.Text)
        End If

        Dim _DuePercent2 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtDuePct2.Text) Then
            _DuePercent2 = Nothing
        Else
            _DuePercent2 = Decimal.Parse(TxtDuePct2.Text)
        End If

        Dim _DuePercent3 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtDuePct3.Text) Then
            _DuePercent3 = Nothing
        Else
            _DuePercent3 = Decimal.Parse(TxtDuePct3.Text)
        End If

        Dim _DuePercent4 As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtDuePct4.Text) Then
            _DuePercent4 = Nothing
        Else
            _DuePercent4 = Decimal.Parse(TxtDuePct4.Text)
        End If

        Dim _LeaseTermMonths As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtLeaseTerm.Text) Then
            _LeaseTermMonths = Nothing
        Else
            _LeaseTermMonths = Decimal.Parse(TxtLeaseTerm.Text)
        End If

        Dim _LeaseConsideration As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtLeaseConsid.Text) Then
            _LeaseConsideration = Nothing
        Else
            _LeaseConsideration = Decimal.Parse(TxtLeaseConsid.Text)
        End If

        Dim _TIAllowance As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtLeaseTIAllow.Text) Then
            _TIAllowance = Nothing
        Else
            _TIAllowance = Decimal.Parse(TxtLeaseTIAllow.Text)
        End If

        Dim _ParkingRatio As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtParkingRatio.Text) Then
            _ParkingRatio = Nothing
        Else
            _ParkingRatio = Decimal.Parse(TxtParkingRatio.Text)
        End If

        Dim _ReservedAmt As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtReservedAmt.Text) Then
            _ReservedAmt = Nothing
        Else
            _ReservedAmt = Decimal.Parse(TxtReservedAmt.Text)
        End If

        Dim _UnreservedAmt As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtUnReservedAmt.Text) Then
            _UnreservedAmt = Nothing
        Else
            _UnreservedAmt = Decimal.Parse(TxtUnReservedAmt.Text)
        End If

        Dim _RooftopAmt As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtRoofTopAmt.Text) Then
            _RooftopAmt = Nothing
        Else
            _RooftopAmt = Decimal.Parse(TxtRoofTopAmt.Text)
        End If

        Dim _ListCommAddition As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtLFCommAdd.Text) Then
            _ListCommAddition = Nothing
        Else
            _ListCommAddition = Decimal.Parse(TxtLFCommAdd.Text)
        End If

        Dim _ListCommDeduction As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtLFCommDed.Text) Then
            _ListCommDeduction = Nothing
        Else
            _ListCommDeduction = Decimal.Parse(TxtLFCommDed.Text)
        End If


        Dim _ProcCommAddition As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtPFCommAdd.Text) Then
            _ProcCommAddition = Nothing
        Else
            _ProcCommAddition = Decimal.Parse(TxtPFCommAdd.Text)
        End If

        Dim _ProcCommDeduction As Decimal? = Nothing
        If String.IsNullOrWhiteSpace(TxtPFCommDed.Text) Then
            _ProcCommDeduction = Nothing
        Else
            _ProcCommDeduction = Decimal.Parse(TxtPFCommDed.Text)
        End If
        Try
            Dim existingBooking As New clsBookings() With {
            .BookingID = BookingID,
            .LeadAgent = txtLeadAgent.Text,
            .BookingDate = txtBookingDate.Value,
            .DivisionType = CmbDivisionType.Text,
            .TransactionType = If(Not String.IsNullOrWhiteSpace(CmbTransactionType.SelectedItem), CmbTransactionType.SelectedItem, ""),
            .TotalConsidation = _TotalConsidation,
            .TotalSf = _TotalSf,
            .PropertyName = If(Not String.IsNullOrWhiteSpace(TxtPropInfoName.Text), TxtPropInfoName.Text, ""),
            .SubMarketID = Convert.ToInt32(ComboBoxSubMarkets.SelectedValue),
            .PropertyAddress = If(Not String.IsNullOrWhiteSpace(TxtPropInfoAddress.Text), TxtPropInfoAddress.Text, ""),
            .PropertyCity = TxtPropInfoCity.Text,
            .PropertyState = If(Not String.IsNullOrWhiteSpace(TxtPropInfoState.Text), TxtPropInfoState.Text, ""),
            .PropertyZIP = If(Not String.IsNullOrWhiteSpace(TxtPropInfoZip.Text), TxtPropInfoZip.Text, ""),
            .ListClientID = Convert.ToInt32(ComboBoxClients.SelectedValue),
            .ListClientAddressID = Convert.ToInt32(CmbCIAttention1.SelectedValue),
            .ListFirmID = ComboBoxLF1.SelectedValue,
            .ListFirmAddressID = CmbBFIAttention1.SelectedValue,
            .ListIsLeeArizona = checkBoxLEELF.Checked,
            .AddRefListFirmID = ComboBoxLF2.SelectedValue,
            .AddRefListFirmAddressID = CmbBFIAttention3.SelectedValue,
            .ListTotalFees = _ListTotalFees,
            .ListTotalFeesManual = _ListTotalFeesManual,
            .ListCommRate = _ListCommRate,
            .ListOutsideBroker1 = _ListOutsideBroker1,
            .ListOutsideBroker2 = _ListOutsideBroker2,
            .ListLeeGrossComm = _ListLeeGrossComm,
            .ListGrossAmt1 = _ListGrossAmt1,
            .ListGrossAmt2 = _ListGrossAmt2,
            .ListGrossAmt3 = _ListGrossAmt3,
            .ListGrossAmt4 = _ListGrossAmt4,
            .ListGrossAmt5 = _ListGrossAmt5,
            .ListGrossAmt6 = _ListGrossAmt6,
            .ListLeeAgent1 = If(Not String.IsNullOrWhiteSpace(TxtLFLEEAgt1.Text), TxtLFLEEAgt1.Text, ""),
            .ListLeeAgent2 = If(Not String.IsNullOrWhiteSpace(TxtLFLEEAgt2.Text), TxtLFLEEAgt2.Text, ""),
            .ListLeeAgent3 = If(Not String.IsNullOrWhiteSpace(TxtLFLEEAgt3.Text), TxtLFLEEAgt3.Text, ""),
            .ListLeeAgent4 = If(Not String.IsNullOrWhiteSpace(TxtLFLEEAgt4.Text), TxtLFLEEAgt4.Text, ""),
            .ListLeeAgent5 = If(Not String.IsNullOrWhiteSpace(TxtLFLEEAgt5.Text), TxtLFLEEAgt5.Text, ""),
            .ListLeeAgent6 = If(Not String.IsNullOrWhiteSpace(TxtLFLEEAgt6.Text), TxtLFLEEAgt6.Text, ""),
            .ListSplitPct1 = _ListSplitPct1,
            .ListSplitPct2 = _ListSplitPct2,
            .ListSplitPct3 = _ListSplitPct3,
            .ListSplitPct4 = _ListSplitPct4,
            .ListSplitPct5 = _ListSplitPct5,
            .ListSplitPct6 = _ListSplitPct6,
            .ProcClientID = ComboBoxClients2.SelectedValue,
            .ProcClientAddressID = Convert.ToInt32(CmbCIAttention2.SelectedValue),
            .ProcIsPropertyAddress = checkBoxPropAddress.Checked,
            .ProcFirmID = ComboBoxPF1.SelectedValue,
            .ProcFirmAddressID = CmbBFIAttention2.SelectedValue,
            .ProcIsLeeArizona = CheckBoxLEEPF.Checked,
            .AddRefProcFirmID = ComboBoxPF2.SelectedValue,
            .AddRefProcFirmAddressID = CmbBFIAttention4.SelectedValue,
            .ProcTotalFees = _ProcTotalFees,
            .ProcTotalFeesManual = _ProcTotalFeesManual,
            .ProcCommRate = _ProcCommRate,
            .ProcOutsideBroker1 = _ProcOutsideBroker1,
            .ProcOutsideBroker2 = _ProcOutsideBroker2,
            .ProcLeeGrossComm = _ProcLeeGrossComm,
            .ProcGrossAmt1 = _ProcGrossAmt1,
            .ProcGrossAmt2 = _ProcGrossAmt2,
            .ProcGrossAmt3 = _ProcGrossAmt3,
            .ProcGrossAmt4 = _ProcGrossAmt4,
            .ProcGrossAmt5 = _ProcGrossAmt5,
            .ProcGrossAmt6 = _ProcGrossAmt6,
            .ProcLeeAgent1 = If(Not String.IsNullOrWhiteSpace(TxtPFLEEAgt1.Text), TxtPFLEEAgt1.Text, ""),
            .ProcLeeAgent2 = If(Not String.IsNullOrWhiteSpace(TxtPFLEEAgt2.Text), TxtPFLEEAgt2.Text, ""),
            .ProcLeeAgent3 = If(Not String.IsNullOrWhiteSpace(TxtPFLEEAgt3.Text), TxtPFLEEAgt3.Text, ""),
            .ProcLeeAgent4 = If(Not String.IsNullOrWhiteSpace(TxtPFLEEAgt4.Text), TxtPFLEEAgt4.Text, ""),
            .ProcLeeAgent5 = If(Not String.IsNullOrWhiteSpace(TxtPFLEEAgt5.Text), TxtPFLEEAgt5.Text, ""),
            .ProcLeeAgent6 = If(Not String.IsNullOrWhiteSpace(TxtPFLEEAgt6.Text), TxtPFLEEAgt6.Text, ""),
            .ProcSplitPct1 = _ProcSplitPct1,
            .ProcSplitPct2 = _ProcSplitPct2,
            .ProcSplitPct3 = _ProcSplitPct3,
            .ProcSplitPct4 = _ProcSplitPct4,
            .ProcSplitPct5 = _ProcSplitPct5,
            .ProcSplitPct6 = _ProcSplitPct6,
            .BillListBroker = ChkBillListing.Checked,
            .BillClient = ChkBillClient.Checked,
            .BillOther = ChkBillOther.Checked,
            .BillCompany = TxtBillCompany.Text,
            .BillAttention = TxtBillInfoAttention.Text,
            .BillAddress = TxtBillInfoAddress.Text,
            .BillCity = TxtBillInfoCity.Text,
            .BillState = TxtBillInfoState.Text,
            .BillZIP = TxtBillInfoZip.Text,
            .BillPhoneNo = TxtBillInfoPhnNo.Text,
            .WireInstructionsSend = CheckBoxSendWireInst.Checked,
            .EscrowNumber = TxtEscrowNo.Text,
            .DoNotSendInvoice = CheckBoxDontSndInv.Checked,
            .InvoiceMail = CheckBoxSndMailInv.Checked,
            .InvoiceEmail = CheckBoxSendEmailInv.Checked,
            .DueDate1 = dtNewDate1.Value,
            .DuePercent1 = _DuePercent1,
            .DueDate2 = dtNewDate2.Value,
            .DuePercent2 = _DuePercent2,
            .DueDate3 = dtNewDate3.Value,
            .DuePercent3 = _DuePercent3,
            .DueDate4 = dtNewDate4.Value,
            .DuePercent4 = _DuePercent4,
            .BillEmailAddress1 = TxtBillEmail1.Text,
            .BillEmailAddress2 = TxtBillEmail2.Text,
            .BillEmailAddress3 = TxtBillEmail3.Text,
            .LeaseType = If(Not String.IsNullOrWhiteSpace(ComboBoxLeaseType.SelectedItem), ComboBoxLeaseType.SelectedItem.ToString(), ""),
            .LeaseRateType = If(Not String.IsNullOrWhiteSpace(CmbLeaseRateType.SelectedItem), CmbLeaseRateType.SelectedItem, ""),
            .LeaseTermMonths = _LeaseTermMonths,
            .LeasedSF = If(Not String.IsNullOrWhiteSpace(TxtLeasedSF.Text), TxtLeasedSF.Text, ""),
            .RSFCalculator = If(Not String.IsNullOrWhiteSpace(cmbRSFCalculator.SelectedItem), cmbRSFCalculator.SelectedItem, ""),
            .LeaseConsideration = _LeaseConsideration,
            .CommencementDate = DTLeaseCommence.Value,
            .ExpDate = DTLeaseExpiry.Value,
            .OccupancyDate = DTLeaseOccupancy.Value,
            .TIAllowance = _TIAllowance,
            .ParkingRatio = _ParkingRatio,
            .ReservedAmt = _ReservedAmt,
            .UnreservedAmt = _UnreservedAmt,
            .RooftopAmt = _RooftopAmt,
            .LeaseNotes = If(Not String.IsNullOrWhiteSpace(TxtLeaseNotes.Text), TxtLeaseNotes.Text, ""),
            .ListCommAddition = _ListCommAddition,
            .ListCommDeduction = _ListCommDeduction,
            .ProcCommAddition = _ProcCommAddition,
            .ProcCommDeduction = _ProcCommDeduction,
            .lastModifiedBy = clsGlobals.gblLoginName,
            .lastModifiedOn = DateTime.Now
        }
            '.createdBy = "",
            '.createdOn = DateTime.Now,

            ' Call the data access method to update the existing Booking
            objBookingsBal.Update(existingBooking)
            MessageBox.Show("Booking updated successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            LoadAllData(SearchQuery)
            TabBookings.SelectedTab = TabPageBL
        Catch ex As Exception
            ' Log the exception if necessary
            ' logger.Log(ex.ToString())
            MessageBox.Show("Failed to update the booking." & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub NavigationButtonEnableDisable()
        Try
            If grdBookings.RowCount > 0 Then
                btnPrev.Enabled = True
                btnFirst.Enabled = True
                btnNext.Enabled = True
                btnLast.Enabled = True

                If grdBookings.CurrentRow IsNot Nothing Then
                    If grdBookings.RowCount = 1 Then
                        btnPrev.Enabled = False
                        btnFirst.Enabled = False
                        btnNext.Enabled = False
                        btnLast.Enabled = False
                    ElseIf grdBookings.CurrentRow.Index = 0 Then
                        btnPrev.Enabled = False
                        btnFirst.Enabled = False
                        btnNext.Enabled = True
                        btnLast.Enabled = True
                    ElseIf grdBookings.CurrentRow.Index = grdBookings.RowCount - 1 Then
                        btnNext.Enabled = False
                        btnLast.Enabled = False
                        btnPrev.Enabled = True
                        btnFirst.Enabled = True
                    End If
                End If
            Else
                btnPrev.Enabled = False
                btnFirst.Enabled = False
                btnNext.Enabled = False
                btnLast.Enabled = False
            End If
        Catch ex As Exception
            ' Handle your exception here, log or throw as needed
            Throw ex
        End Try
    End Sub

    'Private Sub GetLeadAgent(columnName As String)
    '    Dim distinctValues As New List(Of String)()
    '    distinctValues = objBookingsBal.GetDropDownValues("Bookings", columnName)
    '    Select Case columnName.ToLower()
    '        Case "leadAgent"
    '            comboBoxLeadAgent.Items.Clear()
    '            comboBoxLeadAgent.Items.AddRange(distinctValues.ToArray())
    '            'Case "state"
    '            '    ComboBoxState.Items.Clear()
    '            '    ComboBoxState.Items.AddRange(distinctValues.ToArray())
    '    End Select
    'End Sub

    Private Sub AssignComboBox()
        GetDropDownValues("submarkets", ComboBoxSubMarkets)
        GetDropDownValues("clients", ComboBoxClients)
        GetDropDownValues("clients", ComboBoxClients2)
        GetDropDownValues("clients", ComboBoxLF1)
        GetDropDownValues("clients", ComboBoxPF1)
        GetDropDownValues("clients", ComboBoxLF2)
        GetDropDownValues("clients", ComboBoxPF2)
    End Sub

    Private Sub GetDropDownValues(columnName As String, comboBox As ComboBox)
        Select Case columnName.ToLower()
            Case "submarkets"
                Dim subMarkets As List(Of clsSubMarkets) = objBookingsBal.GetSubMarkets(CmbDivisionType.Text)
                Dim defaultSubMarket As New clsSubMarkets With {.SubMarketId = -1, .SubMarket = "Select a SubMarket"}
                SetupComboBox(comboBox, subMarkets, "SubMarket", "SubMarketId", defaultSubMarket)
            Case "clients"
                Dim clients As List(Of clsClients) = objClientsBal.GetAll()
                Dim defaultClient As New clsClients With {.clientID = -1, .clientName = "Select a Client"}
                SetupComboBox(comboBox, clients, "ClientName", "ClientId", defaultClient)
                'Case "listingfirms"
                '    Dim listingFirms As List(Of clsListingFirms) = objListingFirmBal.GetAll()
                '    Dim defaultLF As New clsListingFirms With {.listingFirmID = -1, .firmName = "Select a Listing Firm"}
                '    SetupComboBox(comboBox, listingFirms, "FirmName", "ListingFirmId", defaultLF)
                'Case "procuringfirms"
                '    Dim procuringFirms As List(Of clsProcuringFirms) = objProcuringFirmBal.GetAll()
                '    Dim defaultPF As New clsProcuringFirms With {.procuringFirmID = -1, .firmName = "Select a Procuring Firm"}
                '    SetupComboBox(comboBox, procuringFirms, "FirmName", "ProcuringFirmId", defaultPF)
        End Select
    End Sub

    Private Sub SetupComboBox(Of T)(comboBox As ComboBox, dataSource As List(Of T), displayMember As String, valueMember As String, defaultItem As T)
        dataSource.Insert(0, defaultItem)
        comboBox.DisplayMember = displayMember
        comboBox.ValueMember = valueMember
        comboBox.DataSource = dataSource
        comboBox.SelectedIndex = 0
    End Sub

    Private Sub ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim comboBox As ComboBox = CType(sender, ComboBox)
        Dim textBoxes As TextBox()
        Dim AttnComboBox As ComboBox

        If comboBox Is ComboBoxClients Then
            textBoxes = {TxtCIAddress1, TxtCICity1, TxtCIState1, TxtCIZip1}
            AttnComboBox = CmbCIAttention1
        ElseIf comboBox Is ComboBoxClients2 Then
            textBoxes = {TxtCIAddress2, TxtCICity2, TxtCIState2, TxtCIZip2}
            AttnComboBox = CmbCIAttention2
        ElseIf comboBox Is ComboBoxLF1 Then
            textBoxes = {TxtBFIAddress1, TxtBFICity1, TxtBFIState1, TxtBFIZip1}
            AttnComboBox = CmbBFIAttention1
        ElseIf comboBox Is ComboBoxLF2 Then
            textBoxes = {TxtBFIAddress3, TxtBFICity3, TxtBFIState3, TxtBFIZip3}
            AttnComboBox = CmbBFIAttention3
        ElseIf comboBox Is ComboBoxPF1 Then
            textBoxes = {TxtBFIAddress2, TxtBFICity2, TxtBFIState2, TxtBFIZip2}
            AttnComboBox = CmbBFIAttention2
        ElseIf comboBox Is ComboBoxPF2 Then
            textBoxes = {TxtBFIAddress4, TxtBFICity4, TxtBFIState4, TxtBFIZip4}
            AttnComboBox = CmbBFIAttention4
        Else
            Return ' Exit if it's an unknown ComboBox
        End If

        ClearFields(textBoxes, AttnComboBox)

        Try
            Dim selectedClientId As Integer = CType(comboBox.SelectedValue, Integer)
            If selectedClientId = -1 Then
                ' Do nothing as address controls are cleared already. 
            Else
                Dim objClientAddresses As New List(Of clsClientAddresses)
                objClientAddresses = objClientsBal.GetAddresses(selectedClientId)
                If objClientAddresses.Count = 0 Then
                    ' Do nothing as address controls are cleared already. 
                ElseIf objClientAddresses.Count = 1 Then
                    ' Set Attention drop-down source to address list. 
                    AttnComboBox.DisplayMember = "Attention"
                    AttnComboBox.ValueMember = "ClientAddressID"
                    AttnComboBox.DataSource = objClientAddresses

                    ' Fill address by setting Attention drop-down's selected index = 0
                    AttnComboBox.SelectedIndex = 0
                    ' SetClientFieldValues(chosenClient, textBoxes, AttnComboBox)
                ElseIf objClientAddresses.Count > 1 Then
                    ' Set Attention drop-down source to address list. 
                    AttnComboBox.DisplayMember = "Attention"
                    AttnComboBox.ValueMember = "ClientAddressID"
                    AttnComboBox.DataSource = objClientAddresses

                    AttnComboBox.SelectedIndex = -1
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ' Log the exception if necessary
            ' logger.Log(ex.ToString())
        End Try
    End Sub

    Private Sub ClearFields(textBoxes As TextBox())
        For Each textBox As TextBox In textBoxes
            textBox.Text = String.Empty
        Next
    End Sub

    Private Sub ClearFields(textBoxes As TextBox(), AttentionCombo As ComboBox)
        For Each textBox As TextBox In textBoxes
            textBox.Text = String.Empty
        Next
        AttentionCombo.SelectedIndex = -1
    End Sub

    Private Sub SetClientFieldValues(clientAddress As clsClientAddresses, textBoxes As TextBox())
        textBoxes(0).Text = clientAddress.Address
        textBoxes(1).Text = clientAddress.City
        textBoxes(2).Text = clientAddress.State
        textBoxes(3).Text = clientAddress.ZIP
    End Sub

    Private Sub AttentionCombo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbCIAttention1.SelectedIndexChanged, CmbCIAttention2.SelectedIndexChanged, CmbBFIAttention2.SelectedIndexChanged, CmbBFIAttention1.SelectedIndexChanged, CmbBFIAttention4.SelectedIndexChanged, CmbBFIAttention3.SelectedIndexChanged
        Dim comboBox As ComboBox = CType(sender, ComboBox)
        Dim textBoxes As TextBox()
        Dim objClientAddressBAL As New clsClientsBAL
        Dim objClientAddress As clsClientAddresses

        If comboBox Is CmbCIAttention1 Then
            textBoxes = {TxtCIAddress1, TxtCICity1, TxtCIState1, TxtCIZip1}
        ElseIf comboBox Is CmbCIAttention2 Then
            textBoxes = {TxtCIAddress2, TxtCICity2, TxtCIState2, TxtCIZip2}
        ElseIf comboBox Is CmbBFIAttention1 Then
            textBoxes = {TxtBFIAddress1, TxtBFICity1, TxtBFIState1, TxtBFIZip1}
        ElseIf comboBox Is CmbBFIAttention3 Then
            textBoxes = {TxtBFIAddress3, TxtBFICity3, TxtBFIState3, TxtBFIZip3}
        ElseIf comboBox Is CmbBFIAttention2 Then
            textBoxes = {TxtBFIAddress2, TxtBFICity2, TxtBFIState2, TxtBFIZip2}
        ElseIf comboBox Is CmbBFIAttention4 Then
            textBoxes = {TxtBFIAddress4, TxtBFICity4, TxtBFIState4, TxtBFIZip4}
        Else
            Return ' Exit if it's an unknown ComboBox
        End If

        Try
            Dim selectedAddressId As Integer = CType(comboBox.SelectedValue, Integer)
            objClientAddress = objClientAddressBAL.GetAddressByAddressID(selectedAddressId)
            ClearFields(textBoxes)
            SetClientFieldValues(objClientAddress, textBoxes)
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ' Log the exception if necessary
            ' logger.Log(ex.ToString())
        End Try
    End Sub

    Private Sub PopulateListingComboBoxWithOwnCompany(comboBox As ComboBox, ownCompany As clsClients)
        Dim ownCompanyList As New List(Of clsClients) From {ownCompany}
        comboBox.DisplayMember = "ClientName"
        comboBox.ValueMember = "ClientId"
        comboBox.DataSource = ownCompanyList
        comboBox.SelectedIndex = 0
    End Sub

    Private Sub PopulateProcuringComboBoxWithOwnCompany(comboBox As ComboBox, ownCompany As clsClients)
        Dim ownCompanyList As New List(Of clsClients) From {ownCompany}
        comboBox.DisplayMember = "ClientName"
        comboBox.ValueMember = "ClientId"
        comboBox.DataSource = ownCompanyList
        comboBox.SelectedIndex = 0
    End Sub

    Private Sub IniRentGrid()
        Try
            'DesignGridView(grdRentDetails, 10, False, False, False, True, DataGridViewAutoSizeColumnsMode.Fill)
            GrdRentDetails.Columns.Clear()
            GrdRentDetails.Columns.Add("RentID", "Rent ID")
            GrdRentDetails.Columns.Add("BookingId", "Booking ID")
            GrdRentDetails.Columns.Add("RentMonthFrom", "Rent Month From")
            GrdRentDetails.Columns.Add("RentMonthTo", "Rent Month To")
            GrdRentDetails.Columns.Add("RentPerSF", "Rent Per SF")
            GrdRentDetails.Columns.Add("TotalSFs", "Total SFs")
            GrdRentDetails.Columns.Add("MonthlyRentCalculated", "Monthly Rent Calculated")
            GrdRentDetails.Columns.Add("MonthlyRentOverride", "Monthly Rent Override")
            GrdRentDetails.Columns.Add("TotalRentCalculated", "Total Rent Calculated")
            GrdRentDetails.Columns.Add("TotalRentOverride", "Total Rent Override")

            ' Set the background color for override columns
            Dim yellowCellStyle As New DataGridViewCellStyle()
            yellowCellStyle.BackColor = Color.Yellow

            GrdRentDetails.Columns("TotalSFs").DefaultCellStyle = yellowCellStyle
            GrdRentDetails.Columns("MonthlyRentOverride").DefaultCellStyle = yellowCellStyle
            GrdRentDetails.Columns("TotalRentOverride").DefaultCellStyle = yellowCellStyle

            GrdRentDetails.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            IniCommissionGrid()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub IniCommissionGrid()
        GrdCommDetails.Columns.Clear()
        GrdCommDetails.Columns.Add("TotalRent", "Total Rent")
        GrdCommDetails.Columns.Add("ListCommPct", "List Comm Percent")
        GrdCommDetails.Columns.Add("ListCommTotal", "List Comm Total")
        GrdCommDetails.Columns.Add("ProcCommPct", "Proc Comm Percent")
        GrdCommDetails.Columns.Add("ProcCommTotal", "Proc Comm Total")
        GrdCommDetails.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
    End Sub

    Private Sub LoadAllRentData()
        Try
            rentData = Nothing
            Me.Cursor = Cursors.WaitCursor
            rentData = objRentDetailsBal.GetByBookingID(BookingID)
            ShowRentData(rentData)
            'txtFilter.Focus()
            ButtonEnableDisable(True, True, False, False, True, True, True)
            Me.Cursor = Cursors.Default
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            Throw ex
        End Try
    End Sub

    Private Sub LoadAllCommData()

    End Sub

    Private Sub ShowRentData(data As List(Of clsRentDetails))
        GrdRentDetails.Rows.Clear()
        If (data.Count = 0) Then
            GrdRentDetails.Rows.Clear()
        Else
            Dim totalRent As Decimal = 0
            Dim totalMonths As Int32 = 0
            For Each item As clsRentDetails In data
                Dim values As Object() = {
                item.RentID,
                item.BookingId,
                item.RentMonthFrom,
                item.RentMonthTo,
                item.RentPerSF,
                item.TotalSFs,
                item.MonthlyRentCalculated,
                item.MonthlyRentOverride,
                item.TotalRentCalculated,
                item.TotalRentOverride
            }
                GrdRentDetails.Rows.Add(values)
                'assigning total rent
                If Not String.IsNullOrEmpty(item.TotalRentCalculated) Then
                    totalRent += item.TotalRentCalculated
                End If

                'assigning total months
                If Not String.IsNullOrEmpty(item.RentMonthDifference) Then
                    totalMonths += item.RentMonthDifference
                End If
            Next
            If (totalRent <> 0) Then
                TxtRentSumTotal.Text = totalRent.ToString()
            End If

            If (totalMonths <> 0) Then
                TxtTotalMonthSum.Text = totalMonths.ToString()
            End If
        End If

        If GrdRentDetails.Rows.Count < 12 Then
            Dim Counter As Integer = 0

            For Counter = (GrdRentDetails.Rows.Count + 1) To 12
                Dim values As Object() = {
                0,
                DBNull.Value,
                DBNull.Value,
                DBNull.Value,
                DBNull.Value,
                DBNull.Value,
                DBNull.Value,
                DBNull.Value,
                DBNull.Value,
                DBNull.Value
            }
                GrdRentDetails.Rows.Add(values)
            Next
        End If

        GrdRentDetails.Refresh()
    End Sub

    Private Sub SetRentFieldValues(rentDetails As clsRentDetails)
        chosenRent = rentDetails
        TxtRentFromMonth.Text = rentDetails.RentMonthFrom.ToString()
        TxtRentToMonth.Text = rentDetails.RentMonthTo.ToString()
        TxtRentNoOfMonth.Text = rentDetails.RentMonthDifference.ToString()
        TxtRentPerSF.Text = rentDetails.RentPerSF.ToString()
        TxtRentSfCal.Text = rentDetails.TotalSFs.ToString()
        TxtRentSFOverride.Text = rentDetails.TotalSFs.ToString()
        TxtRentMonthlyCal.Text = rentDetails.MonthlyRentCalculated.ToString()
        TxtRentMonthlyOverride.Text = rentDetails.MonthlyRentOverride.ToString()
        TxtRentTotCal.Text = rentDetails.TotalRentCalculated.ToString()
        TxtRentTotOverride.Text = rentDetails.TotalRentOverride.ToString()
    End Sub

    Private Sub PopulateRentFromFields()
        chosenRent.BookingId = BookingID
        chosenRent.RentMonthFrom = Integer.Parse(TxtRentFromMonth.Text)
        chosenRent.RentMonthTo = Integer.Parse(TxtRentToMonth.Text)
        chosenRent.RentMonthDifference = Integer.Parse(TxtRentNoOfMonth.Text)
        chosenRent.RentPerSF = Decimal.Parse(TxtRentPerSF.Text)
        chosenRent.TotalSFs = Integer.Parse(TxtRentSfCal.Text) ' Assuming TxtRentSfCal is correct
        chosenRent.MonthlyRentCalculated = Decimal.Parse(TxtRentMonthlyCal.Text)
        chosenRent.MonthlyRentOverride = Decimal.Parse(TxtRentMonthlyOverride.Text)
        chosenRent.TotalRentCalculated = Decimal.Parse(TxtRentTotCal.Text)
        chosenRent.TotalRentOverride = Decimal.Parse(TxtRentTotOverride.Text)
        chosenRent.ProcCommAmtCalculated = 0
        chosenRent.ProcCommAmtOverride = 0
        chosenRent.ProcCommPct = 0
        chosenRent.ListCommPct = 0
        chosenRent.ListCommAmtOverride = 0
        chosenRent.ListCommAmtCalculated = 0

        ' Assign other properties as needed
    End Sub

    Private Sub ClearRentFormFields()
        ' Clear or reset form fields after saving
        TxtRentFromMonth.Text = ""
        TxtRentToMonth.Text = ""
        TxtRentNoOfMonth.Text = ""
        TxtRentPerSF.Text = ""
        TxtRentSfCal.Text = ""
        TxtRentSFOverride.Text = ""
        TxtRentMonthlyCal.Text = ""
        TxtRentMonthlyOverride.Text = ""
        TxtRentTotCal.Text = ""
        TxtRentTotOverride.Text = ""
        ' Clear other fields as needed
    End Sub

    Private Sub BtnAddRent_Click(sender As Object, e As EventArgs) Handles BtnAddRent.Click
        ClearRentFormFields()
        chosenRent.RentID = 0
    End Sub

    Private Sub TabBookings_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabBookings.SelectedIndexChanged
        'MessageBox.Show("hi")
    End Sub

    Private Sub DTLeaseCommence_ValueChanged(sender As Object, e As EventArgs) Handles DTLeaseCommence.ValueChanged, TxtLeaseTerm.TextChanged
        Try
            Dim leaseCommenceDate As DateTime = DTLeaseCommence.Value
            Dim leaseTerm As Integer

            If Integer.TryParse(TxtLeaseTerm.Text, leaseTerm) AndAlso leaseTerm > 0 Then
                Dim leaseExpiryDate As DateTime = leaseCommenceDate.AddMonths(leaseTerm).AddDays(-1)
                DTLeaseExpiry.Value = leaseExpiryDate
            Else
                DTLeaseExpiry.Value = leaseCommenceDate
            End If
        Catch ex As Exception
            MessageBox.Show($"Error calculating lease expiry date: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#Region "Change by HB"
    Private Sub CmbTransactionType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbTransactionType.SelectedIndexChanged
        If CmbTransactionType.SelectedItem = "Select" Then
            totalSaleConsidationPanel.Visible = False
            ListingClientPanel.Visible = True
            TabelPanelProcuring.Visible = True
            TxtListingClient.Text = "LandLord"
            TxtProcuringClient.Text = "Tenant"
            lblBrokerageFirm.Text = "LISTING Firm"
            lblFeeCalculation.Text = "LISTING Fee"
            LblAdditionalRef.Text = "Additional / Referral LISTING Firm"
            TxtBFIAddress1.Visible = True
            label75.Visible = True
            label76.Visible = True
            label31.Visible = True
            label74.Visible = True
            CmbBFIAttention1.Visible = True
            label73.Visible = True
            ComboBoxLF1.Visible = True
            label72.Visible = True
            lblBrokerageFirm.Visible = True
            checkBoxLEELF.Visible = True
            TxtBFIState1.Visible = True
            TxtBFICity1.Visible = True
            TxtBFIAddress1.Visible = True
            TxtBFIZip1.Visible = True
            TxtListingCommRate.Visible = False
            LblListingCommRate.Visible = False
            TxtProcuringCommRate.Visible = False
            LblProcuringCommRate.Visible = False
            GrossProcuringPanel.Visible = False
            ListingGrossPanel.Visible = False
            ProcuringClientPanel.Visible = True
        ElseIf CmbTransactionType.SelectedItem = "Lease" Then
            totalSaleConsidationPanel.Visible = False
            ProcuringClientPanel.Visible = True
            TabelPanelProcuring.Visible = True
            TxtListingClient.Text = "LandLord"
            TxtProcuringClient.Text = "Tenant"
            lblBrokerageFirm.Text = "LISTING Firm"
            lblFeeCalculation.Text = "LISTING Fee"
            LblAdditionalRef.Text = "Additional / Referral LISTING Firm"
            TxtBFIAddress1.Visible = True
            label75.Visible = True
            label76.Visible = True
            label31.Visible = True
            label74.Visible = True
            CmbBFIAttention1.Visible = True
            label73.Visible = True
            ComboBoxLF1.Visible = True
            label72.Visible = True
            lblBrokerageFirm.Visible = True
            checkBoxLEELF.Visible = True
            TxtBFIState1.Visible = True
            TxtBFICity1.Visible = True
            TxtBFIAddress1.Visible = True
            TxtBFIZip1.Visible = True
            TxtListingCommRate.Visible = False
            LblListingCommRate.Visible = False
            TxtProcuringCommRate.Visible = False
            LblProcuringCommRate.Visible = False
            GrossProcuringPanel.Visible = False
            ListingGrossPanel.Visible = False
        ElseIf CmbTransactionType.SelectedItem = "Sale" Then
            totalSaleConsidationPanel.Visible = True
            TabelPanelProcuring.Visible = True
            ProcuringClientPanel.Visible = True
            TxtListingClient.Text = "Seller"
            TxtProcuringClient.Text = "Buyer"
            lblBrokerageFirm.Text = "LISTING Firm"
            lblFeeCalculation.Text = "LISTING Fee"
            LblAdditionalRef.Text = "Additional / Referral LISTING Firm"
            TxtListingCommRate.Visible = True
            LblListingCommRate.Visible = True
            TxtProcuringCommRate.Visible = True
            LblProcuringCommRate.Visible = True
            TxtBFIAddress1.Visible = True
            label75.Visible = True
            label76.Visible = True
            label31.Visible = True
            label74.Visible = True
            CmbBFIAttention1.Visible = True
            label73.Visible = True
            ComboBoxLF1.Visible = True
            label72.Visible = True
            lblBrokerageFirm.Visible = True
            checkBoxLEELF.Visible = True
            TxtBFIState1.Visible = True
            TxtBFICity1.Visible = True
            TxtBFIAddress1.Visible = True
            TxtBFIZip1.Visible = True
            GrossProcuringPanel.Visible = False
            ListingGrossPanel.Visible = False
        ElseIf CmbTransactionType.SelectedItem = "Consulting" Then
            totalSaleConsidationPanel.Visible = False
            ProcuringClientPanel.Visible = False
            TxtListingClient.Text = "Client"
            lblBrokerageFirm.Text = "Consulting  Frim"
            lblFeeCalculation.Text = "CONSULTING Fee"
            LblAdditionalRef.Text = "Additional / Referral CONSULTING Firm"
            TxtBFIAddress1.Visible = False
            label75.Visible = False
            label76.Visible = False
            label31.Visible = False
            label74.Visible = False
            CmbBFIAttention1.Visible = False
            label73.Visible = False
            ComboBoxLF1.Visible = False
            label72.Visible = False
            lblBrokerageFirm.Visible = False
            checkBoxLEELF.Visible = False
            TxtBFIState1.Visible = False
            TxtBFICity1.Visible = False
            TxtBFIAddress1.Visible = False
            TxtBFIZip1.Visible = False
            TxtListingCommRate.Visible = False
            LblListingCommRate.Visible = False
            TxtProcuringCommRate.Visible = False
            LblProcuringCommRate.Visible = False
            GrossProcuringPanel.Visible = False
            ListingGrossPanel.Visible = False
        End If
    End Sub

    Private Sub TxtTotSf_TextChanged(sender As Object, e As EventArgs) Handles TxtTotSf.Leave, TxtTotConsidation.Leave
        CalculateTotalSf()
    End Sub
    Private Function CalculateTotalSf() As Boolean
        If String.IsNullOrEmpty(TxtTotConsidation.Text) AndAlso String.IsNullOrEmpty(TxtTotSf.Text) Then
            TotPerSf.Text = "0"
        ElseIf Not String.IsNullOrEmpty(TxtTotConsidation.Text) AndAlso Not String.IsNullOrEmpty(TxtTotSf.Text) Then
            TotPerSf.Text = TxtTotConsidation.Text / TxtTotSf.Text
        End If
    End Function
    Private Sub TxtLFTotalFee_TextChanged(sender As Object, e As EventArgs) Handles TxtLFTotalFee.TextChanged, TxtLFOutBroker1.Leave, TxtLFOutBroker2.Leave
        CalculateGrossComm(TxtLFTotalFee, TxtLFOutBroker1, TxtLFOutBroker2, TxtLFLeeGrossComm, ListingGrossPanel)
        UpdateInvoiceTotal()
    End Sub

    Private Sub TxtPFTotalFee_TextChanged(sender As Object, e As EventArgs) Handles TxtPFTotalFee.TextChanged, TxtPFOutBroker1.Leave, TxtPFOutBroker2.Leave
        CalculateGrossComm(TxtPFTotalFee, TxtPFOutBroker1, TxtPFOutBroker2, TxtPFLeeGrossComm, GrossProcuringPanel)
        UpdateInvoiceTotal()
    End Sub

    Private Sub CalculateGrossComm(totalFeeTextBox As TextBox, outBroker1TextBox As TextBox, outBroker2TextBox As TextBox, leeGrossCommTextBox As TextBox, grossPanel As Panel)
        Dim totalFee As Decimal
        Dim outBroker1 As Decimal
        Dim outBroker2 As Decimal

        Decimal.TryParse(totalFeeTextBox.Text, totalFee)
        Decimal.TryParse(outBroker1TextBox.Text, outBroker1)
        Decimal.TryParse(outBroker2TextBox.Text, outBroker2)

        Dim grossComm As Decimal = totalFee - (outBroker1 + outBroker2)

        If grossComm > 0 Then
            leeGrossCommTextBox.Text = grossComm.ToString("F2")
            grossPanel.Visible = True
        ElseIf grossComm < 0 Then
            leeGrossCommTextBox.Text = grossComm.ToString("F2")
            grossPanel.Visible = False
        Else
            leeGrossCommTextBox.Text = "-"
            grossPanel.Visible = False
        End If
    End Sub

    Private Sub UpdateInvoiceTotal()
        Dim lfTotalFee As Decimal
        Dim pfTotalFee As Decimal

        Decimal.TryParse(TxtLFTotalFee.Text, lfTotalFee)
        Decimal.TryParse(TxtPFTotalFee.Text, pfTotalFee)

        Dim invoiceTotal As Decimal = lfTotalFee + pfTotalFee
        TxtInvoiceTot.Text = invoiceTotal.ToString("F2")
    End Sub

    Private Sub CalculateGrossComm(agentTextBox As TextBox, grossAmtTextBox As TextBox, splitPctTextBox As TextBox, grossCommTextBox As TextBox, netAgtTextBox As TextBox, leeGrossCommTextBox As TextBox)
        Dim grossComm As Decimal
        Dim splitPct As Decimal
        Dim leeGrossComm As Decimal

        Decimal.TryParse(leeGrossCommTextBox.Text, leeGrossComm)

        If Not String.IsNullOrEmpty(agentTextBox.Text) Then
            If Not String.IsNullOrEmpty(splitPctTextBox.Text) AndAlso (String.IsNullOrEmpty(grossAmtTextBox.Text) OrElse grossAmtTextBox.Text = "0" OrElse grossAmtTextBox.Text = " ") Then
                Decimal.TryParse(splitPctTextBox.Text, splitPct)
                grossComm = leeGrossComm * splitPct / 100
            ElseIf Not String.IsNullOrEmpty(grossAmtTextBox.Text) Then
                Decimal.TryParse(grossAmtTextBox.Text, grossComm)
            End If
        Else
            grossCommTextBox.Text = ""
            netAgtTextBox.Text = ""
            Return
        End If

        grossCommTextBox.Text = grossComm.ToString("F2")

        If grossComm <> 0 Then
            netAgtTextBox.Text = (grossComm / 2).ToString("F2")
        Else
            netAgtTextBox.Text = ""
        End If
    End Sub
    Private Sub HandleAgentLeave(sender As Object, e As EventArgs) Handles TxtLFLEEAgt1.TextChanged, TxtLFLEEAgt2.TextChanged, TxtLFLEEAgt3.TextChanged, TxtLFLEEAgt4.TextChanged, TxtLFLEEAgt5.TextChanged, TxtLFLEEAgt6.TextChanged, TxtPFLEEAgt1.TextChanged, TxtPFLEEAgt2.TextChanged, TxtPFLEEAgt3.TextChanged, TxtPFLEEAgt4.TextChanged, TxtPFLEEAgt5.TextChanged, TxtPFLEEAgt6.TextChanged
        Dim agentTextBox As TextBox = CType(sender, TextBox)
        Select Case agentTextBox.Name
            Case "TxtLFLEEAgt1"
                CalculateGrossComm(TxtLFLEEAgt1, TxtLFGrossAmt1, TxtLFSplitPct1, txtLFGrossComm1, TxtLFNetAgt1, TxtLFLeeGrossComm)
            Case "TxtLFLEEAgt2"
                CalculateGrossComm(TxtLFLEEAgt2, TxtLFGrossAmt2, TxtLFSplitPct2, txtLFGrossComm2, TxtLFNetAgt2, TxtLFLeeGrossComm)
            Case "TxtLFLEEAgt3"
                CalculateGrossComm(TxtLFLEEAgt3, TxtLFGrossAmt3, TxtLFSplitPct3, txtLFGrossComm3, TxtLFNetAgt3, TxtLFLeeGrossComm)
            Case "TxtLFLEEAgt4"
                CalculateGrossComm(TxtLFLEEAgt4, TxtLFGrossAmt4, TxtLFSplitPct4, txtLFGrossComm4, TxtLFNetAgt4, TxtLFLeeGrossComm)
            Case "TxtLFLEEAgt5"
                CalculateGrossComm(TxtLFLEEAgt5, TxtLFGrossAmt5, TxtLFSplitPct5, txtLFGrossComm5, TxtLFNetAgt5, TxtLFLeeGrossComm)
            Case "TxtLFLEEAgt6"
                CalculateGrossComm(TxtLFLEEAgt6, TxtLFGrossAmt6, TxtLFSplitPct6, txtLFGrossComm6, TxtLFNetAgt6, TxtLFLeeGrossComm)
            Case "TxtPFLEEAgt1"
                CalculateGrossComm(TxtPFLEEAgt1, TxtPFGrossAmt1, TxtPFSplitPct1, txtPFGrossComm1, TxtPFNetAgt1, TxtPFLeeGrossComm)
            Case "TxtPFLEEAgt2"
                CalculateGrossComm(TxtPFLEEAgt2, TxtPFGrossAmt2, TxtPFSplitPct2, txtPFGrossComm2, TxtPFNetAgt2, TxtPFLeeGrossComm)
            Case "TxtPFLEEAgt3"
                CalculateGrossComm(TxtPFLEEAgt3, TxtPFGrossAmt3, TxtPFSplitPct3, txtPFGrossComm3, TxtPFNetAgt3, TxtPFLeeGrossComm)
            Case "TxtPFLEEAgt4"
                CalculateGrossComm(TxtPFLEEAgt4, TxtPFGrossAmt4, TxtPFSplitPct4, txtPFGrossComm4, TxtPFNetAgt4, TxtPFLeeGrossComm)
            Case "TxtPFLEEAgt5"
                CalculateGrossComm(TxtPFLEEAgt5, TxtPFGrossAmt5, TxtPFSplitPct5, txtPFGrossComm5, TxtPFNetAgt5, TxtPFLeeGrossComm)
            Case "TxtPFLEEAgt6"
                CalculateGrossComm(TxtPFLEEAgt6, TxtPFGrossAmt6, TxtPFSplitPct6, txtPFGrossComm6, TxtPFNetAgt6, TxtPFLeeGrossComm)
        End Select
    End Sub
    'Private Sub HandleAgentLeave(sender As Object, e As EventArgs) Handles TxtLFLEEAgt1.Leave, TxtLFLEEAgt2.Leave, TxtLFLEEAgt3.Leave, TxtLFLEEAgt4.Leave, TxtLFLEEAgt5.Leave, TxtLFLEEAgt6.Leave, TxtPFLEEAgt1.Leave, TxtPFLEEAgt2.Leave, TxtPFLEEAgt3.Leave, TxtPFLEEAgt4.Leave, TxtPFLEEAgt5.Leave, TxtPFLEEAgt6.Leave
    '    Dim agentTextBox As TextBox = CType(sender, TextBox)
    '    Select Case agentTextBox.Name
    '        Case "TxtLFLEEAgt1"
    '            CalculateGrossComm(TxtLFLEEAgt1, TxtLFGrossAmt1, TxtLFSplitPct1, txtLFGrossComm1, TxtLFNetAgt1, TxtLFLeeGrossComm)
    '        Case "TxtLFLEEAgt2"
    '            CalculateGrossComm(TxtLFLEEAgt2, TxtLFGrossAmt2, TxtLFSplitPct2, txtLFGrossComm2, TxtLFNetAgt2, TxtLFLeeGrossComm)
    '        Case "TxtLFLEEAgt3"
    '            CalculateGrossComm(TxtLFLEEAgt3, TxtLFGrossAmt3, TxtLFSplitPct3, txtLFGrossComm3, TxtLFNetAgt3, TxtLFLeeGrossComm)
    '        Case "TxtLFLEEAgt4"
    '            CalculateGrossComm(TxtLFLEEAgt4, TxtLFGrossAmt4, TxtLFSplitPct4, txtLFGrossComm4, TxtLFNetAgt4, TxtLFLeeGrossComm)
    '        Case "TxtLFLEEAgt5"
    '            CalculateGrossComm(TxtLFLEEAgt5, TxtLFGrossAmt5, TxtLFSplitPct5, txtLFGrossComm5, TxtLFNetAgt5, TxtLFLeeGrossComm)
    '        Case "TxtLFLEEAgt6"
    '            CalculateGrossComm(TxtLFLEEAgt6, TxtLFGrossAmt6, TxtLFSplitPct6, txtLFGrossComm6, TxtLFNetAgt6, TxtLFLeeGrossComm)
    '        Case "TxtPFLEEAgt1"
    '            CalculateGrossComm(TxtPFLEEAgt1, TxtPFGrossAmt1, TxtPFSplitPct1, txtPFGrossComm1, TxtPFNetAgt1, TxtPFLeeGrossComm)
    '        Case "TxtPFLEEAgt2"
    '            CalculateGrossComm(TxtPFLEEAgt2, TxtPFGrossAmt2, TxtPFSplitPct2, txtPFGrossComm2, TxtPFNetAgt2, TxtPFLeeGrossComm)
    '        Case "TxtPFLEEAgt3"
    '            CalculateGrossComm(TxtPFLEEAgt3, TxtPFGrossAmt3, TxtPFSplitPct3, txtPFGrossComm3, TxtPFNetAgt3, TxtPFLeeGrossComm)
    '        Case "TxtPFLEEAgt4"
    '            CalculateGrossComm(TxtPFLEEAgt4, TxtPFGrossAmt4, TxtPFSplitPct4, txtPFGrossComm4, TxtPFNetAgt4, TxtPFLeeGrossComm)
    '        Case "TxtPFLEEAgt5"
    '            CalculateGrossComm(TxtPFLEEAgt5, TxtPFGrossAmt5, TxtPFSplitPct5, txtPFGrossComm5, TxtPFNetAgt5, TxtPFLeeGrossComm)
    '        Case "TxtPFLEEAgt6"
    '            CalculateGrossComm(TxtPFLEEAgt6, TxtPFGrossAmt6, TxtPFSplitPct6, txtPFGrossComm6, TxtPFNetAgt6, TxtPFLeeGrossComm)
    '    End Select
    'End Sub

    Private Sub HandleGrossAmtTextChanged(sender As Object, e As EventArgs) Handles TxtLFGrossAmt1.TextChanged, TxtLFGrossAmt2.TextChanged, TxtLFGrossAmt3.TextChanged, TxtLFGrossAmt4.TextChanged, TxtLFGrossAmt5.TextChanged, TxtLFGrossAmt6.TextChanged, TxtPFGrossAmt1.TextChanged, TxtPFGrossAmt2.TextChanged, TxtPFGrossAmt3.TextChanged, TxtPFGrossAmt4.TextChanged, TxtPFGrossAmt5.TextChanged, TxtPFGrossAmt6.TextChanged
        Dim grossAmtTextBox As TextBox = CType(sender, TextBox)
        Select Case grossAmtTextBox.Name
            Case "TxtLFGrossAmt1"
                CalculateGrossComm(TxtLFLEEAgt1, TxtLFGrossAmt1, TxtLFSplitPct1, txtLFGrossComm1, TxtLFNetAgt1, TxtLFLeeGrossComm)
            Case "TxtLFGrossAmt2"
                CalculateGrossComm(TxtLFLEEAgt2, TxtLFGrossAmt2, TxtLFSplitPct2, txtLFGrossComm2, TxtLFNetAgt2, TxtLFLeeGrossComm)
            Case "TxtLFGrossAmt3"
                CalculateGrossComm(TxtLFLEEAgt3, TxtLFGrossAmt3, TxtLFSplitPct3, txtLFGrossComm3, TxtLFNetAgt3, TxtLFLeeGrossComm)
            Case "TxtLFGrossAmt4"
                CalculateGrossComm(TxtLFLEEAgt4, TxtLFGrossAmt4, TxtLFSplitPct4, txtLFGrossComm4, TxtLFNetAgt4, TxtLFLeeGrossComm)
            Case "TxtLFGrossAmt5"
                CalculateGrossComm(TxtLFLEEAgt5, TxtLFGrossAmt5, TxtLFSplitPct5, txtLFGrossComm5, TxtLFNetAgt5, TxtLFLeeGrossComm)
            Case "TxtLFGrossAmt6"
                CalculateGrossComm(TxtLFLEEAgt6, TxtLFGrossAmt6, TxtLFSplitPct6, txtLFGrossComm6, TxtLFNetAgt6, TxtLFLeeGrossComm)
            Case "TxtPFGrossAmt1"
                CalculateGrossComm(TxtPFLEEAgt1, TxtPFGrossAmt1, TxtPFSplitPct1, txtPFGrossComm1, TxtPFNetAgt1, TxtPFLeeGrossComm)
            Case "TxtPFGrossAmt2"
                CalculateGrossComm(TxtPFLEEAgt2, TxtPFGrossAmt2, TxtPFSplitPct2, txtPFGrossComm2, TxtPFNetAgt2, TxtPFLeeGrossComm)
            Case "TxtPFGrossAmt3"
                CalculateGrossComm(TxtPFLEEAgt3, TxtPFGrossAmt3, TxtPFSplitPct3, txtPFGrossComm3, TxtPFNetAgt3, TxtPFLeeGrossComm)
            Case "TxtPFGrossAmt4"
                CalculateGrossComm(TxtPFLEEAgt4, TxtPFGrossAmt4, TxtPFSplitPct4, txtPFGrossComm4, TxtPFNetAgt4, TxtPFLeeGrossComm)
            Case "TxtPFGrossAmt5"
                CalculateGrossComm(TxtPFLEEAgt5, TxtPFGrossAmt5, TxtPFSplitPct5, txtPFGrossComm5, TxtPFNetAgt5, TxtPFLeeGrossComm)
            Case "TxtPFGrossAmt6"
                CalculateGrossComm(TxtPFLEEAgt6, TxtPFGrossAmt6, TxtPFSplitPct6, txtPFGrossComm6, TxtPFNetAgt6, TxtPFLeeGrossComm)
        End Select
    End Sub

    Private Sub HandleSplitPctLeave(sender As Object, e As EventArgs) Handles TxtLFSplitPct1.Leave, TxtLFSplitPct2.Leave, TxtLFSplitPct3.Leave, TxtLFSplitPct4.Leave, TxtLFSplitPct5.Leave, TxtLFSplitPct6.Leave, TxtPFSplitPct1.Leave, TxtPFSplitPct2.Leave, TxtPFSplitPct3.Leave, TxtPFSplitPct4.Leave, TxtPFSplitPct5.Leave, TxtPFSplitPct6.Leave
        Dim splitPctTextBox As TextBox = CType(sender, TextBox)
        Select Case splitPctTextBox.Name
            Case "TxtLFSplitPct1"
                CalculateGrossComm(TxtLFLEEAgt1, TxtLFGrossAmt1, TxtLFSplitPct1, txtLFGrossComm1, TxtLFNetAgt1, TxtLFLeeGrossComm)
            Case "TxtLFSplitPct2"
                CalculateGrossComm(TxtLFLEEAgt2, TxtLFGrossAmt2, TxtLFSplitPct2, txtLFGrossComm2, TxtLFNetAgt2, TxtLFLeeGrossComm)
            Case "TxtLFSplitPct3"
                CalculateGrossComm(TxtLFLEEAgt3, TxtLFGrossAmt3, TxtLFSplitPct3, txtLFGrossComm3, TxtLFNetAgt3, TxtLFLeeGrossComm)
            Case "TxtLFSplitPct4"
                CalculateGrossComm(TxtLFLEEAgt4, TxtLFGrossAmt4, TxtLFSplitPct4, txtLFGrossComm4, TxtLFNetAgt4, TxtLFLeeGrossComm)
            Case "TxtLFSplitPct5"
                CalculateGrossComm(TxtLFLEEAgt5, TxtLFGrossAmt5, TxtLFSplitPct5, txtLFGrossComm5, TxtLFNetAgt5, TxtLFLeeGrossComm)
            Case "TxtLFSplitPct6"
                CalculateGrossComm(TxtLFLEEAgt6, TxtLFGrossAmt6, TxtLFSplitPct6, txtLFGrossComm6, TxtLFNetAgt6, TxtLFLeeGrossComm)
            Case "TxtPFSplitPct1"
                CalculateGrossComm(TxtPFLEEAgt1, TxtPFGrossAmt1, TxtPFSplitPct1, txtPFGrossComm1, TxtPFNetAgt1, TxtPFLeeGrossComm)
            Case "TxtPFSplitPct2"
                CalculateGrossComm(TxtPFLEEAgt2, TxtPFGrossAmt2, TxtPFSplitPct2, txtPFGrossComm2, TxtPFNetAgt2, TxtPFLeeGrossComm)
            Case "TxtPFSplitPct3"
                CalculateGrossComm(TxtPFLEEAgt3, TxtPFGrossAmt3, TxtPFSplitPct3, txtPFGrossComm3, TxtPFNetAgt3, TxtPFLeeGrossComm)
            Case "TxtPFSplitPct4"
                CalculateGrossComm(TxtPFLEEAgt4, TxtPFGrossAmt4, TxtPFSplitPct4, txtPFGrossComm4, TxtPFNetAgt4, TxtPFLeeGrossComm)
            Case "TxtPFSplitPct5"
                CalculateGrossComm(TxtPFLEEAgt5, TxtPFGrossAmt5, TxtPFSplitPct5, txtPFGrossComm5, TxtPFNetAgt5, TxtPFLeeGrossComm)
            Case "TxtPFSplitPct6"
                CalculateGrossComm(TxtPFLEEAgt6, TxtPFGrossAmt6, TxtPFSplitPct6, txtPFGrossComm6, TxtPFNetAgt6, TxtPFLeeGrossComm)
        End Select
    End Sub

    Private Function SumTextBoxValues(ParamArray textboxes() As TextBox) As Decimal
        Dim total As Decimal = 0
        For Each textbox As TextBox In textboxes
            If Decimal.TryParse(textbox.Text, Nothing) Then
                total += Decimal.Parse(textbox.Text)
            End If
        Next
        Return total
    End Function

    Private Sub SumPFGrossComm_TextChanged(sender As Object, e As EventArgs) Handles txtPFGrossComm1.TextChanged, txtPFGrossComm2.TextChanged, txtPFGrossComm3.TextChanged, txtPFGrossComm4.TextChanged, txtPFGrossComm5.TextChanged, txtPFGrossComm6.TextChanged, MyBase.Load
        txtPFGrossCommTot.Text = SumTextBoxValues(txtPFGrossComm1, txtPFGrossComm2, txtPFGrossComm3, txtPFGrossComm4, txtPFGrossComm5, txtPFGrossComm6).ToString()
    End Sub

    Private Sub SumPFNetAgt_TextChanged(sender As Object, e As EventArgs) Handles TxtPFNetAgt1.TextChanged, TxtPFNetAgt2.TextChanged, TxtPFNetAgt3.TextChanged, TxtPFNetAgt4.TextChanged, TxtPFNetAgt5.TextChanged, TxtPFNetAgt6.TextChanged, MyBase.Load
        TxtPFNetAgtTot.Text = SumTextBoxValues(TxtPFNetAgt1, TxtPFNetAgt2, TxtPFNetAgt3, TxtPFNetAgt4, TxtPFNetAgt5, TxtPFNetAgt6).ToString()
    End Sub

    Private Sub SumPFSplitPct_TextChanged(sender As Object, e As EventArgs) Handles TxtPFSplitPct1.TextChanged, TxtPFSplitPct2.TextChanged, TxtPFSplitPct3.TextChanged, TxtPFSplitPct4.TextChanged, TxtPFSplitPct5.TextChanged, TxtPFSplitPct6.TextChanged, MyBase.Load
        TxtPFSplitPctTot.Text = SumTextBoxValues(TxtPFSplitPct1, TxtPFSplitPct2, TxtPFSplitPct3, TxtPFSplitPct4, TxtPFSplitPct5, TxtPFSplitPct6).ToString()
    End Sub

    Private Sub SumGrossComm_TextChanged(sender As Object, e As EventArgs) Handles txtLFGrossComm1.TextChanged, txtLFGrossComm2.TextChanged, txtLFGrossComm3.TextChanged, txtLFGrossComm4.TextChanged, txtLFGrossComm5.TextChanged, txtLFGrossComm6.TextChanged, MyBase.Load
        txtLFGrossCommTot.Text = SumTextBoxValues(txtLFGrossComm1, txtLFGrossComm2, txtLFGrossComm3, txtLFGrossComm4, txtLFGrossComm5, txtLFGrossComm6).ToString()
    End Sub

    Private Sub SumNetAgt_TextChanged(sender As Object, e As EventArgs) Handles TxtLFNetAgt1.TextChanged, TxtLFNetAgt2.TextChanged, TxtLFNetAgt3.TextChanged, TxtLFNetAgt4.TextChanged, TxtLFNetAgt5.TextChanged, TxtLFNetAgt6.TextChanged, MyBase.Load
        TxtLFNetAgtTot.Text = SumTextBoxValues(TxtLFNetAgt1, TxtLFNetAgt2, TxtLFNetAgt3, TxtLFNetAgt4, TxtLFNetAgt5, TxtLFNetAgt6).ToString()
    End Sub

    Private Sub SumSplitPct_TextChanged(sender As Object, e As EventArgs) Handles TxtLFSplitPct1.TextChanged, TxtLFSplitPct2.TextChanged, TxtLFSplitPct3.TextChanged, TxtLFSplitPct4.TextChanged, TxtLFSplitPct5.TextChanged, TxtLFSplitPct6.TextChanged, MyBase.Load
        TxtLFSplitPctTot.Text = SumTextBoxValues(TxtLFSplitPct1, TxtLFSplitPct2, TxtLFSplitPct3, TxtLFSplitPct4, TxtLFSplitPct5, TxtLFSplitPct6).ToString()
    End Sub



    Private Sub TxtRentSumTotal_TextChanged(sender As Object, e As EventArgs) Handles TxtRentSumTotal.TextChanged, MyBase.Load
        TxtLeaseConsid.Text = TxtRentSumTotal.Text
    End Sub

    Private Sub TxtLeasedSF_TextChanged(sender As Object, e As EventArgs) Handles TxtLeasedSF.TextChanged, cmbRSFCalculator.SelectedIndexChanged
        CalculateEffectiveRate()
    End Sub

    Private Function CalculateEffectiveRate() As Boolean
        If cmbRSFCalculator.SelectedItem IsNot Nothing AndAlso Not String.IsNullOrEmpty(TxtLeasedSF.Text) Then
            Try
                Dim rentSumTotal As Decimal
                Dim totalMonthSum As Decimal
                Dim leasedSF As Decimal

                ' Attempt to parse the text into Decimal
                If Decimal.TryParse(TxtRentSumTotal.Text, rentSumTotal) AndAlso
                   Decimal.TryParse(TxtTotalMonthSum.Text, totalMonthSum) AndAlso
                   Decimal.TryParse(TxtLeasedSF.Text, leasedSF) Then

                    If cmbRSFCalculator.SelectedItem.ToString() = "Monthly" Then
                        TextBox92.Text = (rentSumTotal / totalMonthSum / leasedSF).ToString()
                    ElseIf cmbRSFCalculator.SelectedItem.ToString() = "Annually" Then
                        TextBox92.Text = (rentSumTotal / totalMonthSum / leasedSF * 12).ToString()
                    End If
                Else
                    TextBox92.Text = "" ' Clear the TextBox if parsing fails
                End If

            Catch ex As Exception
                TextBox92.Text = ""
            End Try
        Else
            TextBox92.Text = ""
        End If
    End Function
#End Region

    Private Sub BtnClientAdd_Click(sender As Object, e As EventArgs) Handles BtnClientAdd.Click, BtnClientAdd6.Click, BtnClientAdd5.Click, BtnClientAdd4.Click, BtnClientAdd3.Click, BtnClientAdd2.Click
        Dim ClientsForm As New FrmClients
        'ClientsForm.MdiParent = Me.Parent
        ClientsForm.ShowDialog()

        ' TODO - Check and keep already selected company - check options to refresh data source without affecting selected value
        'GetDropDownValues("clients", ComboBoxClients)
        'GetDropDownValues("clients", ComboBoxClients2)
        'GetDropDownValues("clients", ComboBoxLF1)
        'GetDropDownValues("clients", ComboBoxPF1)
        'GetDropDownValues("clients", ComboBoxLF2)
        'GetDropDownValues("clients", ComboBoxPF2)
    End Sub

    Private Sub btnEmail_Click(sender As Object, e As EventArgs) Handles btnEmail.Click
        Dim PDFFilePath As String = ""
        Dim objBookingsBal As New clsBookingsBAL
        Dim objBooking As New List(Of clsBookings)

        ' Report generation
        objBooking.Add(objBookingsBal.GetByID(Me.BookingID))

        Dim reportForm As New frmReportViewer(objBooking, "RptBooking")
        reportForm.MdiParent = Me.MdiParent
        reportForm.Show()

        ' PDF file generation
        Dim deviceInfoStr As String = "<DeviceInfo>" +
           "  <OutputFormat>PDF</OutputFormat>" +
           "  <PageWidth>8.5in</PageWidth>" +
           "  <PageHeight>11in</PageHeight>" +
           "  <MarginTop>0.2in</MarginTop>" +
           "  <MarginLeft>0.2in</MarginLeft>" +
           "  <MarginRight>0.2in</MarginRight>" +
           "  <MarginBottom>0.2in</MarginBottom>" +
           "</DeviceInfo>"
        Dim Bytes As Byte() = reportForm.RptViewer.LocalReport.Render(format:="PDF", deviceInfo:=deviceInfoStr)

        PDFFilePath = Directory.GetCurrentDirectory() + "\PDFs"
        If Not Directory.Exists(PDFFilePath) Then Directory.CreateDirectory(PDFFilePath)
        PDFFilePath = PDFFilePath + "\Booking " & txtLeadAgent.Text & " " & DateTime.Now.ToString("yyyyMMdd HHmmss") & ".pdf"

        Using stream As FileStream = New FileStream(PDFFilePath, FileMode.Create)
            stream.Write(Bytes, 0, Bytes.Length)
        End Using

        reportForm.Close()

        ' Email generation
        ' Create an instance of Outlook application
        Dim outlookApp As New Outlook.Application()

        ' Create a new MailItem
        Dim mailItem As Outlook.MailItem = DirectCast(outlookApp.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)

        ' Set the subject and body of the email
        mailItem.Attachments.Add(PDFFilePath)

        ' Display the Outlook window
        mailItem.Display()

        ' Clean up
        Marshal.ReleaseComObject(mailItem)
        Marshal.ReleaseComObject(outlookApp)
        mailItem = Nothing
        outlookApp = Nothing
    End Sub
#Region "Sale calculation"
    Private Sub TxtListingCommRate_TextChanged(sender As Object, e As EventArgs) Handles TxtListingCommRate.Leave, TxtTotConsidation.Leave, CmbTransactionType.SelectedIndexChanged, TxtListingManual.Leave, TxtProcuringCommRate.Leave, TxtProcuringManual.Leave
        HandleFeeCalculations()
    End Sub

    Private Sub HandleFeeCalculations()
        If CmbTransactionType.SelectedItem = "Sale" Then
            CalculateTotalFee(TxtListingCommRate, TxtTotConsidation, TxtListingManual, TxtLFTotalFee)
            CalculateTotalFee(TxtProcuringCommRate, TxtTotConsidation, TxtProcuringManual, TxtPFTotalFee)
        Else
            TxtLFTotalFee.Text = TxtListingManual.Text
            TxtPFTotalFee.Text = TxtProcuringManual.Text
        End If
    End Sub

    Private Sub CalculateTotalFee(commRateTextBox As TextBox, totalConsiderationTextBox As TextBox, manualFeeTextBox As TextBox, totalFeeTextBox As TextBox)
        Dim commRate As Decimal
        Dim totalConsideration As Decimal
        Dim isCommRateValid As Boolean = Decimal.TryParse(commRateTextBox.Text, commRate)
        Dim isTotalConsiderationValid As Boolean = Decimal.TryParse(totalConsiderationTextBox.Text, totalConsideration)

        If Not String.IsNullOrWhiteSpace(manualFeeTextBox.Text) Then
            totalFeeTextBox.Text = manualFeeTextBox.Text
        ElseIf isCommRateValid AndAlso isTotalConsiderationValid Then
            totalFeeTextBox.Text = (totalConsideration * commRate / 100).ToString()
        Else
            totalFeeTextBox.Text = "0"
        End If
    End Sub


    Private Sub TxtDuePct_Leave(sender As Object, e As EventArgs) Handles TxtDuePct1.Leave, TxtDuePct2.Leave, TxtDuePct3.Leave, TxtDuePct4.Leave
        CalculatePercetageTotal()
    End Sub
    Private Function CalculatePercetageTotal() As Boolean
        Dim duePct1 As Decimal = 0
        Dim duePct2 As Decimal = 0
        Dim duePct3 As Decimal = 0
        Dim duePct4 As Decimal = 0

        Decimal.TryParse(TxtDuePct1.Text, duePct1)
        Decimal.TryParse(TxtDuePct2.Text, duePct2)
        Decimal.TryParse(TxtDuePct3.Text, duePct3)
        Decimal.TryParse(TxtDuePct4.Text, duePct4)

        Dim duePctTotal As Decimal = duePct1 + duePct2 + duePct3 + duePct4
        TxtDuePctTot.Text = duePctTotal.ToString("F2")
    End Function

    Private Sub CheckBoxSendWireInst_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxSendWireInst.CheckedChanged
        If CheckBoxSendWireInst.Checked Then
            TxtEscrowNo.Visible = True
            LblEscrowNo.Visible = True
        Else
            TxtEscrowNo.Visible = False
            LblEscrowNo.Visible = False
        End If
    End Sub

    Private Sub TxtPropInfoAddress_Leave(sender As Object, e As EventArgs) Handles TxtPropInfoAddress.Leave
        TxtPropertyLease.Text = TxtPropInfoAddress.Text
    End Sub
#End Region
#End Region

End Class