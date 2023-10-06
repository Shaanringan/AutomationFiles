Imports Microsoft.Office.Interop

Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst

Imports MySql.Data.MySqlClient

Public Class AHUForm

    Dim swApp As New SldWorks

    Dim MysqlConn As MySqlConnection
    Dim query As String
    Dim command As MySqlCommand
    Dim reader As MySqlDataReader

    Dim EnqNo As String
    Dim ClientName As String
    Dim AHUType As String
    Dim AHUName As String
    Dim PONo As String
    Dim AHUNos As Integer
    Dim BlowerNos As Integer

    Dim AHUBlowerNos As Integer
    Dim AHUArticleNo As String
    Dim FanDia As String
    Dim AHUWallHt As Integer
    Dim AHUWallWth As Integer
    Dim AHUJobNo As String

    Dim TodayDate As Date

    Dim StdFunc As New Standard_Functions

    Private Sub AHUForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        TodayDate = Today

        MysqlConn = New MySqlConnection With {.ConnectionString = "server = 127.0.0.1; userid = root; password = ; database = aadtech_erp_db"}
        MysqlConn.Open()

        'Get Company Names
        query = "SELECT DISTINCT company FROM `counter_master_data_new` WHERE status = 'Size Data Received' ORDER BY company"
        Dim ConnAdp1 As New MySqlDataAdapter(query, MysqlConn)
        Dim DataSetEnq As New DataTable
        ConnAdp1.Fill(DataSetEnq)

        'Populate Company Names DropDown Box
        ClientName_DropBox.DataSource = DataSetEnq
        ClientName_DropBox.DisplayMember = "company"

        'Get Last JobNo
        AHUJobNo = StdFunc.JobNo()

    End Sub

    Private Sub SubmitBtn_Click(sender As Object, e As EventArgs) Handles submitBtn.Click

        swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchInference, False)    'Snaping OFF

        Dim ahuIPfile As Excel.Application
        Dim ahuIPfileWB As Excel.Workbook
        Dim ahuIPfileWBS As Excel.Workbooks
        Dim ahuIPfileSheet As Excel.Worksheet

        Dim TimeStamp As Date

        ahuIPfile = CreateObject("excel.application")
        ahuIPfile.Visible = False

        Try
            ahuIPfileWBS = ahuIPfile.Workbooks
            ahuIPfileWB = ahuIPfileWBS.Open(OpenFileDialog1.FileName)
            ahuIPfileSheet = ahuIPfileWB.ActiveSheet
        Catch ex As Exception
            MsgBox("Please select an Excel Input file.")
            GoTo endLable
        End Try

        'Excel Counting Lines----------------------------------------
        Dim i As Integer
        i = 3

        While ahuIPfileSheet.Range("A" & i).Value IsNot Nothing 'Counting all entries
            i = i + 1
        End While

        Dim entryCount As Integer   'Number of entries
        entryCount = i - 1
        '------------------------------------------------------------

        Dim RemoveChar() As Char = {".", "/", "\"}

        'Start-------------------------------------------------------
        For i = 3 To entryCount
            'Excel Input-------------------------------------------------
            'AHU Type
            AHUType = ahuIPfileSheet.Range("B" & i).Value2.ToString()
            AHUType = UCase(AHUType)

            'Client Name
            Dim TempClientName As String = ahuIPfileSheet.Range("C" & i).Value2.ToString()
            ClientName = TempClientName.TrimEnd(RemoveChar)

            'AHU Name
            Dim TempAhuName As String = ahuIPfileSheet.Range("D" & i).Value2.ToString()
            AHUName = TempAhuName.TrimEnd(RemoveChar)

            ''Job Number
            'Dim ahuJobNo1, ahuJobNo2 As String     'Job No has 2 parts
            'ahuJobNo1 = ahuIPfileSheet.Range("E" & i).Value2
            'ahuJobNo2 = ahuIPfileSheet.Range("F" & i).Value2
            'If ahuJobNo2 Is Nothing Then
            '    ahuJobNo2 = ""
            'End If
            AHUJobNo = ahuIPfileSheet.Range("E" & i).Value2 & ahuIPfileSheet.Range("F" & i).Value2 'ahuJobNo1.ToString() & ahuJobNo2.ToString()

            'AHU Identifier
            Dim ahuIdent As Char     'For Package Identifyer
            ahuIdent = ahuIPfileSheet.Range("G" & i).Value2.ToString()
            ahuIdent = UCase(ahuIdent)

            'Wall Dimentions
            Dim ahuWallHt, ahuWallHtClear, ahuWallWth, ahuWallWthClear As Integer
            ahuWallHt = Convert.ToDecimal(ahuIPfileSheet.Range("H" & i).Value2)
            ahuWallHtClear = Convert.ToDecimal(ahuIPfileSheet.Range("I" & i).Value2)
            ahuWallWth = Convert.ToDecimal(ahuIPfileSheet.Range("J" & i).Value2)
            ahuWallWthClear = Convert.ToDecimal(ahuIPfileSheet.Range("K" & i).Value2)

            'Fan Details
            Dim ahuFanArticleNo As String
            Dim ahuFanNos As Integer
            ahuFanArticleNo = ahuIPfileSheet.Range("L" & i).Value2.ToString()
            ahuFanNos = Convert.ToInt16(ahuIPfileSheet.Range("M" & i).Value2)

            'Door Details
            Dim ahuDoorYesNo, ahuDoorSide As String
            Dim AHUDoor As Boolean
            Dim ahuDoorHt, ahuDoorWth As Integer
            ahuDoorYesNo = ahuIPfileSheet.Range("N" & i).Value2.ToString()
            ahuDoorYesNo = UCase(ahuDoorYesNo)
            If ahuDoorYesNo = "YES" Then
                ahuDoorSide = ahuIPfileSheet.Range("O" & i).Value2.ToString()
                ahuDoorSide = UCase(ahuDoorSide)
                AHUDoor = True

                ahuDoorHt = Convert.ToDecimal(ahuIPfileSheet.Range("P" & i).Value2)
                ahuDoorWth = Convert.ToDecimal(ahuIPfileSheet.Range("Q" & i).Value2)
            Else
                ahuDoorSide = ""
                AHUDoor = False
                ahuDoorHt = 0
                ahuDoorWth = 0
            End If
            '------------------------------------------------------------

            'Start Model Creation----------------------------------------
            If AHUType = "BOX" Then
                StdFunc.FolderCreationBox(ClientName, AHUName, AHUJobNo)
                Dim boxAHU As New BoxTypeAHU
                boxAHU.MainSub(ClientName, AHUName, AHUJobNo, ahuWallHt, ahuWallWth, ahuFanArticleNo, ahuFanNos, ahuDoorYesNo, ahuDoorSide, ahuDoorHt, ahuDoorWth)
            ElseIf AHUType = "PANEL" Then
                StdFunc.FolderCreationPanel(ClientName, AHUJobNo)
                Dim panelAHU As New PanelTypeAHU
                panelAHU.MainSub(ClientName, AHUName, AHUJobNo, ahuIdent, ahuWallHt, ahuWallWth, ahuFanArticleNo, ahuFanNos, AHUDoor)
            Else
                MsgBox("Line Number " & (i - 2) & " does not have an apropriate AHU Type (BOX/PANEL).")
            End If
            '------------------------------------------------------------
        Next
        '------------------------------------------------------------

        'Close Excel-------------------------------------------------
        ahuIPfileWB.Close()
        ahuIPfile.Quit()

        'EXCEL BACKUP------------------------------------------------
        Try
            My.Computer.FileSystem.CopyFile(OpenFileDialog1.FileName, "C:\AHU Automation - Output\AHU Project Input (" & TimeStamp.ToString(Format(Date.Now, "yyyy-MM-dd H-mm")) & ").xlsx")
        Catch ex As Exception

        End Try

endLable:

        Close()

    End Sub

    Private Sub Exit_Btn1_Click(sender As Object, e As EventArgs) Handles Exit_Btn1.Click, Exit_Btn2.Click

        MysqlConn.Close()

        MysqlConn.Dispose()

        Close()

    End Sub

    Private Sub OpenBtn_Click(sender As Object, e As EventArgs) Handles openBtn.Click

        OpenFileDialog1.Filter = "Excel (*.xlsx) | *.xlsx"
        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            InputFileBox.Text = OpenFileDialog1.FileName
        End If

    End Sub

    Private Sub ClientName_DropBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ClientName_DropBox.SelectedIndexChanged

        ClientName = ClientName_DropBox.Text.ToString

        'Get Surface Finish Data
        query = "SELECT DISTINCT sno FROM `counter_master_data_new` WHERE company = '" & ClientName & "'"
        Dim ConnAdp1 As New MySqlDataAdapter(query, MysqlConn)
        Dim DataSetEnq As New DataTable
        ConnAdp1.Fill(DataSetEnq)

        'Populate Surface Finish DropDown Box in Slide Configurator
        ERPSNo_DropBox.DataSource = DataSetEnq
        ERPSNo_DropBox.DisplayMember = "sno"

    End Sub

    Private Sub ERPSNo_DropBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ERPSNo_DropBox.SelectedIndexChanged

        Dim SNo As String = ERPSNo_DropBox.Text.ToString

        ' Enquiry Section ----------------------------------------------------------------------------------------
        ' Enguiry No
        EnqNo = StdFunc.GetFromTable("enquiry_no", "counter_master_data_new", "sno", SNo)
        EngNo_Lable.Text = EnqNo

        ' PO No
        PONo = StdFunc.GetFromTable("supply_po_number", "counter_master_data_new", "sno", SNo)
        PONo_Lable.Text = PONo

        ' AHU Nos
        query = "SELECT COUNT(enquiry_no) FROM `counter_master_data_new` WHERE enquiry_no = '" & EnqNo & "'"
        command = New MySqlCommand(query, MysqlConn)
        AHUNos = command.ExecuteScalar

        AHUCount_Label.Text = AHUNos

        ' Blower Nos
        query = "SELECT SUM(new_drive_quantity) FROM `counter_master_data_new` WHERE enquiry_no = '" & EnqNo & "'"
        command = New MySqlCommand(query, MysqlConn)
        BlowerNos = command.ExecuteScalar

        TotalBlowerCount_Label.Text = BlowerNos

        ' AHU SNo Section ----------------------------------------------------------------------------------------
        ' AHU Name
        AHUName = StdFunc.GetFromTable("product_description", "counter_master_data_new", "sno", SNo)
        AHUName_Label.Text = AHUName

        ' New Blower Qty
        AHUBlowerNos = StdFunc.GetFromTable("new_drive_quantity", "counter_master_data_new", "sno", SNo)
        FanNos_Label.Text = AHUBlowerNos

        ' Fan Dia / Article No
        AHUArticleNo = StdFunc.GetFromTable("article_no", "counter_master_data_new", "sno", SNo)
        FanDia = StdFunc.GetFromTable("blower_diameter", "counter_master_data_new", "sno", SNo)
        FanDiaArtNo_Label.Text = FanDia & " / " & AHUArticleNo

        ' Wall Dimentions
        AHUWallWth = StdFunc.GetFromTable("wall_width", "counter_master_data_new", "sno", SNo)
        AHUWallHt = StdFunc.GetFromTable("wall_height", "counter_master_data_new", "sno", SNo)
        WallDim_Label.Text = AHUWallWth & " x " & AHUWallHt

    End Sub

    Private Sub SubmitSNoFromDB_Btn_Click(sender As Object, e As EventArgs) Handles SubmitSNoFromDB_Btn.Click

        Dim SNo As String = ERPSNo_DropBox.Text

        RunAHU(SNo)

    End Sub

    Private Sub SubmitAllFromDB_Btn_Click(sender As Object, e As EventArgs) Handles SubmitAllFromDB_Btn.Click

        ClientName = ClientName_DropBox.Text

        Dim SNo As String() = StdFunc.GetColumnArray("sno", "counter_master_data_new", "company", ClientName)

        For i = 0 To UBound(SNo)
            RunAHU(SNo(i))
        Next

    End Sub

    Sub RunAHU(AHUSNo As String)

        'Variable
        AHUWallHt = Convert.ToInt32(StdFunc.GetFromTable("wall_height", "counter_master_data_new", "sno", AHUSNo).Trim)
        AHUWallWth = Convert.ToInt32(StdFunc.GetFromTable("wall_width", "counter_master_data_new", "sno", AHUSNo).Trim)
        AHUArticleNo = StdFunc.GetFromTable("article_no", "counter_master_data_new", "sno", AHUSNo)
        AHUBlowerNos = StdFunc.GetFromTable("new_drive_quantity", "counter_master_data_new", "sno", AHUSNo)
        AHUName = StdFunc.GetFromTable("product_description", "counter_master_data_new", "sno", AHUSNo)

        'Job Number
        Dim TodayYear As Integer = Year(TodayDate)
        TodayYear = Convert.ToInt32(TodayYear.ToString.Substring(2))
        Dim newJobNo As Integer
        Dim JobYear As Integer = Convert.ToInt32(AHUJobNo.Substring(3, 2))
        If JobYear = TodayYear Then
            If AHUJobNo.Length > 8 Then
                Dim tempJNo() As String = AHUJobNo.Split("_")
                newJobNo = Convert.ToInt32(tempJNo(0).Substring(5))
            Else
                newJobNo = Convert.ToInt32(AHUJobNo.Substring(5))
            End If
            AHUJobNo = "AAD" & TodayYear & newJobNo + 1
        Else
            AHUJobNo = "AAD" & TodayYear & "001"
        End If

        Dim ahuIdent As Char = "A"


        'Door Details
        'Dim AHUDoor As Boolean
        Dim ahuDoorSize, ahuDoorSide As String
        Dim ahuDoorHt, ahuDoorWth As Integer
        ahuDoorSize = StdFunc.GetFromTable("door_size", "counter_master_data_new", "sno", AHUSNo).Trim
        Dim ahuDoorYesNo As Boolean = True
        If ahuDoorSize = "0x0x0" Then ahuDoorYesNo = False
        If ahuDoorYesNo Then
            ahuDoorSide = "LHS"
            ahuDoorHt = 1000
            ahuDoorWth = 500
        Else
            ahuDoorSide = ""
            ahuDoorHt = 0
            ahuDoorWth = 0
        End If

        'Start Model Creation----------------------------------------
        If AHUBlowerNos > 6 Then 'BOX
            StdFunc.FolderCreationBox(ClientName, AHUName, AHUJobNo)
            Dim boxAHU As New BoxTypeAHU
            boxAHU.AHUSNo = AHUSNo
            boxAHU.MainSub(ClientName, AHUName, AHUJobNo, AHUWallHt, AHUWallWth, AHUArticleNo, AHUBlowerNos, ahuDoorYesNo, ahuDoorSide, ahuDoorHt, ahuDoorWth)
        Else 'PANEL Then
            StdFunc.FolderCreationPanel(ClientName, AHUJobNo)
            Dim panelAHU As New PanelTypeAHU
            panelAHU.MainSub(ClientName, AHUName, AHUJobNo, ahuIdent, AHUWallHt, AHUWallWth, AHUArticleNo, AHUBlowerNos, AHUDoor:=False)
        End If
        '------------------------------------------------------------

    End Sub

End Class