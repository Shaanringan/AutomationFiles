Imports Microsoft.Office.Interop

Public Class BOMExcel

    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet

    'Dim PanelModel As New PanelAHUModels
    'Dim BoxModel As New BoxAHUModels

    Public Sub EnterValuesInCNC(ByVal PartFileName As String, ByVal width As String, ByVal height As String, ByVal thick As String, ByVal qty As Integer, ByVal Client As String, ByVal AHUName As String, ByVal JobNoFull As String)

        xlApp = GetObject("", "Excel.Application")
        xlBook = xlApp.Workbooks.Open("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNoFull & "\" & JobNoFull & "_CNC Part list.xlsx")
        xlSheet = xlBook.Worksheets("Sheet1")
        xlApp.Visible = False

        'Split Name
        Dim tempName() As String = PartFileName.Split("_")
        Dim PartNo As String = tempName(0) & "_" & tempName(1)
        Dim PartName As String = tempName(2)
        If UBound(tempName) > 2 Then
            For i = 3 To UBound(tempName)
                PartName = PartName & " " & tempName(i)
            Next
        End If

        'Get Row Number
        Dim val As String = xlSheet.Range("A9").Value
        Dim x As Integer = 9
        While val <> Nothing
            x += 1
            val = xlSheet.Range("A" & x).Value
        End While

        'Enter Data
        xlSheet.Range("A" & x).Value = PartNo
        xlSheet.Range("B" & x).Value = PartName
        xlSheet.Range("C" & x).Value = height
        xlSheet.Range("D" & x).Value = width
        xlSheet.Range("E" & x).Value = thick
        xlSheet.Range("F" & x).Value = qty.ToString

        'Add Row
        Dim addBefore As Object = x + 1
        Dim RowNos As Object = 1
        xlSheet.Rows(addBefore).Resize(RowNos).Insert()

        'Save As
        xlBook.Save()
        xlBook.Close(False)
        xlApp.Quit()

    End Sub

    Public Sub BOM_AHUEntries(ByVal FanArticleNo As String, ByVal FanDia As Integer, ByVal FanNos As Integer, ByVal WallWth As Decimal, ByVal TotalHoles As Integer, ByVal Client As String, ByVal AHUName As String, ByVal JobNoFull As String)

        xlApp = GetObject("", "Excel.Application")
        xlBook = xlApp.Workbooks.Open("C:\Program Files (x86)\Crescent Engineering\Automation\BOM\AHU_BILL OF MATERIAL.xlsx")
        xlSheet = xlBook.Worksheets("Sheet1")
        xlApp.Visible = False

        xlSheet.Range("A1").Value = "SUMMARY OF BILL OF MATERIAL_JOB NO_" & JobNoFull
        xlSheet.Range("B2").Value = "Client Name : " & Client
        xlSheet.Range("B3").Value = "AHU Name : " & AHUName

        xlSheet.Range("E2").Value = WallWth

        xlSheet.Range("C4").Value = FanArticleNo
        xlSheet.Range("D4").Value = FanDia
        xlSheet.Range("E4").Value = FanNos

        'Dim totalHoles As Integer = 0
        'If PanelModel.TotalHoles() > 0 Then
        '    totalHoles = PanelModel.TotalHoles()
        'Else
        '    'Box AHU Hole Count Pending
        'End If
        xlSheet.Range("F34").Value = TotalHoles

        xlBook.SaveAs("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNoFull & "\" & JobNoFull & "_AHU_BOM.xlsx")
        xlBook.Close(False)
        xlApp.Quit()

    End Sub

End Class