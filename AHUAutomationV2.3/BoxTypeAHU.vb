Imports System.Math

Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst

Imports Microsoft.Office.Interop

Public Class BoxTypeAHU

    Dim swApp As New SldWorks

    Dim StdFunc As New Standard_Functions
    Dim predictivedb As New PredictiveDBInput
    Dim bomData As New BOMExcel

    Dim maxBoxLth, minBoxLth As Integer
    Dim FanNoX, FanNoY As Integer
    Dim boxHt, boxWth, boxDth As Integer
    Dim SideClear, TopClear As Decimal
    Dim lastpart As Integer
    Dim TempSideClr As Decimal

    Public AHUSNo As Integer

    Public Sub MainSub(ClientName As String, AhuName As String, JobNo As String, WallHt As Integer, WallWth As Integer, fanArticleNo As String, fanNos As Integer, doorYesNo As String, doorSide As String, doorHt As Integer, doorWth As Integer)

        maxBoxLth = StdFunc.GetFromTable("Max_BoxLth", "article_no_table", "article_no", fanArticleNo)
        minBoxLth = StdFunc.GetFromTable("Min_BoxLth", "article_no_table", "article_no", fanArticleNo)

        'TEXT FILE TEST DATA----------------------------------------------
        Dim file As IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter("C:\AHU Automation - Output\" & ClientName & "\" & AhuName & "\" & JobNo & "\_" & JobNo & " - AHU Input Data.txt", True)
        file.WriteLine("---------------------------------------------------------- " & Date.Today & ", " & TimeOfDay)
        file.WriteLine("AHU BOX DATA for " & AhuName)
        file.WriteLine("")
        file.WriteLine("Wall Height - " & WallHt)
        file.WriteLine("Wall Width - " & WallWth)
        file.WriteLine("Fan Article No. - " & fanArticleNo)
        file.WriteLine("Total No of fans = " & fanNos)

#Region "Fan Placement Grid-----------------------------------------------"
        Dim WallArea As Integer = WallHt * WallWth
        Dim MaxFanArea As Integer = maxBoxLth * maxBoxLth * fanNos
        Dim MinFanArea As Integer = minBoxLth * minBoxLth * fanNos
        If WallArea - MinFanArea < 0 Then
            file.WriteLine(vbNewLine & "---------------Wall Dimensions Insufficient---------------" & vbNewLine & "----------------------------------------------------------")
            MsgBox("Wall Dimensions Insufficient for " & AhuName)
            Exit Sub
        End If

        FanNoX = Ceiling(Sqrt(fanNos))
        FanNoY = Ceiling(fanNos / FanNoX)

FanCheck:

        Dim CheckX As Boolean = False
        Dim CheckY As Boolean = False

        If WallWth / (maxBoxLth * FanNoX) >= 1 Then CheckX = True
        If CheckX = False Then
            If WallWth / (minBoxLth * FanNoX) >= 1 Then CheckX = True
        End If

        If WallHt / (maxBoxLth * FanNoY) >= 1 Then CheckY = True
        If CheckY = False Then
            If WallHt / (minBoxLth * FanNoY) >= 1 Then CheckY = True
        End If

        If CheckX And CheckY Then

        Else
            If CheckX Then
                FanNoX += 1
                FanNoY -= 1
                GoTo FanCheck
            Else
                FanNoX -= 1
                FanNoY += 1
                GoTo FanCheck
            End If
        End If

        Dim HoleCD As Integer = StdFunc.GetFromTable("Fan_CD", "article_no_table", "article_no", fanArticleNo)
        Dim FanDia As Integer = StdFunc.GetFromTable("diameter", "article_no_table", "article_no", fanArticleNo)
        Dim FanRingCD As Integer = StdFunc.GetFromTable("Fan_Ring_CD", "article_no_table", "article_no", fanArticleNo)

        'Update Notepad File
        file.WriteLine("Fan Grid (W x H) = " & FanNoX & " x " & FanNoY)
        file.WriteLine("")

#End Region

#Region "Box Dimensions---------------------------------------------------"
        boxWth = Truncate(WallWth / FanNoX / 10) * 10
        Select Case boxWth
            Case > maxBoxLth
                boxWth = maxBoxLth
            Case < minBoxLth
                boxWth = minBoxLth
        End Select

        boxHt = Truncate(WallHt / FanNoY / 10) * 10
        Select Case boxHt
            Case > maxBoxLth
                boxHt = maxBoxLth
            Case < minBoxLth
                boxHt = minBoxLth
        End Select

        boxDth = StdFunc.GetFromTable("Box_Dth", "article_no_table", "article_no", fanArticleNo)

        'Update Notepad File
        file.WriteLine("Box Size (W x H x D) = " & boxWth & " x " & boxHt & " x " & boxDth)
        file.WriteLine("")

        SideClear = WallWth - (FanNoX * boxWth) - 100
        TopClear = WallHt - (FanNoY * boxHt) - 50

#End Region

#Region "Frame Dimensions-------------------------------------------------"
        Dim MaxSecLth As Integer = 1100
        Dim MinBlankWth As Integer = 200
        Dim InnerWallHt As Integer = WallHt - 200
        If TopClear < MinBlankWth Then
            InnerWallHt = (FanNoY * boxHt)
        End If
        Dim InnerWallWth As Integer = WallWth
        Dim VerSecNo As Integer = Ceiling(InnerWallHt / MaxSecLth)
        Dim HorSecNo As Integer = Ceiling(InnerWallWth / MaxSecLth)
        Dim VerSecLth(VerSecNo - 1) As Integer
        Dim HorSecLth(HorSecNo - 1) As Integer

        'Outter Section Width
        Dim SecWthSide As Integer = 100
        If SideClear - doorWth < MinBlankWth * 2 Then
            SecWthSide += SideClear / 2
        End If

        Dim SecWthTop As Integer = 100
        If TopClear < MinBlankWth Then
            SecWthTop += TopClear
        End If

        'Vertical Sections - ALL
        Dim VerMaxSecLth As Integer = MaxSecLth
        Dim VerSecLth_Last As Integer = InnerWallHt - (VerMaxSecLth * (VerSecNo - 1))
        While VerSecLth_Last < 300
            VerMaxSecLth -= 10
            VerSecLth_Last = InnerWallHt - (VerMaxSecLth * (VerSecNo - 1))
        End While

        For i = 0 To UBound(VerSecLth)
            If InnerWallHt <= VerMaxSecLth Then
                VerSecLth(i) = InnerWallHt
            Else
                VerSecLth(i) = VerMaxSecLth
            End If
            InnerWallHt -= VerMaxSecLth
        Next

        'Horizontal Sections - TOP & BOT
        Dim HorMaxSecLth As Integer = MaxSecLth
        Dim HorSecLth_Last As Integer = WallWth - (HorMaxSecLth * (HorSecNo - 1))
        While HorSecLth_Last < 300
            HorMaxSecLth -= 10
            HorSecLth_Last = WallWth - (HorMaxSecLth * (HorSecNo - 1))
        End While

        For i = 0 To UBound(HorSecLth)
            If InnerWallWth < HorMaxSecLth Then
                HorSecLth(i) = InnerWallWth
            Else
                HorSecLth(i) = HorMaxSecLth
            End If
            InnerWallWth -= HorMaxSecLth
        Next

#End Region

#Region "CNC Cutlist------------------------------------------------------"
        Dim xlApp As Excel.Application
        Dim xlBook As Excel.Workbook
        Dim xlSheet As Excel.Worksheet

        xlApp = GetObject("", "Excel.Application")
        xlBook = xlApp.Workbooks.Open("C:\Program Files (x86)\Crescent Engineering\Automation\BOM\CNC PART LIST.xlsx")
        xlSheet = xlBook.Worksheets("Sheet1")
        xlApp.Visible = False

        xlSheet.Range("A3").Value = "Client Name:- " & ClientName
        xlSheet.Range("C4").Value = "Job No:- " & JobNo

        xlBook.SaveAs("C:\AHU Automation - Output\" & ClientName & "\" & AhuName & "\" & JobNo & "\" & JobNo & "_CNC Part list.xlsx")
        xlBook.Close(False)
        xlApp.Quit()

#End Region

#Region "BOM--------------------------------------------------------------"
        Dim totalHoles As Integer = 0
        bomData.BOM_AHUEntries(fanArticleNo, FanDia, fanNos, WallWth, totalHoles, ClientName, AhuName, JobNo)

#End Region

#Region "Start Model Creations--------------------------------------------"

        swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchInference, False)    'Snaping OFF

        Dim BoxAHUModel As New BoxAHUModels
        BoxAHUModel.Client = ClientName
        BoxAHUModel.AHUName = AhuName
        BoxAHUModel.JobNo = JobNo
        BoxAHUModel.ArticleNoFan = fanArticleNo

#Region "Box----------------------"
        BoxAHUModel.BackSheet(boxWth / 1000, boxHt / 1000, boxDth / 1000, FanDia, HoleCD / 1000)
        BoxAHUModel.LHSBeam(boxHt / 1000, boxDth / 1000)
        BoxAHUModel.LHSSheet(boxHt / 1000, boxDth / 1000, HoleCD / 1000)
        BoxAHUModel.BottomSupport(boxDth / 1000)
        BoxAHUModel.RHSBeam(boxHt / 1000, boxDth / 1000)
        BoxAHUModel.FrontTopBeam(boxWth / 1000, boxHt / 1000, boxDth / 1000, HoleCD / 1000)
        BoxAHUModel.FanVerticalStand(HoleCD / 1000, FanRingCD / 1000)
        BoxAHUModel.FanHorizontalStand()

        BoxAHUModel.BackSheetDrawing()
        BoxAHUModel.LHSDrawing()
        BoxAHUModel.BotSupp_RHSBeam_FrontTopBeam_Drawing()

        BoxAHUModel.BoxAssembly(boxWth / 1000, boxHt / 1000, boxDth / 1000, HoleCD / 1000)

        BoxAHUModel.BoxAssyDrawing()

#End Region

#Region "Frame--------------------"
        'Horizontal Section -BOT & TOP

        For i = 0 To UBound(HorSecLth)
            lastpart = UBound(HorSecLth) + 1
            BoxAHUModel.FrameBotSec(HorSecLth(i) / 1000, 0.1, HorSecLth, WallWth, boxWth, FanNoX, SecWthSide, doorWth, i + 1, lastpart)
            BoxAHUModel.FrameTopSec(HorSecLth(i) / 1000, SecWthTop / 1000, HorSecLth, WallWth, boxWth, FanNoX, SecWthSide, doorWth, i + 1, lastpart)
        Next

        'Vertical Sections -SIDE
        For i = 0 To UBound(VerSecLth)
            lastpart = UBound(VerSecLth) + 1
            BoxAHUModel.FrameSideSecRHS(SecWthSide / 1000, VerSecLth(i) / 1000, VerSecLth, WallHt, boxHt, i + 1, lastpart)
            BoxAHUModel.FrameSideSecLHS(SecWthSide / 1000, VerSecLth(i) / 1000, VerSecLth, WallHt, boxHt, i + 1, lastpart)
        Next

        'Vertical Sections -MID
        For i = 0 To UBound(VerSecLth)
            lastpart = UBound(VerSecLth) + 1
            BoxAHUModel.FrameMidVerSec(VerSecLth(i) / 1000, VerSecLth, WallHt, boxHt, i + 1, lastpart)
        Next

        'Horizontal Sections -MID
        BoxAHUModel.FrameMidHorSec(boxWth / 1000, "")
        If doorYesNo = "YES" Then
            BoxAHUModel.FrameMidHorSec(doorWth / 1000, "_Door")
        End If


        If SideClear / 2000 > boxWth / 1000 Then
            TempSideClr = SideClear / 2000
            While TempSideClr > boxWth / 1000
                TempSideClr = TempSideClr - boxWth / 1000
            End While
            BoxAHUModel.FrameMidHorSec(TempSideClr - 0.102, "_BLANK")
        End If


        'If (SideClear - doorWth) >= (MinBlankWth * 2) Then

        'End If

        'Frame Sub Assy
        BoxAHUModel.FrameSubAssy(WallWth / 1000, WallHt / 1000, boxWth / 1000, boxHt / 1000, FanNoX, FanNoY, HorSecLth, VerSecLth, SideClear / 2000, TopClear, doorSide, doorWth / 1000, doorHt / 1000)

        'Frame Drawings
        BoxAHUModel.FrameHorSecDrawings(HorSecNo)
        BoxAHUModel.FrameVerSecDrawings(VerSecNo)

        BoxAHUModel.FrameAssyDrawings()

#End Region

#Region "Blanks-------------------"
        'Box Blank
        If fanNos < FanNoY * FanNoX Then
            BoxAHUModel.BlankSheet(boxHt / 1000, boxWth / 1000, "Box")
            BoxAHUModel.BlankSheetDrawing("Box")
        End If

        'Side Blanks
        If (SideClear - doorWth) >= (MinBlankWth * 2) Then
            BoxAHUModel.BlankSheet(2 * boxHt / 1000, (SideClear - doorWth) / 2000, "Side Clearance -1")
            BoxAHUModel.BlankSheetDrawing("Side Clearance -1")
            If (FanNoY Mod 2) > 0 Then
                BoxAHUModel.BlankSheet(boxHt / 1000, (SideClear - doorWth) / 2000, "Side Clearance -2")
                BoxAHUModel.BlankSheetDrawing("Side Clearance -2")
            End If
        End If

        'Top Blanks
        If TopClear >= MinBlankWth Then
            BoxAHUModel.BlankSheet(TopClear / 1000, 2 * boxWth / 1000, "Top Clearance -1")
            BoxAHUModel.BlankSheetDrawing("Top Clearance -1")
            If (FanNoX Mod 2) = 0 Then
                If doorYesNo = "YES" Then
                    BoxAHUModel.BlankSheet(TopClear / 1000, doorWth / 1000, "Top Clearance -2")
                    BoxAHUModel.BlankSheetDrawing("Top Clearance -2")
                End If
            Else
                If doorYesNo = "YES" Then
                    BoxAHUModel.BlankSheet(TopClear / 1000, (doorWth + boxWth) / 1000, "Top Clearance -2")
                    BoxAHUModel.BlankSheetDrawing("Top Clearance -2")
                Else
                    BoxAHUModel.BlankSheet(TopClear / 1000, boxWth / 1000, "Top Clearance -2")
                    BoxAHUModel.BlankSheetDrawing("Top Clearance -2")
                End If
            End If
        End If

        'Top Side Corner
        If (SideClear - doorWth) >= (MinBlankWth * 2) And TopClear >= MinBlankWth Then
            BoxAHUModel.BlankSheet(TopClear / 1000, SideClear / 2000, "Top Corner")
            BoxAHUModel.BlankSheetDrawing("Top Corner")
        End If

        'Door Blank
        If doorYesNo = "YES" Then
            BoxAHUModel.BlankSheet(((FanNoY * boxHt) - doorHt) / 1000, doorWth / 1000, "Door Blank")
            BoxAHUModel.BlankSheetDrawing("Door Blank")
        End If

#End Region

#Region "Door---------------------"
        If doorYesNo = "YES" Then
            BoxAHUModel.DoorFrontSheet((doorWth - 55) / 1000, doorHt / 1000, ClientName, AhuName, JobNo)
            BoxAHUModel.DoorVerticalCSec(doorHt / 1000, ClientName, AhuName, JobNo, doorSide)
            BoxAHUModel.DoorHorizontalCSec((doorWth - 55) / 1000, ClientName, AhuName, JobNo)
            BoxAHUModel.DoorVerSupportCSec(doorHt / 1000, ClientName, AhuName, JobNo)

            BoxAHUModel.DoorSubAssy((doorWth - 55) / 1000, doorHt / 1000, ClientName, AhuName, JobNo, doorSide)
        End If

#End Region

        'AHU Final Assy-----------
        BoxAHUModel.BoxAHUFinalAssy(boxWth / 1000, boxHt / 1000, fanNos, FanNoY, FanNoX, SideClear / 2000, TopClear / 1000, doorYesNo, doorSide, doorWth / 1000, doorHt / 1000)
        BoxAHUModel.BoxAHUFinalAssyDrawings()

#End Region

#Region "Add To Database--------------------------------------------------"

        For j = 1 To fanNos
            ''Box Parts
            'predictivedb.AHUPartCount("BoxBackSheet_" & Convert.ToString(Truncate(boxWth)) & "x" & Convert.ToString(Truncate(boxHt)) & "x" & Convert.ToString(Truncate(boxDth)) & "_" & Convert.ToString(Truncate(FanDia)))
            'predictivedb.AHUPartCount("BoxLHSSheet_" & Convert.ToString(Truncate(boxHt)) & "x" & Convert.ToString(Truncate(boxDth)))
            'predictivedb.AHUPartCount("BoxLHSSheet_" & Convert.ToString(Truncate(boxDth)))
            'predictivedb.AHUPartCount("BoxRHSBeam_" & Convert.ToString(Truncate(boxHt)) & "x" & Convert.ToString(Truncate(boxDth)))
            'predictivedb.AHUPartCount("BoxFrontTopBeam_" & Convert.ToString(Truncate(boxWth)))
            'For i = 1 To 4
            '    predictivedb.AHUPartCount("Fan-" & fanArticleNo & "_Z-Stand_" & Truncate(HoleCD - 50 - 20 - FanRingCD))
            'Next
            'For i = 1 To 2
            '    predictivedb.AHUPartCount("Fan-" & fanArticleNo & "_Mid-Stand_" & Truncate(Convert.ToDecimal(StdFunc.GetFromTable("Stnd_Dist", "article_no_table", "article_no", fanArticleNo))))
            'Next

            ''Fan Assembly
            'predictivedb.AHUFanCount("Fan-" & fanArticleNo)
            'predictivedb.AHUPartCount("Fan-" & fanArticleNo & "_MotorMountingPlate")
            'For i = 1 To 4
            '    predictivedb.AHUPartCount("Fan-" & fanArticleNo & "_SupportArms")
            'Next
        Next
        '-----------------------------------------------------------------

        'Close Notepad File
        file.WriteLine("---------------------------------------------------------- " & Date.Today & ", " & TimeOfDay)
        file.Close()

        swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchInference, True)     'Snaping ON

#End Region

    End Sub

End Class