Imports System.IO
Imports System.Math

Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst

Imports Microsoft.Office.Interop

Public Class PanelTypeAHU

    Dim swApp As New SldWorks

    Dim Part As ModelDoc2
    Dim Draw As DrawingDoc
    Dim Assy As AssemblyDoc
    Dim boolstatus As Boolean
    Dim longstatus As Long, longwarnings As Long
    Dim myDimension As Dimension

    Dim ExitSub As Boolean

    Dim maxPanelLth, minPanelLth As Decimal
    Dim FanNoX, FanNoY As Integer
    Dim panelHt, panelWth As Decimal
    Dim SideClear, TopClear As Decimal
    Dim BlkNosX, BlkNosY As Integer
    Dim CrnBlkHt As Decimal
    Dim CrnBlkWth As Decimal

    Dim TopXHoles As Boolean
    Dim SideYHoles As Boolean

    Dim PushSide As Boolean
    Dim DoorBlkWth As Decimal
    Dim DoorBlkHt As Decimal

    Dim SideBlkWth() As Decimal

    Dim StdFunc As New Standard_Functions
    Dim predictivedb As New PredictiveDBInput
    Dim bomData As New BOMExcel

    Dim MaxSecLth3mm As Decimal = 3000
    Dim MinBlankLth As Decimal = 120
    Dim SideLWth As Decimal = 50
    Dim TopLWth As Decimal = 50
    Dim MinDoorPnlWth As Decimal = 490    'Door = 350
    Dim MaxDoorPnlWth As Decimal = 890    'Door = 750

    Public Sub MainSub(ClientName As String, AhuName As String, JobNo As String, Identifier As Char, WallHt As Decimal, WallWth As Decimal, fanArticleNo As String, FanNos As Integer, AHUDoor As Boolean)

        MsgBox("GITHUB")

        maxPanelLth = StdFunc.GetFromTable("Max_PanLth", "article_no_table", "article_no", fanArticleNo)
        minPanelLth = StdFunc.GetFromTable("Min_PanLth", "article_no_table", "article_no", fanArticleNo)

        FanNoX = FanNos
        FanNoY = 1

        'TEXT FILE TEST DATA----------------------------------------------
        Dim file As IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter("C:\AHU Automation - Output\" & ClientName & "\" & JobNo & "\_" & JobNo & " - AHU Input Data.txt", True)
        file.WriteLine("---------------------------------------------------------- " & Date.Today & ", " & TimeOfDay)
        file.WriteLine("AHU PANEL DATA for " & AhuName)
        file.WriteLine("")
        file.WriteLine("Wall Height - " & WallHt)
        file.WriteLine("Wall Width - " & WallWth)
        file.WriteLine("Fan Article No. - " & fanArticleNo)
        file.WriteLine("Total No of fans = " & FanNos)

#Region "Fan Placement Grid-----------------------------------------------"

        Dim WallArea As Decimal = WallHt * WallWth
        Dim MaxFanArea As Decimal = maxPanelLth * maxPanelLth * FanNos
        Dim MinFanArea As Decimal = minPanelLth * minPanelLth * FanNos

        If WallArea - MinFanArea < 0 Then
            file.WriteLine(vbNewLine & "---------------Wall Dimensions Insufficient---------------" & vbNewLine & "----------------------------------------------------------")
            MsgBox("Wall Dimensions Insufficient for " & AhuName)
            Exit Sub
        End If

        '3 Fan Special Condition
        If FanNos = 3 Then
            If WallWth - (minPanelLth * FanNos) > 0 Then
                FanNoX = 3
                FanNoY = 1


                If AHUDoor = True Then
                    While WallWth - (minPanelLth * FanNoX) < MinDoorPnlWth
                        FanNoX -= 1
                        FanNoY += 1
                    End While

                    If WallHt - (minPanelLth * FanNoY) < 0 Then
                        file.WriteLine(vbNewLine & "---------------Wall Dimensions Insufficient---------------" & vbNewLine & "----------------------------------------------------------")
                        MsgBox("Wall Width Insufficient for " & AhuName)
                        Exit Sub
                    End If
                End If

                GoTo ContinueProcess
            Else
                file.WriteLine(vbNewLine & "---------------Wall Dimensions Insufficient---------------" & vbNewLine & "----------------------------------------------------------")
                MsgBox("Wall Dimensions are insufficient" & AhuName & vbNewLine & "AHU Automation will continue in 3 Fan L Type")
            End If
        End If

        FanNoX = Ceiling(Sqrt(FanNos))
        FanNoY = Ceiling(FanNos / FanNoX)


        Dim CheckX As Boolean = False
        Dim CheckY As Boolean = False

        Dim LastFanNoX As Integer
        Dim LastFanNoY As Integer
        Dim LoopCount As Integer = 0

FanCheck:
        If WallWth / (maxPanelLth * FanNos) >= 1 Then
            CheckX = True
            FanNoX = FanNos
            FanNoY = 1
            GoTo CheckY
        End If


        If WallWth / (maxPanelLth * FanNoX) >= 1 Then CheckX = True
        If CheckX = False Then
            If WallWth / (minPanelLth * FanNoX) >= 1 Then CheckX = True
        End If

CheckY:
        If WallHt / (maxPanelLth * FanNoY) >= 1 Then CheckY = True
        If CheckY = False Then
            If WallHt / (minPanelLth * FanNoY) >= 1 Then CheckY = True
        End If

        If AHUDoor = True Then
            If WallWth - (minPanelLth * FanNos) < 0 Then
                While WallWth - (minPanelLth * FanNoX) < MinDoorPnlWth
                    FanNoX -= 1
                FanNoY += 1
                End While
            Else
                While WallWth - (minPanelLth * FanNoX) < MinDoorPnlWth
                    FanNoX -= 1
                    FanNoY += 1
                End While
            End If


            If WallHt - (minPanelLth * FanNoY) < 0 Then
                file.WriteLine(vbNewLine & "---------------Wall Dimensions Insufficient---------------" & vbNewLine & "----------------------------------------------------------")
                MsgBox("Wall Width Insufficient for " & AhuName)
                Exit Sub
            End If
        End If

        If CheckX = True And CheckY = True Then
            GoTo ContinueProcess
        End If

        If CheckX And CheckY Then
            file.WriteLine(vbNewLine & "---------------Wall Dimensions Insufficient---------------" & vbNewLine & "----------------------------------------------------------")
            MsgBox("Wall Dimensions Insufficient for " & AhuName)
            Exit Sub
        Else

            If FanNoX = LastFanNoY And FanNoY = LastFanNoX Then
                file.WriteLine(vbNewLine & "---------------Wall Dimensions Insufficient---------------" & vbNewLine & "----------------------------------------------------------")
                MsgBox("Wall Dimensions Insufficient for " & AhuName)
                Exit Sub
            End If

            If CheckX Then
                LastFanNoX = FanNoX
                LastFanNoY = FanNoY
                FanNoX += 1
                FanNoY -= 1
            Else
                LastFanNoX = FanNoX
                LastFanNoY = FanNoY
                FanNoX -= 1
                FanNoY += 1
            End If


            ' 1 fan condition (FanNoX = 0 or FanNoY = 0)
            If FanNoX = 0 Then
                file.WriteLine(vbNewLine & "---------------Wall Dimensions Insufficient---------------" & vbNewLine & "----------------------------------------------------------")
                MsgBox("Wall Width Insufficient for " & AhuName)
                Exit Sub
            ElseIf FanNoY = 0 Then
                file.WriteLine(vbNewLine & "---------------Wall Dimensions Insufficient---------------" & vbNewLine & "----------------------------------------------------------")
                MsgBox("Wall Height Insufficient for " & AhuName)
                Exit Sub
            End If

            CheckX = False
            CheckY = False

            LoopCount = LoopCount + 1

            If LoopCount > 5 Then
                file.WriteLine(vbNewLine & "---------------Wall Dimensions Insufficient---------------" & vbNewLine & "----------------------------------------------------------")
                MsgBox("Wall Height or Width is insufficient for " & AhuName)
                Exit Sub
            End If

            GoTo FanCheck
        End If

        'CHECK TOTAL NO OF FANS VS FANS ASSEMBLE
ContinueProcess:
        Dim FanDia As Integer = StdFunc.GetFromTable("diameter", "article_no_table", "article_no", fanArticleNo)

        'Update Notepad File
        file.WriteLine("Fan Grid (W x H) = " & FanNoX & " x " & FanNoY)
        file.WriteLine("")

#End Region

#Region "Panel Dimensions-------------------------------------------------"

        If AHUDoor = True Then
            panelWth = minPanelLth
            GoTo pnlhtcalc
        End If

        panelWth = (WallWth / FanNoX / 10) * 10
        Select Case panelWth
            Case < maxPanelLth
                If panelWth > minPanelLth Then
                    If MinBlankLth < WallWth - (minPanelLth * FanNoY) Then
                        panelHt = minPanelLth
                    Else
                        GoTo pnlhtcalc
                    End If
                End If
            Case < minPanelLth
                panelWth = minPanelLth
            Case Else
                panelWth = minPanelLth
        End Select

pnlhtcalc:
        panelHt = (WallHt / FanNoY / 10) * 10
        Select Case panelHt
            Case < maxPanelLth
                If panelHt > minPanelLth Then
                    If MinBlankLth < WallHt - (minPanelLth * FanNoY) Then
                        panelHt = minPanelLth
                    Else
                        GoTo clrcalc
                    End If

                End If
            Case < minPanelLth
                panelHt = minPanelLth
            Case Else
                panelHt = minPanelLth
        End Select

clrcalc:
        SideClear = WallWth - (FanNoX * panelWth) '- 100
        TopClear = WallHt - (FanNoY * panelHt) '- 50


        If AHUDoor = True Then
            If SideClear < MinDoorPnlWth Then
                MsgBox("Insufficient space for Door")
                Exit Sub
            End If
        End If


        'Clearance Y less than 120 condition -
        If MinBlankLth > TopClear And TopClear > 0 Then
            BlkNosY = 1
            panelHt = minPanelLth
            TopClear = WallHt - (panelHt * FanNoY)
        ElseIf TopClear = 0 Then
            panelHt = WallHt / FanNoY
            BlkNosY = 0
        End If

        If SideClear < MinBlankLth And SideClear > 0 Then
            panelWth = WallWth / FanNoX
        End If

        SideClear = WallWth - (FanNoX * panelWth)

NotePad:
        'Update Notepad File
        file.WriteLine("Motor Panel Size (W x H) = " & panelWth & " x " & panelHt)
        file.WriteLine("")

#End Region

#Region "Blank Calculations ----------------------------------------------"

        If SideClear = 0 Then
            GoTo TopBlanks
        End If

        'Side Blanks
        Dim EndBlkWth As Decimal

        If AHUDoor = True Then
            If SideClear > MaxDoorPnlWth Then
                BlkNosX = 2
                EndBlkWth = SideClear - MaxDoorPnlWth

                If EndBlkWth < MinBlankLth Then
                    EndBlkWth = MinBlankLth
                    SideClear = SideClear - MinBlankLth
                Else
                    SideClear = MaxDoorPnlWth
                End If

                GoTo TopBlanks
            Else
                BlkNosX = 1
                EndBlkWth = SideClear
                PushSide = True
                GoTo TopBlanks
            End If

        End If

        If SideClear > MinBlankLth * 2 Then
            BlkNosX = Ceiling(SideClear / panelWth / 2)
            EndBlkWth = (SideClear / 2) - (panelWth * (BlkNosX - 1))
            SideClear /= 2
        ElseIf SideClear >= MinBlankLth Then
            BlkNosX = Ceiling(SideClear / panelWth)
            EndBlkWth = SideClear - (panelWth * (BlkNosX - 1))
            PushSide = True
        Else
            BlkNosX = 0
        End If

        If BlkNosX > 1 Then
            If EndBlkWth <= MinBlankLth Then
                BlkNosX -= 1
            End If
        End If

TopBlanks:
        If TopClear = 0 Then
            GoTo StartModelCreation
        End If

        'Top Blanks
        Dim EndBlkHt As Decimal
        If TopClear >= MinBlankLth Then
            BlkNosY = Ceiling(TopClear / panelHt)
            EndBlkHt = TopClear - (panelHt * (BlkNosY - 1))
        Else
            EndBlkHt = TopClear
            BlkNosY = 1
        End If

        Dim Crnblk As Boolean = False
        If BlkNosX > 0 And BlkNosY > 0 Then
            Crnblk = True
        End If

#End Region

StartModelCreation:
#Region "Start Model Creations--------------------------------------------"

        swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchInference, False)    'Snaping OFF
        swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swInputDimValOnCreate, False)  'Input Dimension OFF

        Dim PanelAHUModel As New PanelAHUModels
        PanelAHUModel.Client = ClientName
        PanelAHUModel.AHUName = AhuName
        PanelAHUModel.JobNo = JobNo
        PanelAHUModel.ArticleNoFan = fanArticleNo
        PanelAHUModel.WallWth = WallWth
        PanelAHUModel.WallHt = WallHt
        PanelAHUModel.SaveFolder = "C:\AHU Automation - Output\" & ClientName & "\" & JobNo

        'Start Model Creations--------------------------------------------
        PanelAHUModel.MotorSelection(FanDia, fanArticleNo)
        PanelAHUModel.MotorSheet(panelWth / 1000, panelHt / 1000, FanDia, FanNoX, FanNoY, FanNos)
        PanelAHUModel.MotorSubAssembly(fanArticleNo)

        ''Blanks-----------------------------------------------------------

        If AHUDoor = True Then
            If BlkNosX > 1 Then
                If FanDia = "560" Or FanDia = "500" Or FanDia = "450" Then
                    SideYHoles = False
                Else
                    SideYHoles = True
                End If
                TopXHoles = False

                PanelAHUModel.Blanks(EndBlkWth / 1000, panelHt / 1000, "_07_Side Blank_A1", 1, Crnblk, BlkNosY, TopXHoles, SideYHoles, 1)

                If FanNoY > 1 Then
                    PanelAHUModel.Blanks(EndBlkWth / 1000, panelHt / 1000, "_07_Side Blank_B1", 1, Crnblk, BlkNosY, TopXHoles, SideYHoles, 1)
                End If

                GoTo TopBlankCalculation
            Else
                GoTo TopBlankCalculation
            End If
        End If

        Dim BlkQty As Integer = 0

        If SideClear <= 0 Then
            GoTo TopBlankCalculation
        End If

        '---------Side Blanks------------
        Dim BlkWth As Decimal = SideClear
        Dim BlankWth(BlkNosX - 1) As Decimal
        If BlkNosX > 0 Then
            For i = 1 To BlkNosX
                If i = BlkNosX And EndBlkWth <= MinBlankLth Then
                    BlankWth(i - 1) = BlkWth
                Else
                    If BlkWth > panelWth Then
                        If BlkWth > minPanelLth And BlkWth < maxPanelLth Then
                            BlankWth(i - 1) = BlkWth
                        Else
                            BlankWth(i - 1) = panelWth
                        End If
                    Else
                        BlankWth(i - 1) = BlkWth
                    End If
                End If

                If PushSide = False Then
                    BlkQty = BlkNosX * 2
                Else
                    BlkQty = BlkNosX
                End If

                If FanDia = "560" Or FanDia = "500" Or FanDia = "450" Then
                    SideYHoles = False
                Else
                    SideYHoles = True
                End If
                TopXHoles = False

                PanelAHUModel.Blanks(BlankWth(i - 1) / 1000, panelHt / 1000, "_07_Side Blank_A" & i, BlkQty, Crnblk, BlkNosY, TopXHoles, SideYHoles, i)

                If FanNoY > 1 Then
                    PanelAHUModel.Blanks(BlankWth(i - 1) / 1000, panelHt / 1000, "_07_Side Blank_B" & i, BlkQty, Crnblk, BlkNosY, TopXHoles, SideYHoles, i)
                End If

                BlkWth -= panelWth
            Next
        End If

TopBlankCalculation:

        Dim BlkHt As Decimal = TopClear
        Dim BY As Decimal = BlkNosY
        Dim BlankHt(3) As Decimal
        BlankHt(0) = 0
        BlankHt(1) = 0
        BlankHt(2) = 0

        If TopClear <= 0 And AHUDoor = True Then
            GoTo DoorModel
        ElseIf TopClear <= 0 Then
            GoTo LandCSections
        End If

        '----------Top Blanks-----------
        If BlkNosY > 0 Then
            For i = 1 To BlkNosY
                'If i = BlkNosY And EndBlkHt <= MinBlankLth Then
                If i = BlkNosY Then
                    'If EndBlkHt <= MinBlankLth Then
                    '    EndBlkHt = TopClear
                    'End If
                    BlankHt(i - 1) = BlkHt
                Else
                    If BlkHt > panelHt Then
                        If BlkHt - panelHt >= MinBlankLth Then
                            BlankHt(i - 1) = panelHt
                            BlkHt -= panelHt
                        Else
                            BlankHt(i - 1) = 400
                            BlkHt -= 400
                        End If
                    Else
                        BlankHt(i - 1) = BlkHt
                    End If
                End If

                If FanDia = "560" Or FanDia = "500" Or FanDia = "450" Then
                    TopXHoles = False
                Else
                    TopXHoles = True
                End If
                SideYHoles = False

                BlkQty = FanNoX
                PanelAHUModel.Blanks(panelWth / 1000, BlankHt(i - 1) / 1000, "_08_Top Blank_" & i, BlkQty, Crnblk, BlkNosY, TopXHoles, SideYHoles, i)

                ' BlkHt -= panelHt
            Next
        End If

DoorModel:
        Dim DoorCase As String
        If AHUDoor = True Then
            DoorBlkWth = SideClear / 1000

            If FanNoY >= 1 And BlkNosY = 0 Then
                DoorBlkHt = WallHt / 1000
                DoorCase = "Door 1"
                PanelAHUModel.DoorParts(DoorBlkWth, DoorBlkHt, BlkNosY, FanNoY, DoorCase, panelHt, BlankHt)

            ElseIf FanNoY > 1 And BlkNosY >= 1 Then
                DoorBlkHt = (panelHt * FanNoY) / 1000
                DoorCase = "Door 2"
                PanelAHUModel.DoorParts(DoorBlkWth, DoorBlkHt, BlkNosY, FanNoY, DoorCase, panelHt, BlankHt)

                TopXHoles = False
                SideYHoles = False

                'If EndBlkWth <> SideClear Then
                '    For i = 1 To FanNoY
                '        PanelAHUModel.Blanks(EndBlkWth, panelHt / 1000, "_07_Side Blank_A" & i, 1, Crnblk, BlkNosY, TopXHoles, SideYHoles, i)
                '    Next
                'End If

                'If BlkNosX < 2 Then
                '    BY = BlkNosY - 1
                'End If

                For i = 1 To BY
                    PanelAHUModel.Blanks(DoorBlkWth, BlankHt(i - 1) / 1000, "_09_Corner Blank_" & i, 1, Crnblk, BlkNosY, TopXHoles, SideYHoles, i)
                Next

                For j = BY + 1 To BY * 2
                    PanelAHUModel.Blanks(EndBlkWth / 1000, BlankHt(j - (BY + 1)) / 1000, "_09_Corner Blank_" & j, 1, Crnblk, BlkNosY, TopXHoles, SideYHoles, j)
                Next

            ElseIf FanNoY = 1 And BlkNosY >= 1 Then
                DoorBlkHt = (panelHt + BlankHt(0)) / 1000
                DoorCase = "Door 3"
                PanelAHUModel.DoorParts(DoorBlkWth, DoorBlkHt, BlkNosY, FanNoY, DoorCase, panelHt, BlankHt)

                TopXHoles = False
                SideYHoles = False

                If BlkNosY > 1 Then

                    PanelAHUModel.Blanks(DoorBlkWth, BlankHt(1) / 1000, "_09_Corner Blank_" & 1, 1, Crnblk, BlkNosY, TopXHoles, SideYHoles, 1)

                    If BlkNosX >= 2 Then
                        If SideClear + EndBlkWth > MaxDoorPnlWth Then
                            PanelAHUModel.Blanks(EndBlkWth / 1000, BlankHt(0) / 1000, "_09_Corner Blank_" & 2, 1, Crnblk, BlkNosY, TopXHoles, SideYHoles, 2)
                            PanelAHUModel.Blanks(EndBlkWth / 1000, BlankHt(1) / 1000, "_09_Corner Blank_" & 3, 1, Crnblk, BlkNosY, TopXHoles, SideYHoles, 3)
                            If EndBlkWth > 220 Then
                                PanelAHUModel.HorCChannel3(((EndBlkWth - 53) / 1000), FanNoY, BlkNosY, FanNoX, BlkNosX)
                            End If
                        End If
                    End If

                End If
            Else
                    DoorBlkHt = WallHt / 1000
                DoorCase = "Door 4"
                PanelAHUModel.DoorParts(DoorBlkWth, DoorBlkHt, BlkNosY, FanNoY, DoorCase, panelHt, BlankHt)

            End If

            PanelAHUModel.DoorSubAssembly()

            GoTo LandCSections
        End If

        If TopClear <= 0 Then
            BlkHt = WallHt
        End If

        If BlkNosY > 1 Then
            BlkHt = panelHt + 0.5
        Else

        End If

        If TopClear <= 0 And SideClear <= 0 Then
            GoTo LandCSections
        End If

        '------------Corner Blanks-------------
        Dim a As Integer
        If BlkNosX > 0 And BlkNosY > 0 Then
            For i = 1 To BlkNosX
                a = i - 1
                If a = BlkNosX - 1 Then
                    If a = 0 Then
                        a = 0
                    Else
                        a = BlkNosX
                    End If
                End If

                For j = 1 To BlkNosY
                    If EndBlkHt <= MinBlankLth Then
                        EndBlkHt = TopClear
                    End If

                    If PushSide = False Then
                        BlkQty = BlkNosX * 2
                    Else
                        BlkQty = BlkNosX
                    End If

                    TopXHoles = False
                    SideYHoles = False

                    PanelAHUModel.Blanks(BlankWth(i - 1) / 1000, BlankHt(j - 1) / 1000, "_09_Corner Blank_" & a + j, BlkQty, Crnblk, BlkNosY, TopXHoles, SideYHoles, i)
                    CrnBlkHt = BlankHt(0)
                    CrnBlkWth = BlankWth(0)
                Next
            Next
        End If

        'GoTo abc         '----------------------------------------------------------------------------
LandCSections:
        'L & C Sections-------------------------------------------------------

        PanelAHUModel.Side_L(SideLWth / 1000, FanNoY, panelHt, BlkNosY, TopClear, CrnBlkHt, BlankHt)
        PanelAHUModel.Top_Bot_L(TopLWth / 1000, BlkNosX, BlkNosY, BlkWth, panelWth, FanNoX, CrnBlkWth, PushSide, SideClear, AHUDoor, DoorBlkWth, EndBlkWth / 1000)
        PanelAHUModel.VerCChannels(FanNoY, FanNoX, panelHt, BlkNosY, TopClear, SideClear, BlankHt)

        If TopClear > 0 Or FanNoY > 1 Then
            PanelAHUModel.HorCChannel(panelWth / 100, FanNoY, BlkNosY, FanNoX)
        End If

        If SideClear > 0 Or FanNoX > 1 Then

            If PushSide = True Then
                PanelAHUModel.HorCChannel2((SideClear - 53) / 1000, FanNoY, BlkNosY, FanNoX, PushSide, BlkNosX, AHUDoor)
                PanelAHUModel.HorCChannel3(((panelWth - 53) / 1000), FanNoY, BlkNosY, FanNoX, BlkNosX)
            Else
                If BlkNosX >= 2 Then
                    PanelAHUModel.HorCChannel2((SideClear - 53) / 1000, FanNoY, BlkNosY, FanNoX, PushSide, BlkNosX, AHUDoor)
                    If EndBlkWth > 220 Then
                        PanelAHUModel.HorCChannel3(((EndBlkWth - 53) / 1000), FanNoY, BlkNosY, FanNoX, BlkNosX)
                    End If
                Else
                        PanelAHUModel.HorCChannel2((SideClear - 53) / 1000, FanNoY, BlkNosY, FanNoX, PushSide, BlkNosX, AHUDoor)
                End If
            End If
        End If

        PanelAHUModel.BaseStand()
abc:

        ''Assembly---------------------------------------------------------
        PanelAHUModel.Assembly(FanNos, FanNoX, FanNoY, SideClear / 1000, TopClear / 1000, panelWth / 1000, panelHt / 1000, BlkNosX, BlkNosY,
                              PushSide, BlankHt(0), BlankHt(1), BlankHt(2), ClientName, JobNo, AHUDoor, fanArticleNo, FanDia, EndBlkWth / 1000,
                              FanDia, fanArticleNo, DoorBlkWth)

#End Region

        swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swInputDimValOnCreate, True)  'Input Dimension ON
        swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchInference, True)    'Snapping ON
    End Sub


End Class
