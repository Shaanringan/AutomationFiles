Imports System.Collections.Generic
Imports System.IO
Imports System.Math

Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Excel

Public Class PanelAHUModels

    Dim swApp As New SldWorks

    Dim Part As ModelDoc2
    Dim Draw As DrawingDoc
    Dim Assy As AssemblyDoc
    Dim boolstatus As Boolean
    Dim longstatus As Long, longwarnings As Long

    Dim myFeature As Object
    Dim BlockUtil As MathUtility
    Dim Blockpoint As MathPoint
    Dim InsPoint As MathPoint
    Dim myBlockDefinition As SketchBlockDefinition

    Dim myMate As Object
    Dim InsComp As Component2
    Dim FittedFans As Integer
    Dim RemFans As Integer

    Dim swSheet As Sheet
    Dim myView As View
    Dim DrawView As View
    Dim BaseView As View
    Dim excludedComponents As Object
    Dim skSegment As Object
    Dim myDimension As Dimension
    Dim myDisplayDim As Object
    Dim myNote As Note
    Dim isLeft As Boolean

    ReadOnly LibPath As String = "C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library"
    Public SaveFolder As String

    ReadOnly DrawTemp As String = "C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Drawing\AADTech Drawing Template.DRWDOT"
    ReadOnly DrawSheet As String = "C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Drawing\AADTech Sheet Format - A4 - Landscape.slddrt"
    ReadOnly BOMTemp As String = "C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Drawing\BOM Template.sldbomtbt"

    Dim StdFunc As New Standard_Functions
    Dim predictivedb As New PredictiveDBInput
    Dim bomData As New BOMExcel
    Dim swBOMTable As BomTableAnnotation
    Dim swTable As TableAnnotation
    Dim autoballoonParams As AutoBalloonOptions
    Dim vBaloon As Object

    Public Client As String
    Public AHUName As String
    Public JobNo As String
    Public ArticleNoFan As String

    Public WallWth As Decimal
    Public WallHt As Decimal

    Dim SideLDist As Decimal
    Dim MaxSecLth3mm As Integer = 3000
    Dim LastLgt As Decimal
    Dim xhole As Integer
    Dim xdist As Decimal
    Dim ydist As Decimal
    Dim vAnnotations As Object

    Dim swLocalSketchPatternFeat As LocalSketchPatternFeatureData

    Dim skPoint As Object
    Dim swFeat As Feature
    Dim swFeatMgr As IFeatureManager
    Dim swFeatData As Object

    Dim DoorPnlWth As Decimal
    Dim DoorPnlHt As Decimal
    Dim DoorHt As Decimal
    Dim DoorWth As Decimal


    Dim MotorShtXHole As Integer                'Motor Sheet Holes
    Dim MotorShtXHoleDist As Decimal
    Dim MotorShtYHole As Integer
    Dim MotorShtYHoleDist As Decimal

    Dim HoleX As Integer                'Blanks holes
    Dim HoleXDist As Decimal
    Dim HoleY As Integer
    Dim HoleYDist As Decimal

    Dim HoleSpc As Decimal
    Dim HoleYBlk(3) As Integer
    Dim HoleYBlkDist(3) As Decimal

    Dim PartNameList As New List(Of String)
    Dim PartList As New List(Of String)
    Dim QtyList As New List(Of Integer)
    Dim ResultIndex As Integer
    Dim PartName As String
    Dim QuantityDictionary As New Dictionary(Of String, Integer)
    Dim AssyComponents As New List(Of String)

    Dim BTLth As Decimal
    Dim BTL1 As Decimal
    Dim BTL2 As Decimal

    Dim TopBlkYHoles(3) As Integer
    Dim TopBlkYDist(3) As Decimal
    Dim SideBlkXHoles(4) As Integer
    Dim SideBlkXDist(4) As Decimal
    Dim DoorXHoles As Integer
    Dim DoorXHoleDist As Decimal

    Dim NameList As New List(Of String)         'name of the part after splitting
    Dim PartNoList As New List(Of String)       ' job number after splitting
    Dim swComp As Component2
    Dim swSelMgr As SelectionMgr
    Dim AllComp As Object
    Dim selCount As Integer = 0
    Dim swEntity As Entity
    Dim JobNumber As String

    Dim XValueList As New List(Of Decimal)
    Dim YValueList As New List(Of Decimal)
    Dim ZValueList As New List(Of Decimal)


#Region "Motor"

    Public Sub MotorSelection(MotorDia As Integer, ArtNo As String)

        Part = swApp.OpenDoc6(LibPath & "\Panel - Motor\" & MotorDia & " mm Motor\" & MotorDia & "mm.SLDASM", 2, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2(MotorDia & "mm", False, longstatus)
        Assy = Part
        Part = swApp.ActiveDoc
        Assy.ViewZoomtofit2()
        boolstatus = Assy.EditRebuild3()
        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_" & ArtNo & ".SLDASM", 0, 0)

        swApp.CloseAllDocuments(True)

    End Sub

    Public Sub MotorSheet(PnlWth As Decimal, PnlHt As Decimal, FanDia As String, FanNoX As Integer, FanNoY As Integer, FanNos As Integer)

        If FanDia = "560" Or FanDia = "500" Or FanDia = "450" Or FanDia = "400" Then
            'Open Part
            Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\_01A_Motor Blank_01.SLDPRT", 1, 0, "", longstatus, longwarnings)
            swApp.ActivateDoc2("_01A_Motor Blank_01.SLDPRT", False, longstatus)
            Part = swApp.ActiveDoc

            'Re-Dim
            boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_01A_Motor Blank_01.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("Width@BaseFlange")
            myDimension.SystemValue = PnlWth
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Height@BaseFlange@_01A_Motor Blank_01.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("Height@BaseFlange")
            myDimension.SystemValue = PnlHt
            Part.ClearSelection2(True)

            boolstatus = Part.EditRebuild3()
            Part.ViewZoomtofit2()

        Else
            'Open Part
            Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\_01A_Motor Blank_01_2mm.SLDPRT", 1, 0, "", longstatus, longwarnings)
            swApp.ActivateDoc2("_01A_Motor Blank_01_2mm.SLDPRT", False, longstatus)
            Part = swApp.ActiveDoc

            'Re-Dim
            boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_01A_Motor Blank_01_2mm.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("Width@BaseFlange")
            myDimension.SystemValue = PnlWth
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Height@BaseFlange@_01A_Motor Blank_01_2mm.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("Height@BaseFlange")
            myDimension.SystemValue = PnlHt
            Part.ClearSelection2(True)

            boolstatus = Part.EditRebuild3()
            Part.ViewZoomtofit2()
        End If

        ' 9.2mm Holes --------------------------------------------------------------
        If FanDia = "560" Or FanDia = "500" Or FanDia = "450" Then
            MotorShtXHole = InterBoltingNumber(PnlWth * 1000, 25)
            MotorShtXHoleDist = BoltDistance(MotorShtXHole, PnlWth * 1000, 25)
            MotorShtYHole = InterBoltingNumber(PnlHt * 1000, 25)
            MotorShtYHoleDist = BoltDistance(MotorShtYHole, PnlHt * 1000, 25)
        Else
            MotorShtXHole = SmolInterBoltingNumber(PnlWth * 1000, 25)
            MotorShtXHoleDist = SmolBoltDistance(MotorShtXHole, PnlWth * 1000, 25)
            MotorShtYHole = SmolInterBoltingNumber(PnlHt * 1000, 25)
            MotorShtYHoleDist = SmolBoltDistance(MotorShtYHole, PnlHt * 1000, 25)
        End If

        boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("BotHole", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
        myFeature = Part.FeatureManager.FeatureLinearPattern2(MotorShtXHole, MotorShtXHoleDist, 0, 0, False, True, "NULL", "NULL", False)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("BotHole", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
        myFeature = Part.FeatureManager.FeatureLinearPattern2(0, 0, MotorShtYHole, MotorShtYHoleDist, True, True, "NULL", "NULL", False)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("TopHole", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
        myFeature = Part.FeatureManager.FeatureLinearPattern2(MotorShtXHole, MotorShtXHoleDist, 0, 0, True, True, "NULL", "NULL", False)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("TopHole", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
        myFeature = Part.FeatureManager.FeatureLinearPattern2(0, 0, MotorShtYHole, MotorShtYHoleDist, True, False, "NULL", "NULL", False)
        Part.ClearSelection2(True)

        'Motor Cutout Block
        Dim Ardata(2) As Double
        Ardata(0) = 0
        Ardata(1) = 0
        Ardata(2) = 0

        BlockUtil = swApp.GetMathUtility
        Blockpoint = BlockUtil.CreatePoint(Ardata)
        boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)

        Part.SketchManager.InsertSketch(True)
        myBlockDefinition = Part.SketchManager.MakeSketchBlockFromFile(Blockpoint, LibPath & "\Motor Cutout\" & FanDia & "_motor blank cutout.SLDBLK", False, 1, 0)

        myFeature = Part.FeatureManager.FeatureCut4(True, False, True, 0, 0, 0.01, 0.01, False, False, False, False, 2, 2, False, False, False, False, True, True, True, True, True, False, 0, 0, False, True)
        Part.SelectionManager.EnableContourSelection = False
        Part.ClearSelection2(True)
        boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)
        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_01A_Motor Blank_01.SLDPRT", 0, 0)
        MotorPlateDrawings(AHUName & "_01A_Motor Blank_01", FanNos, FanDia, PnlWth, PnlHt)
        swApp.CloseAllDocuments(True)


        FittedFans = FanNoX * FanNoY
        RemFans = (FanNoX * FanNoY) - FanNos

        If FittedFans > FanNos Then
            Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_05_Outer Blank.SLDPRT", 1, 0, "", longstatus, longwarnings)
            swApp.ActivateDoc2("_05_Outer Blank.SLDPRT", False, longstatus)
            Part = swApp.ActiveDoc

            'Re-Dim
            boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_05_Outer Blank.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("Width@BaseFlange")
            myDimension.SystemValue = PnlWth
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Height@BaseFlange@_05_Outer Blank.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("Height@BaseFlange")
            myDimension.SystemValue = PnlHt
            Part.ClearSelection2(True)
            boolstatus = Part.EditRebuild3()

            ' Zoom To Fit
            Part.ViewZoomtofit2()

            boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)


            ' 9.2mm Holes --------------------------------------------------------------
            MotorShtXHole = InterBoltingNumber(PnlWth * 1000, 25)
            MotorShtXHoleDist = BoltDistance(MotorShtXHole, PnlWth * 1000, 25)
            MotorShtYHole = InterBoltingNumber(PnlHt * 1000, 25)
            MotorShtYHoleDist = BoltDistance(MotorShtYHole, PnlHt * 1000, 25)

            boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("BotHole", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2(MotorShtXHole, MotorShtXHoleDist, 0, 0, False, True, "NULL", "NULL", False)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("TopHole", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2(MotorShtXHole, MotorShtXHoleDist, 0, 0, True, True, "NULL", "NULL", False)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("BotHole", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2(0, 0, MotorShtYHole, MotorShtYHoleDist, True, False, "NULL", "NULL", False)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("TopHole", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2(0, 0, MotorShtYHole, MotorShtYHoleDist, True, True, "NULL", "NULL", False)
            Part.ClearSelection2(True)


            ' Save
            longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_06_BlankOff.SLDPRT", 0, 0)
            BlankDrawings(AHUName & "_06_BlankOff", RemFans)

            ' Close Document
            swApp.CloseAllDocuments(True)
        End If

    End Sub

    Public Sub MotorSubAssembly(ArtNo As String) 'fan subassembly

        'Open Fan & Motor Blank
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_" & ArtNo & ".SLDASM", 2, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_01A_Motor Blank_01.SLDPRT", 1, 0, "", longstatus, longwarnings)

        ' New Assy File
        Dim assyTemp As String = swApp.GetUserPreferenceStringValue(9)
        Part = swApp.NewDocument(assyTemp, 0, 0, 0)
        Assy = Part

        ' Insert Component
        InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_" & ArtNo & ".SLDASM", 0, "", False, "", 1, 1, 1)
        InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_01A_Motor Blank_01.SLDPRT", 0, "", False, "", 1.5, 0.6, 0.7)

        ' Zoom To Fit
        Assy.ViewZoomtofit2()

        ' Save As
        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_Motor Sub Assembly.SLDASM", 0, 0)
        swApp.CloseAllDocuments(True)

        ' Open Assy
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_Motor Sub Assembly.SLDASM", 2, 32, "", longstatus, longwarnings)
        Assy = Part
        Assy.ViewZoomtofit2()

        'Mates
        'Origin Mate #1
        Part = swApp.ActiveDoc
        boolstatus = Part.Extension.SelectByID2(AHUName & "_" & ArtNo & "-1@" & AHUName & "_Motor Sub Assembly", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        Part.UnFixComponent
        Part.ClearSelection2(True)
        boolstatus = Part.Extension.SelectByID2("Point1@Origin@" & AHUName & "_01A_Motor Blank_01-1@" & AHUName & "_Motor Sub Assembly", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Point1@Origin", "EXTSKETCHPOINT", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        'Fan and Assembly Mates
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_" & ArtNo & "-1@" & AHUName & "_Motor Sub Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_" & ArtNo & "-1@" & AHUName & "_Motor Sub Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_" & ArtNo & "-1@" & AHUName & "_Motor Sub Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        Part.EditRebuild3()

        'Motor Blank and Motor Mates
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_01A_Motor Blank_01-1@" & AHUName & "_Motor Sub Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_01A_Motor Blank_01-1@" & AHUName & "_Motor Sub Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_01A_Motor Blank_01-1@" & AHUName & "_Motor Sub Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        Part.EditRebuild3()
        Part.ViewZoomtofit2()

        ' Save As
        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_Motor Sub Assembly.SLDASM", 0, 0)
        swApp.CloseAllDocuments(True)
    End Sub

#End Region

#Region "Blanks"
    Public Sub Blanks(BlkWth As Decimal, BlkHt As Decimal, BlankName As String, BlkQty As Integer, Crnblk As Boolean, BlkNoY As Integer, TopXHoles As Boolean, SideYHoles As Boolean, i As Integer)

        'BLANKS      
        Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_05_Outer Blank.SLDPRT", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("_05_Outer Blank", False, longstatus)
        Part = swApp.ActiveDoc

        'Re-Dim
        boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_05_Outer Blank", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        myDimension = Part.Parameter("Width@BaseFlange")
        myDimension.SystemValue = BlkWth
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Height@BaseFlange@_05_Outer Blank", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        myDimension = Part.Parameter("Height@BaseFlange")
        myDimension.SystemValue = BlkHt
        Part.ClearSelection2(True)
        boolstatus = Part.EditRebuild3()

        ' Zoom To Fit
        Part.ViewZoomtofit2()

        boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

        ' 9.2mm Holes --------------------------------------------------------------
        If TopXHoles = True Then
            HoleX = SmolInterBoltingNumber(BlkWth * 1000, 25)
            HoleXDist = SmolBoltDistance(HoleX, BlkWth * 1000, 25)
        Else
            HoleX = InterBoltingNumber(BlkWth * 1000, 25)
            HoleXDist = BoltDistance(HoleX, BlkWth * 1000, 25)
        End If

        If SideYHoles = True Then
            HoleY = SmolInterBoltingNumber(BlkHt * 1000, 25)
            HoleYDist = SmolBoltDistance(HoleY, BlkHt * 1000, 25)
        Else
            HoleY = InterBoltingNumber(BlkHt * 1000, 25)
            HoleYDist = BoltDistance(HoleY, BlkHt * 1000, 25)
        End If

        boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("BotHole", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
        If HoleX = 1 Then
            myFeature = Part.FeatureManager.FeatureLinearPattern2(HoleX + 1, HoleXDist, 0, 0, False, True, "NULL", "NULL", False)
        Else
            If HoleX = 2 Then
                myFeature = Part.FeatureManager.FeatureLinearPattern2(HoleX + 1, HoleXDist / 2, 0, 0, False, True, "NULL", "NULL", False)
            Else
                myFeature = Part.FeatureManager.FeatureLinearPattern2(HoleX, HoleXDist, 0, 0, False, True, "NULL", "NULL", False)
            End If
        End If
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("TopHole", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
        If HoleX = 1 Then
            myFeature = Part.FeatureManager.FeatureLinearPattern2(HoleX + 1, HoleXDist, 0, 0, True, True, "NULL", "NULL", False)
        Else
            If HoleX = 2 Then
                myFeature = Part.FeatureManager.FeatureLinearPattern2(HoleX + 1, HoleXDist / 2, 0, 0, True, True, "NULL", "NULL", False)
            Else
                myFeature = Part.FeatureManager.FeatureLinearPattern2(HoleX, HoleXDist, 0, 0, True, True, "NULL", "NULL", False)
            End If
        End If
        Part.ClearSelection2(True)


        If HoleX <= 2 And HoleY <= 2 Then
            GoTo saveblank
        End If

        If HoleX <= 2 Then
            If HoleYDist > 0 Then
                boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("TopHole", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
                myFeature = Part.FeatureManager.FeatureLinearPattern2(0, 0, HoleY + 1, HoleYDist, True, True, "NULL", "NULL", False)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("BotHole", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
                myFeature = Part.FeatureManager.FeatureLinearPattern2(0, 0, HoleY + 1, HoleYDist, True, False, "NULL", "NULL", False)
                Part.ClearSelection2(True)

                'Yvalues(BlkNoY) = HoleY + 1
            End If
        Else

            If HoleY <= 2 Then
                'Yvalues(BlkNoY) = HoleY
                GoTo saveblank
            Else
                boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("TopHole", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
                myFeature = Part.FeatureManager.FeatureLinearPattern2(0, 0, HoleY, HoleYDist, True, True, "NULL", "NULL", False)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("BotHole", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
                myFeature = Part.FeatureManager.FeatureLinearPattern2(0, 0, HoleY, HoleYDist, True, False, "NULL", "NULL", False)

                'Yvalues(BlkNoY) = HoleY

                'boolstatus = Part.Extension.SelectByID2("LPattern2", "BODYFEATURE", 0, 0, 0, False, 4, Nothing, 0)
                'Part.EditSuppress2()
                Part.ClearSelection2(True)
            End If
        End If

saveblank:
        ' Save
        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & BlankName & ".SLDPRT", 0, 0)
        BlankDrawings(AHUName & BlankName, BlkQty)

        If BlankName = "_08_Top Blank_" & i Then
            TopBlkYHoles(i) = HoleY
            TopBlkYDist(i) = HoleYDist

        ElseIf BlankName = "_07_Side Blank_A" & i Then
            SideBlkXHoles(i) = HoleX
            SideBlkXDist(i) = HoleXDist

        ElseIf BlankName = "_09_Corner Blank_" & i Then
            SideBlkXHoles(i) = HoleX
            SideBlkXDist(i) = HoleXDist
        End If

        ' Close Document
        swApp.CloseAllDocuments(True)

        If AssyComponents.Contains(AHUName & "_07_Side Blank_A1.SLDPRT") Then
            SideBlankUpdates(Crnblk)
        End If

        If AssyComponents.Contains(AHUName & "_07_Side Blank_A2.SLDPRT") Then
            SideBlankUpdates(Crnblk)
        End If

        If AssyComponents.Contains(AHUName & "_08_Top Blank_1.SLDPRT") Then
            TopBlankUpdates()
        End If

        If AssyComponents.Contains(AHUName & "_08_Top Blank_2.SLDPRT") Then
            TopBlankUpdates()
        End If

    End Sub

    Public Sub DoorParts(BlkWth As Decimal, BlkHt As Decimal, BlkNoY As Integer, FansY As Integer, DoorCase As String, PnlHt As Decimal, TopBlankHt() As Decimal)

        DoorPnlWth = BlkWth
        DoorPnlHt = BlkHt

        Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_11_Door Blank.SLDPRT", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("_11_Door Blank", False, longstatus)
        Part = swApp.ActiveDoc

        boolstatus = Part.Extension.SelectByID2("Height@DoorBase@_11_Door Blank.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        myDimension = Part.Parameter("Height@DoorBase")
        myDimension.SystemValue = DoorPnlHt
        Part.ClearSelection2(True)


        boolstatus = Part.Extension.SelectByID2("Width@DoorBase@_11_Door Blank.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        myDimension = Part.Parameter("Width@DoorBase")
        myDimension.SystemValue = DoorPnlWth
        Part.ClearSelection2(True)

        boolstatus = Part.EditRebuild3()


        ' 9.2mm Holes --------------------------------------------------------------
#Region "X Direction Holes"
        HoleX = InterBoltingNumber(DoorPnlWth * 1000, 25)
        HoleXDist = BoltDistance(HoleX, DoorPnlWth * 1000, 25)

        DoorXHoles = HoleX
        DoorXHoleDist = HoleXDist

        HoleY = InterBoltingNumber(DoorPnlHt * 1000, 25)
        HoleYDist = BoltDistance(HoleY, DoorPnlHt * 1000, 25)

        boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("BotHole", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
        myFeature = Part.FeatureManager.FeatureLinearPattern2(HoleX, HoleXDist, 0, 0, False, True, "NULL", "NULL", False)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("TopHole", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
        myFeature = Part.FeatureManager.FeatureLinearPattern2(HoleX, HoleXDist, 0, 0, True, True, "NULL", "NULL", False)
        Part.ClearSelection2(True)
#End Region

#Region "Y Direction Holes"

        For a = 0 To BlkNoY - 1
            HoleYBlk(a) = SlotsInterBoltingNumber(TopBlankHt(a), 25)
            HoleYBlkDist(a) = SlotsBoltDistance(HoleYBlk(a), TopBlankHt(a), 25)
        Next


        swApp.ActivateDoc2("_11_Door Blank.SLDPRT", False, longstatus)
        Part = swApp.ActiveDoc

        boolstatus = Part.Extension.SelectByID2("YHole", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditSketch()

        Dim NoOfFans As Integer = Math.Floor(FansY / 2)
        Dim XCord As Decimal = (DoorPnlWth / 2) - 0.025
        Dim StartZero As Decimal = (DoorPnlHt / 2) - 0.025
        Select Case DoorCase
            Case = "Door1"                       'Fans more than 1 // No Blanks
                For b = 0 To NoOfFans
                    For a = 2 To MotorShtYHole + FansY - 1
                        If a = 3 Then
                            skPoint = Part.SketchManager.CreatePoint(XCord, (MotorShtYHoleDist * a) - StartZero, 0)
                            skPoint = Part.SketchManager.CreatePoint(-XCord, StartZero - (MotorShtYHoleDist * a) - StartZero, 0)

                            skPoint = Part.SketchManager.CreatePoint(XCord, (MotorShtYHoleDist * a) + 0.05 - StartZero, 0)
                            skPoint = Part.SketchManager.CreatePoint(-XCord, (MotorShtYHoleDist * a) + 0.05 - StartZero, 0)
                        Else
                            skPoint = Part.SketchManager.CreatePoint(XCord, (MotorShtYHoleDist * a) + (b * 0.05) - StartZero, 0)
                            skPoint = Part.SketchManager.CreatePoint(-XCord, (MotorShtYHoleDist * a) + (b * 0.05) - StartZero, 0)
                        End If
                    Next
                Next

            Case = "Door 2"                     'Fans more than 1 //  Blanks >= 1
                'Fans Calculation
                Dim NewFan As Integer = 0

                For a = 1 To MotorShtYHole + FansY - 1
                    If a >= 4 Then
                        NewFan = 1
                    End If
                    If a = 3 Then
                        skPoint = Part.SketchManager.CreatePoint(XCord, (MotorShtYHoleDist * a) - StartZero, 0)
                        skPoint = Part.SketchManager.CreatePoint(-XCord, (MotorShtYHoleDist * a) - StartZero, 0)

                        skPoint = Part.SketchManager.CreatePoint(XCord, (MotorShtYHoleDist * a) + 0.05 - StartZero, 0)
                        skPoint = Part.SketchManager.CreatePoint(-XCord, (MotorShtYHoleDist * a) + 0.05 - StartZero, 0)
                    Else
                        skPoint = Part.SketchManager.CreatePoint(XCord, (MotorShtYHoleDist * a) + (NewFan * 0.05) - StartZero, 0)
                        skPoint = Part.SketchManager.CreatePoint(-XCord, (MotorShtYHoleDist * a) + (NewFan * 0.05) - StartZero, 0)
                    End If
                Next


            Case = "Door 3"                     'Fans = 1 //  Blanks >= 1
                'Fans Calculation
                For a = 1 To MotorShtYHole - 1
                    If a = 3 Then
                        skPoint = Part.SketchManager.CreatePoint(XCord, (MotorShtYHoleDist * a) - StartZero, 0)
                        skPoint = Part.SketchManager.CreatePoint(-XCord, (MotorShtYHoleDist * a) - StartZero, 0)

                        skPoint = Part.SketchManager.CreatePoint(XCord, (MotorShtYHoleDist * a) + 0.05 - StartZero, 0)
                        skPoint = Part.SketchManager.CreatePoint(-XCord, (MotorShtYHoleDist * a) + 0.05 - StartZero, 0)
                    Else
                        skPoint = Part.SketchManager.CreatePoint(XCord, (MotorShtYHoleDist * a) - StartZero, 0)
                        skPoint = Part.SketchManager.CreatePoint(-XCord, (MotorShtYHoleDist * a) - StartZero, 0)
                    End If
                Next

                'Blank Calculation
                For a = 1 To HoleYBlk(0) - 1
                    skPoint = Part.SketchManager.CreatePoint(XCord, ((PnlHt / 1000) * FansY) + (HoleYBlkDist(0) * a) - StartZero, 0)
                    skPoint = Part.SketchManager.CreatePoint(-XCord, ((PnlHt / 1000) * FansY) + (HoleYBlkDist(0) * a) - StartZero, 0)
                Next

            Case = "Door 4"                     'Case else
                For b = 0 To NoOfFans
                    For a = 1 To MotorShtYHole - 1
                        If a = 3 Then
                            skPoint = Part.SketchManager.CreatePoint(XCord, (MotorShtYHoleDist * a) - StartZero, 0)
                            skPoint = Part.SketchManager.CreatePoint(-XCord, (MotorShtYHoleDist * a) - StartZero, 0)

                            skPoint = Part.SketchManager.CreatePoint(XCord, (MotorShtYHoleDist * a) + 0.05 - StartZero, 0)
                            skPoint = Part.SketchManager.CreatePoint(-XCord, (MotorShtYHoleDist * a) + 0.05 - StartZero, 0)
                        Else
                            skPoint = Part.SketchManager.CreatePoint(XCord, (MotorShtYHoleDist * a) + (b * 0.05) - StartZero, 0)
                            skPoint = Part.SketchManager.CreatePoint(-XCord, (MotorShtYHoleDist * a) + (b * 0.05) - StartZero, 0)
                        End If
                    Next
                Next
        End Select

        Part.ClearSelection2(True)
        boolstatus = Part.EditRebuild3()

#End Region

SaveDoorBlank:

        ' Zoom To Fit
        Part.ViewZoomtofit2()

        ' Save
        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_11_Door Blank.SLDPRT", 0, 0)
        DoorDrawings(AHUName & "_11_Door Blank", DoorPnlWth, DoorPnlHt)
        swApp.CloseDoc(True)

        Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_10_Door.SLDPRT", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("_10_Door", False, longstatus)
        Part = swApp.ActiveDoc

        boolstatus = Part.Extension.SelectByID2("Height@Door@_10_Door.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        myDimension = Part.Parameter("Height@Door")
        myDimension.SystemValue = (DoorPnlHt - 0.15 - 0.1)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Width@Door@_10_Door.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        myDimension = Part.Parameter("Width@Door")
        myDimension.SystemValue = (DoorPnlWth - 0.14)
        Part.ClearSelection2(True)

        boolstatus = Part.EditRebuild3()

        ' Zoom To Fit
        Part.ViewZoomtofit2()

        ' Save
        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_10_Door.SLDPRT", 0, 0)
        DoorDrawings(AHUName & "_10_Door", DoorPnlWth - 0.14, DoorPnlHt - 0.15 - 0.1)
        swApp.CloseDoc(True)

    End Sub

    Public Sub DoorSubAssembly()

        'Open / Insert Parts
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_11_Door Blank.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_10_Door.SLDPRT", 1, 32, "", longstatus, longwarnings)

        'New Assy File -----------------------------------------------------------------------------------------
        Dim assyTemp As String = swApp.GetUserPreferenceStringValue(9)
        Part = swApp.NewDocument(assyTemp, 0, 0, 0)
        Assy = Part

        'Insert Parts--------------------------------------------------------------------------------------------------------------
        InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_11_Door Blank.SLDPRT", 0, "", False, "", 1, 1, 1)
        InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_10_Door.SLDPRT", 0, "", False, "", 1.3, 1.3, 1.3)

        Assy = Part
        Assy.ViewZoomtofit2()

        ' Save As----------------------------------------------------------------------------------
        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_Door SubAssembly.SLDASM", 0, 0)
        swApp.CloseAllDocuments(True)

        ' Open Assy--------------------------------------------------------------------------------
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_Door SubAssembly.SLDASM", 2, 32, "", longstatus, longwarnings)
        Assy = Part
        Assy.ViewZoomtofit2()

        'Mates
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_10_Door-1@" & AHUName & "_Door SubAssembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_11_Door Blank-1@" & AHUName & "_Door SubAssembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '-------- Right Plane Coincident
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_11_Door Blank-1@" & AHUName & "_Door SubAssembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_10_Door-1@" & AHUName & "_Door SubAssembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.002, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Front Plane coincidence 
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("DoorBlkBottom@" & AHUName & "_11_Door Blank-1@" & AHUName & "_Door SubAssembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("DoorBottom@" & AHUName & "_10_Door-1@" & AHUName & "_Door SubAssembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, True, 0.1, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '----------- Bottom Plane Distance
        Part.ClearSelection2(True)

        Part.EditRebuild3()
        Assy.ViewZoomtofit2()


        boolstatus = Part.Extension.SelectByID2(AHUName & "_11_Door Blank-1@" & AHUName & "_Door SubAssembly", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        Part.UnFixComponent()
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_11_Door Blank-1@" & AHUName & "_Door SubAssembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_11_Door Blank-1@" & AHUName & "_Door SubAssembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_11_Door Blank-1@" & AHUName & "_Door SubAssembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        Part.EditRebuild3()
        Assy.ViewZoomtofit2()

        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_Door SubAssembly.SLDASM", 0, 0)
        swApp.CloseAllDocuments(True)

    End Sub

    Public Sub SideBlankUpdates(Crnblk)

        'Side Blanks
        If AssyComponents.Contains(AHUName & "_07_Side Blank_B1.SLDPRT") Then
            Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_A1.SLDPRT", 1, 32, "", longstatus, longwarnings)
            swApp.ActivateDoc2(AHUName & "_07_Side Blank_A1", False, longstatus)
            Part = swApp.ActiveDoc
            'boolstatus = Part.Extension.SelectByID2("LPattern3", "BODYFEATURE", 0, 0, 0, False, 4, Nothing, 0)
            'Part.EditSuppress2()

            Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_A1.SLDDRW", 3, 0, "", longstatus, longwarnings)
            boolstatus = Part.EditRebuild3()
            longstatus = Part.SaveAs3(SaveFolder & "\PDF Drawings\" & AHUName & "_07_Side Blank_A1.PDF", 0, 2)
            swApp.CloseAllDocuments(True)


            Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_A2.SLDPRT", 1, 32, "", longstatus, longwarnings)
            swApp.ActivateDoc2(AHUName & "_07_Side Blank_A2", False, longstatus)
            Part = swApp.ActiveDoc
            'boolstatus = Part.Extension.SelectByID2("LPattern3", "BODYFEATURE", 0, 0, 0, False, 4, Nothing, 0)
            'Part.EditSuppress2()

            Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_A2.SLDDRW", 3, 0, "", longstatus, longwarnings)
            boolstatus = Part.EditRebuild3()
            longstatus = Part.SaveAs3(SaveFolder & "\PDF Drawings\" & AHUName & "_07_Side Blank_A2.PDF", 0, 2)
            swApp.CloseAllDocuments(True)

        End If

        If Crnblk = True Then
            Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_A1.SLDPRT", 1, 32, "", longstatus, longwarnings)
            swApp.ActivateDoc2(AHUName & "_07_Side Blank_A1", False, longstatus)
            Part = swApp.ActiveDoc
            'boolstatus = Part.Extension.SelectByID2("LPattern2", "BODYFEATURE", 0, 0, 0, False, 4, Nothing, 0)
            'Part.EditSuppress2()

            Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_A1.SLDDRW", 3, 0, "", longstatus, longwarnings)
            boolstatus = Part.EditRebuild3()
            longstatus = Part.SaveAs3(SaveFolder & "\PDF Drawings\" & AHUName & "_07_Side Blank_A1.PDF", 0, 2)
            swApp.CloseAllDocuments(True)

            If AssyComponents.Contains(AHUName & "_07_Side Blank_A2.SLDPRT") Then
                Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_A2.SLDPRT", 1, 32, "", longstatus, longwarnings)
                swApp.ActivateDoc2(AHUName & "_07_Side Blank_A2", False, longstatus)
                Part = swApp.ActiveDoc
                'boolstatus = Part.Extension.SelectByID2("LPattern2", "BODYFEATURE", 0, 0, 0, False, 4, Nothing, 0)
                'Part.EditSuppress2()

                Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_A2.SLDDRW", 3, 0, "", longstatus, longwarnings)
                boolstatus = Part.EditRebuild3()
                longstatus = Part.SaveAs3(SaveFolder & "\PDF Drawings\" & AHUName & "_07_Side Blank_A2.PDF", 0, 2)
                swApp.CloseAllDocuments(True)
            End If

        End If

        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_B1.SLDPRT", 1, 32, "", longstatus, longwarnings)
        swApp.ActivateDoc2(AHUName & "_07_Side Blank_B1", False, longstatus)
        Part = swApp.ActiveDoc
        'boolstatus = Part.Extension.SelectByID2("LPattern4", "BODYFEATURE", 0, 0, 0, False, 4, Nothing, 0)
        'Part.EditSuppress2()
        Part.ClearSelection2(True)

        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_B1.SLDDRW", 3, 0, "", longstatus, longwarnings)
        boolstatus = Part.EditRebuild3()
        longstatus = Part.SaveAs3(SaveFolder & "\PDF Drawings\" & AHUName & "_07_Side Blank_B1.PDF", 0, 2)
        swApp.CloseAllDocuments(True)


        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_B2.SLDPRT", 1, 32, "", longstatus, longwarnings)
        swApp.ActivateDoc2(AHUName & "_07_Side Blank_B2", False, longstatus)
        Part = swApp.ActiveDoc
        'boolstatus = Part.Extension.SelectByID2("LPattern4", "BODYFEATURE", 0, 0, 0, False, 4, Nothing, 0)
        'Part.EditSuppress2()
        Part.ClearSelection2(True)

        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_B2.SLDDRW", 3, 0, "", longstatus, longwarnings)
        boolstatus = Part.EditRebuild3()
        longstatus = Part.SaveAs3(SaveFolder & "\PDF Drawings\" & AHUName & "_07_Side Blank_B2.PDF", 0, 2)
        swApp.CloseAllDocuments(True)

    End Sub

    Public Sub TopBlankUpdates()

        'Suppress Linear patterns
        If AssyComponents.Contains(AHUName & "_08_Top Blank_2.SLDPRT") Then
            Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_08_Top Blank_1.SLDPRT", 1, 0, "", longstatus, longwarnings)
            swApp.ActivateDoc2(AHUName & "_08_Top Blank_1", False, longstatus)
            Part = swApp.ActiveDoc
            boolstatus = Part.Extension.SelectByID2("LPattern4", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditSuppress2()

            Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "08_Top Blank_1.SLDDRW", 3, 0, "", longstatus, longwarnings)
            boolstatus = Part.EditRebuild3()
            longstatus = Part.SaveAs3(SaveFolder & "\PDF Drawings\" & AHUName & "08_Top Blank_1.PDF", 0, 2)
            swApp.CloseAllDocuments(True)

            '-------------------------------------------------------------------------------------------------------------------

            Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_08_Top Blank_2.SLDPRT", 1, 0, "", longstatus, longwarnings)
            swApp.ActivateDoc2(AHUName & "_08_Top Blank_2", False, longstatus)
            Part = swApp.ActiveDoc
            boolstatus = Part.Extension.SelectByID2("LPattern3", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditSuppress2()

            Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "08_Top Blank_2.SLDDRW", 3, 0, "", longstatus, longwarnings)
            boolstatus = Part.EditRebuild3()
            longstatus = Part.SaveAs3(SaveFolder & "\PDF Drawings\" & AHUName & "08_Top Blank_2.PDF", 0, 2)
            swApp.CloseAllDocuments(True)

        End If

    End Sub

#End Region

#Region "Frame"

#Region "Side L"

    Public Sub Side_L(LWidth As Decimal, FansY As Integer, PnlHgt As Decimal, BlkNoY As Integer, TopClear As Decimal, CrnBlkHt As Decimal, BlkHt() As Decimal)

        For a = 0 To BlkNoY - 1
            HoleYBlk(a) = SlotsInterBoltingNumber(BlkHt(a), 25)
            HoleYBlkDist(a) = SlotsBoltDistance(HoleYBlk(a), BlkHt(a), 25)
        Next

        'OPEN SIDE L -------------------------------------X------------------------------------------------
        If WallHt <= MaxSecLth3mm Then
#Region "Single Side L"
            '--------------------------------------------LEFT----------------------------------------------
            Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_03A_Side_L.SLDPRT", 1, 0, "", longstatus, longwarnings)
            swApp.ActivateDoc2("_03A_Side_L.SLDPRT", False, longstatus)
            Part = swApp.ActiveDoc
            boolstatus = Part.Extension.SelectByID2("Height@BaseFlange@_03A_Side_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("Height@BaseFlange")
            myDimension.SystemValue = (WallHt - 100) / 1000
            HoleSpc = (WallHt - 100) / 1000
            Part.ClearSelection2(True)
            boolstatus = Part.EditRebuild3()

            Part.ViewZoomtofit2()
            SideLHoles(HoleSpc)                  '----------4.2mm holes

#Region "Slots"
            Dim Minus25 As Decimal = 0.025
            ' --------------------------------------------Slots---------------------------------------------------------------
            swApp.ActivateDoc2("_03A_Side_L.SLDPRT", False, longstatus)
            Part = swApp.ActiveDoc

            boolstatus = Part.Extension.SelectByID2("slot1", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditSketch()

            boolstatus = Part.Extension.SelectByID2("Dist@slot1@_03A_Side_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("Dist@slot1")
            myDimension.SystemValue = MotorShtYHoleDist - 0.025

            Part.ClearSelection2(True)
            boolstatus = Part.EditRebuild3()

            swApp.ActivateDoc2("_03A_Side_L.SLDPRT", False, longstatus)
            Part = swApp.ActiveDoc
            boolstatus = Part.Extension.SelectByID2("Slots", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditSketch()

            Dim NoOfFans As Integer = Math.Floor(FansY / 2)
            Dim NextFan As Integer = 0
            '-------------------------- Blank Y  = 0 ------------------------------------
            If BlkNoY = 0 Then

                If NoOfFans > 0 Then
                    For a = 2 To MotorShtYHole + FansY - 1
                        If a = 3 Then
                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) - Minus25, 0)
                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) + 0.05 - Minus25, 0)
                        Else
                            If FansY = 1 And a = 4 Then
                                GoTo SaveSideL
                            Else
                                If a > 3 Then
                                    skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) + 0.05 - Minus25, 0)
                                Else
                                    skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) - Minus25, 0)
                                End If
                            End If
                        End If
                    Next
                Else                                    'Fans = 1
                    For a = 2 To MotorShtXHole - 2
                        skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) - Minus25, 0)
                    Next
                End If

            ElseIf BlkNoY = 1 Then

                '-------------------------- Blank Y = 1 --------------------------------------
                '------ Fans
                If NoOfFans > 0 Then                       'Fans more than 1

                    For a = 2 To MotorShtYHole + FansY
                        If a Mod 3 = 0 Then
                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) + (NextFan * 0.05) - Minus25, 0)

                            If a Mod 3 = 0 Then
                                NextFan = NextFan + 1
                            End If

                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) + (NextFan * 0.05) - Minus25, 0)
                        Else
                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) + (NextFan * 0.05) - Minus25, 0)
                        End If
                    Next

                Else                                    'Fans = 1
                    For a = 2 To MotorShtYHole - 1
                        If a = 3 Then
                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) - Minus25, 0)
                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) + 0.05 - Minus25, 0)
                        Else
                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) - Minus25, 0)
                        End If
                    Next
                End If
                '-x-x-x-x-x-x-x-x-x-x
                '------ Blanks
                For a = 1 To HoleYBlk(0) - 2
                    skPoint = Part.SketchManager.CreatePoint(0, ((PnlHgt / 1000) * FansY) + (HoleYBlkDist(0) * a) - Minus25, 0)
                Next

            Else

                '-------------------------- Blank Y > 1 --------------------------------------
                '------ Fans

                If NoOfFans > 0 Then                       'Fans more than 1

                    For a = 2 To MotorShtYHole + FansY
                        If a Mod 3 = 0 Then
                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) + (NextFan * 0.05) - Minus25, 0)

                            If a Mod 3 = 0 Then
                                NextFan = NextFan + 1
                            End If

                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) + (NextFan * 0.05) - Minus25, 0)
                        Else
                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) + (NextFan * 0.05) - Minus25, 0)
                        End If
                    Next

                Else                                    'Fans = 1
                    For a = 2 To MotorShtYHole - 1
                        If a = 3 Then
                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) - Minus25, 0)
                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) + 0.05 - Minus25, 0)
                        Else
                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) - Minus25, 0)
                        End If
                    Next
                End If

                '-x-x-x-x-x-x-x-x-x-x-
                '------ Blanks
                For a = 1 To HoleYBlk(0) - 1
                    If a = HoleYBlk(0) - 1 Then
                        skPoint = Part.SketchManager.CreatePoint(0, ((PnlHgt / 1000) * FansY) + (HoleYBlkDist(0) * a) - Minus25, 0)
                        skPoint = Part.SketchManager.CreatePoint(0, ((PnlHgt / 1000) * FansY) + (HoleYBlkDist(0) * a) + 0.05 - Minus25, 0)
                    Else
                        skPoint = Part.SketchManager.CreatePoint(0, ((PnlHgt / 1000) * FansY) + (HoleYBlkDist(0) * a) - Minus25, 0)
                    End If
                Next

                For a = 1 To HoleYBlk(1) - 1
                    skPoint = Part.SketchManager.CreatePoint(0, ((PnlHgt / 1000) * FansY) + BlkHt(0) / 1000 + (HoleYBlkDist(1) * a) - Minus25, 0)
                Next

                If BlkNoY > 2 Then

                    If HoleYBlk(2) = 1 Then
                        skPoint = Part.SketchManager.CreatePoint(0, ((PnlHgt / 1000) * FansY) + BlkHt(0) / 1000 + BlkHt(1) / 1000 - Minus25, 0)
                    End If

                    For a = 1 To HoleYBlk(2) - 1
                        If a = 1 Then
                            skPoint = Part.SketchManager.CreatePoint(0, ((PnlHgt / 1000) * FansY) + BlkHt(0) / 1000 + BlkHt(1) / 1000 - Minus25, 0)
                        Else
                            skPoint = Part.SketchManager.CreatePoint(0, ((PnlHgt / 1000) * FansY) + BlkHt(0) / 1000 + BlkHt(1) / 1000 + (HoleYBlkDist(2) * a) - Minus25, 0)
                        End If
                    Next
                End If

            End If
            '-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-

            Part.ClearSelection2(True)
            boolstatus = Part.EditRebuild3()
            Part.ViewZoomtofit2()
#End Region

SaveSideL:
            boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)
            longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_03A_Side_L_Left.SLDPRT", 0, 0)         ' Save As
            swApp.CloseAllDocuments(True)

            '--------------------------------------------RIGHT-----------------------------------------------------------

            Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_03A_Side_L_Right.SLDPRT", 1, 0, "", longstatus, longwarnings)
            swApp.ActivateDoc2("_03A_Side_L_Right.SLDPRT", False, longstatus)
            Part = swApp.ActiveDoc
            boolstatus = Part.Extension.SelectByID2("Height@BaseFlange@_03A_Side_L_Right.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("Height@BaseFlange")
            myDimension.SystemValue = (WallHt - 100) / 1000
            HoleSpc = (WallHt - 100) / 1000
            Part.ClearSelection2(True)
            boolstatus = Part.EditRebuild3()

            Part.ViewZoomtofit2()
            SideLHoles(HoleSpc)                  '----------4.2mm holes

#Region "Slots"

            ' --------------------------------------------Slots---------------------------------------------------------------
            swApp.ActivateDoc2("_03A_Side_L_Right.SLDPRT", False, longstatus)
            Part = swApp.ActiveDoc

            boolstatus = Part.Extension.SelectByID2("slot1", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditSketch()

            boolstatus = Part.Extension.SelectByID2("Dist@slot1@_03A_Side_L_Right.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("Dist@slot1")
            myDimension.SystemValue = MotorShtYHoleDist - 0.025

            Part.ClearSelection2(True)
            boolstatus = Part.EditRebuild3()

            swApp.ActivateDoc2("_03A_Side_L_Right.SLDPRT", False, longstatus)
            Part = swApp.ActiveDoc
            boolstatus = Part.Extension.SelectByID2("Slots", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditSketch()

            NextFan = 0
            '-------------------------- Blank Y  = 0 ------------------------------------
            If BlkNoY = 0 Then

                If NoOfFans > 0 Then
                    For a = 2 To MotorShtYHole + FansY - 1
                        If a = 3 Then
                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) - Minus25, 0)
                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) + 0.05 - Minus25, 0)
                        Else
                            If FansY = 1 And a = 4 Then
                                GoTo SaveSideL
                            Else
                                If a > 3 Then
                                    skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) + 0.05 - Minus25, 0)
                                Else
                                    skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) - Minus25, 0)
                                End If
                            End If
                        End If
                    Next
                Else                                    'Fans = 1
                    For a = 2 To MotorShtXHole - 2
                        skPoint = Part.SketchManager.CreatePoint(0, MotorShtYHoleDist * a - Minus25, 0)
                    Next
                End If


            ElseIf BlkNoY = 1 Then

                '-------------------------- Blank Y = 1 --------------------------------------
                '------ Fans
                If NoOfFans > 0 Then                       'Fans more than 1

                    For a = 2 To MotorShtYHole + FansY
                        If a Mod 3 = 0 Then
                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) + (NextFan * 0.05) - Minus25, 0)

                            If a Mod 3 = 0 Then
                                NextFan = NextFan + 1
                            End If

                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) + (NextFan * 0.05) - Minus25, 0)
                        Else
                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) + (NextFan * 0.05) - Minus25, 0)
                        End If
                    Next

                Else                                    'Fans = 1
                    For a = 2 To MotorShtYHole - 1
                        If a = 3 Then
                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) - Minus25, 0)
                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) + 0.05 - Minus25, 0)
                        Else
                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) - Minus25, 0)
                        End If
                    Next
                End If
                '-x-x-x-x-x-x-x-x-x-x
                '------ Blanks
                For a = 1 To HoleYBlk(0) - 2
                    skPoint = Part.SketchManager.CreatePoint(0, ((PnlHgt / 1000) * FansY) + (HoleYBlkDist(0) * a) - Minus25, 0)
                Next

            Else

                '-------------------------- Blank Y > 1 --------------------------------------
                '------ Fans

                If NoOfFans > 0 Then                       'Fans more than 1

                    For a = 2 To MotorShtYHole + FansY
                        If a Mod 3 = 0 Then
                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) + (NextFan * 0.05) - Minus25, 0)

                            If a Mod 3 = 0 Then
                                NextFan = NextFan + 1
                            End If

                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) + (NextFan * 0.05) - Minus25, 0)
                        Else
                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) + (NextFan * 0.05) - Minus25, 0)
                        End If
                    Next

                Else                                    'Fans = 1
                    For a = 2 To MotorShtYHole - 1
                        If a = 3 Then
                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) - Minus25, 0)
                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) + 0.05 - Minus25, 0)
                        Else
                            skPoint = Part.SketchManager.CreatePoint(0, (MotorShtYHoleDist * a) - Minus25, 0)
                        End If
                    Next
                End If

                '-x-x-x-x-x-x-x-x-x-x-
                '------ Blanks
                For a = 1 To HoleYBlk(0) - 1
                    If a = HoleYBlk(0) - 1 Then
                        skPoint = Part.SketchManager.CreatePoint(0, ((PnlHgt / 1000) * FansY) + (HoleYBlkDist(0) * a) - Minus25, 0)
                        skPoint = Part.SketchManager.CreatePoint(0, ((PnlHgt / 1000) * FansY) + (HoleYBlkDist(0) * a) + 0.05 - Minus25, 0)
                    Else
                        skPoint = Part.SketchManager.CreatePoint(0, ((PnlHgt / 1000) * FansY) + (HoleYBlkDist(0) * a) - Minus25, 0)
                    End If
                Next


                For a = 1 To HoleYBlk(1) - 1
                    skPoint = Part.SketchManager.CreatePoint(0, ((PnlHgt / 1000) * FansY) + BlkHt(0) / 1000 + (HoleYBlkDist(1) * a) - Minus25, 0)
                Next

                If BlkNoY > 2 Then

                    If HoleYBlk(2) = 1 Then
                        skPoint = Part.SketchManager.CreatePoint(0, ((PnlHgt / 1000) * FansY) + BlkHt(0) / 1000 + BlkHt(1) / 1000 - Minus25, 0)
                    End If

                    For a = 1 To HoleYBlk(2) - 1
                        If a = 1 Then
                            skPoint = Part.SketchManager.CreatePoint(0, ((PnlHgt / 1000) * FansY) + BlkHt(0) / 1000 + BlkHt(1) / 1000 - Minus25, 0)
                        Else
                            skPoint = Part.SketchManager.CreatePoint(0, ((PnlHgt / 1000) * FansY) + BlkHt(0) / 1000 + BlkHt(1) / 1000 + (HoleYBlkDist(2) * a) - Minus25, 0)
                        End If
                    Next
                End If

            End If
            '-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-

            Part.ClearSelection2(True)
            boolstatus = Part.EditRebuild3()
            Part.ViewZoomtofit2()
#End Region

SaveSideL2:
            boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)
            longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_03A_Side_L_Right.SLDPRT", 0, 0)

            SideLDrawings()              'Create Drawings for both Ls
            swApp.CloseAllDocuments(True)
#End Region
        Else
#Region "Two Side Ls"

            '    '03A Left-------------------------------------------------------------------
            '    Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_03A_Side_L.SLDPRT", 1, 0, "", longstatus, longwarnings)
            '    swApp.ActivateDoc2("_03A_Side_L.SLDPRT", False, longstatus)
            '    Part = swApp.ActiveDoc
            '    boolstatus = Part.Extension.SelectByID2("Height@BaseFlange@_03A_Side_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            '    myDimension = Part.Parameter("Height@BaseFlange")
            '    myDimension.SystemValue = (PnlHgt + PnlHgt / 2) / 1000 - 0.047
            '    HoleSpc = (PnlHgt + PnlHgt / 2) / 1000 - 0.047
            '    Part.ClearSelection2(True)
            '    boolstatus = Part.EditRebuild3()

            '    'boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_03A_Side_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            '    'myDimension = Part.Parameter("Width@BaseFlange")
            '    'myDimension.SystemValue = LWidth
            '    'Part.ClearSelection2(True)
            '    'boolstatus = Part.EditRebuild3()
            '    ' Zoom To Fit
            '    Part.ViewZoomtofit2()

            '    SideLHoles(HoleSpc)

            '    boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

            '    ' Save As
            '    longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_03A_Side_L_Left.SLDPRT", 0, 0)
            '    swApp.CloseAllDocuments(True)

            '    '03B Left-------------------------------------------------------------------

            '    Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_03A_Side_L.SLDPRT", 1, 0, "", longstatus, longwarnings)
            '    swApp.ActivateDoc2("_03A_Side_L.SLDPRT", False, longstatus)
            '    Part = swApp.ActiveDoc
            '    boolstatus = Part.Extension.SelectByID2("Height@BaseFlange@_03A_Side_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            '    myDimension = Part.Parameter("Height@BaseFlange")
            '    myDimension.SystemValue = (WallHt - (PnlHgt + PnlHgt / 2) - 6) / 1000 - 0.047
            '    HoleSpc = (WallHt - (PnlHgt + PnlHgt / 2) - 6) / 1000 - 0.047
            '    Part.ClearSelection2(True)
            '    boolstatus = Part.EditRebuild3()

            '    'boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_03A_Side_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            '    'myDimension = Part.Parameter("Width@BaseFlange")
            '    'myDimension.SystemValue = LWidth
            '    'Part.ClearSelection2(True)
            '    'boolstatus = Part.EditRebuild3()
            '    ' Zoom To Fit
            '    Part.ViewZoomtofit2()

            '    SideLHoles(HoleSpc)

            '    boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

            '    ' Save As
            '    longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_03B_Side_L2_Left.SLDPRT", 0, 0)
            '    swApp.CloseAllDocuments(True)

            '    '--------x---------x----------x----------x---------x---------x---------x--------x--------x--------x--------x--------

            '    '03A Right-------------------------------------------------------------------

            '    Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_03A_Side_L_Right.SLDPRT", 1, 0, "", longstatus, longwarnings)
            '    swApp.ActivateDoc2("_03A_Side_L_Right.SLDPRT", False, longstatus)
            '    Part = swApp.ActiveDoc
            '    boolstatus = Part.Extension.SelectByID2("Height@BaseFlange@_03A_Side_L_Right.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            '    myDimension = Part.Parameter("Height@BaseFlange")
            '    myDimension.SystemValue = (PnlHgt + PnlHgt / 2) / 1000 - 0.047
            '    HoleSpc = (PnlHgt + PnlHgt / 2) / 1000 - 0.047
            '    Part.ClearSelection2(True)
            '    boolstatus = Part.EditRebuild3()

            '    'boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_03A_Side_L_Right.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            '    'myDimension = Part.Parameter("Width@BaseFlange")
            '    'myDimension.SystemValue = LWidth
            '    'Part.ClearSelection2(True)
            '    'boolstatus = Part.EditRebuild3()
            '    ' Zoom To Fit
            '    Part.ViewZoomtofit2()

            '    SideLHoles(HoleSpc)

            '    boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

            '    ' Save As
            '    longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_03A_Side_L_Right.SLDPRT", 0, 0)
            '    swApp.CloseAllDocuments(True)

            '    '03B Right-------------------------------------------------------------------

            '    Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_03A_Side_L_Right.SLDPRT", 1, 0, "", longstatus, longwarnings)
            '    swApp.ActivateDoc2("_03A_Side_L_Right.SLDPRT", False, longstatus)
            '    Part = swApp.ActiveDoc
            '    boolstatus = Part.Extension.SelectByID2("Height@BaseFlange@_03A_Side_L_Right.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            '    myDimension = Part.Parameter("Height@BaseFlange")
            '    myDimension.SystemValue = (WallHt - (PnlHgt + PnlHgt / 2) - 6) / 1000 - 0.047
            '    HoleSpc = (WallHt - (PnlHgt + PnlHgt / 2) - 6) / 1000 - 0.047
            '    Part.ClearSelection2(True)
            '    boolstatus = Part.EditRebuild3()

            '    'boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_03A_Side_L_Right.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            '    'myDimension = Part.Parameter("Width@BaseFlange")
            '    'myDimension.SystemValue = LWidth
            '    'Part.ClearSelection2(True)
            '    'boolstatus = Part.EditRebuild3()

            '    ' Zoom To Fit
            '    Part.ViewZoomtofit2()

            '    SideLHoles(HoleSpc)

            '    boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

            '    ' Save As
            '    longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_03B_Side_L2_Right.SLDPRT", 0, 0)
            '    swApp.CloseAllDocuments(True)

        End If

        'SideLSlotsLeft(FansY, TopClear, PnlHgt, CrnBlkHt)
        'SideLSlotsRight(FansY, TopClear, PnlHgt, CrnBlkHt)

        '' Close Document   
        'swApp.CloseAllDocuments(True)
#End Region

    End Sub

    Public Sub SideLHoles(HoleSpc As Decimal)

        '4.2mm holes--------------------------------------------------------
        If WallHt <= MaxSecLth3mm Then
            Dim HoleY As Integer = SideHoleNumber(HoleSpc, 0.04)
            Dim HoleYDist As Decimal = SideHoleDist(HoleY, HoleSpc, 0.04)

            boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("4.2mm Hole", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2(HoleY + 1, HoleYDist, 0, 0, False, True, "NULL", "NULL", False)
            Part.ClearSelection2(True)
        Else
            Dim HoleY As Integer = SideHoleNumber(HoleSpc, 0.04)
            Dim HoleYDist As Decimal = SideHoleDist(HoleY, HoleSpc, 0.04)

            boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("4.2mm Hole", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2(HoleY + 1, HoleYDist, 0, 0, False, True, "NULL", "NULL", False)
            Part.ClearSelection2(True)
        End If

    End Sub

#End Region

#Region "Bot/Top L"
    Public Sub Top_Bot_L(LHeight As Decimal, BlkNosX As Integer, BlkNosY As Integer, BlkWth As Decimal, PanelWth As Decimal, FanNoX As Integer,
                         CrnBlkWth As Decimal, Pushside As Boolean, SideClear As Decimal, AHUDoor As Boolean, DoorBlkWth As Decimal, SBWth As Decimal)

        If LastLgt < 250 Then
            LastLgt = CrnBlkWth - 50
            xdist = LastLgt / 1000
        Else
            xdist = LastLgt / 1000
        End If

        ''OPEN TOP & BOTTOM L -----------------------------X-----------------------------------------------
        If WallWth <= (MaxSecLth3mm - 120) Then

            If AHUDoor = True Then
                GoTo AHUDoor
            End If

            Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_02A_Top & Bot_L.SLDPRT", 1, 0, "", longstatus, longwarnings)
            swApp.ActivateDoc2("_02A_Top & Bot_L.SLDPRT", False, longstatus)
            Part = swApp.ActiveDoc
            boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_02A_Top & Bot_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("Width@BaseFlange")
            myDimension.SystemValue = WallWth / 1000
            Part.ClearSelection2(True)

            TopBotLHoles()                                                     '4.2mm Holes

#Region "Slots"
            '--------------------------------------------Slots---------------------------------------------------------------
            swApp.ActivateDoc2("_02A_Top & Bot_L.SLDPRT", False, longstatus)
            Part = swApp.ActiveDoc

            boolstatus = Part.Extension.SelectByID2("Slots", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditSketch()


            '----------------------------------- Push Side = TRUE ------------------------------------------------

            If Pushside = True Then
                'Fans 
                For b = 1 To FanNoX
                    For a = 1 To MotorShtXHole - 1
                        skPoint = Part.SketchManager.CreatePoint(-((MotorShtXHoleDist * a) - ((WallWth / 2000) - 0.025 - (PanelWth / 1000 * (b - 1)))), 0, 0)
                    Next
                    skPoint = Part.SketchManager.CreatePoint(-((0.05 * b) + (b * (MotorShtXHoleDist * (MotorShtXHole - 1))) + ((-WallWth / 2000) + 0.025)), 0, 0)
                Next

                'Blanks
                If BlkNosX = 1 Then
                    For a = 1 To SideBlkXHoles(1) - 2
                        skPoint = Part.SketchManager.CreatePoint(-((PanelWth / 1000 * FanNoX) - (WallWth / 2000) + (SideBlkXDist(1) * a) + 0.025), 0, 0)
                    Next
                Else

                    MsgBox("WIP")
                    'For a = 1 To SideBlkXHoles(1) - 1
                    '    skPoint = Part.SketchManager.CreatePoint((PanelWth / 1000 * FanNoX) - (WallWth / 2000) + (SideBlkXDist(1) * a) + 0.025, 0, 0)
                    'Next

                    'For a = 1 To SideBlkXHoles(2) - 1
                    '    skPoint = Part.SketchManager.CreatePoint((PanelWth / 1000 * FanNoX) - (WallWth / 2000) + (SideBlkXDist(1) * a) + 0.025, 0, 0)
                    'Next

                End If

                Part.ClearSelection2(True)

                ' Zoom To Fit
                boolstatus = Part.EditRebuild3()
                Part.ViewZoomtofit2()

                boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)
                ' Save As
                longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02A_Bot_L.SLDPRT", 0, 0)
                swApp.CloseDoc(True)


                '-----------------------------------------Top L Pushside ----------------------------------------
                Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_02A_Top_L - PushSide.SLDPRT", 1, 0, "", longstatus, longwarnings)
                swApp.ActivateDoc2("_02A_Top_L - PushSide.SLDPRT", False, longstatus)
                Part = swApp.ActiveDoc
                boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_02A_Top_L - PushSide.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                myDimension = Part.Parameter("Width@BaseFlange")
                myDimension.SystemValue = WallWth / 1000
                Part.ClearSelection2(True)

                TopBotLHoles()                                                     '4.2mm Holes

                swApp.ActivateDoc2("_02A_Top_L - PushSide.SLDPRT", False, longstatus)
                Part = swApp.ActiveDoc

                boolstatus = Part.Extension.SelectByID2("Slots", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditSketch()

                'Fans 
                For b = 1 To FanNoX
                    For a = 1 To MotorShtXHole - 1
                        skPoint = Part.SketchManager.CreatePoint(-((MotorShtXHoleDist * a) - ((WallWth / 2000) - 0.025 - (PanelWth / 1000 * (b - 1)))), 0, 0)
                    Next
                    skPoint = Part.SketchManager.CreatePoint(-((0.05 * b) + (b * (MotorShtXHoleDist * (MotorShtXHole - 1))) + ((-WallWth / 2000) + 0.025)), 0, 0)
                Next

                'Blanks
                If BlkNosX = 1 Then
                    For a = 1 To SideBlkXHoles(1) - 2
                        skPoint = Part.SketchManager.CreatePoint(-((PanelWth / 1000 * FanNoX) - (WallWth / 2000) + (SideBlkXDist(1) * a) + 0.025), 0, 0)
                    Next
                Else

                    MsgBox("WIP")
                    'For a = 1 To SideBlkXHoles(1) - 1
                    '    skPoint = Part.SketchManager.CreatePoint((PanelWth / 1000 * FanNoX) - (WallWth / 2000) + (SideBlkXDist(1) * a) + 0.025, 0, 0)
                    'Next

                    'For a = 1 To SideBlkXHoles(2) - 1
                    '    skPoint = Part.SketchManager.CreatePoint((PanelWth / 1000 * FanNoX) - (WallWth / 2000) + (SideBlkXDist(1) * a) + 0.025, 0, 0)
                    'Next

                End If

                Part.ClearSelection2(True)

                ' Zoom To Fit
                boolstatus = Part.EditRebuild3()
                Part.ViewZoomtofit2()

                boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)
                ' Save As
                longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02A_Top_L.SLDPRT", 0, 0)
                swApp.CloseDoc(True)

                TopBotLDrawings()
                swApp.CloseDoc(True)

                Exit Sub
            End If

            If AHUDoor = True Then
AHUDoor:
                Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_02A_Top & Bot_L.SLDPRT", 1, 0, "", longstatus, longwarnings)
                swApp.ActivateDoc2("_02A_Top & Bot_L.SLDPRT", False, longstatus)
                Part = swApp.ActiveDoc
                boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_02A_Top & Bot_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                myDimension = Part.Parameter("Width@BaseFlange")
                myDimension.SystemValue = WallWth / 1000
                Part.ClearSelection2(True)

                TopBotLHoles()                                                     '4.2mm Holes

                swApp.ActivateDoc2("_02A_Top & Bot_L.SLDPRT", False, longstatus)
                Part = swApp.ActiveDoc

                boolstatus = Part.Extension.SelectByID2("Slots", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditSketch()

                'Blank Holes
                If BlkNosX > 1 Then
                    Dim d As Integer
                    If BlkNosY > 1 Then
                        d = 3
                    Else
                        d = 2
                    End If

                    For a = 1 To SideBlkXHoles(d)
                        skPoint = Part.SketchManager.CreatePoint(((WallWth / 2000) - (0.025 + (SideBlkXDist(d) * a))), 0, 0)
                    Next
                    skPoint = Part.SketchManager.CreatePoint(((WallWth / 2000) - (0.025 + SBWth)), 0, 0)
                End If

                Dim LastWth As Decimal
                If BlkNosX > 1 Then
                    LastWth = SBWth
                Else
                    LastWth = 0
                End If

                'Fan Holes
                For b = 1 To FanNoX
                    For a = 1 To MotorShtXHole - 1
                        skPoint = Part.SketchManager.CreatePoint(((WallWth / 2000) - (0.025 + LastWth + (MotorShtXHoleDist * a) + ((PanelWth / 1000) * (b - 1)))), 0, 0)
                    Next
                    skPoint = Part.SketchManager.CreatePoint(((WallWth / 2000) - (0.025 + LastWth + ((PanelWth / 1000) * b))), 0, 0)
                Next

                'Door Blank Holes
                For a = 1 To DoorXHoles - 2
                    skPoint = Part.SketchManager.CreatePoint(((WallWth / 2000) - (0.025 + LastWth + ((PanelWth / 1000) * FanNoX) + (DoorXHoleDist * a))), 0, 0)
                Next

                ' Zoom To Fit
                boolstatus = Part.EditRebuild3()
                Part.ViewZoomtofit2()

                boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)
                ' Save As
                longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02A_Bot_L.SLDPRT", 0, 0)
                swApp.CloseAllDocuments(True)

                '------------------------------------ Top L -----------------------------------------------

                Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_02A_Top & Bot_L.SLDPRT", 1, 0, "", longstatus, longwarnings)
                swApp.ActivateDoc2("_02A_Top & Bot_L.SLDPRT", False, longstatus)
                Part = swApp.ActiveDoc
                boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_02A_Top & Bot_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                myDimension = Part.Parameter("Width@BaseFlange")
                myDimension.SystemValue = WallWth / 1000
                Part.ClearSelection2(True)

                TopBotLHoles()                                                     '4.2mm Holes

                swApp.ActivateDoc2("_02A_Top & Bot_L.SLDPRT", False, longstatus)
                Part = swApp.ActiveDoc

                boolstatus = Part.Extension.SelectByID2("Slots", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditSketch()

                'Blank Holes
                If BlkNosX > 1 Then
                    Dim d As Integer
                    If BlkNosY > 1 Then
                        d = 3
                    Else
                        d = 2
                    End If
                    For a = 1 To SideBlkXHoles(d)
                        skPoint = Part.SketchManager.CreatePoint(((-WallWth / 2000) + (0.025 + (SideBlkXDist(d) * a))), 0, 0)
                    Next
                    skPoint = Part.SketchManager.CreatePoint(((-WallWth / 2000) + (0.025 + SBWth)), 0, 0)
                End If


                'Fan Holes
                For b = 1 To FanNoX
                    For a = 1 To MotorShtXHole - 1
                        skPoint = Part.SketchManager.CreatePoint(((-WallWth / 2000) + (0.025 + LastWth + (MotorShtXHoleDist * a) + ((PanelWth / 1000) * (b - 1)))), 0, 0)
                    Next
                    skPoint = Part.SketchManager.CreatePoint(((-WallWth / 2000) + (0.025 + LastWth + ((PanelWth / 1000) * b))), 0, 0)
                Next

                'Door Blank Holes
                For a = 1 To DoorXHoles - 2
                    skPoint = Part.SketchManager.CreatePoint(((-WallWth / 2000) + (0.025 + LastWth + ((PanelWth / 1000) * FanNoX) + (DoorXHoleDist * a))), 0, 0)
                Next

                ' Zoom To Fit
                boolstatus = Part.EditRebuild3()
                Part.ViewZoomtofit2()

                longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02A_Top_L.SLDPRT", 0, 0)
                swApp.CloseAllDocuments(True)

                TopBotLDrawings()
                swApp.CloseDoc(True)

                Exit Sub
            End If

            If FanNoX Mod 2 <> 0 Then                                      'Check if ODD
                '----------------------------------- ODD FANS ------------------------------------------------
                '-------------------------- ODD with 1 fan/ multiple fans ------------------------------------

                For b = 0 To Math.Floor(FanNoX / 2)
                    If b > 0 Then
                        For a = 1 To MotorShtXHole
                            skPoint = Part.SketchManager.CreatePoint((-MotorShtXHoleDist / 2) + (-MotorShtXHoleDist * a) - (b * 0.05), 0, 0)
                            skPoint = Part.SketchManager.CreatePoint((MotorShtXHoleDist / 2) + (MotorShtXHoleDist * a) + (b * 0.05), 0, 0)
                        Next
                    Else
                        Dim x As Integer = MotorShtXHole - 1
                        For a = 1 To x Step 2
                            skPoint = Part.SketchManager.CreatePoint((MotorShtXHoleDist / 2) * a, 0, 0)
                            skPoint = Part.SketchManager.CreatePoint((-MotorShtXHoleDist / 2) * a, 0, 0)
                        Next
                    End If
                Next

                '-------------------------- ODD Fans with SIDE BLANKS ------------------------------
                If BlkNosX > 0 Then                                     'if >1 SIDE BLANKS

                    If SideBlkXHoles(1) = 0 Or SideBlkXHoles(1) = 1 Then

                        skPoint = Part.SketchManager.CreatePoint(PanelWth * (FanNoX / 2000) + 0.025, 0, 0)
                        skPoint = Part.SketchManager.CreatePoint(-PanelWth * (FanNoX / 2000) - 0.025, 0, 0)
                        GoTo clrslc
                    End If

                    skPoint = Part.SketchManager.CreatePoint(PanelWth * (FanNoX / 2000) + 0.025, 0, 0)
                    skPoint = Part.SketchManager.CreatePoint(-PanelWth * (FanNoX / 2000) - 0.025, 0, 0)

                    For a = 1 To SideBlkXHoles(1)
                        If SideBlkXHoles(1) = 2 Then
                            skPoint = Part.SketchManager.CreatePoint(PanelWth * (FanNoX / 2000) + 0.025 + ((SideBlkXDist(1) / 2) * a), 0, 0)
                            skPoint = Part.SketchManager.CreatePoint(-PanelWth * (FanNoX / 2000) - 0.025 + ((-SideBlkXDist(1) / 2) * a), 0, 0)
                        Else
                            skPoint = Part.SketchManager.CreatePoint(PanelWth * (FanNoX / 2000) + 0.025 + (SideBlkXDist(1) * a), 0, 0)
                            skPoint = Part.SketchManager.CreatePoint(-PanelWth * (FanNoX / 2000) - 0.025 + (-SideBlkXDist(1) * a), 0, 0)
                        End If
                    Next

                    If SideBlkXHoles(2) = 0 Then
                        GoTo clrslc
                    ElseIf SideBlkXHoles(2) = 1 Then
                        skPoint = Part.SketchManager.CreatePoint(PanelWth * (FanNoX / 2000) + 0.025 + (PanelWth / 1000) + 0.025, 0, 0)
                        skPoint = Part.SketchManager.CreatePoint(-PanelWth * (FanNoX / 2000) - 0.025 + (-PanelWth / 1000) - 0.025, 0, 0)
                    End If

                    For a = 1 To SideBlkXHoles(2) - 1
                        If SideBlkXHoles(2) = 2 Then
                            skPoint = Part.SketchManager.CreatePoint(PanelWth * (FanNoX / 2000) + 0.025 + (PanelWth / 1000) + ((SideBlkXDist(2) / 2) * a), 0, 0)
                            skPoint = Part.SketchManager.CreatePoint(-PanelWth * (FanNoX / 2000) - 0.025 + ((-SideBlkXDist(2) / 2) * a) + (-PanelWth / 1000), 0, 0)
                        Else
                            skPoint = Part.SketchManager.CreatePoint(PanelWth * (FanNoX / 2000) + 0.025 + (PanelWth / 1000) + (SideBlkXDist(2) * a), 0, 0)
                            skPoint = Part.SketchManager.CreatePoint(-PanelWth * (FanNoX / 2000) - 0.025 + (-SideBlkXDist(2) * a) + (-PanelWth / 1000), 0, 0)
                        End If
                    Next

                End If

                '-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-

                '------------------------------------- EVEN FANS ---------------------------------------------------
            ElseIf FanNoX Mod 2 = 0 Then                                  'Check if EVEN

                'EVEN 
                For b = 1 To (FanNoX / 2)
                    If b = 1 Then
                        For a = 0 To MotorShtXHole - 1
                            skPoint = Part.SketchManager.CreatePoint(-0.025 + (-MotorShtXHoleDist * a), 0, 0)
                            skPoint = Part.SketchManager.CreatePoint(0.025 + (MotorShtXHoleDist * a), 0, 0)
                        Next
                    Else
                        For a = 0 To MotorShtXHole - 1
                            skPoint = Part.SketchManager.CreatePoint(-0.025 + (-MotorShtXHoleDist * a) - (b * 0.025), 0, 0)
                            skPoint = Part.SketchManager.CreatePoint(0.025 + (MotorShtXHoleDist * a) + (b * 0.025), 0, 0)
                        Next
                    End If
                Next

                '---------------------------- EVEN Fans with SIDE BLANKS ----------------------------------
                If BlkNosX > 0 Then

                    If SideBlkXHoles(1) = 0 Or SideBlkXHoles(1) = 1 Then
                        GoTo clrslc
                    End If

                    skPoint = Part.SketchManager.CreatePoint(PanelWth * (FanNoX / 2000) + 0.025, 0, 0)
                    skPoint = Part.SketchManager.CreatePoint(-PanelWth * (FanNoX / 2000) - 0.025, 0, 0)

                    For a = 1 To SideBlkXHoles(1)
                        If SideBlkXHoles(1) = 2 Then
                            skPoint = Part.SketchManager.CreatePoint(PanelWth * (FanNoX / 2000) + 0.025 + ((SideBlkXDist(1) / 2) * a), 0, 0)
                            skPoint = Part.SketchManager.CreatePoint(-PanelWth * (FanNoX / 2000) - 0.025 + ((-SideBlkXDist(1) / 2) * a), 0, 0)
                        Else
                            skPoint = Part.SketchManager.CreatePoint(PanelWth * (FanNoX / 2000) + 0.025 + (SideBlkXDist(1) * a), 0, 0)
                            skPoint = Part.SketchManager.CreatePoint(-PanelWth * (FanNoX / 2000) - 0.025 + (-SideBlkXDist(1) * a), 0, 0)
                        End If

                    Next

                    If SideBlkXHoles(2) = 0 Then
                        GoTo clrslc
                    ElseIf SideBlkXHoles(2) = 1 Then
                        skPoint = Part.SketchManager.CreatePoint(PanelWth * (FanNoX / 2000) + 0.025 + (PanelWth / 1000) + 0.025, 0, 0)
                        skPoint = Part.SketchManager.CreatePoint(-PanelWth * (FanNoX / 2000) - 0.025 + (-PanelWth / 1000) - 0.025, 0, 0)
                    End If

                    For a = 1 To SideBlkXHoles(2) - 1
                        If SideBlkXHoles(2) = 2 Then
                            skPoint = Part.SketchManager.CreatePoint(PanelWth * (FanNoX / 2000) + 0.025 + (PanelWth / 1000) + ((SideBlkXDist(2) / 2) * a), 0, 0)
                            skPoint = Part.SketchManager.CreatePoint(-PanelWth * (FanNoX / 2000) - 0.025 + ((-SideBlkXDist(2) / 2) * a) + (-PanelWth / 1000), 0, 0)

                        Else
                            skPoint = Part.SketchManager.CreatePoint(PanelWth * (FanNoX / 2000) + 0.025 + (PanelWth / 1000) + (SideBlkXDist(2) * a), 0, 0)
                            skPoint = Part.SketchManager.CreatePoint(-PanelWth * (FanNoX / 2000) - 0.025 + (-SideBlkXDist(2) * a) + (-PanelWth / 1000), 0, 0)
                        End If
                    Next

                End If

            End If

            Part.ClearSelection2(True)
clrslc:

#End Region

            ' Zoom To Fit
            boolstatus = Part.EditRebuild3()
            Part.ViewZoomtofit2()

            boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)
            ' Save As
            longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02A_Bot_L.SLDPRT", 0, 0)
            longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02A_Top_L.SLDPRT", 0, 0)
            swApp.CloseAllDocuments(True)

            TopBotLDrawings()
            swApp.CloseAllDocuments(True)


        Else

            Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_02A_Top & Bot_L.SLDPRT", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("_02A_Top & Bot_L.SLDPRT", False, longstatus)
        Part = swApp.ActiveDoc

        If AHUDoor = True Then

            If FanNoX Mod 2 = 0 Then
#Region "Even FanX"
                'Even fans
                boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_02A_Top & Bot_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                myDimension = Part.Parameter("Width@BaseFlange")
                myDimension.SystemValue = DoorBlkWth + (PanelWth * (FanNoX / 2000))
                BTLth = DoorBlkWth + (PanelWth * (FanNoX / 2000))
                Part.ClearSelection2(True)


                TopBotLHoles()                                                     '4.2mm Holes

                '--------------------------------------------------------- Bot L2 ---------------------------------------------------
                swApp.ActivateDoc2("_02A_Top & Bot_L.SLDPRT", False, longstatus)
                Part = swApp.ActiveDoc

                boolstatus = Part.Extension.SelectByID2("Slots", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditSketch()


                'Fan Holes
                For b = 1 To FanNoX / 2
                    For a = 1 To MotorShtXHole - 1
                        skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - 0.025 - (MotorShtXHoleDist * a) - (PanelWth / 1000 * (b - 1))), 0, 0)
                    Next
                    skPoint = Part.SketchManager.CreatePoint((BTLth / 2) - 0.025 - (b * (MotorShtXHoleDist * (MotorShtXHole - 1))) - (0.05 * b), 0, 0)
                Next

                'Door Blank Holes
                For a = 1 To DoorXHoles - 2
                    skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - (((PanelWth / 1000 * (FanNoX / 2)))) - (DoorXHoleDist * a) - 0.025), 0, 0)
                Next

                ' Zoom To Fit
                boolstatus = Part.EditRebuild3()
                Part.ViewZoomtofit2()

                boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)
                ' Save As
                longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02B_Bot_L2.SLDPRT", 0, 0)
                swApp.CloseAllDocuments(True)


                '------------------------------------------------- Top L2 ----------------------------------------------------------------

                Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_02A_Top & Bot_L.SLDPRT", 1, 0, "", longstatus, longwarnings)
                swApp.ActivateDoc2("_02A_Top & Bot_L.SLDPRT", False, longstatus)
                Part = swApp.ActiveDoc

                boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_02A_Top & Bot_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                myDimension = Part.Parameter("Width@BaseFlange")
                myDimension.SystemValue = DoorBlkWth + (PanelWth * (FanNoX / 2000))
                BTLth = DoorBlkWth + (PanelWth * (FanNoX / 2000))
                Part.ClearSelection2(True)


                TopBotLHoles()                                                     '4.2mm Holes

                swApp.ActivateDoc2("_02A_Top & Bot_L.SLDPRT", False, longstatus)
                Part = swApp.ActiveDoc

                boolstatus = Part.Extension.SelectByID2("Slots", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditSketch()

                'Fan Holes
                For b = 1 To FanNoX / 2
                    For a = 1 To MotorShtXHole - 1
                        skPoint = Part.SketchManager.CreatePoint(((-BTLth / 2) + 0.025 + (MotorShtXHoleDist * a) + (PanelWth / 1000 * (b - 1))), 0, 0)
                    Next
                    skPoint = Part.SketchManager.CreatePoint((-BTLth / 2) + 0.025 + (b * (MotorShtXHoleDist * (MotorShtXHole - 1))) + (0.05 * b), 0, 0)
                Next

                'Door Blank Holes
                For a = 1 To DoorXHoles - 2
                    skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - 0.025 - (((PanelWth / 1000 * (FanNoX / 2)))) - (SideBlkXDist(1) * a)), 0, 0)
                Next

                longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02B_Top_L2.SLDPRT", 0, 0)
                swApp.CloseAllDocuments(True)


                '------------------------------------ Bot L -------------------------------------------

                Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_02A_Top & Bot_L.SLDPRT", 1, 0, "", longstatus, longwarnings)
                swApp.ActivateDoc2("_02A_Top & Bot_L.SLDPRT", False, longstatus)
                Part = swApp.ActiveDoc

                boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_02A_Top & Bot_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                myDimension = Part.Parameter("Width@BaseFlange")
                myDimension.SystemValue = (WallWth / 1000) - (DoorBlkWth + (PanelWth * (FanNoX / 2000)))
                BTLth = (WallWth / 1000) - (DoorBlkWth + (PanelWth * (FanNoX / 2000)))
                Part.ClearSelection2(True)


                TopBotLHoles()                                                     '4.2mm Holes

                boolstatus = Part.Extension.SelectByID2("Slots", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditSketch()

                'Fan Holes
                For b = 1 To FanNoX / 2
                    For a = 1 To MotorShtXHole - 1
                        skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - 0.025 - (MotorShtXHoleDist * a) - (PanelWth / 1000 * (b - 1))), 0, 0)
                    Next
                    skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - 0.025 - (b * (MotorShtXHoleDist * (MotorShtXHole - 1))) - (0.05 * b)), 0, 0)
                Next

                    'Blank Holes
                    If BlkNosX = 1 Then
                        For a = 1 To SideBlkXHoles(1) - 2
                            skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - (((PanelWth / 1000 * (FanNoX / 2)))) - (SideBlkXDist(1) * a) - 0.025), 0, 0)
                        Next
                    Else
                        Dim d As Integer
                        If BlkNosY > 1 Then
                            d = 3
                        Else
                            d = 2
                        End If
                        For a = 1 To SideBlkXHoles(d) - 1
                            skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - (((PanelWth / 1000 * (FanNoX / 2)))) - (SideBlkXDist(d) * a) - 0.025), 0, 0)
                        Next
                    End If


                ' Zoom To Fit
                boolstatus = Part.EditRebuild3()
                Part.ViewZoomtofit2()

                boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

                ' Save As
                longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02A_Bot_L.SLDPRT", 0, 0)
                swApp.CloseAllDocuments(True)


                '------------------------------- Top L -----------------------------------------

                Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_02A_Top & Bot_L.SLDPRT", 1, 0, "", longstatus, longwarnings)
                swApp.ActivateDoc2("_02A_Top & Bot_L.SLDPRT", False, longstatus)
                Part = swApp.ActiveDoc

                boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_02A_Top & Bot_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                myDimension = Part.Parameter("Width@BaseFlange")
                myDimension.SystemValue = (WallWth / 1000) - (DoorBlkWth + (PanelWth * (FanNoX / 2000)))
                BTLth = (WallWth / 1000) - (DoorBlkWth + (PanelWth * (FanNoX / 2000)))
                Part.ClearSelection2(True)


                TopBotLHoles()                                                     '4.2mm Holes

                boolstatus = Part.Extension.SelectByID2("Slots", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditSketch()


                'Fan Holes
                For b = 1 To FanNoX / 2
                    For a = 1 To MotorShtXHole - 1
                        skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - 0.025 - (MotorShtXHoleDist * a) - (PanelWth / 1000 * (b - 1))), 0, 0)
                    Next
                    skPoint = Part.SketchManager.CreatePoint((BTLth / 2) - 0.025 - (b * (MotorShtXHoleDist * (MotorShtXHole - 1))) - (0.05 * b), 0, 0)

                Next

                    'Blank Holes
                    If BlkNosX = 1 Then
                        For a = 1 To SideBlkXHoles(1) - 2
                            skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - 0.025 - (((PanelWth / 1000 * (FanNoX / 2)))) - (SideBlkXDist(1) * a)), 0, 0)
                        Next
                    Else
                        Dim d As Integer
                        If BlkNosY > 1 Then
                            d = 3
                        Else
                            d = 2
                        End If
                        For a = 1 To SideBlkXHoles(d) - 1
                            skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - 0.025 - (((PanelWth / 1000 * (FanNoX / 2)))) - (SideBlkXDist(d) * a)), 0, 0)
                        Next

                        'MsgBox("WIP")
                        'For a = 1 To SideBlkXHoles(1) - 1
                        '    skPoint = Part.SketchManager.CreatePoint((PanelWth / 1000 * FanNoX) - (WallWth / 2000) + (SideBlkXDist(1) * a) + 0.025, 0, 0)
                        'Next

                        'For a = 1 To SideBlkXHoles(2) - 1
                        '    skPoint = Part.SketchManager.CreatePoint((PanelWth / 1000 * FanNoX) - (WallWth / 2000) + (SideBlkXDist(1) * a) + 0.025, 0, 0)
                        'Next
                    End If

                ' Zoom To Fit
                boolstatus = Part.EditRebuild3()
                Part.ViewZoomtofit2()

                boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)
                ' Save As
                longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02A_Top_L.SLDPRT", 0, 0)
                swApp.CloseAllDocuments(True)
#End Region
            Else
#Region "Odd FanX"
                'Odd fans
                boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_02A_Top & Bot_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                myDimension = Part.Parameter("Width@BaseFlange")
                myDimension.SystemValue = DoorBlkWth + (PanelWth / 1000 * Truncate(FanNoX / 2))
                BTLth = DoorBlkWth + (PanelWth / 1000 * Truncate(FanNoX / 2))
                Part.ClearSelection2(True)


                TopBotLHoles()                                                     '4.2mm Holes

                boolstatus = Part.Extension.SelectByID2("Slots", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditSketch()

                'Fan Holes
                For b = 1 To Truncate(FanNoX / 2)
                    For a = 1 To MotorShtXHole - 1
                        skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - 0.025 - (MotorShtXHoleDist * a) - (PanelWth / 1000 * (b - 1))), 0, 0)
                    Next
                    skPoint = Part.SketchManager.CreatePoint((BTLth / 2) - 0.025 - (b * (MotorShtXHoleDist * (MotorShtXHole - 1))) - (0.05 * b), 0, 0)
                Next

                'Door Blank Holes
                For a = 1 To DoorXHoles - 2
                    skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - (((PanelWth / 1000 * Truncate(FanNoX / 2)))) - (DoorXHoleDist * a) - 0.025), 0, 0)
                Next


                ' Zoom To Fit
                boolstatus = Part.EditRebuild3()
                Part.ViewZoomtofit2()

                boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)
                ' Save As
                longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02B_Bot_L2.SLDPRT", 0, 0)
                swApp.CloseAllDocuments(True)


                '-x-x-x-xx-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x--x-x-x-xx-x-x-x-
                '------------- Top L2 -----------------
                Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_02A_Top & Bot_L.SLDPRT", 1, 0, "", longstatus, longwarnings)
                swApp.ActivateDoc2("_02A_Top & Bot_L.SLDPRT", False, longstatus)
                Part = swApp.ActiveDoc

                boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_02A_Top & Bot_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                myDimension = Part.Parameter("Width@BaseFlange")
                myDimension.SystemValue = DoorBlkWth + (PanelWth / 1000 * Truncate(FanNoX / 2))
                BTLth = DoorBlkWth + (PanelWth / 1000 * Truncate(FanNoX / 2))
                Part.ClearSelection2(True)


                TopBotLHoles()                                                     '4.2mm Holes

                boolstatus = Part.Extension.SelectByID2("Slots", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditSketch()

                'Fan Holes
                For b = 1 To Truncate(FanNoX / 2)
                    For a = 1 To MotorShtXHole - 1
                        skPoint = Part.SketchManager.CreatePoint(((-BTLth / 2) + 0.025 + (MotorShtXHoleDist * a) + (PanelWth / 1000 * (b - 1))), 0, 0)
                    Next
                    skPoint = Part.SketchManager.CreatePoint((-BTLth / 2) + 0.025 + (b * (MotorShtXHoleDist * (MotorShtXHole - 1))) + (0.05 * b), 0, 0)

                Next

                'Door Blank Holes
                For a = 1 To DoorXHoles - 2
                    skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - 0.025 - (((PanelWth / 1000 * Truncate(FanNoX / 2)))) - (DoorXHoleDist * a)), 0, 0)
                Next

                ' Zoom To Fit
                boolstatus = Part.EditRebuild3()
                Part.ViewZoomtofit2()
                boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

                ' Save As
                longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02B_Top_L2.SLDPRT", 0, 0)
                swApp.CloseAllDocuments(True)


                '-x-x-x-x--x-x-x-x-x-x-x-x-x-x-x- 02A Bot L x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-

                Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_02A_Top & Bot_L.SLDPRT", 1, 0, "", longstatus, longwarnings)
                swApp.ActivateDoc2("_02A_Top & Bot_L.SLDPRT", False, longstatus)
                Part = swApp.ActiveDoc

                boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_02A_Top & Bot_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                myDimension = Part.Parameter("Width@BaseFlange")
                myDimension.SystemValue = (WallWth / 1000) - (DoorBlkWth + (PanelWth / 1000 * Truncate(FanNoX / 2)))
                BTLth = (WallWth / 1000) - (DoorBlkWth + (PanelWth / 1000 * Truncate(FanNoX / 2)))
                Part.ClearSelection2(True)


                TopBotLHoles()                                                     '4.2mm Holes

                boolstatus = Part.Extension.SelectByID2("Slots", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditSketch()

                'Fan Holes
                For b = 1 To Ceiling(FanNoX / 2)
                    For a = 1 To MotorShtXHole - 1
                        skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - 0.025 - (MotorShtXHoleDist * a) - (PanelWth / 1000 * (b - 1))), 0, 0)
                    Next
                    skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - 0.025 - (b * (MotorShtXHoleDist * (MotorShtXHole - 1))) - (0.05 * b)), 0, 0)
                Next

                'Blank Holes
                If BlkNosX = 1 Then
                    For a = 1 To SideBlkXHoles(1) - 2
                        skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - (((PanelWth / 1000 * Ceiling(FanNoX / 2)))) - (SideBlkXDist(1) * a) - 0.025), 0, 0)
                    Next
                Else
                    For a = 1 To SideBlkXHoles(2) - 1
                        skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - (((PanelWth / 1000 * Ceiling(FanNoX / 2)))) - ((SideBlkXDist(2) / 2) * a) - 0.025), 0, 0)
                    Next

                End If


                ' Zoom To Fit
                boolstatus = Part.EditRebuild3()
                Part.ViewZoomtofit2()

                boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

                ' Save As
                longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02A_Bot_L.SLDPRT", 0, 0)
                swApp.CloseAllDocuments(True)


                '-x-x-x-x--x-x-x-x-x-x-x-x-x-x-x- 02A Top L x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-

                Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_02A_Top & Bot_L.SLDPRT", 1, 0, "", longstatus, longwarnings)
                swApp.ActivateDoc2("_02A_Top & Bot_L.SLDPRT", False, longstatus)
                Part = swApp.ActiveDoc

                boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_02A_Top & Bot_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                myDimension = Part.Parameter("Width@BaseFlange")
                myDimension.SystemValue = (WallWth / 1000) - (DoorBlkWth + (PanelWth / 1000 * Truncate(FanNoX / 2)))
                BTLth = (WallWth / 1000) - (DoorBlkWth + (PanelWth / 1000 * Truncate(FanNoX / 2)))
                Part.ClearSelection2(True)


                TopBotLHoles()                                                     '4.2mm Holes

                boolstatus = Part.Extension.SelectByID2("Slots", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditSketch()


                'Fan Holes
                For b = 1 To Ceiling(FanNoX / 2)
                    For a = 1 To MotorShtXHole - 1
                        skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - 0.025 - (MotorShtXHoleDist * a) - (PanelWth / 1000 * (b - 1))), 0, 0)
                    Next
                    skPoint = Part.SketchManager.CreatePoint((BTLth / 2) - 0.025 - (b * (MotorShtXHoleDist * (MotorShtXHole - 1))) - (0.05 * b), 0, 0)
                Next

                'Blank Holes
                If BlkNosX = 1 Then
                    For a = 1 To SideBlkXHoles(1) - 2
                        skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - 0.025 - (((PanelWth / 1000 * Ceiling(FanNoX / 2)))) - (SideBlkXDist(1) * a)), 0, 0)
                    Next
                Else
                    For a = 1 To SideBlkXHoles(2) - 1
                        skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - 0.025 - (((PanelWth / 1000 * Ceiling(FanNoX / 2)))) - ((SideBlkXDist(2) / 2) * a)), 0, 0)
                    Next

                End If

                longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02A_Top_L.SLDPRT", 0, 0)
                swApp.CloseAllDocuments(True)

            End If      'Even & Odd fansX
#End Region
        Else
            'NO AHU Door
            If FanNoX Mod 2 = 0 Then
                'Even fans
#Region "Even Fans"

                boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_02A_Top & Bot_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                myDimension = Part.Parameter("Width@BaseFlange")
                myDimension.SystemValue = SideClear / 1000 + (PanelWth * (FanNoX / 2000))
                BTLth = SideClear / 1000 + (PanelWth * (FanNoX / 2000))
                Part.ClearSelection2(True)


                TopBotLHoles()                                                     '4.2mm Holes

                '------- Bot L2 ----------
                swApp.ActivateDoc2("_02A_Top & Bot_L.SLDPRT", False, longstatus)
                Part = swApp.ActiveDoc

                boolstatus = Part.Extension.SelectByID2("Slots", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditSketch()

                'Fan Holes
                For b = 1 To FanNoX / 2
                    For a = 1 To MotorShtXHole - 1
                        skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - 0.025 - (MotorShtXHoleDist * a) - (PanelWth / 1000 * (b - 1))), 0, 0)
                    Next
                    skPoint = Part.SketchManager.CreatePoint((BTLth / 2) - 0.025 - (b * (MotorShtXHoleDist * (MotorShtXHole - 1))) - (0.05 * b), 0, 0)
                Next

                'Blank Holes
                If BlkNosX = 1 Then
                    For a = 1 To SideBlkXHoles(1) - 2
                        skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - (((PanelWth / 1000 * (FanNoX / 2)))) - (SideBlkXDist(1) * a) - 0.025), 0, 0)
                    Next
                Else

                    For a = 1 To SideBlkXHoles(1) - 2
                        skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - (((PanelWth / 1000 * (FanNoX / 2)))) - (SideBlkXDist(1) * a) - 0.025), 0, 0)
                    Next

                    For a = 1 To SideBlkXHoles(2) - 1
                        skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - (((PanelWth / 1000 * (FanNoX / 2)))) - (SideBlkXDist(2) * a) - 0.025), 0, 0)
                    Next

                    'MsgBox("WIP")
                    'For a = 1 To SideBlkXHoles(1) - 1
                    '    skPoint = Part.SketchManager.CreatePoint((PanelWth / 1000 * FanNoX) - (WallWth / 2000) + (SideBlkXDist(1) * a) + 0.025, 0, 0)
                    'Next

                    'For a = 1 To SideBlkXHoles(2) - 1
                    '    skPoint = Part.SketchManager.CreatePoint((PanelWth / 1000 * FanNoX) - (WallWth / 2000) + (SideBlkXDist(1) * a) + 0.025, 0, 0)
                    'Next
                End If


                ' Zoom To Fit
                boolstatus = Part.EditRebuild3()
                Part.ViewZoomtofit2()

                boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)
                ' Save As
                longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02B_Bot_L2.SLDPRT", 0, 0)
                swApp.CloseAllDocuments(True)

                '-x-x-x-xx-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x--x-x-x-xx-x-x-x-
                '------------- Top L2 -----------------
                Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_02A_Top & Bot_L.SLDPRT", 1, 0, "", longstatus, longwarnings)
                swApp.ActivateDoc2("_02A_Top & Bot_L.SLDPRT", False, longstatus)
                Part = swApp.ActiveDoc

                boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_02A_Top & Bot_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                myDimension = Part.Parameter("Width@BaseFlange")
                myDimension.SystemValue = SideClear / 1000 + (PanelWth * (FanNoX / 2000))
                BTLth = SideClear / 1000 + (PanelWth * (FanNoX / 2000))
                Part.ClearSelection2(True)


                TopBotLHoles()                                                     '4.2mm Holes

                boolstatus = Part.Extension.SelectByID2("Slots", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditSketch()

                'Fan Holes
                For b = 1 To FanNoX / 2
                    For a = 1 To MotorShtXHole - 1
                        skPoint = Part.SketchManager.CreatePoint(((-BTLth / 2) + 0.025 + (MotorShtXHoleDist * a) + (PanelWth / 1000 * (b - 1))), 0, 0)
                    Next
                    skPoint = Part.SketchManager.CreatePoint((-BTLth / 2) + 0.025 + (b * (MotorShtXHoleDist * (MotorShtXHole - 1))) + (0.05 * b), 0, 0)

                Next

                'Blank Holes
                If BlkNosX = 1 Then
                    For a = 1 To SideBlkXHoles(1) - 2
                        skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - 0.025 - (((PanelWth / 1000 * (FanNoX / 2)))) - (SideBlkXDist(1) * a)), 0, 0)
                    Next
                Else

                    For a = 1 To SideBlkXHoles(1) - 2
                        skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - 0.025 - (((PanelWth / 1000 * (FanNoX / 2)))) - (SideBlkXDist(1) * a)), 0, 0)
                    Next

                    For a = 1 To SideBlkXHoles(2) - 1
                        skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - 0.025 - (((PanelWth / 1000 * (FanNoX / 2)))) - (SideBlkXDist(2) * a)), 0, 0)
                    Next
                    'MsgBox("WIP")
                    'For a = 1 To SideBlkXHoles(1) - 1
                    '    skPoint = Part.SketchManager.CreatePoint((PanelWth / 1000 * FanNoX) - (WallWth / 2000) + (SideBlkXDist(1) * a) + 0.025, 0, 0)
                    'Next

                    'For a = 1 To SideBlkXHoles(2) - 1
                    '    skPoint = Part.SketchManager.CreatePoint((PanelWth / 1000 * FanNoX) - (WallWth / 2000) + (SideBlkXDist(1) * a) + 0.025, 0, 0)
                    'Next

                End If


                ' Zoom To Fit
                boolstatus = Part.EditRebuild3()
                Part.ViewZoomtofit2()

                boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)
                ' Save As
                longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02B_Top_L2.SLDPRT", 0, 0)
                swApp.CloseAllDocuments(True)


                '-x-x-x-xx-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x--x-x-x-xx-x-x-x-
                '------------- Bot L -----------------

                Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_02A_Top & Bot_L.SLDPRT", 1, 0, "", longstatus, longwarnings)
                swApp.ActivateDoc2("_02A_Top & Bot_L.SLDPRT", False, longstatus)
                Part = swApp.ActiveDoc

                boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_02A_Top & Bot_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                myDimension = Part.Parameter("Width@BaseFlange")
                myDimension.SystemValue = (WallWth / 1000) - (SideClear / 1000 + (PanelWth * (FanNoX / 2000)))
                BTLth = (WallWth / 1000) - (SideClear / 1000 + (PanelWth * (FanNoX / 2000)))
                Part.ClearSelection2(True)


                TopBotLHoles()                                                     '4.2mm Holes

                boolstatus = Part.Extension.SelectByID2("Slots", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditSketch()

                'Fan Holes
                For b = 1 To FanNoX / 2
                    For a = 1 To MotorShtXHole - 1
                        skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - 0.025 - (MotorShtXHoleDist * a) - (PanelWth / 1000 * (b - 1))), 0, 0)
                    Next
                    skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - 0.025 - (b * (MotorShtXHoleDist * (MotorShtXHole - 1))) - (0.05 * b)), 0, 0)
                Next

                'Blank Holes
                If BlkNosX = 1 Then
                    For a = 1 To SideBlkXHoles(1) - 2
                        skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - (((PanelWth / 1000 * (FanNoX / 2)))) - (SideBlkXDist(1) * a) - 0.025), 0, 0)
                    Next
                Else
                    For a = 1 To SideBlkXHoles(1) - 2
                        skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - (((PanelWth / 1000 * (FanNoX / 2)))) - (SideBlkXDist(1) * a) - 0.025), 0, 0)
                    Next

                    For a = 1 To SideBlkXHoles(2) - 1
                        skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - (((PanelWth / 1000 * (FanNoX / 2)))) - (SideBlkXDist(2) * a) - 0.025), 0, 0)
                    Next
                    'MsgBox("WIP")
                    'For a = 1 To SideBlkXHoles(1) - 1
                    '    skPoint = Part.SketchManager.CreatePoint((PanelWth / 1000 * FanNoX) - (WallWth / 2000) + (SideBlkXDist(1) * a) + 0.025, 0, 0)
                    'Next

                    'For a = 1 To SideBlkXHoles(2) - 1
                    '    skPoint = Part.SketchManager.CreatePoint((PanelWth / 1000 * FanNoX) - (WallWth / 2000) + (SideBlkXDist(1) * a) + 0.025, 0, 0)
                    'Next
                End If


                ' Zoom To Fit
                boolstatus = Part.EditRebuild3()
                Part.ViewZoomtofit2()

                boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

                ' Save As
                longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02A_Bot_L.SLDPRT", 0, 0)
                swApp.CloseAllDocuments(True)

                '-x-x-x-xx-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x--x-x-x-xx-x-x-x-
                '------------- Top L -----------------

                Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_02A_Top & Bot_L.SLDPRT", 1, 0, "", longstatus, longwarnings)
                swApp.ActivateDoc2("_02A_Top & Bot_L.SLDPRT", False, longstatus)
                Part = swApp.ActiveDoc

                boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_02A_Top & Bot_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                myDimension = Part.Parameter("Width@BaseFlange")
                myDimension.SystemValue = (WallWth / 1000) - (SideClear / 1000 + (PanelWth * (FanNoX / 2000)))
                BTLth = (WallWth / 1000) - (SideClear / 1000 + (PanelWth * (FanNoX / 2000)))
                Part.ClearSelection2(True)


                TopBotLHoles()                                                     '4.2mm Holes

                boolstatus = Part.Extension.SelectByID2("Slots", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditSketch()


                'Fan Holes
                For b = 1 To FanNoX / 2
                    For a = 1 To MotorShtXHole - 1
                        skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - 0.025 - (MotorShtXHoleDist * a) - (PanelWth / 1000 * (b - 1))), 0, 0)
                    Next
                    skPoint = Part.SketchManager.CreatePoint((BTLth / 2) - 0.025 - (b * (MotorShtXHoleDist * (MotorShtXHole - 1))) - (0.05 * b), 0, 0)

                Next

                'Blank Holes
                If BlkNosX = 1 Then
                    For a = 1 To SideBlkXHoles(1) - 2
                        skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - 0.025 - (((PanelWth / 1000 * (FanNoX / 2)))) - (SideBlkXDist(1) * a)), 0, 0)
                    Next
                Else
                    For a = 1 To SideBlkXHoles(1) - 2
                        skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - 0.025 - (((PanelWth / 1000 * (FanNoX / 2)))) - (SideBlkXDist(1) * a)), 0, 0)
                    Next

                    For a = 1 To SideBlkXHoles(2) - 1
                        skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - 0.025 - (((PanelWth / 1000 * (FanNoX / 2)))) - (SideBlkXDist(2) * a)), 0, 0)
                    Next

                    'MsgBox("WIP")
                    'For a = 1 To SideBlkXHoles(1) - 1
                    '    skPoint = Part.SketchManager.CreatePoint((PanelWth / 1000 * FanNoX) - (WallWth / 2000) + (SideBlkXDist(1) * a) + 0.025, 0, 0)
                    'Next

                    'For a = 1 To SideBlkXHoles(2) - 1
                    '    skPoint = Part.SketchManager.CreatePoint((PanelWth / 1000 * FanNoX) - (WallWth / 2000) + (SideBlkXDist(1) * a) + 0.025, 0, 0)
                    'Next

                End If

                ' Zoom To Fit
                boolstatus = Part.EditRebuild3()
                Part.ViewZoomtofit2()

                boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)
                ' Save As
                longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02A_Top_L.SLDPRT", 0, 0)
                swApp.CloseAllDocuments(True)
#End Region
            Else
                'Odd fans
#Region "Odd Fans"
                boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_02A_Top & Bot_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                myDimension = Part.Parameter("Width@BaseFlange")
                myDimension.SystemValue = SideClear / 1000 + ((PanelWth / 1000) * Truncate(FanNoX / 2))
                BTLth = SideClear / 1000 + ((PanelWth / 1000) * Truncate(FanNoX / 2))
                Part.ClearSelection2(True)


                TopBotLHoles()                                                     '4.2mm Holes

                '------- Bot L2 ----------
                swApp.ActivateDoc2("_02A_Top & Bot_L.SLDPRT", False, longstatus)
                Part = swApp.ActiveDoc

                boolstatus = Part.Extension.SelectByID2("Slots", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditSketch()

                'Fan Holes
                For b = 1 To Truncate(FanNoX / 2)
                    For a = 1 To MotorShtXHole - 1
                        skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - 0.025 - (MotorShtXHoleDist * a) - (PanelWth / 1000 * (b - 1))), 0, 0)
                    Next
                    skPoint = Part.SketchManager.CreatePoint((BTLth / 2) - 0.025 - (b * (MotorShtXHoleDist * (MotorShtXHole - 1))) - (0.05 * b), 0, 0)
                Next

                'Blank Holes
                If BlkNosX = 1 Then
                    For a = 1 To SideBlkXHoles(1) - 2
                        skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - (((PanelWth / 1000 * Truncate(FanNoX / 2)))) - (SideBlkXDist(1) * a) - 0.025), 0, 0)
                    Next
                Else
                    For a = 1 To SideBlkXHoles(1) - 2
                        skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - (((PanelWth / 1000 * Truncate(FanNoX / 2)))) - (SideBlkXDist(1) * a) - 0.025), 0, 0)
                    Next

                    For a = 1 To SideBlkXHoles(2) - 1
                        skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - (((PanelWth / 1000 * Truncate(FanNoX / 2)))) - (SideBlkXDist(2) * a) - 0.025), 0, 0)
                    Next

                    'MsgBox("WIP")
                    'For a = 1 To SideBlkXHoles(1) - 1
                    '    skPoint = Part.SketchManager.CreatePoint((PanelWth / 1000 * FanNoX) - (WallWth / 2000) + (SideBlkXDist(1) * a) + 0.025, 0, 0)
                    'Next

                    'For a = 1 To SideBlkXHoles(2) - 1
                    '    skPoint = Part.SketchManager.CreatePoint((PanelWth / 1000 * FanNoX) - (WallWth / 2000) + (SideBlkXDist(1) * a) + 0.025, 0, 0)
                    'Next
                End If


                ' Zoom To Fit
                boolstatus = Part.EditRebuild3()
                Part.ViewZoomtofit2()

                boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)
                ' Save As
                longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02B_Bot_L2.SLDPRT", 0, 0)
                swApp.CloseAllDocuments(True)


                '-x-x-x-xx-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x--x-x-x-xx-x-x-x-
                '------------- Top L2 -----------------
                Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_02A_Top & Bot_L.SLDPRT", 1, 0, "", longstatus, longwarnings)
                swApp.ActivateDoc2("_02A_Top & Bot_L.SLDPRT", False, longstatus)
                Part = swApp.ActiveDoc

                boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_02A_Top & Bot_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                myDimension = Part.Parameter("Width@BaseFlange")
                myDimension.SystemValue = SideClear / 1000 + ((PanelWth / 1000) * Truncate(FanNoX / 2))
                BTLth = SideClear / 1000 + ((PanelWth / 1000) * Truncate(FanNoX / 2))
                Part.ClearSelection2(True)


                TopBotLHoles()                                                     '4.2mm Holes

                boolstatus = Part.Extension.SelectByID2("Slots", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditSketch()

                'Fan Holes
                For b = 1 To Truncate(FanNoX / 2)
                    For a = 1 To MotorShtXHole - 1
                        skPoint = Part.SketchManager.CreatePoint(((-BTLth / 2) + 0.025 + (MotorShtXHoleDist * a) + (PanelWth / 1000 * (b - 1))), 0, 0)
                    Next
                    skPoint = Part.SketchManager.CreatePoint((-BTLth / 2) + 0.025 + (b * (MotorShtXHoleDist * (MotorShtXHole - 1))) + (0.05 * b), 0, 0)

                Next

                'Blank Holes
                If BlkNosX = 1 Then
                    For a = 1 To SideBlkXHoles(1) - 2
                        skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - 0.025 - (((PanelWth / 1000 * Truncate(FanNoX / 2)))) - (SideBlkXDist(1) * a)), 0, 0)
                    Next
                Else
                    For a = 1 To SideBlkXHoles(1) - 2
                        skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - 0.025 - (((PanelWth / 1000 * Truncate(FanNoX / 2)))) - (SideBlkXDist(1) * a)), 0, 0)
                    Next

                    For a = 1 To SideBlkXHoles(2) - 1
                        skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - 0.025 - (((PanelWth / 1000 * Truncate(FanNoX / 2)))) - (SideBlkXDist(2) * a)), 0, 0)
                    Next
                    ' MsgBox("WIP")
                    'For a = 1 To SideBlkXHoles(1) - 1
                    '    skPoint = Part.SketchManager.CreatePoint((PanelWth / 1000 * FanNoX) - (WallWth / 2000) + (SideBlkXDist(1) * a) + 0.025, 0, 0)
                    'Next

                    'For a = 1 To SideBlkXHoles(2) - 1
                    '    skPoint = Part.SketchManager.CreatePoint((PanelWth / 1000 * FanNoX) - (WallWth / 2000) + (SideBlkXDist(1) * a) + 0.025, 0, 0)
                    'Next

                End If


                ' Zoom To Fit
                boolstatus = Part.EditRebuild3()
                Part.ViewZoomtofit2()

                boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)
                ' Save As
                longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02B_Top_L2.SLDPRT", 0, 0)
                swApp.CloseAllDocuments(True)


                '-x-x-x-xx-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x--x-x-x-xx-x-x-x--x-x-x--x-x-x--x-x-x-x-x--x-x-x-x--x-x-x-x-x-x-x-x-x-x-x-x-x-x-
                '------------- Bot L -----------------

                Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_02A_Top & Bot_L.SLDPRT", 1, 0, "", longstatus, longwarnings)
                swApp.ActivateDoc2("_02A_Top & Bot_L.SLDPRT", False, longstatus)
                Part = swApp.ActiveDoc

                boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_02A_Top & Bot_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                myDimension = Part.Parameter("Width@BaseFlange")
                myDimension.SystemValue = (WallWth / 1000) - (SideClear / 1000 + ((PanelWth / 1000) * Truncate(FanNoX / 2)))
                BTLth = (WallWth / 1000) - (SideClear / 1000 + ((PanelWth / 1000) * Truncate(FanNoX / 2)))
                Part.ClearSelection2(True)


                TopBotLHoles()                                                     '4.2mm Holes

                boolstatus = Part.Extension.SelectByID2("Slots", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditSketch()

                'Fan Holes
                For b = 1 To Ceiling(FanNoX / 2)
                    For a = 1 To MotorShtXHole - 1
                        skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - 0.025 - (MotorShtXHoleDist * a) - (PanelWth / 1000 * (b - 1))), 0, 0)
                    Next
                    skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - 0.025 - (b * (MotorShtXHoleDist * (MotorShtXHole - 1))) - (0.05 * b)), 0, 0)
                Next

                'Blank Holes
                If BlkNosX = 1 Then
                    For a = 1 To SideBlkXHoles(1) - 2
                        skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - (((PanelWth / 1000 * Ceiling(FanNoX / 2)))) - (SideBlkXDist(1) * a) - 0.025), 0, 0)
                    Next
                Else
                    For a = 1 To SideBlkXHoles(1) - 2
                        skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - (((PanelWth / 1000 * Ceiling(FanNoX / 2)))) - (SideBlkXDist(1) * a) - 0.025), 0, 0)
                    Next

                    For a = 1 To SideBlkXHoles(2) - 1
                        skPoint = Part.SketchManager.CreatePoint(-((BTLth / 2) - (((PanelWth / 1000 * Ceiling(FanNoX / 2)))) - (SideBlkXDist(2) * a) - 0.025), 0, 0)
                    Next
                    'MsgBox("WIP")
                    'For a = 1 To SideBlkXHoles(1) - 1
                    '    skPoint = Part.SketchManager.CreatePoint((PanelWth / 1000 * FanNoX) - (WallWth / 2000) + (SideBlkXDist(1) * a) + 0.025, 0, 0)
                    'Next

                    'For a = 1 To SideBlkXHoles(2) - 1
                    '    skPoint = Part.SketchManager.CreatePoint((PanelWth / 1000 * FanNoX) - (WallWth / 2000) + (SideBlkXDist(1) * a) + 0.025, 0, 0)
                    'Next
                End If


                ' Zoom To Fit
                boolstatus = Part.EditRebuild3()
                Part.ViewZoomtofit2()

                boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

                ' Save As
                longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02A_Bot_L.SLDPRT", 0, 0)
                swApp.CloseAllDocuments(True)



                '-x-x-x-xx-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x--x-x-x-xx-x-x-x-
                '------------- Top L -----------------

                Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_02A_Top & Bot_L.SLDPRT", 1, 0, "", longstatus, longwarnings)
                swApp.ActivateDoc2("_02A_Top & Bot_L.SLDPRT", False, longstatus)
                Part = swApp.ActiveDoc

                boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_02A_Top & Bot_L.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                myDimension = Part.Parameter("Width@BaseFlange")
                myDimension.SystemValue = (WallWth / 1000) - (SideClear / 1000 + ((PanelWth / 1000) * Truncate(FanNoX / 2)))
                BTLth = (WallWth / 1000) - (SideClear / 1000 + ((PanelWth / 1000) * Truncate(FanNoX / 2)))
                Part.ClearSelection2(True)


                TopBotLHoles()                                                     '4.2mm Holes

                boolstatus = Part.Extension.SelectByID2("Slots", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditSketch()


                'Fan Holes
                For b = 1 To Ceiling(FanNoX / 2)
                    For a = 1 To MotorShtXHole - 1
                        skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - 0.025 - (MotorShtXHoleDist * a) - (PanelWth / 1000 * (b - 1))), 0, 0)
                    Next
                    skPoint = Part.SketchManager.CreatePoint((BTLth / 2) - 0.025 - (b * (MotorShtXHoleDist * (MotorShtXHole - 1))) - (0.05 * b), 0, 0)

                Next

                'Blank Holes
                If BlkNosX = 1 Then
                    For a = 1 To SideBlkXHoles(1) - 2
                        skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - 0.025 - (((PanelWth / 1000 * Ceiling(FanNoX / 2)))) - (SideBlkXDist(1) * a)), 0, 0)
                    Next
                Else
                    For a = 1 To SideBlkXHoles(1) - 2
                        skPoint = Part.SketchManager.CreatePoint(((BTLth / 2) - 0.025 - (((PanelWth / 1000 * Ceiling(FanNoX / 2)))) - (SideBlkXDist(1) * a)), 0, 0)
                    Next
                    'MsgBox("WIP")
                    'For a = 1 To SideBlkXHoles(1) - 1
                    '    skPoint = Part.SketchManager.CreatePoint((PanelWth / 1000 * FanNoX) - (WallWth / 2000) + (SideBlkXDist(1) * a) + 0.025, 0, 0)
                    'Next

                    'For a = 1 To SideBlkXHoles(2) - 1
                    '    skPoint = Part.SketchManager.CreatePoint((PanelWth / 1000 * FanNoX) - (WallWth / 2000) + (SideBlkXDist(1) * a) + 0.025, 0, 0)
                    'Next

                End If

                ' Zoom To Fit
                boolstatus = Part.EditRebuild3()
                Part.ViewZoomtofit2()

                boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)
                ' Save As
                longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02A_Top_L.SLDPRT", 0, 0)
                swApp.CloseAllDocuments(True)

#End Region
            End If

        End If

        End If

#Region "Suppress Edge Flanges"
        If WallWth > (MaxSecLth3mm - 120) Then
            'Suppress Edge Flanges
            Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_02A_Bot_L.SLDPRT", 1, 0, "", longstatus, longwarnings)
            swApp.ActivateDoc2("_02A_Bot_L.SLDPRT", False, longstatus)
            Part = swApp.ActiveDoc
            boolstatus = Part.Extension.SelectByID2("Left Edge", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditSuppress2()
            Part.EditRebuild3()
            longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02A_Bot_L.SLDPRT", 0, 0)
            swApp.CloseAllDocuments(True)


            Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_02A_Top_L.SLDPRT", 1, 0, "", longstatus, longwarnings)
            swApp.ActivateDoc2("_02A_Top_L.SLDPRT", False, longstatus)
            Part = swApp.ActiveDoc
            boolstatus = Part.Extension.SelectByID2("Right Edge", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditSuppress2()
            Part.EditRebuild3()
            longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02A_Top_L.SLDPRT", 0, 0)
            swApp.CloseAllDocuments(True)

            Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_02B_Bot_L2.SLDPRT", 1, 0, "", longstatus, longwarnings)
            swApp.ActivateDoc2("_02B_Bot_L2.SLDPRT", False, longstatus)
            Part = swApp.ActiveDoc
            boolstatus = Part.Extension.SelectByID2("Right Edge", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditSuppress2()
            Part.EditRebuild3()
            longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02B_Bot_L2.SLDPRT", 0, 0)
            swApp.CloseAllDocuments(True)


            Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_02B_Top_L2.SLDPRT", 1, 0, "", longstatus, longwarnings)
            swApp.ActivateDoc2("_02B_Top_L2.SLDPRT", False, longstatus)
            Part = swApp.ActiveDoc
            boolstatus = Part.Extension.SelectByID2("Left Edge", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditSuppress2()
            Part.EditRebuild3()
            longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02B_Top_L2.SLDPRT", 0, 0)
            swApp.CloseAllDocuments(True)

            TopBotLDrawings()
            TopBot2BLDrawings()
            swApp.CloseAllDocuments(True)
        End If
#End Region

    End Sub

    Public Sub TopBotLHoles()

        '4.2mm holes--------------------------------------------------------
        If WallWth <= MaxSecLth3mm Then
            Dim HoleX As Integer = SideHoleNumberTopBotL(WallWth, 40)
            Dim HoleXDist As Decimal = SideHoleDistTopBotL(HoleX, WallWth, 40)

            boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, False, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("4.2mm Hole", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2(HoleX + 1, HoleXDist, 0, 0, True, True, "NULL", "NULL", False)
            Part.ClearSelection2(True)
        Else
            Dim HoleX As Integer = SideHoleNumberTopBotL(WallWth / 2, 40)
            Dim HoleXDist As Decimal = SideHoleDistTopBotL(HoleX, WallWth / 2, 40)

            boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, False, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("4.2mm Hole", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2(HoleX + 1, HoleXDist, 0, 0, True, True, "NULL", "NULL", False)
            Part.ClearSelection2(True)
        End If

    End Sub

#End Region

#Region "Vertical C"
    Public Sub VerCChannels(FanNoY As Integer, FanNoX As Integer, PnlHgt As Decimal, BlkNoY As Integer, TopClear As Decimal, SideClear As Decimal, BlkHt() As Decimal)

        Dim FanY As Integer = FanNoY
        If FanNoX = 1 And SideClear = 0 Then
            Exit Sub
        End If

        'OPEN Vertical Channel -------------------------------------------------------------------------------------
        If WallHt <= (MaxSecLth3mm - 220) Then
            Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_04A_Vertical Channel.SLDPRT", 1, 0, "", longstatus, longwarnings)
            swApp.ActivateDoc2("_04A_Vertical Channel.SLDPRT", False, longstatus)
            Part = swApp.ActiveDoc
            boolstatus = Part.Extension.SelectByID2("Height@BaseFlange@_04A_Vertical Channel.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("Height@BaseFlange")
            myDimension.SystemValue = (WallHt - 4) / 1000
            Part.ClearSelection2(True)

            VerticalCSlots04A(FanY, BlkNoY, PnlHgt, BlkHt)

            ' Zoom To Fit
            boolstatus = Part.EditRebuild3()
            Part.ViewZoomtofit2()
            boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

            ' Save As
            longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_04A_Vertical Channel.SLDPRT", 0, 0)
            VerCDrawings(FanNoX)

            swApp.CloseDoc(True)
        Else

            MsgBox("Wall Height exceeds Maximum Height of the Vertical C" & vbNewLine & "Automation will continue but Vertical C Flat Pattern won't fit on the nesting sheets")

#Region "Two vertical C"

            'Part 1--------------------------------------------------------------------------------------------
            Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_04A_Vertical Channel.SLDPRT", 1, 0, "", longstatus, longwarnings)
            swApp.ActivateDoc2("_04A_Vertical Channel.SLDPRT", False, longstatus)
            Part = swApp.ActiveDoc
            boolstatus = Part.Extension.SelectByID2("Height@BaseFlange@_04A_Vertical Channel.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("Height@BaseFlange")
            myDimension.SystemValue = (WallHt - 4) / 1000
            Part.ClearSelection2(True)

            VerticalCSlots04A(FanY, BlkNoY, PnlHgt, BlkHt)

            'boolstatus = Part.Extension.SelectByID2("Height@Slot 02@_04A_Vertical Channel.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            'myDimension = Part.Parameter("Height@Slot 02")
            'myDimension.SystemValue = PnlHgt / 1000
            'Part.ClearSelection2(True)

            'VerticalCSlots04A(FanY, BlkNoY)

            ' Zoom To Fit
            boolstatus = Part.EditRebuild3()
            Part.ViewZoomtofit2()
            boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

            ' Save As
            longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_04A_Vertical Channel.SLDPRT", 0, 0)
            VerCDrawings(FanNoX)

            swApp.CloseDoc(True)


            ''04A & 04B -- 2 Parts--------------------------------------------------------------------------------------------
            'Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_04A_Vertical Channel.SLDPRT", 1, 0, "", longstatus, longwarnings)
            'swApp.ActivateDoc2("_04A_Vertical Channel.SLDPRT", False, longstatus)
            'Part = swApp.ActiveDoc
            'boolstatus = Part.Extension.SelectByID2("Height@BaseFlange@_04A_Vertical Channel.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            'myDimension = Part.Parameter("Height@BaseFlange")
            'myDimension.SystemValue = (WallHt - (PnlHgt + PnlHgt / 2) - 6) / 1000
            'Part.ClearSelection2(True)


            '' Zoom To Fit
            'boolstatus = Part.EditRebuild3()
            'Part.ViewZoomtofit2()
            'boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

            '' Save As
            'longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_04B_Vertical Channel_02.SLDPRT", 0, 0)
            'VerCDrawings(FanNoX)

            'swApp.CloseDoc(True)
#End Region

        End If
        ' Close Document
        swApp.CloseAllDocuments(True)

    End Sub

    Public Sub VerticalCSlots04A(FansY As Integer, BlkNoY As Integer, PnlHgt As Decimal, BlkHt() As Decimal)

        For a = 0 To BlkNoY - 1
            HoleYBlk(a) = SlotsInterBoltingNumber(BlkHt(a), 25)
            HoleYBlkDist(a) = SlotsBoltDistance(HoleYBlk(a), BlkHt(a), 25)
        Next

        Dim NoOfFans As Integer = Math.Floor(FansY / 2)
        Dim NextFan As Integer = 0
        ' --------------------------------------------Slots---------------------------------------------------------------
        swApp.ActivateDoc2("_04A_Vertical Channel.SLDPRT", False, longstatus)
        Part = swApp.ActiveDoc

        boolstatus = Part.Extension.SelectByID2("slot1", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditSketch()

        boolstatus = Part.Extension.SelectByID2("Dist@slot1@_04A_Vertical Channel.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        myDimension = Part.Parameter("Dist@slot1")
        myDimension.SystemValue = MotorShtYHoleDist

        Part.ClearSelection2(True)
        boolstatus = Part.EditRebuild3()

        swApp.ActivateDoc2("_04A_Vertical Channel.SLDPRT", False, longstatus)
        Part = swApp.ActiveDoc
        boolstatus = Part.Extension.SelectByID2("Slots", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditSketch()

        Dim Plus23 As Decimal = 0.023
        '-------------------------- Blank Y  = 0 ------------------------------------
        If BlkNoY = 0 Then

            If NoOfFans > 0 Then
                For a = 2 To MotorShtYHole + FansY - 1
                    If a = 3 Then
                        skPoint = Part.SketchManager.CreatePoint(-0.025, (MotorShtYHoleDist * a) + Plus23, 0)
                        skPoint = Part.SketchManager.CreatePoint(-0.025, (MotorShtYHoleDist * a) + 0.05 + Plus23, 0)
                    Else
                        If FansY = 1 And a = 4 Then
                            GoTo SaveVerC
                        Else
                            If a > 3 Then
                                skPoint = Part.SketchManager.CreatePoint(-0.025, (MotorShtYHoleDist * a) + 0.05 + Plus23, 0)
                            Else
                                skPoint = Part.SketchManager.CreatePoint(-0.025, (MotorShtYHoleDist * a) + Plus23, 0)
                            End If
                        End If
                    End If
                Next
            Else                                    'Fans = 1
                For a = 2 To MotorShtXHole - 2
                    skPoint = Part.SketchManager.CreatePoint(-0.025, (MotorShtYHoleDist * a) + Plus23, 0)
                Next
            End If

        ElseIf BlkNoY = 1 Then

            '-------------------------- Blank Y = 1 --------------------------------------
            '------ Fans
            If NoOfFans > 0 Then                       'Fans more than 1

                For a = 2 To MotorShtYHole + FansY
                    If a Mod 3 = 0 Then
                        skPoint = Part.SketchManager.CreatePoint(-0.025, (MotorShtYHoleDist * a) + (NextFan * 0.05) + Plus23, 0)

                        If a Mod 3 = 0 Then
                            NextFan = NextFan + 1
                        End If

                        skPoint = Part.SketchManager.CreatePoint(-0.025, (MotorShtYHoleDist * a) + (NextFan * 0.05) + Plus23, 0)
                    Else
                        skPoint = Part.SketchManager.CreatePoint(-0.025, (MotorShtYHoleDist * a) + (NextFan * 0.05) + Plus23, 0)
                    End If
                Next

            Else                                    'Fans = 1
                For a = 2 To MotorShtYHole - 1
                    If a = 3 Then
                        skPoint = Part.SketchManager.CreatePoint(-0.025, (MotorShtYHoleDist * a) + Plus23, 0)
                        skPoint = Part.SketchManager.CreatePoint(-0.025, (MotorShtYHoleDist * a) + 0.05 + Plus23, 0)
                    Else
                        skPoint = Part.SketchManager.CreatePoint(-0.025, (MotorShtYHoleDist * a) + Plus23, 0)
                    End If
                Next
            End If
            '-x-x-x-x-x-x-x-x-x-x
            '------ Blanks
            For a = 1 To HoleYBlk(0) - 2
                skPoint = Part.SketchManager.CreatePoint(-0.025, ((PnlHgt / 1000) * FansY) + (HoleYBlkDist(0) * a) + Plus23, 0)
            Next

        Else

            '-------------------------- Blank Y > 1 --------------------------------------
            '------ Fans

            If NoOfFans > 0 Then                       'Fans more than 1

                For a = 2 To MotorShtYHole + FansY
                    If a Mod 3 = 0 Then
                        skPoint = Part.SketchManager.CreatePoint(-0.025, (MotorShtYHoleDist * a) + (NextFan * 0.05) + Plus23, 0)

                        If a Mod 3 = 0 Then
                            NextFan = NextFan + 1
                        End If

                        skPoint = Part.SketchManager.CreatePoint(-0.025, (MotorShtYHoleDist * a) + (NextFan * 0.05) + Plus23, 0)
                    Else
                        skPoint = Part.SketchManager.CreatePoint(-0.025, (MotorShtYHoleDist * a) + (NextFan * 0.05) + Plus23, 0)
                    End If
                Next

            Else                                    'Fans = 1
                For a = 2 To MotorShtYHole - 1
                    If a = 3 Then
                        skPoint = Part.SketchManager.CreatePoint(-0.025, (MotorShtYHoleDist * a) + Plus23, 0)
                        skPoint = Part.SketchManager.CreatePoint(-0.025, (MotorShtYHoleDist * a) + 0.05 + Plus23, 0)
                    Else
                        skPoint = Part.SketchManager.CreatePoint(-0.025, (MotorShtYHoleDist * a) + Plus23, 0)
                    End If
                Next
            End If

            '-x-x-x-x-x-x-x-x-x-x-
            '------ Blanks
            For a = 1 To HoleYBlk(0) - 1
                If a = HoleYBlk(0) - 1 Then
                    skPoint = Part.SketchManager.CreatePoint(-0.025, ((PnlHgt / 1000) * FansY) + (HoleYBlkDist(0) * a) + Plus23, 0)
                    skPoint = Part.SketchManager.CreatePoint(-0.025, ((PnlHgt / 1000) * FansY) + (HoleYBlkDist(0) * a) + 0.05 + Plus23, 0)
                Else
                    skPoint = Part.SketchManager.CreatePoint(-0.025, ((PnlHgt / 1000) * FansY) + (HoleYBlkDist(0) * a) + Plus23, 0)
                End If
            Next

            For a = 1 To HoleYBlk(1) - 1
                skPoint = Part.SketchManager.CreatePoint(-0.025, ((PnlHgt / 1000) * FansY) + BlkHt(0) / 1000 + (HoleYBlkDist(1) * a) + Plus23, 0)
            Next

            If BlkNoY > 2 Then
                If HoleYBlk(2) = 1 Then
                    skPoint = Part.SketchManager.CreatePoint(-0.025, ((PnlHgt / 1000) * FansY) + BlkHt(0) / 1000 + BlkHt(1) / 1000 + Plus23, 0)
                End If

                For a = 1 To HoleYBlk(2) - 1
                    If a = 1 Then
                        skPoint = Part.SketchManager.CreatePoint(-0.025, ((PnlHgt / 1000) * FansY) + BlkHt(0) / 1000 + BlkHt(1) / 1000 + Plus23, 0)
                    Else
                        skPoint = Part.SketchManager.CreatePoint(-0.025, ((PnlHgt / 1000) * FansY) + BlkHt(0) / 1000 + BlkHt(1) / 1000 + (HoleYBlkDist(2) * a) + Plus23, 0)
                    End If
                Next
            End If

        End If

        Part.ClearSelection2(True)
        boolstatus = Part.EditRebuild3()

        '------------------------------------------------- Side Holes -------------------------------------------------
        If FansY = 1 And BlkNoY <= 1 Then
            boolstatus = Part.Extension.SelectByID2("SideHoles", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditSuppress2()
            Part.ClearSelection2(True)

            GoTo SaveVerC
        End If

        Part.ClearSelection2(True)
        boolstatus = Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("SideHole", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditSketch()

        boolstatus = Part.Extension.SelectByID2("FirstHole@SideHole@_04A_Vertical Channel.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        myDimension = Part.Parameter("FirstHole@SideHole")
        myDimension.SystemValue = (PnlHgt / 1000) - 0.002
        Part.ClearSelection2(True)
        boolstatus = Part.EditRebuild3()

        Part = swApp.ActiveDoc
        boolstatus = Part.Extension.SelectByID2("SideSkt", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditSketch()

        'For Fans
        If FansY > 1 Then
            For a = 2 To FansY
                skPoint = Part.SketchManager.CreatePoint(-0.013, (((PnlHgt / 1000) * a) - 0.002), 0)
            Next
        End If

        'For Blanks
        If BlkNoY > 1 Then
            If BlkHt(0) = 400 Then
                skPoint = Part.SketchManager.CreatePoint(-0.013, (((PnlHgt / 1000) * FansY) + (BlkHt(0) / 1000) - 0.002), 0)
            Else
                For a = 1 To BlkNoY - 1
                    skPoint = Part.SketchManager.CreatePoint(-0.013, (((PnlHgt / 1000) * FansY) + ((BlkHt(0) / 1000) * a) - 0.002), 0)
                Next
            End If

        End If

        Part.ClearSelection2(True)
        boolstatus = Part.EditRebuild3()
        Part.ViewZoomtofit2()

SaveVerC:
        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_04A_Vertical Channel.SLDPRT", 0, 0)

    End Sub

    Public Sub VerticalCSlots04B(FanNoY As Integer, BlkNoY As Integer)
        Exit Sub

    End Sub

#End Region

#Region "Horizontal C"
    Public Sub HorCChannel(PnlWth As Decimal, FanNoY As Integer, BlkNosY As Integer, FanNoX As Integer)

        If BlkNosY >= 1 Or FanNoY > 1 Then
            Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_06_Hor_C_Support.SLDPRT", 1, 0, "", longstatus, longwarnings)
            swApp.ActivateDoc2("_06_Hor_C_Support.SLDPRT", False, longstatus)
            Part = swApp.ActiveDoc
            boolstatus = Part.Extension.SelectByID2("Width@Base@_06_Hor_C_Support.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("Width@Base")
            myDimension.SystemValue = (PnlWth - 1.01) / 10
            Part.ClearSelection2(True)

            'Re-Dim
            boolstatus = Part.Extension.SelectByID2("Dist@HorCHoles@_06_Hor_C_Support.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("Dist@HorCHoles")
            myDimension.SystemValue = MotorShtXHoleDist
            Part.ClearSelection2(True)

            Part.ViewZoomtofit2()
            boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)


            ' Save As
            longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_05_Hor_C_Support.SLDPRT", 0, 0)
            swApp.CloseAllDocuments(True)

            Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_05_Hor_C_Support.SLDPRT", 1, 0, "", longstatus, longwarnings)

            Dim DwgQty As Integer
            If BlkNosY > 1 Then
                DwgQty = FanNoX * FanNoY + (BlkNosY / 2)
            Else
                DwgQty = FanNoX * FanNoY
            End If

            CreateDrawings(AHUName & "_05_Hor_C_Support", DwgQty, (PnlWth - 1.01) / 10)
            swApp.CloseAllDocuments(True)
        End If
    End Sub

    Public Sub HorCChannel2(Sdclr As Decimal, FanNoY As Integer, BlkNosY As Integer, FanNoX As Integer, PushSide As Boolean, BlkNosX As Integer, AHUDoor As Boolean)

        Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_06A_Hor_C_Support_2.SLDPRT", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("_06A_Hor_C_Support_2.SLDPRT", False, longstatus)
        Part = swApp.ActiveDoc

        boolstatus = Part.Extension.SelectByID2("Width@Base@_06_Hor_C_Support_2.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        myDimension = Part.Parameter("Width@Base")
        myDimension.SystemValue = Sdclr
        Part.ClearSelection2(True)

        If AHUDoor = True Then
            If DoorXHoles > 3 Then
                'Re-Dim
                boolstatus = Part.Extension.SelectByID2("Dist@HorCHoles@_06_Hor_C_Support.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                myDimension = Part.Parameter("Dist@HorCHoles")
                myDimension.SystemValue = DoorXHoleDist
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("SlotPattern", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditSuppress2()
                Part.ClearSelection2(True)
            Else
                boolstatus = Part.Extension.SelectByID2("HorCSlots", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditSuppress2()
                Part.ClearSelection2(True)
            End If
        Else
            If BlkNosX >= 1 Then
                If SideBlkXHoles(BlkNosX) > 3 Then
                    'Re-Dim
                    boolstatus = Part.Extension.SelectByID2("Dist@HorCHoles@_06_Hor_C_Support.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                    myDimension = Part.Parameter("Dist@HorCHoles")
                    myDimension.SystemValue = SideBlkXDist(BlkNosX)
                    Part.ClearSelection2(True)

                    boolstatus = Part.Extension.SelectByID2("SlotPattern", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
                    Part.EditSuppress2()
                    Part.ClearSelection2(True)
                Else
                    boolstatus = Part.Extension.SelectByID2("HorCSlots", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
                    Part.EditSuppress2()
                    Part.ClearSelection2(True)
                End If
            End If
        End If

        Part.ViewZoomtofit2()
        boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)


        Part.ClearSelection2(True)
        Part.ViewZoomtofit2()

        ' Save As
        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_05A_Hor_C_Support_2.SLDPRT", 0, 0)
        swApp.CloseAllDocuments(True)

        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_05A_Hor_C_Support_2.SLDPRT", 1, 0, "", longstatus, longwarnings)

        Dim DwgQty As Integer
        If AHUDoor = True Then
            DwgQty = 1
        Else
            If PushSide = True Then
                DwgQty = FanNoY + BlkNosY - 1
            Else
                DwgQty = (FanNoY + BlkNosY - 1) * 2
            End If
        End If

        CreateDrawings(AHUName & "_05A_Hor_C_Support_2", DwgQty, Sdclr)
        swApp.CloseAllDocuments(True)

    End Sub

    Public Sub HorCChannel3(EdBlkWth As Decimal, FanNoY As Integer, BlkNosY As Integer, FanNoX As Integer, BlkNosX As Integer)

        Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "_06A_Hor_C_Support_2.SLDPRT", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("_06A_Hor_C_Support_2.SLDPRT", False, longstatus)
        Part = swApp.ActiveDoc

        boolstatus = Part.Extension.SelectByID2("Width@Base@_06_Hor_C_Support_2.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        myDimension = Part.Parameter("Width@Base")
        myDimension.SystemValue = EdBlkWth
        Part.ClearSelection2(True)

        'Re-Dim
        If BlkNosX > 1 Then
            Dim d As Integer
            If BlkNosY > 1 Then
                d = 3
            Else
                d = 2
            End If

            If SideBlkXHoles(d) > 3 Then
                'Re-Dim
                boolstatus = Part.Extension.SelectByID2("Dist@HorCHoles@_06_Hor_C_Support.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                myDimension = Part.Parameter("Dist@HorCHoles")
                myDimension.SystemValue = SideBlkXDist(BlkNosX)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("SlotPattern", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditSuppress2()
                Part.ClearSelection2(True)
            Else
                boolstatus = Part.Extension.SelectByID2("HorCSlots", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditSuppress2()
                Part.ClearSelection2(True)
            End If
        Else
            boolstatus = Part.Extension.SelectByID2("Dist@HorCHoles@_06_Hor_C_Support.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("Dist@HorCHoles")
            myDimension.SystemValue = MotorShtXHoleDist
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("SlotPattern", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditSuppress2()
            Part.ClearSelection2(True)
        End If

        Part.ViewZoomtofit2()
        boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)


        Part.ClearSelection2(True)
        Part.ViewZoomtofit2()

        ' Save As
        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_05B_Hor_C_Support_3.SLDPRT", 0, 0)
        swApp.CloseAllDocuments(True)

        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_05B_Hor_C_Support_3.SLDPRT", 1, 0, "", longstatus, longwarnings)

        Dim DwgQty As Integer
        DwgQty = FanNoY + BlkNosY - 1

        CreateDrawings(AHUName & "_05B_Hor_C_Support_3", DwgQty, EdBlkWth)
        swApp.CloseAllDocuments(True)

    End Sub



#End Region

#Region "Base Stand"
    Public Sub BaseStand()
        Part = swApp.OpenDoc6(LibPath & "\Panel - Sample\" & "Base_Stand.SLDPRT", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("Base_Stand.SLDPRT", False, longstatus)
        Part = swApp.ActiveDoc

        Dim BaseStdLth As Decimal
        Dim BaseQty As Integer

        If WallWth <= MaxSecLth3mm Then
            boolstatus = Part.Extension.SelectByID2("Dist@Base@Base_Stand.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("Dist@Base")
            myDimension.SystemValue = WallWth / 1000
            BaseStdLth = WallWth / 1000
            BaseQty = 1
            Part.ClearSelection2(True)
        Else
            boolstatus = Part.Extension.SelectByID2("Dist@Base@Base_Stand.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("Dist@Base")
            myDimension.SystemValue = WallWth / 2000
            BaseStdLth = WallWth / 2000
            BaseQty = 2
            Part.ClearSelection2(True)
        End If

        Part.ViewZoomtofit2()
        boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

        ' Save As
        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_10_Base_Stand.SLDPRT", 0, 0)
        swApp.CloseAllDocuments(True)

        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_10_Base_Stand.SLDPRT", 1, 0, "", longstatus, longwarnings)
        BaseStandDrawings(AHUName & "_10_Base_Stand", BaseQty, BaseStdLth)
        swApp.CloseAllDocuments(True)

    End Sub

#End Region

#End Region

    Public Sub Assembly(FanNos As Integer, FanNoX As Integer, FanNoY As Integer, SideClear As Decimal, TopClear As Decimal, PnlWth As Decimal, PnlHt As Decimal, BlkNosX As Integer,
                        BlkNosY As Integer, PushSide As Boolean, BlkOne As Decimal, BlkTwo As Decimal, BlkThree As Decimal, ClientName As String, JobNum As String, AHUDoor As Boolean,
                        fanArticleNo As String, FanDia As Integer, EndBlankWth As Decimal, MotorDia As Integer, ArtNo As String, DoorBlkWth As Decimal)

        ListofAllParts()

        'Variables
        Dim FittedFans As Integer = FanNoX * FanNoY
        Dim RemFans As Integer = FittedFans - FanNos

        SideLDist = WallWth / 2000

        'Open / Insert Parts
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_02A_Bot_L.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_02A_Top_L.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_02B_Bot_L2.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_02B_Top_L2.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_03A_Side_L_Left.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_03A_Side_L_Right.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_03B_Side_L2_Left.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_03B_Side_L2_Right.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_04A_Vertical Channel.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_04B_Vertical Channel_02.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_05_Hor_C_Support.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_05A_Hor_C_Support_2.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_05B_Hor_C_Support_3.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_06_BlankOff.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_A1.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_A2.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_B1.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_B2.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_08_Top Blank_1.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_08_Top Blank_2.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_08_Top Blank_3.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_09_Corner Blank_1.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_09_Corner Blank_2.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_09_Corner Blank_3.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_09_Corner Blank_4.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_Door SubAssembly.SLDASM", 2, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_10_Base_Stand.SLDPRT", 1, 32, "", longstatus, longwarnings)

        For f = 0 To FanNos
            Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_Motor Sub Assembly.SLDASM", 2, 0, "", longstatus, longwarnings)
        Next

        'New Assy File -----------------------------------------------------------------------------------------
        Dim assyTemp As String = swApp.GetUserPreferenceStringValue(9)
        Part = swApp.NewDocument(assyTemp, 0, 0, 0)
        Assy = Part

        'Insert Parts--------------------------------------------------------------------------------------------------------------
        InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_02A_Bot_L.SLDPRT", 0, "", False, "", 1, 1, 1)
        InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_02A_Top_L.SLDPRT", 0, "", False, "", 1.3, 1.3, 1.3)
        InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_02B_Bot_L2.SLDPRT", 0, "", False, "", 1.5, 1.5, 1.5)
        InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_02B_Top_L2.SLDPRT", 0, "", False, "", 1.7, 1.7, 1.7)
        InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_03A_Side_L_Left.SLDPRT", 0, "", False, "", 1.9, 1.9, 1.9)
        InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_03A_Side_L_Right.SLDPRT", 0, "", False, "", 2.1, 2.1, 2.1)
        InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_03B_Side_L2_Left.SLDPRT", 0, "", False, "", 2.3, 2.3, 2.3)
        InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_03B_Side_L2_Right.SLDPRT", 0, "", False, "", 2.5, 2.5, 2.5)
        InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_Motor Sub Assembly.SLDASM", 0, "", False, "", 3.0, 3.0, 3.0)
        InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_04A_Vertical Channel.SLDPRT", 0, "", False, "", 3.5, 3.5, 3.5)
        InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_04B_Vertical Channel_02.SLDPRT", 0, "", False, "", 3.7, 3.7, 3.7)
        InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_08_Top Blank_1.SLDPRT", 0, "", False, "", 4, 4, 4)
        InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_08_Top Blank_2.SLDPRT", 0, "", False, "", 0.5, 0.5, 0.5)
        InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_08_Top Blank_3.SLDPRT", 0, "", False, "", 0.5, 0.5, 0.5)
        InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_10_Base_Stand.SLDPRT", 0, "", False, "", 0.5, 1.5, 2.5)

        If PushSide = True Then
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_09_Corner Blank_1.SLDPRT", 0, "", False, "", -0.2, -0.2, -0.2)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_09_Corner Blank_2.SLDPRT", 0, "", False, "", -0.8, -0.8, -0.8)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_09_Corner Blank_3.SLDPRT", 0, "", False, "", -0.8, -0.8, -0.8)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_09_Corner Blank_4.SLDPRT", 0, "", False, "", -0.8, -0.8, -0.8)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_A1.SLDPRT", 0, "", False, "", 2.8, 2.8, 2.8)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_B1.SLDPRT", 0, "", False, "", 2.8, 2.8, 2.8)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_05B_Hor_C_Support_3.SLDPRT", 0, "", False, "", 2.5, 2.3, 2.2)
        Else
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_09_Corner Blank_1.SLDPRT", 0, "", False, "", -0.2, -0.2, -0.2)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_09_Corner Blank_1.SLDPRT", 0, "", False, "", -0.4, -0.4, -0.4)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_09_Corner Blank_2.SLDPRT", 0, "", False, "", -0.8, -0.8, -0.8)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_09_Corner Blank_2.SLDPRT", 0, "", False, "", -0.5, -0.5, -0.5)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_09_Corner Blank_3.SLDPRT", 0, "", False, "", -0.7, -0.7, -0.7)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_09_Corner Blank_3.SLDPRT", 0, "", False, "", -0.8, -0.8, -0.8)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_09_Corner Blank_4.SLDPRT", 0, "", False, "", -0.8, -0.8, -0.8)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_09_Corner Blank_4.SLDPRT", 0, "", False, "", -0.8, -0.8, -0.8)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_A1.SLDPRT", 0, "", False, "", 2.8, 2.8, 2.8)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_A1.SLDPRT", 0, "", False, "", 3.2, 3.2, 3.2)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_A2.SLDPRT", 0, "", False, "", 3.5, 3.5, 3.5)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_A2.SLDPRT", 0, "", False, "", 3.7, 3.7, 3.7)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_B1.SLDPRT", 0, "", False, "", 2.8, 2.8, 2.8)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_B1.SLDPRT", 0, "", False, "", 3.2, 3.2, 3.2)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_B2.SLDPRT", 0, "", False, "", 3.5, 3.5, 3.5)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_07_Side Blank_B2.SLDPRT", 0, "", False, "", 3.7, 3.7, 3.7)
        End If


        If BlkOne = 400 Then
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_05_Hor_C_Support.SLDPRT", 0, "", False, "", 3.9, 3.9, 3.9)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_05_Hor_C_Support.SLDPRT", 0, "", False, "", 4.5, 4.5, 3)
        Else
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_05_Hor_C_Support.SLDPRT", 0, "", False, "", 3.9, 3.9, 3.9)
        End If

        'If (FanNoY + BlkNosY - 1) > 0 Then
        '    For a = 1 To (FanNoY + BlkNosY - 1)
        '        InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_05A_Hor_C_Support_2.SLDPRT", 0, "", False, "", 2.2, 2.2, 2.2)
        '    Next
        '    'InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_05A_Hor_C_Support_2.SLDPRT", 0, "", False, "", 2.3, 2.3, 2.3)
        'End If

        If BlkNosY = 0 Then
            If FanNoY > 1 Then
                InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_05A_Hor_C_Support_2.SLDPRT", 0, "", False, "", 2.2, 2.2, 2.2)
                InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_05A_Hor_C_Support_2.SLDPRT", 0, "", False, "", 2.3, 2.3, 2.2)
            End If
        ElseIf BlkNosY = 1 Then
            If FanNoY >= 1 Then
                InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_05A_Hor_C_Support_2.SLDPRT", 0, "", False, "", 2.2, 2.2, 2.2)
                InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_05A_Hor_C_Support_2.SLDPRT", 0, "", False, "", 2.3, 2.3, 2.2)
            End If
        Else
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_05A_Hor_C_Support_2.SLDPRT", 0, "", False, "", 2.2, 2.2, 2.2)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_05A_Hor_C_Support_2.SLDPRT", 0, "", False, "", 2.3, 2.3, 2.2)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_05A_Hor_C_Support_2.SLDPRT", 0, "", False, "", 2.4, 2.2, 2.2)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_05A_Hor_C_Support_2.SLDPRT", 0, "", False, "", 2.5, 2.3, 2.2)
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_05B_Hor_C_Support_3.SLDPRT", 0, "", False, "", 2.5, 2.3, 2.2)
        End If

        If AHUDoor = True Then
            InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_Door SubAssembly.SLDASM", 0, "", False, "", 1.9, 1.9, 1.9)

            If BlkNosY = 1 Then
                InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_05B_Hor_C_Support_3.SLDPRT", 0, "", False, "", 2.5, 2.3, 2.2)
            End If

            'delete extra parts
            boolstatus = Part.Extension.SelectByID2(AHUName & "_07_Side Blank_A1-2@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(AHUName & "_07_Side Blank_B1-2@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(AHUName & "_09_Corner Blank_1-2@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(AHUName & "_09_Corner Blank_2-2@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(AHUName & "_09_Corner Blank_3-2@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(AHUName & "_09_Corner Blank_4-2@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            Part.EditDelete()

            For a = 2 To 6
                boolstatus = Part.Extension.SelectByID2(AHUName & "_05A_Hor_C_Support_2-" & a & "@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2(AHUName & "_05_Hor_C_Support-" & a + 1 & "@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                Part.EditDelete()
            Next
        End If

        If FittedFans > FanNos Then
            For d = 1 To RemFans
                InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_06_BlankOff.SLDPRT", 0, "", False, "", 1.6, 1.6, 1.6)
            Next
        End If

        Assy = Part
        Assy.ViewZoomtofit2()

        ' Save As----------------------------------------------------------------------------------
        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_AHU Final Assembly.SLDASM", 0, 0)
        swApp.CloseAllDocuments(True)

        ' Open Assy--------------------------------------------------------------------------------
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_AHU Final Assembly.SLDASM", 2, 32, "", longstatus, longwarnings)
        Assy = Part
        Assy.ViewZoomtofit2()

        ' Create Axis------------------------------------------------------------------------------
        boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        boolstatus = Part.InsertAxis2(True)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Axis1", "AXIS", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.SelectedFeatureProperties(0, 0, 0, 0, 0, 0, 0, 1, 0, "X-Axis")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        boolstatus = Part.InsertAxis2(True)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Axis2", "AXIS", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.SelectedFeatureProperties(0, 0, 0, 0, 0, 0, 0, 1, 0, "Y-Axis")
        Part.ClearSelection2(True)

        boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

        ' Assemble Parts --------------------------------------------------------------------------
        'Bot L
        Part = swApp.ActiveDoc
        boolstatus = Part.Extension.SelectByID2(AHUName & "_02A_Bot_L-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        Part.UnFixComponent
        Part.ClearSelection2(True)

        '-------------------------------------------------------------------- Bot L Assembly ---------------------------------------------------------------------------------

        If AssyComponents.Contains(AHUName & "_02B_Bot_L2.SLDPRT") Then
            If AHUDoor = True Then

                'IF AHUDOOR TRUE

                If FanNoX Mod 2 = 0 Then
                    BTL2 = DoorBlkWth + (PnlWth * Truncate(FanNoX / 2))                          '02B
                    BTL1 = (WallWth / 1000) - (DoorBlkWth + (PnlWth * Truncate(FanNoX / 2)))     '02A 
                Else
                    BTL2 = DoorBlkWth + (PnlWth * Truncate(FanNoX / 2))                          '02B
                    BTL1 = (WallWth / 1000) - (DoorBlkWth + (PnlWth * Truncate(FanNoX / 2)))     '02A 
                End If

                Part = swApp.ActiveDoc
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_02A_Bot_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, True, BTL2 / 2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane coincidence 
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_02A_Bot_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane coincidence 
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_02A_Bot_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)

                ''------------------------------------------------------------- Add 2nd Bot L -------------------------------------------------------------------------------
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_02B_Bot_L2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, BTL1 / 2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane coincidence 
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_02B_Bot_L2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane coincidence 
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_02B_Bot_L2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)

                Part.EditRebuild3()

                boolstatus = Part.Extension.SelectByID2(AHUName & "_07_Side Blank_B1-2@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditDelete()

            Else
                If FanNoX Mod 2 = 0 Then
                    BTLth = (WallWth / 1000) - (SideClear + (PnlWth * (FanNoX / 2)))
                Else
                    BTLth = (WallWth / 1000) - (SideClear + (PnlWth * Truncate(FanNoX / 2)))
                End If


                Part = swApp.ActiveDoc
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_02A_Bot_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, True, (WallWth / 2000) - (BTLth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane coincidence 
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_02A_Bot_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane coincidence 
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_02A_Bot_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)


                ''------------------------------------------------------------- Add 2nd Bot L -------------------------------------------------------------------------------

                If FanNoX Mod 2 = 0 Then
                    BTLth = (WallWth / 1000) - (SideClear + (PnlWth * (FanNoX / 2)))
                Else
                    BTLth = SideClear + (PnlWth * Truncate(FanNoX / 2))
                End If

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_02B_Bot_L2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, (WallWth / 2000) - (BTLth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane coincidence 
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_02B_Bot_L2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane coincidence 
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_02B_Bot_L2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)

                Part.EditRebuild3()

            End If

            GoTo MtrAssem
        End If

        If AssyComponents.Contains(AHUName & "_02B_Bot_L2.SLDPRT") Then
            'Part = swApp.ActiveDoc
            'boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_02A_Bot_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            'boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            'myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, True, WallWth / 4000, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane coincidence 
            'Part.ClearSelection2(True)

            'boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_02A_Bot_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            'boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            'myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane coincidence 
            'Part.ClearSelection2(True)

            'boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_02A_Bot_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            'boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            'myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
            'Part.ClearSelection2(True)

            '''------------------------------------------------------------- Add 2nd Bot L -------------------------------------------------------------------------------
            'boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_02B_Bot_L2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            'boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            'myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, WallWth / 4000, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane coincidence 
            'Part.ClearSelection2(True)

            'boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_02B_Bot_L2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            'boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            'myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane coincidence 
            'Part.ClearSelection2(True)

            'boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_02B_Bot_L2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            'boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            'myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
            'Part.ClearSelection2(True)

            'Part.EditRebuild3()
        Else
            boolstatus = Part.Extension.SelectByID2("Point1@Origin@" & AHUName & "_02A_Bot_L-1@" & AHUName & "_AHU Final Assembly", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Point1@Origin@" & AHUName & "_AHU Final Assembly", "EXTSKETCHPOINT", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            'L and Assembly Mates
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_02A_Bot_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane coincidence 
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_02A_Bot_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane coincidence 
            Part.ClearSelection2(True)

            Part.EditRebuild3()
        End If

        '-------------------------------------------------------------------- Motor Assembly -------------------------------------------------------------------------------
MtrAssem:
        Part = swApp.ActiveDoc

        If AHUDoor = True Then
            If PushSide = False Then
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_02A_Top & Bot_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                If (WallWth / 2000) - (DoorBlkWth + (PnlWth / 2)) < 0 Then
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, True, Math.Abs((WallWth / 2000 - (DoorBlkWth + (PnlWth / 2)))), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)

                Else
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, ((WallWth / 2000 - (DoorBlkWth + (PnlWth / 2)))), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                End If
                Part.ClearSelection2(True)

                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, (PnlHt / 2) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                    Part.ClearSelection2(True)

                Else
                    boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_02A_Top & Bot_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, True, WallWth / 2000 - PnlWth / 2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, (PnlHt / 2) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)
            End If

            Part.EditRebuild3()
            GoTo FansPattern
        End If

        If PushSide = True Then
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_02A_Top & Bot_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, True, WallWth / 2000 - PnlWth / 2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, (PnlHt / 2) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

        Else
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_02A_Top & Bot_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, (WallWth / 2000) - SideClear - (PnlWth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, (PnlHt / 2) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

        End If

        '------------------------------------------------------------- Remaining Fans ---------------------------------------------------------------------------------
FansPattern:
        If PushSide = True Then
            boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, True, 4, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX, PnlWth, FanNoY, PnlHt, False, True, "NULL", "NULL", False)
            Part.ClearSelection2(True)
            Part.ViewZoomtofit2()
        Else
            boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, True, 4, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX, PnlWth, FanNoY, PnlHt, True, True, "NULL", "NULL", False)
            Part.ClearSelection2(True)
            Part.ViewZoomtofit2()
        End If

        'Delete Extra Fans
        While FittedFans > FanNos
            boolstatus = Part.Extension.SelectByID2(AHUName & "_Motor Sub Assembly-" & FittedFans & "@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
            FittedFans -= 1
        End While

        FittedFans = FanNoX * FanNoY
        RemFans = (FanNoX * FanNoY) - FanNos

        If FittedFans > FanNos Then
            For d = 1 To RemFans
                'Mates
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_06_BlankOff-" & d & "@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_Motor Sub Assembly-" & FanNos & "@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_06_BlankOff-" & d & "@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_Motor Sub Assembly-" & FanNos & "@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, True, PnlWth * (d - 1), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)   '-----------Right Plane coincidence 
                Part.ClearSelection2(True)

                If PushSide = True Then
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_06_BlankOff-" & d & "@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_Motor Sub Assembly-" & FanNos & "@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, True, PnlHt * (FanNoY - 1), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane Distance = (780/2 - 0.025)
                    Part.ClearSelection2(True)
                Else
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_06_BlankOff-" & d & "@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_Motor Sub Assembly-" & FanNos & "@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, True, PnlHt * (FanNoY - 1), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane Distance = (780/2 - 0.025)
                    Part.ClearSelection2(True)
                End If

                Part.EditRebuild3()
            Next
        End If                                   'Assemble BlankOff


        'Hide Parts
        If FanNoY > 1 Then
            For a = 1 To FanNos
                If a Mod 2 = 0 Then
                    boolstatus = Part.Extension.SelectByID2(AHUName & "_Motor Sub Assembly-" & a & "@" & AHUName & "_AHU Final Assembly/" & AHUName & "_" & fanArticleNo & "-1@" & AHUName & "_Motor Sub Assembly/" & FanDia & "mm_C Channel_2-2@" & AHUName & "_" & fanArticleNo, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2(AHUName & "_Motor Sub Assembly-" & a & "@" & AHUName & "_AHU Final Assembly/" & AHUName & "_" & fanArticleNo & "-1@" & AHUName & "_Motor Sub Assembly/" & FanDia & "mm_C Channel_2-1@" & AHUName & "_" & fanArticleNo, "COMPONENT", 0, 0, 0, True, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2(AHUName & "_Motor Sub Assembly-" & a & "@" & AHUName & "_AHU Final Assembly/" & AHUName & "_" & fanArticleNo & "-1@" & AHUName & "_Motor Sub Assembly/" & FanDia & "mm_10_C Channel-2@" & AHUName & "_" & fanArticleNo, "COMPONENT", 0, 0, 0, True, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2(AHUName & "_Motor Sub Assembly-" & a & "@" & AHUName & "_AHU Final Assembly/" & AHUName & "_" & fanArticleNo & "-1@" & AHUName & "_Motor Sub Assembly/" & FanDia & "mm_10_C Channel-1@" & AHUName & "_" & fanArticleNo, "COMPONENT", 0, 0, 0, True, 0, Nothing, 0)
                    Part.HideComponent2()
                End If
            Next
        Else
            For a = 1 To FanNos
                boolstatus = Part.Extension.SelectByID2(AHUName & "_Motor Sub Assembly-" & a & "@" & AHUName & "_AHU Final Assembly/" & AHUName & "_" & fanArticleNo & "-1@" & AHUName & "_Motor Sub Assembly/" & FanDia & "mm_C Channel_2-2@" & AHUName & "_" & fanArticleNo, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2(AHUName & "_Motor Sub Assembly-" & a & "@" & AHUName & "_AHU Final Assembly/" & AHUName & "_" & fanArticleNo & "-1@" & AHUName & "_Motor Sub Assembly/" & FanDia & "mm_C Channel_2-1@" & AHUName & "_" & fanArticleNo, "COMPONENT", 0, 0, 0, True, 0, Nothing, 0)
                Part.HideComponent2()
            Next
        End If

        Part.ClearSelection2(True)
        Part.EditRebuild3()

        '-------------------------------------------------------------------- Horizontal C Channel ----------------------------------------------------------------------

        If AssyComponents.Contains(AHUName & "_05_Hor_C_Support.SLDPRT") Then

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_05_Hor_C_Support-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            If PushSide = True Then
                If FanNoY > 1 Then
                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_05_Hor_C_Support-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_Motor Sub Assembly-3@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                    Part.ClearSelection2(True)
                Else
                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_05_Hor_C_Support-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_Motor Sub Assembly-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                    Part.ClearSelection2(True)
                End If
            Else
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_05_Hor_C_Support-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)
            End If

            If PushSide = True Then
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_05_Hor_C_Support-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, PnlHt - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                If TopClear < 0 Then
                    boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, True, 4, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2(AHUName & "_05_Hor_C_Support-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                    myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX, PnlWth, FanNoY, PnlHt, False, True, "NULL", "NULL", False)
                    Part.ClearSelection2(True)
                    Part.ViewZoomtofit2()
                ElseIf FanNoY > 1 Then
                    boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, True, 4, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2(AHUName & "_05_Hor_C_Support-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                    myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX - 1, PnlWth, FanNoY + BlkNosY - 1, PnlHt, False, True, "NULL", "NULL", False)
                    Part.ClearSelection2(True)
                    Part.ViewZoomtofit2()
                ElseIf BlkNosY > 1 Then
                    boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, True, 4, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2(AHUName & "_05_Hor_C_Support-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                    If BlkOne = 400 Then
                        myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX - 1, PnlWth, FanNoY + BlkNosY - 1, 0.4, False, True, "NULL", "NULL", False)
                    Else
                        myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX - 1, PnlWth, FanNoY + BlkNosY - 1, PnlHt, False, True, "NULL", "NULL", False)
                    End If
                    Part.ClearSelection2(True)
                    Part.ViewZoomtofit2()
                Else
                    boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, True, 4, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2(AHUName & "_05_Hor_C_Support-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                    myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX, PnlWth, FanNoY + BlkNosY - 1, PnlHt / 2, False, True, "NULL", "NULL", False)
                    Part.ClearSelection2(True)
                    Part.ViewZoomtofit2()
                End If

            Else
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_05_Hor_C_Support-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, PnlHt - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                If BlkOne = 400 Then
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_05_Hor_C_Support-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, (PnlHt * FanNoY) + 0.4 - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                    Part.ClearSelection2(True)

                    boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, True, 2, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2(AHUName & "_05_Hor_C_Support-2@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                    myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX, PnlWth, 0, 0, True, True, "NULL", "NULL", False)
                    Part.ClearSelection2(True)
                    Part.ViewZoomtofit2()
                End If

                If TopClear < 0 Then
                    boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, True, 4, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2(AHUName & "_05_Hor_C_Support-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                    myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX, PnlWth, FanNoY, PnlHt, True, True, "NULL", "NULL", False)
                    Part.ClearSelection2(True)
                    Part.ViewZoomtofit2()
                ElseIf FanNoY > 1 Then
                    boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, True, 4, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2(AHUName & "_05_Hor_C_Support-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                    If BlkOne = 400 Then
                        myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX, PnlWth, FanNoY + BlkNosY - 2, PnlHt, True, True, "NULL", "NULL", False)
                    Else
                        myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX, PnlWth, FanNoY + BlkNosY - 1, PnlHt, True, True, "NULL", "NULL", False)
                    End If
                    Part.ClearSelection2(True)
                    Part.ViewZoomtofit2()
                Else
                    boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, True, 4, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2(AHUName & "_05_Hor_C_Support-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                    myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX, PnlWth, FanNoY + BlkNosY - 1, BlkOne / 1000, True, True, "NULL", "NULL", False)
                    Part.ClearSelection2(True)
                    Part.ViewZoomtofit2()
                End If

            End If

            If BlkOne = 400 Then
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_05_Hor_C_Support-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_05_Hor_C_Support-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

            End If

            Part.EditRebuild3()
        End If

        '------------------------------------------------------------------ DOOR Mates-----------------------------------------------------------------------------------

        If AHUDoor = True Then

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_Door SubAssembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_Door SubAssembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (WallWth / 2000) - (SideClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("DoorBlkBottom@" & AHUName & "_Door SubAssembly-1@" & AHUName & "_AHU Final Assembly/" & AHUName & "_11_Door Blank-1@" & AHUName & "_Door SubAssembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_02A_Bot_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignANTI_ALIGNED, True, 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Top Plane distance 
            Part.ClearSelection2(True)

            Part.EditRebuild3()
        End If

        '------------------------------------------------------------------SIDE BLANK Mates-----------------------------------------------------------------------------------

        If PushSide = True Then

            If AHUDoor = True Then
                'Mates LEFT SIDE
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_07_Side Blank_A1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt / 2) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_07_Side Blank_A1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_07_Side Blank_A1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (WallWth / 2000) - (EndBlankWth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                Part.ClearSelection2(True)

                Part.EditRebuild3()

                'Fans Pattern
                boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2(AHUName & "_07_Side Blank_A1-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoY, PnlHt, 0, 0, False, False, "NULL", "NULL", False)
                Part.ClearSelection2(True)
                Part.ViewZoomtofit2()


                GoTo TopBlankAssem
            End If

            'ONE SIDE BLANK Mates
            'Mates LEFT SIDE
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_07_Side Blank_A1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt / 2) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_07_Side Blank_A1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_07_Side Blank_A1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (WallWth / 2000) - (SideClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
            Part.ClearSelection2(True)

            Part.EditRebuild3()

            'SideBlank B
            If AssyComponents.Contains(AHUName & "_07_Side Blank_B1.SLDPRT") Then
                'Mates LEFT SIDE
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_07_Side Blank_B1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, PnlHt + (PnlHt / 2) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_07_Side Blank_B1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_07_Side Blank_B1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (WallWth / 2000) - (SideClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                Part.ClearSelection2(True)

                Part.EditRebuild3()
            End If

            GoTo TopBlankAssem
        Else

            If AHUDoor = True Then
                'Mates LEFT SIDE
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_07_Side Blank_A1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt / 2) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_07_Side Blank_A1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_07_Side Blank_A1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (WallWth / 2000) - (EndBlankWth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                Part.ClearSelection2(True)

                Part.EditRebuild3()

                ''Fans Pattern
                'boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                'boolstatus = Part.Extension.SelectByID2(AHUName & "_07_Side Blank_A1-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                'myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoY, PnlHt, 0, 0, True, False, "NULL", "NULL", False)
                'Part.ClearSelection2(True)
                'Part.ViewZoomtofit2()

                If AssyComponents.Contains(AHUName & "_07_Side Blank_B1.SLDPRT") Then
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_07_Side Blank_B1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, PnlHt + (PnlHt / 2) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                    Part.ClearSelection2(True)

                    boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_07_Side Blank_B1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                    Part.ClearSelection2(True)

                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_07_Side Blank_B1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (WallWth / 2000) - (EndBlankWth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                    Part.ClearSelection2(True)

                    Part.EditRebuild3()
                End If



                GoTo TopBlankAssem
            End If

        End If

#Region " Side Blanks A1-A2_B1-B2"

#Region "SideBlank_A"
        If AssyComponents.Contains(AHUName & "_07_Side Blank_A2.SLDPRT") Then
            '-----------------A1
            'Mates RIGHT SIDE
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_07_Side Blank_A1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt / 2) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_07_Side Blank_A1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_07_Side Blank_A1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (PnlWth * FanNoX / 2) + (PnlWth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
            Part.ClearSelection2(True)

            Part.EditRebuild3()

            Assy.ViewZoomtofit2()
            '--------------------------------------------------------------- 2nd Blank -------------------------------------------------------------------------------  
            'Mates Left SIDE
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_07_Side Blank_A1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt / 2) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_07_Side Blank_A1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_07_Side Blank_A1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlWth * FanNoX / 2) + (PnlWth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
            Part.ClearSelection2(True)

            Part.EditRebuild3()

            '-----------------A2
            'Mates RIGHT SIDE
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_07_Side Blank_A2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt / 2) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_07_Side Blank_A2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_07_Side Blank_A2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (PnlWth * FanNoX / 2) + PnlWth + ((SideClear - PnlWth) / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
            Part.ClearSelection2(True)

            Part.EditRebuild3()

            Assy.ViewZoomtofit2()

            '--------------------------------------------------------------- 2nd Blank -------------------------------------------------------------------------------  
            'Mates LEFT SIDE
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_07_Side Blank_A2-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt / 2) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_07_Side Blank_A2-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_07_Side Blank_A2-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlWth * FanNoX / 2) + PnlWth + ((SideClear - PnlWth) / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
            Part.ClearSelection2(True)

            Part.EditRebuild3()

            Assy.ViewZoomtofit2()

            If AssyComponents.Contains(AHUName & "_07_Side Blank_B2.SLDPRT") Then

            Else
                GoTo TopBlankAssem
            End If
        End If
#End Region

#Region "SideBlank_B"
        If AssyComponents.Contains(AHUName & "_07_Side Blank_B2.SLDPRT") Then
            '--------- B1
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_07_Side Blank_B1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, PnlHt + (PnlHt / 2) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_07_Side Blank_B1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_07_Side Blank_B1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (PnlWth * FanNoX / 2) + (PnlWth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
            Part.ClearSelection2(True)

            Part.EditRebuild3()

            'Mates LEFT SIDE
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_07_Side Blank_B1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, PnlHt + (PnlHt / 2) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_07_Side Blank_B1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_07_Side Blank_B1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlWth * FanNoX / 2) + (PnlWth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
            Part.ClearSelection2(True)

            Part.EditRebuild3()


            '----------------- B2
            'Mates RIGHT SIDE
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_07_Side Blank_B2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, PnlHt + (PnlHt / 2) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_07_Side Blank_B2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_07_Side Blank_B2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (PnlWth * FanNoX / 2) + PnlWth + ((SideClear - PnlWth) / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
            Part.ClearSelection2(True)

            Part.EditRebuild3()

            'Mates RIGHT SIDE
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_07_Side Blank_B2-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, PnlHt + (PnlHt / 2) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_07_Side Blank_B2-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_07_Side Blank_B2-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlWth * FanNoX / 2) + PnlWth + ((SideClear - PnlWth) / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
            Part.ClearSelection2(True)

            Part.EditRebuild3()

            Assy.ViewZoomtofit2()

            GoTo TopBlankAssem
        End If

#End Region

#End Region

#Region "Side Blanks A1_B1"
        'ONE SIDE BLANK Mates
        'Mates RIGHT SIDE
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_07_Side Blank_A1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt / 2) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_07_Side Blank_A1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_07_Side Blank_A1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (WallWth / 2000) - (SideClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
        Part.ClearSelection2(True)

        Part.EditRebuild3()

        'Mates LEFT SIDE
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_07_Side Blank_A1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt / 2) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_07_Side Blank_A1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_07_Side Blank_A1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (WallWth / 2000) - (SideClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
        Part.ClearSelection2(True)

        Part.EditRebuild3()


        'SideBlank B
        If AssyComponents.Contains(AHUName & "_07_Side Blank_B1.SLDPRT") Then
            'Mates RIGHT SIDE
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_07_Side Blank_B1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, PnlHt + (PnlHt / 2) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_07_Side Blank_B1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_07_Side Blank_B1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (WallWth / 2000) - (SideClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
            Part.ClearSelection2(True)

            Part.EditRebuild3()

            'Mates LEFT SIDE
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_07_Side Blank_B1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, PnlHt + (PnlHt / 2) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_07_Side Blank_B1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_07_Side Blank_B1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (WallWth / 2000) - (SideClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
            Part.ClearSelection2(True)

            Part.EditRebuild3()
        End If

#End Region

TopBlankAssem:
        '
        '-------------------------------------------------------------------TOP BLANK Mates----------------------------------------------------------------------------------
        If AHUDoor = True Then
            Part = swApp.ActiveDoc
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_08_Top Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) + (BlkOne / 2000) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Top Plane coincidence 
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_08_Top Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_08_Top Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
            Part.ClearSelection2(True)

            Part.EditRebuild3()

            Assy.ViewZoomtofit2()
            ''--------------------------------------------------------------- 2nd Blank -------------------------------------------------------------------------------  
            If AssyComponents.Contains(AHUName & "_08_Top Blank_2.SLDPRT") Then
                Part = swApp.ActiveDoc
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_08_Top Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + BlkOne / 1000 + (BlkTwo / 2000), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane coincidence 
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_08_Top Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_08_Top Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_08_Top Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Top Plane Coincident
                Part.ClearSelection2(True)

                Part.EditRebuild3()
            End If

            'Fans Pattern
            boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2(AHUName & "_08_Top Blank_1-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(AHUName & "_08_Top Blank_2-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            If PushSide = True Then
                myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX, PnlWth, 0, 0, False, False, "NULL", "NULL", False)
            Else
                myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX, PnlWth, 0, 0, True, False, "NULL", "NULL", False)
            End If

            Part.ClearSelection2(True)
                Part.ViewZoomtofit2()

            GoTo VerticalC
        End If

        If PushSide = True Then
            If AssyComponents.Contains(AHUName & "_08_Top Blank_2.SLDPRT") Then

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_08_Top Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_08_Top Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_08_Top Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, (PnlHt * FanNoY) - 0.025 + BlkOne / 2000, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                Part.ClearSelection2(True)


                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_08_Top Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_08_Top Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_08_Top Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, (PnlHt * FanNoY) + BlkOne / 1000 + BlkTwo / 2000 - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                Part.ClearSelection2(True)


                'Fans Pattern
                boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2(AHUName & "_08_Top Blank_1-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2(AHUName & "_08_Top Blank_2-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX, PnlWth, 0, 0, False, False, "NULL", "NULL", False)
                Part.ClearSelection2(True)
                Part.ViewZoomtofit2()

            Else
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_08_Top Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_08_Top Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_08_Top Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, (PnlHt * FanNoY) - 0.025 + (TopClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                Part.ClearSelection2(True)

                'Fans Pattern
                boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2(AHUName & "_08_Top Blank_1-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2(AHUName & "_08_Top Blank_2-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX, PnlWth, 0, 0, False, False, "NULL", "NULL", False)
                Part.ClearSelection2(True)
                Part.ViewZoomtofit2()

            End If
            GoTo VerticalC
        End If

        If AssyComponents.Contains(AHUName & "_08_Top Blank_2.SLDPRT") Then
            'TWO TOP BLANK Mates
            Part = swApp.ActiveDoc
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_08_Top Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) + (BlkOne / 2000) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Top Plane coincidence 
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_08_Top Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_08_Top Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, True, (WallWth / 2000) - SideClear - (PnlWth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Right Plane coincidence 
            Part.ClearSelection2(True)

            Part.EditRebuild3()

            Assy.ViewZoomtofit2()
            ''--------------------------------------------------------------- 2nd Blank -------------------------------------------------------------------------------  

            Part = swApp.ActiveDoc
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_08_Top Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + BlkOne / 1000 + (BlkTwo / 2000), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane coincidence 
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_08_Top Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_08_Top Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, True, (WallWth / 2000) - SideClear - (PnlWth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane coincidence 
            Part.ClearSelection2(True)

            Part.EditRebuild3()

            'Fans Pattern
            boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(AHUName & "_08_Top Blank_1-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(AHUName & "_08_Top Blank_2-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX, PnlWth, 0, 0, True, False, "NULL", "NULL", False)
            Part.ClearSelection2(True)
            Part.ViewZoomtofit2()
        Else
            'ONE TOP BLANK Mates
            Part = swApp.ActiveDoc
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_08_Top Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + (TopClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Top Plane coincidence 
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_08_Top Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_08_Top Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, True, (WallWth / 2000) - SideClear - (PnlWth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Right Plane coincidence 
            Part.ClearSelection2(True)

            Part.EditRebuild3()

            'Fans Pattern
            boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, True, 4, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(AHUName & "_08_Top Blank_1-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX, PnlWth, BlkNosY, (0.4 + ((TopClear - 0.4) / 2)), True, False, "NULL", "NULL", False)
            Part.ClearSelection2(True)
            Part.ViewZoomtofit2()

            Part.Save()
        End If

        If AssyComponents.Contains(AHUName & "_08_Top Blank_3.SLDPRT") Then
            Part = swApp.ActiveDoc
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_08_Top Blank_3-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + BlkOne / 1000 + (BlkTwo / 1000) + (BlkThree / 2000), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane coincidence 
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_08_Top Blank_3-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_08_Top Blank_3-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, True, (WallWth / 2000) - SideClear - (PnlWth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane coincidence 
            Part.ClearSelection2(True)

            Part.EditRebuild3()

            'Fans Pattern
            boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(AHUName & "_08_Top Blank_3-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX, PnlWth, 0, 0, True, False, "NULL", "NULL", False)
            Part.ClearSelection2(True)
            Part.ViewZoomtofit2()

        End If

VerticalC:
        '-----------------------------------------------------------------Vertical C Channel Mates----------------------------------------------------------------------------
        If PushSide = True Then

            If AssyComponents.Contains(AHUName & "_04B_Vertical Channel_02.SLDPRT") Then
                '1st Vertical Channel
                Part = swApp.ActiveDoc
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.023, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '----------- Top Plane Distance 
                'myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt - PnlHt / 4 - 0.022), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.002, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '----------- Front Plane Distance 
                'myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, True, PnlWth / 2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                Part.EditRebuild3()

                '2nd Vertical Channel
                Part = swApp.ActiveDoc
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_04B_Vertical Channel_02-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.023, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '----------- Top Plane Distance 
                'myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, ((PnlHt + PnlHt / 2) + ((WallHt / 1000) - (PnlHt + PnlHt / 2)) / 2) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_04B_Vertical Channel_02-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.002, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '----------- Front Plane Distance 
                'myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_04B_Vertical Channel_02-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, True, PnlWth / 2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                Part.EditRebuild3()

                'Mates
                'Vertical Channel Pattern
                If BlkNosX = 0 Then
                    boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2(AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                    myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX - 1, PnlWth, 0, 0, True, False, "NULL", "NULL", False)
                    boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2(AHUName & "_04B_Vertical Channel_02-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                    myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX - 1, PnlWth, 0, 0, True, False, "NULL", "NULL", False)
                Else
                    boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2(AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                    myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX + 1, PnlWth, 0, 0, True, False, "NULL", "NULL", False)
                    boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2(AHUName & "_04B_Vertical Channel_02-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                    myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX + 1, PnlWth, 0, 0, True, False, "NULL", "NULL", False)
                End If

            Else
                '1st Vertical Channel
                Part = swApp.ActiveDoc
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.023, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '----------- Top Plane Distance 
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.002, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '----------- Front Plane Distance 
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, True, PnlWth / 2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                Part.EditRebuild3()

                boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2(AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX, PnlWth, 0, 0, False, False, "NULL", "NULL", False)

            End If

            GoTo CornerBlanks
        Else
            If AHUDoor = True Then

                Part = swApp.ActiveDoc
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.023, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '----------- Top Plane Distance 
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.002, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '----------- Front Plane Distance 
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, True, PnlWth / 2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Right Plane coincidence
                Part.ClearSelection2(True)

                Part.EditRebuild3()

                'Vertical Channel Pattern
                boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2(AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX + 1, PnlWth, 0, 0, True, False, "NULL", "NULL", False)

                GoTo CornerBlanks
            End If
        End If

        If BlkNosX = 0 Then
            Select Case FanNoX
                Case 1
                    GoTo CornerBlanks
                Case 2
                    Part = swApp.ActiveDoc
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.023, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '----------- Top Plane Distance 
                    'myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '-------- Top Plane Coincident
                    Part.ClearSelection2(True)

                    boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.002, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '----------- Front Plane Distance 
                    'myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                    Part.ClearSelection2(True)

                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane coincidence 
                    Part.ClearSelection2(True)

                    Part.EditRebuild3()
                Case Else
                    If AssyComponents.Contains(AHUName & "_04B_Vertical Channel_02.SLDPRT") Then
                        TwoVerticalChannel(FanNoX, PnlWth, BlkNosX, PnlHt)
                    Else
                        Part = swApp.ActiveDoc
                        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                        boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.023, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '----------- Top Plane Distance 
                        'myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '-------- Top Plane Coincident
                        Part.ClearSelection2(True)

                        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                        boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.002, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '----------- Front Plane Distance 
                        'myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                        Part.ClearSelection2(True)

                        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, PnlWth / 2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Right Plane coincidence 
                        Part.ClearSelection2(True)

                        Part.EditRebuild3()

                        Part.Save()

                        'Linear Pattern
                        boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                        boolstatus = Part.Extension.SelectByID2(AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                        myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX - 1, PnlWth, 0, 0, True, False, "NULL", "NULL", False)
                    End If
            End Select
        Else
            If AssyComponents.Contains(AHUName & "_04B_Vertical Channel_02.SLDPRT") Then
                TwoVerticalChannel(FanNoX, PnlWth, BlkNosX, PnlHt)
            Else
                OneVerticalChannel(FanNoX, SideClear, PnlWth)
            End If
        End If

        Part.ClearSelection2(True)
        Part.ViewZoomtofit2()

        If BlkNosX >= 2 Then
            For f = FanNoX + 2 To FanNoX + 3
                InsComp = Assy.AddComponent5(SaveFolder & "\AHU\" & AHUName & "_04A_Vertical Channel.SLDPRT", 0, "", False, "", 3.5, 3.5, 3.5)
            Next

            'Left Side
            Part = swApp.ActiveDoc
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_04A_Vertical Channel-" & FanNoX + 2 & "@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.023, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '----------- Top Plane Distance 
            'myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '-------- Top Plane Coincident
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_04A_Vertical Channel-" & FanNoX + 2 & "@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.002, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '----------- Front Plane Distance 
            'myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_04A_Vertical Channel-" & FanNoX + 2 & "@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_05_Side Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, PnlWth * 2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '----------- Right Plane coincidence 
            Part.ClearSelection2(True)

            Part.EditRebuild3()

            'Right Side
            Part = swApp.ActiveDoc
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_04A_Vertical Channel-" & FanNoX + 2 + 1 & "@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.023, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '----------- Top Plane Distance 
            'myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '-------- Top Plane Coincident
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_04A_Vertical Channel-" & FanNoX + 2 + 1 & "@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.002, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '----------- Front Plane Distance 
            'myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_04A_Vertical Channel-" & FanNoX + 2 + 1 & "@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_05_Side Blank_1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, True, PnlWth * 2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane coincidence 
            Part.ClearSelection2(True)

            Part.EditRebuild3()
        End If

        Part.ClearSelection2(True)
        Part.ViewZoomtofit2()

        '-----------------------------------------------------------------Corner Blanks-----------------------------------------------------------------------------------------
CornerBlanks:
        If PushSide = True Then
            If AHUDoor = True Then
                If FanNoY = 1 Then
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) + (BlkOne / 1000) + (BlkTwo / 2000) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                    Part.ClearSelection2(True)
                Else
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) + (BlkOne / 2000) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                    Part.ClearSelection2(True)
                End If

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (WallWth / 2000) - (SideClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)
                Part.EditRebuild3()

                If AssyComponents.Contains(AHUName & "_09_Corner Blank_2.SLDPRT") Then

                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) + (BlkOne / 2000) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                    Part.ClearSelection2(True)

                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (WallWth / 2000) - (EndBlankWth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                    Part.ClearSelection2(True)

                    boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                    Part.ClearSelection2(True)
                    Part.EditRebuild3()

                End If

                If AssyComponents.Contains(AHUName & "_09_Corner Blank_3.SLDPRT") Then

                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_3-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) + (BlkOne / 1000) + (BlkTwo / 2000) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                    Part.ClearSelection2(True)

                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_3-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (WallWth / 2000) - (EndBlankWth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                    Part.ClearSelection2(True)

                    boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_3-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                    Part.ClearSelection2(True)
                    Part.EditRebuild3()

                End If

                GoTo TopLAssem
            End If

            If AssyComponents.Contains(AHUName & "_09_Corner Blank_2.SLDPRT") Then
                'LEFT SIDE
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + BlkOne / 2000, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (WallWth / 2000) - (SideClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)
                Part.EditRebuild3()


                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) + (BlkOne / 1000) + (BlkTwo / 2000) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (WallWth / 2000) - (SideClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)
                Part.EditRebuild3()

            Else
                'LEFT SIDE
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + (TopClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (WallWth / 2000) - (SideClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)
                Part.EditRebuild3()
            End If
            GoTo TopLAssem
        Else
            If AHUDoor = True Then
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                'If BlkTwo = 0 Then
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) + (BlkOne / 2000) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                    'Else
                    '    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) + (BlkOne / 1000) + (BlkTwo / 2000) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                    'End If
                    Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (WallWth / 2000) - (SideClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)
                Part.EditRebuild3()

                If AssyComponents.Contains(AHUName & "_09_Corner Blank_2.SLDPRT") Then

                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    If AssyComponents.Contains(AHUName & "_09_Corner Blank_3.SLDPRT") Then
                        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) + (BlkOne / 1000) + (BlkTwo / 2000) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                    Else
                        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) + (BlkOne / 2000) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                    End If
                    Part.ClearSelection2(True)

                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    If AssyComponents.Contains(AHUName & "_09_Corner Blank_3.SLDPRT") Then
                        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (WallWth / 2000) - (SideClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                    Else
                        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (WallWth / 2000) - (EndBlankWth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                    End If
                    Part.ClearSelection2(True)

                    boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                    Part.ClearSelection2(True)
                    Part.EditRebuild3()

                End If

                If AssyComponents.Contains(AHUName & "_09_Corner Blank_3.SLDPRT") Then

                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_3-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) + (BlkOne / 2000) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                    'myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) + (BlkOne / 1000) + (BlkTwo / 2000) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                    Part.ClearSelection2(True)

                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_3-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (WallWth / 2000) - (EndBlankWth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                    Part.ClearSelection2(True)

                    boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_3-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                    Part.ClearSelection2(True)
                    Part.EditRebuild3()

                End If

                If AssyComponents.Contains(AHUName & "_09_Corner Blank_4.SLDPRT") Then

                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_4-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) + (BlkOne / 1000) + (BlkTwo / 2000) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                    Part.ClearSelection2(True)

                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_4-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (WallWth / 2000) - (EndBlankWth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                    Part.ClearSelection2(True)

                    boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_4-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                    Part.ClearSelection2(True)
                    Part.EditRebuild3()

                End If

                GoTo TopLAssem
            End If
        End If


        If SideClear > PnlWth Then
            If TopClear > PnlHt Then
#Region "If TopClr & SideClr Is Greater = 4 corner blanks"
                ''--------------------------------------------------------------- 1st Corner Blank ------------------------------------------------------------------------------- 
                'RIGHT SIDE
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + (BlkOne / 2000), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (PnlWth * FanNoX / 2) + (PnlWth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                Part.ClearSelection2(True)

                'LEFT SIDE
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + (BlkOne / 2000), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlWth * FanNoX / 2) + (PnlWth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                Part.ClearSelection2(True)

                Part.EditRebuild3()

                ''--------------------------------------------------------------- 2nd Corner Blank ------------------------------------------------------------------------------- 
                'RIGHT SIDE
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + BlkOne / 1000 + (BlkTwo / 2000), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (PnlWth * FanNoX / 2) + (PnlWth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                Part.ClearSelection2(True)

                'LEFT SIDE
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_2-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + BlkOne / 1000 + (BlkTwo / 2000), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_2-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlWth * FanNoX / 2) + (PnlWth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                Part.ClearSelection2(True)

                Part.EditRebuild3()

                ''--------------------------------------------------------------- 3rd Corner Blank ------------------------------------------------------------------------------- 
                'RIGHT SIDE
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_3-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + (BlkOne / 2000), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_3-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (PnlWth * FanNoX / 2) + PnlWth + ((SideClear - PnlWth) / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                Part.ClearSelection2(True)

                'LEFT SIDE
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_3-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + (BlkOne / 2000), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_3-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlWth * FanNoX / 2) + PnlWth + ((SideClear - PnlWth) / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                Part.ClearSelection2(True)

                Part.EditRebuild3()

                ''--------------------------------------------------------------- 4th Corner Blank ------------------------------------------------------------------------------- 
                'RIGHT SIDE
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_4-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + BlkOne / 1000 + (BlkTwo / 2000), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_4-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (PnlWth * FanNoX / 2) + PnlWth + ((SideClear - PnlWth) / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                Part.ClearSelection2(True)

                'LEFT SIDE
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_4-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + BlkOne / 1000 + (BlkTwo / 2000), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_4-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlWth * FanNoX / 2) + PnlWth + ((SideClear - PnlWth) / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                Part.ClearSelection2(True)

                Part.EditRebuild3()

#Region "FRONT PLANE Mates"
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)
                Part.EditRebuild3()

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)
                Part.EditRebuild3()

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)
                Part.EditRebuild3()

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_2-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)
                Part.EditRebuild3()

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_3-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)
                Part.EditRebuild3()

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_3-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)
                Part.EditRebuild3()

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_4-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)
                Part.EditRebuild3()

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_4-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)
                Part.EditRebuild3()
#End Region
#End Region
            Else
#Region "If TopClr Is Less than PnlHt"
                'Corner Blank-1-1"
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + (BlkOne / 2000), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                If AssyComponents.Contains(AHUName & "_09_Corner Blank_2.SLDPRT") Then
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (PnlWth * FanNoX / 2) + (PnlWth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                Else
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (PnlWth * FanNoX / 2) + (SideClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                End If
                Part.ClearSelection2(True)

                'Corner Blank-1-2"
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + (BlkOne / 2000), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                If AssyComponents.Contains(AHUName & "_09_Corner Blank_2.SLDPRT") Then
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlWth * FanNoX / 2) + (PnlWth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                Else
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlWth * FanNoX / 2) + (SideClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                End If
                Part.ClearSelection2(True)

                Part.EditRebuild3()

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)
                Part.EditRebuild3()

                '--------------------------------------------------------------- 2nd Corner Blank ------------------------------------------------------------------------------- 
                If AssyComponents.Contains(AHUName & "_09_Corner Blank_2.SLDPRT") Then
                    '2nd Corner Blank
                    'RIGHT SIDE
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + (TopClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                    Part.ClearSelection2(True)

                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (PnlWth * FanNoX / 2) + PnlWth + ((SideClear - PnlWth) / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                    Part.ClearSelection2(True)

                    'LEFT SIDE
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_2-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + (TopClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                    Part.ClearSelection2(True)

                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_2-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlWth * FanNoX / 2) + PnlWth + ((SideClear - PnlWth) / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                    Part.ClearSelection2(True)

                    boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                    Part.ClearSelection2(True)

                    boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_2-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                    Part.ClearSelection2(True)

                    Part.EditRebuild3()
                End If
#End Region
            End If

            GoTo TopLAssem

        ElseIf TopClear > PnlHt Then
#Region "If TopClr is Greater than PnlHt"
            'Corner Blank-1-1"
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + (BlkOne / 2000), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (WallWth / 2000) - (SideClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
            Part.ClearSelection2(True)

            'Corner Blank-1-2"
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + (BlkOne / 2000), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (WallWth / 2000) - (SideClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
            Part.ClearSelection2(True)
            Part.EditRebuild3()

            If AssyComponents.Contains(AHUName & "_09_Corner Blank_2.SLDPRT") Then
                '2nd Corner Blank
                'RIGHT SIDE
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + (BlkOne / 1000) + (BlkTwo / 2000), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (WallWth / 2000) - (SideClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                Part.ClearSelection2(True)

                'LEFT SIDE
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_2-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + (BlkOne / 1000) + (BlkTwo / 2000), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_2-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (WallWth / 2000) - (SideClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_2-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)
                Part.EditRebuild3()

                Part.EditRebuild3()
            End If

            If AssyComponents.Contains(AHUName & "_09_Corner Blank_3.SLDPRT") Then
                '3rd Corner Blank
                'RIGHT SIDE
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_3-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + (BlkOne / 1000) + (BlkTwo / 1000) + (BlkThree / 2000), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_3-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (WallWth / 2000) - (SideClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                Part.ClearSelection2(True)

                'LEFT SIDE
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_3-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + (BlkOne / 1000) + (BlkTwo / 1000) + (BlkThree / 2000), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_3-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (WallWth / 2000) - (SideClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
                Part.ClearSelection2(True)



                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_3-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_3-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
                Part.ClearSelection2(True)
                Part.EditRebuild3()

                Part.EditRebuild3()
            End If

#End Region
            GoTo TopLAssem
        End If


#Region "SINGLE CORNER BLANK Mates"
        'RIGHT SIDE
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + (TopClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, False, (PnlWth * FanNoX / 2) + (SideClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        'LEFT SIDE
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt * FanNoY) - 0.025 + (TopClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_09_Corner Blank_1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlWth * FanNoX / 2) + (SideClear / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
        Part.ClearSelection2(True)
        Part.EditRebuild3()
#End Region


        '--------------------------------------------------------------------Top L----------------------------------------------------------------------------------------------
TopLAssem:
        If AssyComponents.Contains(AHUName & "_02B_Top_L2.SLDPRT") Then
            If AHUDoor = True Then
                If FanNoX Mod 2 = 0 Then
                    BTL2 = DoorBlkWth + (PnlWth * Truncate(FanNoX / 2))                          '02B
                    BTL1 = (WallWth / 1000) - (DoorBlkWth + (PnlWth * Truncate(FanNoX / 2)))     '02A 
                Else
                    BTL2 = DoorBlkWth + (PnlWth * Truncate(FanNoX / 2))                          '02B
                    BTL1 = (WallWth / 1000) - (DoorBlkWth + (PnlWth * Truncate(FanNoX / 2)))     '02A 
                End If

                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_02A_Top_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, (WallHt - 50) / 1000, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_02A_Top_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                Part = swApp.ActiveDoc
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_02A_Top_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, BTL2 / 2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                Part.EditRebuild3()

                Assy.ViewZoomtofit2()

                '------------------------------------------------------------- Add 2nd Top L -------------------------------------------------------------------------------   
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_02B_Top_L2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, (WallHt - 50) / 1000, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_02B_Top_L2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_02B_Top_L2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, True, BTL1 / 2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                Part.EditRebuild3()

            Else

                If FanNoX Mod 2 = 0 Then
                    BTLth = (WallWth / 1000) - (SideClear + (PnlWth * (FanNoX / 2)))
                Else
                    BTLth = (WallWth / 1000) - (SideClear + (PnlWth * Truncate(FanNoX / 2)))
                End If

                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_02A_Top_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, (WallHt - 50) / 1000, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_02A_Top_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                Part = swApp.ActiveDoc
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_02A_Top_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, (WallWth / 2000) - (BTLth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                Part.EditRebuild3()

                Assy.ViewZoomtofit2()


                '------------------------------------------------------------- Add 2nd Top L -------------------------------------------------------------------------------

                If FanNoX Mod 2 = 0 Then
                    BTLth = (WallWth / 1000) - (SideClear + (PnlWth * (FanNoX / 2)))
                Else
                    BTLth = SideClear + (PnlWth * Truncate(FanNoX / 2))
                End If

                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_02B_Top_L2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, (WallHt - 50) / 1000, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_02B_Top_L2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_02B_Top_L2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, True, (WallWth / 2000) - (BTLth / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                Part.EditRebuild3()

            End If

            GoTo SideLAssem
        End If

        If AssyComponents.Contains(AHUName & "_02B_Top_L2.SLDPRT") Then
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_02A_Top_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, (WallHt - 50) / 1000, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_02A_Top_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            Part = swApp.ActiveDoc
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_02A_Top_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, WallWth / 4000, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            Part.EditRebuild3()

            Assy.ViewZoomtofit2()

            '------------------------------------------------------------- Add 2nd Top L -------------------------------------------------------------------------------   
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_02B_Top_L2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, (WallHt - 50) / 1000, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_02B_Top_L2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_02B_Top_L2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, True, WallWth / 4000, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            Part.EditRebuild3()
        Else
            If PushSide = False Then
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_02A_Top_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, (WallHt - 50) / 1000, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_02A_Top_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_02A_Top_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)
            Else
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_02A_Top_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, (WallHt - 50) / 1000, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                'myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, True, (WallHt - 50) / 1000, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_02A_Top_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_02A_Top_L-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)


            End If

            Part.EditRebuild3()
        End If



SideLAssem:
        '--------------------------------------------------------------------Side L---------------------------------------------------------------------------------------------
        If AssyComponents.Contains(AHUName & "_03B_Side_L2_Right.SLDPRT") Then
            TwoSideLs(PnlHt)
        Else

            'Mates LEFT SIDE
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_03A_Side_L_Left-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            'myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_03A_Side_L_Left-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_03A_Side_L_Left-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, SideLDist - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            Part.EditRebuild3()

            'Mates RIGHT SIDE
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_03A_Side_L_Right-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_03A_Side_L_Right-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            'myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, True, 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_03A_Side_L_Right-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, SideLDist - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            Part.EditRebuild3()
        End If

        '-------------------------------------------------------------------- Base Stand ---------------------------------------------------------------------------------------------
        Part = swApp.ActiveDoc
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_10_Base_Stand-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, True, WallWth / 2000, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_10_Base_Stand-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, True, 0.027, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_10_Base_Stand-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly/" & AHUName & "_" & ArtNo & "-1@" & AHUName & "_Motor Sub Assembly/" & MotorDia & "mm_10_C Channel-1@" & AHUName & "_" & ArtNo & "", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        'If AHUDoor = True Then
        '    myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Right Plane Coincident
        'Else
        myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 1, True, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Right Plane Coincident
        'End If
        Part.ClearSelection2(True)

        If WallWth > MaxSecLth3mm Then
            'Fans Pattern
            boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(AHUName & "_10_Base_Stand-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2(2, WallWth / 2000, 0, 0, False, False, "NULL", "NULL", False)
        End If
        Part.ClearSelection2(True)
        Part.ViewZoomtofit2()

        '-------------------------------------------------------------------- Horizontal C Channel 2 & 3 ----------------------------------------------------------------------
        If AssyComponents.Contains(AHUName & "_05B_Hor_C_Support_3.SLDPRT") Then

            If AHUDoor = True Then
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_05B_Hor_C_Support_3-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_Motor Sub Assembly_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.002, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_05B_Hor_C_Support_3-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, True, PnlHt - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                If PushSide = True Then
                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_05B_Hor_C_Support_3-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, 0.024, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                    Part.ClearSelection2(True)

                    'For b = 1 To FanNoY + BlkNosY - 1
                    '    boolstatus = Part.Extension.SelectByID2(AHUName & "_05_Hor_C_Support-" & b & "@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
                    '    Part.HideComponent2()
                    '    Part.ClearSelection2(True)
                    'Next

                Else
                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_05B_Hor_C_Support_3-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_07_Side Blank_A1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.024, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                    Part.ClearSelection2(True)
                End If

                boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2(AHUName & "_05B_Hor_C_Support_3-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoY + BlkNosY - 1, PnlHt, 0, 0, True, True, "NULL", "NULL", False)
                Part.ClearSelection2(True)
                Part.ViewZoomtofit2()
            End If

            If BlkOne = 400 Then
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_05B_Hor_C_Support_3-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_Motor Sub Assembly_2-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.002, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_05B_Hor_C_Support_3-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_09_Corner Blank_3-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, True, 0.2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_05B_Hor_C_Support_3-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_07_Side Blank_A1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.024, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)
            Else
                boolstatus = Part.Extension.SelectByID2(AHUName & "_05B_Hor_C_Support_3-2@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                Part.EditDelete()
            End If

            Part.ClearSelection2(True)
            Part.ViewZoomtofit2()

        End If

        If AssyComponents.Contains(AHUName & "_05A_Hor_C_Support_2.SLDPRT") Then

            If AHUDoor = True Then
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_05A_Hor_C_Support_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_Motor Sub Assembly_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.002, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_05A_Hor_C_Support_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, DoorPnlHt - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_05A_Hor_C_Support_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_Door SubAssembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, 0.024, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                If BlkNosY >= 2 Then
                    boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2(AHUName & "_05A_Hor_C_Support_2-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                    If BlkOne = 400 Then
                        myFeature = Part.FeatureManager.FeatureLinearPattern2(2, 0.4, 0, 0, True, True, "NULL", "NULL", False)
                    Else
                        myFeature = Part.FeatureManager.FeatureLinearPattern2(BlkNosY - 1, PnlHt, 0, 0, True, True, "NULL", "NULL", False)
                    End If
                    Part.ClearSelection2(True)
                    Part.ViewZoomtofit2()
                End If

                    GoTo SaveAssem
                End If


                '------------------------------------------ Aligned with Fans ------------------------------------------------
                'Right Side
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_05A_Hor_C_Support_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_Motor Sub Assembly_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.002, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_05A_Hor_C_Support_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_07_Side Blank_A1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.024, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            If FanNoY >= 1 And BlkNosY >= 1 Then

                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_05A_Hor_C_Support_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, True, PnlHt - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2(AHUName & "_05A_Hor_C_Support_2-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoY, PnlHt, 0, 0, True, True, "NULL", "NULL", False)
                Part.ClearSelection2(True)
                Part.ViewZoomtofit2()

            ElseIf FanNoY >= 1 And BlkNosY = 0 Then
                If (FanNoY - 1) > 0 Then
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_05A_Hor_C_Support_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, True, PnlHt - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                    Part.ClearSelection2(True)

                    boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2(AHUName & "_05A_Hor_C_Support_2-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                    myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoY - 1, PnlHt, 0, 0, True, True, "NULL", "NULL", False)
                    Part.ClearSelection2(True)
                    Part.ViewZoomtofit2()

                    GoTo SaveAssem
                End If
            End If


            'Left Side
            If BlkNosY = 1 Then
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_05A_Hor_C_Support_2-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_Motor Sub Assembly_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.002, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_05A_Hor_C_Support_2-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_07_Side Blank_A1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, 0.024, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                If FanNoY >= 1 And BlkNosY >= 1 Then

                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_05A_Hor_C_Support_2-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, PnlHt - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                    Part.ClearSelection2(True)

                    boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2(AHUName & "_05A_Hor_C_Support_2-2@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                    myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoY, PnlHt, 0, 0, True, True, "NULL", "NULL", False)
                    Part.ClearSelection2(True)
                    Part.ViewZoomtofit2()

                ElseIf FanNoY >= 1 And BlkNosY = 0 Then
                    If (FanNoY - 1) > 0 Then
                        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_05A_Hor_C_Support_2-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, PnlHt - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                        Part.ClearSelection2(True)

                        boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                        boolstatus = Part.Extension.SelectByID2(AHUName & "_05A_Hor_C_Support_2-2@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                        myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoY - 1, PnlHt, 0, 0, True, True, "NULL", "NULL", False)
                        Part.ClearSelection2(True)
                        Part.ViewZoomtofit2()

                        GoTo SaveAssem
                    End If

                End If

            Else
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_05A_Hor_C_Support_2-3@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_Motor Sub Assembly_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.002, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_05A_Hor_C_Support_2-3@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_07_Side Blank_A1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, 0.024, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                If FanNoY >= 1 And BlkNosY >= 1 Then

                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_05A_Hor_C_Support_2-3@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, PnlHt - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                    Part.ClearSelection2(True)

                    boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2(AHUName & "_05A_Hor_C_Support_2-3@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                    myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoY, PnlHt, 0, 0, True, True, "NULL", "NULL", False)
                    Part.ClearSelection2(True)
                    Part.ViewZoomtofit2()

                ElseIf FanNoY >= 1 And BlkNosY = 0 Then
                    If (FanNoY - 1) > 0 Then
                        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_05A_Hor_C_Support_2-3@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, PnlHt - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                        Part.ClearSelection2(True)

                        boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                        boolstatus = Part.Extension.SelectByID2(AHUName & "_05A_Hor_C_Support_2-3@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                        myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoY - 1, PnlHt, 0, 0, True, True, "NULL", "NULL", False)
                        Part.ClearSelection2(True)
                        Part.ViewZoomtofit2()

                        GoTo SaveAssem
                    End If

                End If

            End If


            '-------------------------------------------- Aligned with Top Blanks -----------------------------------------------------
            'Right Side
            If BlkNosY > 1 Then
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_05A_Hor_C_Support_2-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_Motor Sub Assembly_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.002, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_05A_Hor_C_Support_2-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_1-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.024, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)


                If BlkOne = 400 Then
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_05A_Hor_C_Support_2-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, True, (PnlHt * FanNoY) - 0.025 + 0.4, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                    Part.ClearSelection2(True)

                    GoTo LeftSideHorC
                Else
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_05A_Hor_C_Support_2-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, True, (PnlHt * FanNoY) - 0.025 + PnlHt, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                    Part.ClearSelection2(True)

                    'boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                    'boolstatus = Part.Extension.SelectByID2(AHUName & "_05A_Hor_C_Support_2-2@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                    'myFeature = Part.FeatureManager.FeatureLinearPattern2(BlkNosY - 1, PnlHt, 0, 0, True, True, "NULL", "NULL", False)
                    'Part.ClearSelection2(True)
                    'Part.ViewZoomtofit2()

                End If

            End If

            'Left Side
LeftSideHorC:
            If BlkNosY > 1 Then
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_05A_Hor_C_Support_2-4@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_Motor Sub Assembly_2-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.002, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_05A_Hor_C_Support_2-4@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_09_Corner Blank_1-2@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, 0.024, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Part.ClearSelection2(True)

                If BlkOne = 400 Then
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_05A_Hor_C_Support_2-4@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, (PnlHt * FanNoY) - 0.025 + 0.4, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                    Part.ClearSelection2(True)

                    GoTo SaveAssem
                Else
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_05A_Hor_C_Support_2-4@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, (PnlHt * FanNoY) - 0.025 + PnlHt, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                    Part.ClearSelection2(True)

                    'boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
                    'boolstatus = Part.Extension.SelectByID2(AHUName & "_05A_Hor_C_Support_2-4@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                    'myFeature = Part.FeatureManager.FeatureLinearPattern2(BlkNosY - 1, PnlHt, 0, 0, True, True, "NULL", "NULL", False)
                    'Part.ClearSelection2(True)
                    'Part.ViewZoomtofit2()
                End If

            End If

            boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(AHUName & "_05A_Hor_C_Support_2-2@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2(BlkNosY - 1, PnlHt, 0, 0, True, True, "NULL", "NULL", False)
            Part.ClearSelection2(True)
            Part.ViewZoomtofit2()

            boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(AHUName & "_05A_Hor_C_Support_2-4@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2(BlkNosY - 1, PnlHt, 0, 0, True, True, "NULL", "NULL", False)
            Part.ClearSelection2(True)
            Part.ViewZoomtofit2()

            '------------------WIP-------------------

            If PushSide = True Then


            End If

            Part.EditRebuild3()
        End If

        Part.EditRebuild3()

        MsgBox("hi")
SaveAssem:
        ' Zoom To Fit
        Part.ViewZoomtofit2()

        ' Save As
        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_AHU Final Assembly.SLDASM", 0, 0)
        longstatus = Part.SaveAs3(SaveFolder & "\" & AHUName & "_E Drawing\" & AHUName & "_AHU.EASM", 0, 2)

        'PreDXF()

        swApp.CloseAllDocuments(True)

        AssemDrawing()

        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_AHU Final Assembly.SLDASM", 2, 32, "", longstatus, longwarnings)
        CreateCNCPartList(ClientName, JobNum)

        swApp.CloseAllDocuments(True)

    End Sub

    '-------------------------------x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x------------------------------------------------------------------------x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x----------------------------------


#Region "VERTICAL CHANNELS & Two Side L Condition - Assembly"

    Public Sub TwoVerticalChannel(FanNoX As Integer, PnlWth As Decimal, BlkNosX As Integer, PnlHt As Decimal)

Two_Ver_Channel:
        '1st Vertical Channel
        Part = swApp.ActiveDoc
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt - PnlHt / 4 - 0.022), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, True, PnlWth / 2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        Part.EditRebuild3()

        '2nd Vertical Channel
        Part = swApp.ActiveDoc
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_04B_Vertical Channel_02-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, ((PnlHt + PnlHt / 2) + ((WallHt / 1000) - (PnlHt + PnlHt / 2)) / 2) - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_04B_Vertical Channel_02-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_04B_Vertical Channel_02-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_Motor Sub Assembly-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, True, PnlWth / 2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        Part.EditRebuild3()

        'Mates
        'Vertical Channel Pattern
        If BlkNosX = 0 Then
            boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX - 1, PnlWth, 0, 0, True, False, "NULL", "NULL", False)
            boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(AHUName & "_04B_Vertical Channel_02-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX - 1, PnlWth, 0, 0, True, False, "NULL", "NULL", False)
        Else
            boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX + 1, PnlWth, 0, 0, True, False, "NULL", "NULL", False)
            boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(AHUName & "_04B_Vertical Channel_02-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX + 1, PnlWth, 0, 0, True, False, "NULL", "NULL", False)
        End If

    End Sub

    Public Sub OneVerticalChannel(FanNoX As Integer, SideClear As Decimal, PnlWth As Decimal)

Single_Ver_Channel:
        Part = swApp.ActiveDoc
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.023, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '----------- Top Plane Distance 
        'myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '-------- Top Plane Coincident
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, 0.002, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '----------- Front Plane Distance 
        'myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, (WallWth / 2000) - SideClear, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        Part.EditRebuild3()

        'Vertical Channel Pattern
        boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2(AHUName & "_04A_Vertical Channel-1@" & AHUName & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
        myFeature = Part.FeatureManager.FeatureLinearPattern2(FanNoX + 1, PnlWth, 0, 0, False, False, "NULL", "NULL", False)

    End Sub

    Public Sub TwoSideLs(PnlHt As Decimal)
TWO_SIDE_Ls:
        '1st SIDE L Mates LEFT SIDE
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_03A_Side_L_Left-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt - PnlHt / 4 + 0.0015), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_03A_Side_L_Left-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_03A_Side_L_Left-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, SideLDist, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '-----------Right Plane distance 
        Part.ClearSelection2(True)

        Part.EditRebuild3()

        '1st SIDE L Mates RIGHT SIDE
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_03A_Side_L_Right-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_03A_Side_L_Right-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, (PnlHt - PnlHt / 4 + 0.0015), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_03A_Side_L_Right-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, SideLDist, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
        Part.ClearSelection2(True)

        Part.EditRebuild3()

        '----------------------------------------------------------------2nd SIDE L------------------------------------------------------------------------------------------
        ' Mates LEFT SIDE
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_03B_Side_L2_Left-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, ((PnlHt + PnlHt / 2) + ((WallHt / 1000) - (PnlHt + PnlHt / 2)) / 2) - 0.0485, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_03B_Side_L2_Left-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_03B_Side_L2_Left-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 0, False, SideLDist, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '-----------Right Plane distance 
        Part.ClearSelection2(True)

        Part.EditRebuild3()

        '2nd SIDE L Mates RIGHT SIDE

        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_03B_Side_L2_Right-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateCOINCIDENT, 1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)  '--------Front Plane Coincident
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_03B_Side_L2_Right-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, swMateAlign_e.swMateAlignALIGNED, True, ((PnlHt + PnlHt / 2) + ((WallHt / 1000) - (PnlHt + PnlHt / 2)) / 2) - 0.0485, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '------Top Plane distance
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_03B_Side_L2_Right-1@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & AHUName & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        myMate = Assy.AddMate5(swMateType_e.swMateDISTANCE, 1, False, SideLDist, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus) '-----------Right Plane distance 
        Part.ClearSelection2(True)

        Part.EditRebuild3()

    End Sub

#End Region

#Region "Drawings"
    Public Sub CreateDrawings(PartName As String, DwgQty As Integer, HorCSize As Decimal)
        'Exit Sub
        swApp.ActivateDoc2(PartName, False, longstatus)
        Part = swApp.ActiveDoc

        If Part Is Nothing Then
            MsgBox("Please open an appropriate Part Document")
            Exit Sub
        End If

        'Bounding Box
        Dim BBox As Object = StdFunc.BoundingBox()
        Dim xDim As Decimal = Abs(BBox(0)) + Abs(BBox(3))
        Dim yDim As Decimal = Abs(BBox(1)) + Abs(BBox(4))
        Dim zDim As Decimal = Abs(BBox(2)) + Abs(BBox(5))

        'Bounding Box - Flat
        boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditUnsuppress2()

        Dim BBoxFlat As Object = StdFunc.BoundingBox()
        Dim xDimFlat As Decimal = Abs(BBoxFlat(0)) + Abs(BBoxFlat(3))
        Dim yDimFlat As Decimal = Abs(BBoxFlat(1)) + Abs(BBoxFlat(4))
        Dim zDimFlat As Decimal = Abs(BBoxFlat(2)) + Abs(BBoxFlat(5))

        boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditSuppress2()

        'swApp.CloseAllDocuments(True)

        ' Sheet Scale
        Dim SScaleX As Integer = Ceiling((xDimFlat + zDim + xDim) / (0.297 - (0.02 + 0.04 + 0.04 + 0.02)))
        Dim SScaleY As Integer = Ceiling((zDim + yDimFlat) / (0.21 - (0.03 + 0.04 + 0.03)))
        Dim SScale As Integer = SScaleX
        If SScaleX < SScaleY Then
            SScale = SScaleY
        End If

        ' Adjust Values for Scale
        xDim /= SScale
        yDim /= SScale
        zDim /= SScale

        xDimFlat /= SScale
        yDimFlat /= SScale
        zDimFlat /= SScale

        ' Get Margins
        Dim ClearX As Decimal = (0.297 - (xDimFlat + 0.04 + zDim + 0.04 + xDim)) / 2
        If ClearX < 0.03 Then ClearX = 0.03
        Dim ClearY As Decimal = (0.21 - (yDimFlat + 0.04 + zDim)) / 2
        If ClearY < 0.03 Then ClearY = 0.03

        ' Calculate View Placements
        Dim xFrontFlat As Decimal = ClearX + xDimFlat / 2 + 0.005
        Dim yTopSec As Decimal = ClearY + zDim / 2 + 0.005
        Dim yFrontFlat As Decimal = yTopSec + zDim / 2 + 0.04 + yDimFlat / 2
        Dim xRightSec As Decimal = xFrontFlat + xDimFlat / 2 + 0.02 + zDim / 2
        Dim xIso As Decimal = xRightSec + zDim / 2 + 0.04 + xDim / 2

        ' New Drawing
        Part = swApp.NewDocument(DrawSheet, 2, 0.297, 0.21)
        Draw = Part

        swSheet = Draw.GetCurrentSheet()
        swSheet.SetProperties2(12, 12, 1, SScale, False, 0.297, 0.21, True)
        swSheet.SetTemplateName(DrawTemp)
        swSheet.ReloadTemplate(True)
        swApp.ActivateDoc2("Draw1 - Sheet1", False, longstatus)
        Part = swApp.ActiveDoc


        ' Views
        boolstatus = Draw.CreateFlatPatternViewFromModelView(PartName, "Default", xFrontFlat + 0.005, yFrontFlat, 0)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Draw.ActivateView("Drawing View1")
        Part.ClearSelection2(True)

        'Front - Outside
        myView = Part.CreateDrawViewFromModelView3(PartName, "*Front", -xFrontFlat, yFrontFlat, 0)
        boolstatus = Draw.ActivateView("Drawing View2")
        Part.ClearSelection2(True)

        'Top - Section
        boolstatus = Draw.ActivateView("Drawing View2")
        skSegment = Part.SketchManager.CreateLine((xDim * SScale / 2) + 0.15, 0, 0, -(xDim * SScale / 2) - 0.15, 0, 0)
        boolstatus = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xFrontFlat, yTopSec, 0, "A", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View3")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View3", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewVerticalCenter, BaseView)
        Part.ClearSelection2(True)
        boolstatus = Part.ActivateSheet("Sheet1")

        'Right - Section
        boolstatus = Draw.ActivateView("Drawing View2")
        skSegment = Part.SketchManager.CreateLine(0, (yDim * SScale / 2) + 0.15, 0, 0, -(yDim * SScale / 2) - 0.15, 0)
        boolstatus = Part.Extension.SelectByID2("Line2", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xRightSec + 0.02, yFrontFlat, 0, "B", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View4")
        Part.ClearSelection2(True)

        'Isometric
        myView = Part.CreateDrawViewFromModelView3(PartName, "*Dimetric", xRightSec + 0.085, yFrontFlat, 0)
        boolstatus = Draw.ActivateView("Drawing View2")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View5", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject6(1, -1)
        boolstatus = DrawView.SetDisplayMode4(False, swDisplayMode_e.swSHADED, True, True, False)
        Part.ClearSelection2(True)


        For i = 5 To 25
            boolstatus = Part.ActivateView("Drawing View3")
            boolstatus = Part.Extension.SelectByID2("DetailItem" & i & "@Drawing View3", "NOTE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
            Part.ClearSelection2(True)
            boolstatus = Part.ActivateView("Drawing View4")
            boolstatus = Part.Extension.SelectByID2("DetailItem" & i & "@Drawing View4", "NOTE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
            Part.ClearSelection2(True)
        Next

        'Dimensions--------------------------------------------------------------------------------------------------------------
        boolstatus = Draw.ActivateSheet("Sheet1")

        'Height  
        boolstatus = Draw.ActivateView("Drawing View1")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, yFrontFlat + (yDimFlat / 2), -5000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, yFrontFlat - (yDimFlat / 2), -5000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat - (xDimFlat / 2) - 0.015, yFrontFlat, 0)
        Part.ClearSelection2(True)

        'Width 
        boolstatus = Part.ActivateView("Drawing View1")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat + 0.005 + (xDimFlat / 2), yFrontFlat, -5000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat + 0.005 - (xDimFlat / 2), yFrontFlat, -5000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat, yFrontFlat + (yDimFlat / 2) + 0.015, 0)
        Part.ClearSelection2(True)



        'Width - Bottom section   
        boolstatus = Draw.ActivateView("Drawing View3")
        boolstatus = Part.Extension.SelectByRay((xFrontFlat + 0.005) - (HorCSize / (2 * SScale)), yTopSec - (0.01 / SScale), -5500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay((xFrontFlat + 0.005) + (HorCSize / (2 * SScale)), yTopSec - (0.01 / SScale), -5500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat + 0.005, yTopSec + 0.01, 0)
        Part.ClearSelection2(True)

        'Height - Bottom section
        boolstatus = Part.ActivateView("Drawing View3")
        boolstatus = Part.Extension.SelectByRay((xFrontFlat + 0.005) - ((HorCSize - 0.002) / (2 * SScale)), yTopSec - (0.015 / SScale), -5500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay((xFrontFlat + 0.005), yTopSec + (0.015 / SScale), -5500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat + 0.005 - 0.05, yTopSec, 0)
        Part.ClearSelection2(True)



        'Width - Right   
        boolstatus = Draw.ActivateView("Drawing View4")
        boolstatus = Part.Extension.SelectByRay((xRightSec + 0.02) - (0.015 / SScale), yFrontFlat, -5500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay((xRightSec + 0.02) + (0.015 / SScale), yFrontFlat + (0.045 / SScale), -5500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xRightSec + 0.02, yFrontFlat + 0.02, 0)
        Part.ClearSelection2(True)

        'Height - Right
        boolstatus = Part.ActivateView("Drawing View4")
        boolstatus = Part.Extension.SelectByRay((xRightSec + 0.02), yFrontFlat - (0.05 / SScale), -5500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay((xRightSec + 0.02), yFrontFlat + (0.05 / SScale), -5500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xRightSec + 0.02 - 0.012, yFrontFlat, 0)
        Part.ClearSelection2(True)


        'Height - Right
        boolstatus = Part.ActivateView("Drawing View4")
        boolstatus = Part.Extension.SelectByRay((xRightSec + 0.02) + (0.014 / SScale), yFrontFlat + (0.04 / SScale), -5500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay((xRightSec + 0.02), yFrontFlat + (0.05 / SScale), -5500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xRightSec + 0.02 + 0.015, yFrontFlat - (0.01 / SScale), 0)
        Part.ClearSelection2(True)

        ' Note
        Dim X, Y, Z As Decimal
        boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingBalloonStyle, 0, swBalloonStyle_e.swBS_None)

        X = Round(xDimFlat * SScale * 1000, 2)
        Y = Round(yDimFlat * SScale * 1000, 2)
        Z = Round(zDimFlat * SScale * 1000, 2)
        myNote = Draw.CreateText2(PartName & vbNewLine & "Qty - " & DwgQty & vbNewLine & X & "mm x " & Y & "mm x " & "2.00mm", xRightSec + 0.035, yTopSec, 0, 0.004, 0)

        boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

        XValueList.Add(X)
        YValueList.Add(Y)
        ZValueList.Add(Z)

        ' Save As
        Part.ViewZoomtofit2()
        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & PartName & ".SLDDRW", 0, 2)
        swApp.CloseAllDocuments(True)

        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & PartName & ".SLDDRW", 3, 0, "", longstatus, longwarnings)
        boolstatus = Part.EditRebuild3()
        longstatus = Part.SaveAs3(SaveFolder & "\PDF Drawings\" & PartName & ".PDF", 0, 2)
        swApp.CloseAllDocuments(True)

    End Sub

    Public Sub BlankDrawings(PartName As String, BlkQty As Integer)
        'Exit Sub
        swApp.ActivateDoc2(PartName, False, longstatus)
        Part = swApp.ActiveDoc

        If Part Is Nothing Then
            MsgBox("Please open an appropriate Part Document")
            Exit Sub
        End If

        'Bounding Box
        Dim BBox As Object = StdFunc.BoundingBox()
        Dim xDim As Decimal = Abs(BBox(0)) + Abs(BBox(3))
        Dim yDim As Decimal = Abs(BBox(1)) + Abs(BBox(4))
        Dim zDim As Decimal = Abs(BBox(2)) + Abs(BBox(5))

        'Bounding Box - Flat
        boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditUnsuppress2()

        Dim BBoxFlat As Object = StdFunc.BoundingBox()
        Dim xDimFlat As Decimal = Abs(BBoxFlat(0)) + Abs(BBoxFlat(3))
        Dim yDimFlat As Decimal = Abs(BBoxFlat(1)) + Abs(BBoxFlat(4))
        Dim zDimFlat As Decimal = Abs(BBoxFlat(2)) + Abs(BBoxFlat(5))

        boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditSuppress2()

        'swApp.CloseAllDocuments(True)

        ' Sheet Scale
        Dim SScaleX As Integer = Ceiling((xDimFlat + zDim + xDim) / (0.297 - (0.03 + 0.04 + 0.04 + 0.03)))
        Dim SScaleY As Integer = Ceiling((zDim + yDimFlat) / (0.21 - (0.03 + 0.04 + 0.03)))
        Dim SScale As Integer = SScaleX
        If SScaleX < SScaleY Then
            SScale = SScaleY
        End If

        ' Adjust Values for Scale
        xDim /= SScale
        yDim /= SScale
        zDim /= SScale

        xDimFlat /= SScale
        yDimFlat /= SScale
        zDimFlat /= SScale

        ' Get Margins
        Dim ClearX As Decimal = (0.297 - (xDimFlat + 0.04 + zDim + 0.04 + xDim)) / 2
        If ClearX < 0.03 Then ClearX = 0.03
        Dim ClearY As Decimal = (0.21 - (yDimFlat + 0.04 + zDim)) / 2
        If ClearY < 0.03 Then ClearY = 0.03

        ' Calculate View Placements
        Dim xFrontFlat As Decimal = ClearX + xDimFlat / 2
        Dim yTopSec As Decimal = ClearY + zDim / 2
        Dim yFrontFlat As Decimal = yTopSec + zDim / 2 + 0.04 + yDimFlat / 2
        Dim xRightSec As Decimal = xFrontFlat + xDimFlat / 2 + 0.04 + zDim / 2
        Dim xIso As Decimal = xRightSec + zDim / 2 + 0.04 + xDim / 2

        ' New Drawing
        Part = swApp.NewDocument(DrawSheet, 2, 0.297, 0.21)
        Draw = Part

        swSheet = Draw.GetCurrentSheet()
        swSheet.SetProperties2(12, 12, 1, SScale, False, 0.297, 0.21, True)
        swSheet.SetTemplateName(DrawTemp)
        swSheet.ReloadTemplate(True)
        swApp.ActivateDoc2("Draw1 - Sheet1", False, longstatus)
        Part = swApp.ActiveDoc

        ' Views
        boolstatus = Draw.CreateFlatPatternViewFromModelView(PartName, "Default", xFrontFlat, yFrontFlat, 0)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Draw.ActivateView("Drawing View1")
        Part.ClearSelection2(True)

        'Front - Outside
        myView = Part.CreateDrawViewFromModelView3(PartName, "*Front", -xFrontFlat, yFrontFlat, 0)
        boolstatus = Draw.ActivateView("Drawing View2")
        Part.ClearSelection2(True)

        'Top - Section
        boolstatus = Draw.ActivateView("Drawing View2")
        skSegment = Part.SketchManager.CreateLine((xDim * SScale / 2) + 0.15, 0, 0, -(xDim * SScale / 2) - 0.15, 0, 0)
        boolstatus = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xFrontFlat, yTopSec, 0, "A", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View3")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View3", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewVerticalCenter, BaseView)
        Part.ClearSelection2(True)
        boolstatus = Part.ActivateSheet("Sheet1")

        'Right - Section
        boolstatus = Draw.ActivateView("Drawing View2")
        skSegment = Part.SketchManager.CreateLine(0, (yDim * SScale / 2) + 0.15, 0, 0, -(yDim * SScale / 2) - 0.15, 0)
        boolstatus = Part.Extension.SelectByID2("Line2", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xRightSec, yFrontFlat, 0, "B", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View4")
        Part.ClearSelection2(True)

        'Isometric
        myView = Part.CreateDrawViewFromModelView3(PartName, "*Dimetric", xIso, yFrontFlat, 0)
        boolstatus = Draw.ActivateView("Drawing View2")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View5", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject6(1, -1)
        boolstatus = DrawView.SetDisplayMode4(False, swDisplayMode_e.swSHADED, True, True, False)
        Part.ClearSelection2(True)

        'Dimensions--------------------------------------------------------------------------------------------------------------
        boolstatus = Draw.ActivateSheet("Sheet1")

        'Height - Flat  
        boolstatus = Draw.ActivateView("Drawing View1")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, yFrontFlat + (yDimFlat / 2), -5000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, yFrontFlat - (yDimFlat / 2), -5000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat - (xDimFlat / 2) - 0.015, yFrontFlat, 0)
        Part.ClearSelection2(True)

        'Width - Flat
        boolstatus = Part.ActivateView("Drawing View1")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat + (xDimFlat / 2), yFrontFlat, -5000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat - (xDimFlat / 2), yFrontFlat, -5000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat, yFrontFlat + (yDimFlat / 2) + 0.015, 0)
        Part.ClearSelection2(True)

        'Height - Top Section
        boolstatus = Part.ActivateView("Section View A-A")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, yTopSec - (zDim / 2), -5000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat - (xDim / 2) + (0.001 / SScale), yTopSec + (zDim / 2), -5000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat - (xDim / 2) - 0.015, yTopSec + (zDim / 2) + 0.015, 0)
        Part.ClearSelection2(True)

        'Width - Top Section
        boolstatus = Part.ActivateView("Section View A-A")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat - (xDim / 2), yTopSec, -5000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat + (xDim / 2), yTopSec, -5000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat, yTopSec + (zDim / 2) + 0.015, 0)
        Part.ClearSelection2(True)

        boolstatus = Part.ActivateView("Drawing View4")
        boolstatus = Part.Extension.SelectByID2("Drawing View4", "DRAWINGVIEW", xRightSec, yFrontFlat, 0, True, 0, Nothing, 0)
        Dim vAnn As Object
        vAnn = Draw.InsertModelAnnotations3(0, 32776, False, False, False, True)
        Part.ClearSelection2(True)

        'Clean Dimensions - Remove Angle Dimentions
        For i = 1 To 25
            boolstatus = Part.ActivateView("Drawing View4")
            boolstatus = Part.Extension.SelectByID2("D4@EdgeBend" & i & "@Draw3-SectionAssembly-2-1@Drawing View4", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
            boolstatus = Part.Extension.SelectByID2("DetailItem10@Drawing View4", "NOTE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
            boolstatus = Part.Extension.SelectByID2("D7@Sheet-Metal1@Draw2-SectionAssembly-2-1@Drawing View4", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
            boolstatus = Part.Extension.SelectByID2("D7@Sheet-Metal1@Draw1-SectionAssembly-2-1@Drawing View4", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
            boolstatus = Part.Extension.SelectByID2("D4@EdgeBend1@Draw1-SectionAssembly-2-1@Drawing View4", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
        Next i


        boolstatus = Part.ActivateView("Drawing View3")
        boolstatus = Part.Extension.SelectByID2("Drawing View3", "DRAWINGVIEW", xFrontFlat, yTopSec, 0, True, 0, Nothing, 0)
        vAnn = Draw.InsertModelAnnotations3(0, 32776, False, False, False, True)
        Part.ClearSelection2(True)

        For j = 1 To 5
            boolstatus = Part.ActivateView("Drawing View3")
            boolstatus = Part.Extension.SelectByID2("DetailItem9@Drawing View3", "NOTE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
            boolstatus = Part.Extension.SelectByID2("RD1@Drawing View3", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
            boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@Draw1-SectionAssembly-1-1@Drawing View3", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
            boolstatus = Part.Extension.SelectByID2("D7@Sheet-Metal1@Draw1-SectionAssembly-1-1@Drawing View3", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
        Next j

        ' Note
        Dim X, Y, Z As Decimal
        boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingBalloonStyle, 0, swBalloonStyle_e.swBS_None)

        X = Round(xDimFlat * SScale * 1000, 2)
        Y = Round(yDimFlat * SScale * 1000, 2)
        Z = Round(zDimFlat * SScale * 1000, 2)
        myNote = Draw.CreateText2(PartName & vbNewLine & "Qty - " & BlkQty & vbNewLine & X & "mm x " & Y & "mm x " & Z & "mm", xRightSec, yTopSec, 0, 0.004, 0)

        boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

        XValueList.Add(X)
        YValueList.Add(Y)
        ZValueList.Add(Z)

        ' Save As
        Part.ViewZoomtofit2()
        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & PartName & ".SLDDRW", 0, 2)
        swApp.CloseAllDocuments(True)

        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & PartName & ".SLDDRW", 3, 0, "", longstatus, longwarnings)
        boolstatus = Part.EditRebuild3()
        longstatus = Part.SaveAs3(SaveFolder & "\PDF Drawings\" & PartName & ".PDF", 0, 2)
        swApp.CloseAllDocuments(True)
    End Sub

    Public Sub DoorDrawings(PartName As String, DoorWth As Decimal, DoorHt As Decimal)
        'Exit Sub
        swApp.ActivateDoc2(PartName, False, longstatus)
        Part = swApp.ActiveDoc

        If Part Is Nothing Then
            MsgBox("Please open an appropriate Part Document")
            Exit Sub
        End If

        'Bounding Box
        Dim BBox As Object = StdFunc.BoundingBox()
        Dim xDim As Decimal = Abs(BBox(0)) + Abs(BBox(3))
        Dim yDim As Decimal = Abs(BBox(1)) + Abs(BBox(4))
        Dim zDim As Decimal = Abs(BBox(2)) + Abs(BBox(5))

        'Bounding Box - Flat
        boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditUnsuppress2()

        Dim BBoxFlat As Object = StdFunc.BoundingBox()
        Dim xDimFlat As Decimal = Abs(BBoxFlat(0)) + Abs(BBoxFlat(3))
        Dim yDimFlat As Decimal = Abs(BBoxFlat(1)) + Abs(BBoxFlat(4))
        Dim zDimFlat As Decimal = Abs(BBoxFlat(2)) + Abs(BBoxFlat(5))

        boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditSuppress2()

        'swApp.CloseAllDocuments(True)

        ' Sheet Scale
        Dim SScaleX As Integer = Ceiling((xDimFlat + zDim + xDim) / (0.297 - (0.03 + 0.04 + 0.04 + 0.03)))
        Dim SScaleY As Integer = Ceiling((zDim + yDimFlat) / (0.21 - (0.03 + 0.04 + 0.03)))
        Dim SScale As Integer = SScaleX
        If SScaleX < SScaleY Then
            SScale = SScaleY
        End If

        ' Adjust Values for Scale
        xDim /= SScale
        yDim /= SScale
        zDim /= SScale

        xDimFlat /= SScale
        yDimFlat /= SScale
        zDimFlat /= SScale

        ' Get Margins
        Dim ClearX As Decimal = (0.297 - (xDimFlat + 0.04 + zDim + 0.04 + xDim)) / 2
        If ClearX < 0.03 Then ClearX = 0.03
        Dim ClearY As Decimal = (0.21 - (yDimFlat + 0.04 + zDim)) / 2
        If ClearY < 0.03 Then ClearY = 0.03

        ' Calculate View Placements
        Dim xFrontFlat As Decimal = ClearX + xDimFlat / 2
        Dim yTopSec As Decimal = ClearY + zDim / 2
        Dim yFrontFlat As Decimal = yTopSec + zDim / 2 + 0.04 + yDimFlat / 2
        Dim xRightSec As Decimal = xFrontFlat + xDimFlat / 2 + 0.04 + zDim / 2
        Dim xIso As Decimal = xRightSec + zDim / 2 + 0.04 + xDim / 2

        ' New Drawing
        Part = swApp.NewDocument(DrawSheet, 2, 0.297, 0.21)
        Draw = Part

        swSheet = Draw.GetCurrentSheet()
        swSheet.SetProperties2(12, 12, 1, SScale, False, 0.297, 0.21, True)
        swSheet.SetTemplateName(DrawTemp)
        swSheet.ReloadTemplate(True)
        swApp.ActivateDoc2("Draw1 - Sheet1", False, longstatus)
        Part = swApp.ActiveDoc

        ' Views
        boolstatus = Draw.CreateFlatPatternViewFromModelView(PartName, "Default", xFrontFlat, yFrontFlat, 0)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Draw.ActivateView("Drawing View1")
        Part.ClearSelection2(True)

        'Front - Outside
        myView = Part.CreateDrawViewFromModelView3(PartName, "*Front", -xFrontFlat, yFrontFlat, 0)
        boolstatus = Draw.ActivateView("Drawing View2")
        Part.ClearSelection2(True)

        'Top - Section
        boolstatus = Draw.ActivateView("Drawing View2")
        skSegment = Part.SketchManager.CreateLine((xDim * SScale / 2) + 0.15, 0, 0, -(xDim * SScale / 2) - 0.15, 0, 0)
        boolstatus = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xFrontFlat, yTopSec, 0, "A", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View3")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View3", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewVerticalCenter, BaseView)
        Part.ClearSelection2(True)
        boolstatus = Part.ActivateSheet("Sheet1")

        'Right - Section
        boolstatus = Draw.ActivateView("Drawing View2")
        skSegment = Part.SketchManager.CreateLine(0, (yDim * SScale / 2) + 0.15, 0, 0, -(yDim * SScale / 2) - 0.15, 0)
        boolstatus = Part.Extension.SelectByID2("Line2", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xRightSec, yFrontFlat, 0, "B", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View4")
        Part.ClearSelection2(True)

        'Isometric
        myView = Part.CreateDrawViewFromModelView3(PartName, "*Dimetric", xIso, yFrontFlat, 0)
        boolstatus = Draw.ActivateView("Drawing View2")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View5", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject6(1, -1)
        boolstatus = DrawView.SetDisplayMode4(False, swDisplayMode_e.swSHADED, True, True, False)
        Part.ClearSelection2(True)

        'Dimensions--------------------------------------------------------------------------------------------------------------
        boolstatus = Draw.ActivateSheet("Sheet1")

        'Height - Flat  
        boolstatus = Draw.ActivateView("Drawing View1")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, yFrontFlat + (yDimFlat / 2), -5000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, yFrontFlat - (yDimFlat / 2), -5000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat - (xDimFlat / 2) - 0.015, yFrontFlat, 0)
        Part.ClearSelection2(True)

        'Width - Flat
        boolstatus = Part.ActivateView("Drawing View1")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat + (xDimFlat / 2), yFrontFlat, -5000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat - (xDimFlat / 2), yFrontFlat, -5000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat, yFrontFlat + (yDimFlat / 2) + 0.015, 0)
        Part.ClearSelection2(True)

        'Height - Top Section 
        boolstatus = Part.ActivateView("Section View A-A")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat - (((DoorWth / 2) - 0.02) / SScale), yTopSec - (0.0075 / SScale), -5000, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat - (((DoorWth / 2) - 0.0005) / SScale), yTopSec + (0.0075 / SScale), -5000, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat - 0.04, yTopSec, 0)
        Part.ClearSelection2(True)

        'Width - Top Section
        boolstatus = Part.ActivateView("Section View A-A")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat - (DoorWth / (2 * SScale)), yTopSec, -5000, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat + (DoorWth / (2 * SScale)), yTopSec, -5000, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat, yTopSec - 0.015, 0)
        Part.ClearSelection2(True)

        'Width - Right Section  
        boolstatus = Part.ActivateView("Section View B-B")
        boolstatus = Part.Extension.SelectByRay(xRightSec + (0.0075 / SScale), yFrontFlat + (((DoorHt / 2) - 0.04) / SScale), -5000, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xRightSec - (0.0075 / SScale), yFrontFlat + (((DoorHt / 2) - 0.0005) / SScale), -5000, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xRightSec, yFrontFlat + 0.055, 0)
        Part.ClearSelection2(True)

        'Height - Right Section
        boolstatus = Part.ActivateView("Section View B-B")
        boolstatus = Part.Extension.SelectByRay(xRightSec, yFrontFlat - (DoorHt / (2 * SScale)), -5000, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xRightSec, yFrontFlat + (DoorHt / (2 * SScale)), -5000, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xRightSec + 0.01, yFrontFlat, 0)
        Part.ClearSelection2(True)



        For a = 1 To 70
            boolstatus = Part.ActivateView("Drawing View4")
            boolstatus = Part.Extension.SelectByID2("DetailItem" & a & "@Drawing View4", "NOTE", 0, 0, 0, True, 0, Nothing, 0)
            Part.EditDelete()
            boolstatus = Part.ActivateView("Drawing View3")
            boolstatus = Part.Extension.SelectByID2("DetailItem" & a & "@Drawing View3", "NOTE", 0, 0, 0, True, 0, Nothing, 0)
            Part.EditDelete()
        Next



        ' Note
        Dim X, Y, Z As Decimal
        boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingBalloonStyle, 0, swBalloonStyle_e.swBS_None)

        X = Round(xDimFlat * SScale * 1000, 2)
        Y = Round(yDimFlat * SScale * 1000, 2)
        Z = Round(zDimFlat * SScale * 1000, 2)
        myNote = Draw.CreateText2(PartName & vbNewLine & "Qty - 1 " & vbNewLine & X & "mm x " & Y & "mm x " & Z & "mm", xRightSec, yTopSec, 0, 0.004, 0)

        boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

        XValueList.Add(X)
        YValueList.Add(Y)
        ZValueList.Add(Z)

        ' Save As
        Part.ViewZoomtofit2()
        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & PartName & ".SLDDRW", 0, 2)
        swApp.CloseAllDocuments(True)

        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & PartName & ".SLDDRW", 3, 0, "", longstatus, longwarnings)
        boolstatus = Part.EditRebuild3()
        longstatus = Part.SaveAs3(SaveFolder & "\PDF Drawings\" & PartName & ".PDF", 0, 2)
        swApp.CloseAllDocuments(True)
    End Sub

    Public Sub MotorPlateDrawings(Partname As String, FanNos As Integer, FanDia As Integer, PanelWth As Decimal, PanelHT As Decimal)
        'Exit Sub
        swApp.ActivateDoc2(Partname, False, longstatus)
        Part = swApp.ActiveDoc

        If Part Is Nothing Then
            MsgBox("Please open an appropriate Part Document")
            Exit Sub
        End If

        'Bounding Box
        Dim BBox As Object = StdFunc.BoundingBox()
        Dim xDim As Decimal = Abs(BBox(0)) + Abs(BBox(3))
        Dim yDim As Decimal = Abs(BBox(1)) + Abs(BBox(4))
        Dim zDim As Decimal = Abs(BBox(2)) + Abs(BBox(5))

        'Bounding Box - Flat
        boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditUnsuppress2()

        Dim BBoxFlat As Object = StdFunc.BoundingBox()
        Dim xDimFlat As Decimal = Abs(BBoxFlat(0)) + Abs(BBoxFlat(3))
        Dim yDimFlat As Decimal = Abs(BBoxFlat(1)) + Abs(BBoxFlat(4))
        Dim zDimFlat As Decimal = Abs(BBoxFlat(2)) + Abs(BBoxFlat(5))

        boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditSuppress2()

        'swApp.CloseAllDocuments(True)

        ' Sheet Scale
        Dim SScaleX As Integer = Ceiling((xDimFlat + zDim + xDim) / (0.297 - (0.03 + 0.04 + 0.04 + 0.03)))
        Dim SScaleY As Integer = Ceiling((zDim + yDimFlat) / (0.21 - (0.03 + 0.04 + 0.03)))
        Dim SScale As Integer = SScaleX
        If SScaleX < SScaleY Then
            SScale = SScaleY
        End If

        ' Adjust Values for Scale
        xDim /= SScale
        yDim /= SScale
        zDim /= SScale

        xDimFlat /= SScale
        yDimFlat /= SScale
        zDimFlat /= SScale

        ' Get Margins
        Dim ClearX As Decimal = (0.297 - (xDimFlat + 0.04 + zDim + 0.04 + xDim)) / 2
        If ClearX < 0.03 Then ClearX = 0.03
        Dim ClearY As Decimal = (0.21 - (yDimFlat + 0.04 + zDim)) / 2
        If ClearY < 0.03 Then ClearY = 0.03

        ' Calculate View Placements
        Dim xFrontFlat As Decimal = ClearX + xDimFlat / 2
        Dim yTopSec As Decimal = ClearY + zDim / 2
        Dim yFrontFlat As Decimal = yTopSec + zDim / 2 + 0.04 + yDimFlat / 2
        Dim xRightSec As Decimal = xFrontFlat + xDimFlat / 2 + 0.04 + zDim / 2
        Dim xIso As Decimal = xRightSec + zDim / 2 + 0.04 + xDim / 2

        ' New Drawing
        Part = swApp.NewDocument(DrawSheet, 2, 0.297, 0.21)
        Draw = Part

        swSheet = Draw.GetCurrentSheet()
        swSheet.SetProperties2(12, 12, 1, SScale, False, 0.297, 0.21, True)
        swSheet.SetTemplateName(DrawTemp)
        swSheet.ReloadTemplate(True)
        swApp.ActivateDoc2("Draw1 - Sheet1", False, longstatus)
        Part = swApp.ActiveDoc

        ' Views
        boolstatus = Draw.CreateFlatPatternViewFromModelView(Partname, "Default", xFrontFlat, yFrontFlat, 0)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Draw.ActivateView("Drawing View1")
        Part.ClearSelection2(True)

        'Front - Outside
        myView = Part.CreateDrawViewFromModelView3(Partname, "*Front", -xFrontFlat, yFrontFlat, 0)
        boolstatus = Draw.ActivateView("Drawing View2")
        Part.ClearSelection2(True)

        'Top - Section
        boolstatus = Draw.ActivateView("Drawing View2")
        'skSegment = Part.SketchManager.CreateLine((xDim * SScale / 2) + 0.15, 0, 0, -(xDim * SScale / 2) - 0.15, 0, 0)
        skSegment = Part.SketchManager.CreateLine((xDim * SScale / 2) + 0.15, (PanelWth / 2 * 87) / 100, 0, -(xDim * SScale / 2) - 0.15, (PanelWth / 2 * 87) / 100, 0)
        'skSegment = Part.SketchManager.CreateLine((xDim * SScale / 2) + 0.15, 0.345, 0, -(xDim * SScale / 2) - 0.15, 0.345, 0)
        boolstatus = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xFrontFlat, yTopSec, 0, "A", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View3")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View3", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewVerticalCenter, BaseView)
        Part.ClearSelection2(True)
        boolstatus = Part.ActivateSheet("Sheet1")

        'Right - Section
        boolstatus = Draw.ActivateView("Drawing View2")
        skSegment = Part.SketchManager.CreateLine((PanelHT / 2 * 87) / 100, (yDim * SScale / 2) + 0.15, 0, (PanelHT / 2 * 87) / 100, -(yDim * SScale / 2) - 0.15, 0)
        boolstatus = Part.Extension.SelectByID2("Line2", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xRightSec, yFrontFlat, 0, "B", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View4")
        Part.ClearSelection2(True)

        'Isometric
        myView = Part.CreateDrawViewFromModelView3(Partname, "*Dimetric", xIso, yFrontFlat, 0)
        boolstatus = Draw.ActivateView("Drawing View2")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View5", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject6(1, -1)
        boolstatus = DrawView.SetDisplayMode4(False, swDisplayMode_e.swSHADED, True, True, False)
        Part.ClearSelection2(True)

        'Dimensions--------------------------------------------------------------------------------------------------------------
        boolstatus = Draw.ActivateSheet("Sheet1")

        Select Case FanDia
            Case = "310"
                boolstatus = Part.ActivateView("Drawing View1")
                boolstatus = Part.Extension.SelectByRay(0.0793769120675188, 0.136042810878273, -4999.999, 0, 0, -1, 0.000470660536879686, 1, False, 0, 0)
                myDisplayDim = Part.AddDimension2(0.124851701433891, 0.179078974161163, 0)
                Part.ClearSelection2(True)

            Case = "350"
                boolstatus = Part.ActivateView("Drawing View1")
                boolstatus = Part.Extension.SelectByRay(0.0801934912410846, 0.137107886344444, -4999.999, 0, 0, -1, 0.000436206866635698, 1, False, 0, 0)
                myDisplayDim = Part.AddDimension2(0.124851701433891, 0.179078974161163, 0)
                Part.ClearSelection2(True)

            Case = "400"
                boolstatus = Part.ActivateView("Drawing View1")
                boolstatus = Part.Extension.SelectByRay(0.086323111375232, 0.137397749515128, -4999.999, 0, 0, -1, 0.000682971434621279, 1, False, 0, 0)
                myDisplayDim = Part.AddDimension2(0.124851701433891, 0.179078974161163, 0)
                Part.ClearSelection2(True)

            Case = "450"
                boolstatus = Part.ActivateView("Drawing View1")
                boolstatus = Part.Extension.SelectByRay(0.0823312252964427, 0.141546640281291, -4999.999, 0, 0, -1, 0.000959525691699605, 1, False, 0, 0)
                myDisplayDim = Part.AddDimension2(0.124851701433891, 0.179078974161163, 0)
                Part.ClearSelection2(True)

            Case = "500"
                boolstatus = Part.ActivateView("Drawing View1")
                boolstatus = Part.Extension.SelectByRay(0.0822106270040131, 0.143174781508755, -4999.999, 0, 0, -1, 0.000682971434621279, 1, False, 0, 0)
                myDisplayDim = Part.AddDimension2(0.124851701433891, 0.179078974161163, 0)
                Part.ClearSelection2(True)

            Case = "560"
                boolstatus = Part.ActivateView("Drawing View1")
                boolstatus = Part.Extension.SelectByRay(0.0852274384087651, 0.144804898553848, -5000, 0, 0, -1, 0.001, 1, False, 0, 0)
                myDisplayDim = Part.AddDimension2(0.124701013579276, 0.178928286306549, 0)
                Part.ClearSelection2(True)

        End Select


        'Height - Flat  
        boolstatus = Draw.ActivateView("Drawing View1")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, yFrontFlat + (yDimFlat / 2), -4500, 0, 0, -1, 0.001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, yFrontFlat - (yDimFlat / 2), -4500, 0, 0, -1, 0.001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddVerticalDimension2(xFrontFlat - (xDimFlat / 2) - 0.015, yFrontFlat, 0)
        Part.ClearSelection2(True)

        'Width - Flat
        boolstatus = Part.ActivateView("Drawing View1")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat + (xDimFlat / 2), yFrontFlat, -5500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat - (xDimFlat / 2), yFrontFlat, -5500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat, yFrontFlat + (yDimFlat / 2) + 0.015, 0)
        Part.ClearSelection2(True)

        'Height - Top Section
        boolstatus = Part.ActivateView("Section View A-A")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat + (xDim / 2), yTopSec + (zDim / 2), -5000, 0, 0, -1, 0.001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat + (xDim / 2) - (zDim * 20), yTopSec - (zDim / 2), -5000, 0, 0, -1, 0.001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat + (xDim / 2) + 0.015, yTopSec - (zDim / 2) + 0.015, 0)
        Part.ClearSelection2(True)

        'Width - Top Section
        boolstatus = Part.ActivateView("Section View A-A")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat - (xDim / 2), yTopSec, -5000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat + (xDim / 2), yTopSec, -5000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat, yTopSec + (zDim / 2) + 0.015, 0)
        Part.ClearSelection2(True)

        'Height - Right Section
        boolstatus = Part.ActivateView("Section View B-B")
        boolstatus = Part.Extension.SelectByRay(xRightSec, yFrontFlat + (yDim / 2), -5000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xRightSec, yFrontFlat - (yDim / 2), -5000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xRightSec + (zDim / 2) + 0.01, yFrontFlat, 0)
        Part.ClearSelection2(True)

        'Width - Right Section
        boolstatus = Part.ActivateView("Section View B-B")
        boolstatus = Part.Extension.SelectByRay(xRightSec - (zDim * 5), yFrontFlat + (yDim / 2) - (zDim * 10), -5000, 0, 0, -1, 0.001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xRightSec + (zDim * 5), yFrontFlat + (yDim / 2) - zDim, -5000, 0, 0, -1, 0.001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xRightSec - (zDim / 2) - 0.015, yFrontFlat + (yDim / 2) + 0.015, 0)
        Part.ClearSelection2(True)



        Dim vAnnotations As Object
        boolstatus = Part.ActivateView("Drawing View3")
        boolstatus = Part.Extension.SelectByID2("Drawing View3", "DRAWINGVIEW", xFrontFlat, yTopSec, 0, True, 0, Nothing, 0)
        vAnnotations = Draw.InsertModelAnnotations3(0, 32776, False, False, False, True)
        Part.ClearSelection2(True)
        boolstatus = Part.ActivateView("Drawing View4")
        boolstatus = Part.Extension.SelectByID2("Drawing View4", "DRAWINGVIEW", xRightSec, yFrontFlat, 0, True, 0, Nothing, 0)
        vAnnotations = Draw.InsertModelAnnotations3(0, 32776, False, False, False, True)
        Part.ClearSelection2(True)


        ' Remove "SECTION" Note
        For i = 1 To 80
            boolstatus = Part.ActivateView("Drawing View3")
            boolstatus = Part.Extension.SelectByID2("DetailItem" & i & "@Drawing View3", "NOTE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
            boolstatus = Part.Extension.SelectByID2("RD" & i & "@Drawing View3", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
        Next i

        For j = 1 To 80
            boolstatus = Part.ActivateView("Drawing View4")
            boolstatus = Part.Extension.SelectByID2("DetailItem" & j & "@Drawing View4", "NOTE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
            boolstatus = Part.Extension.SelectByID2("D4@EdgeBend" & j & "@Draw1-SectionAssembly-2-1@Drawing View4", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
            boolstatus = Part.Extension.SelectByID2("RD" & j & "@Drawing View4", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
        Next


        ' Note
        Dim X, Y, Z As Decimal
        boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingBalloonStyle, 0, swBalloonStyle_e.swBS_None)

        X = Round(xDimFlat * SScale * 1000, 2)
        Y = Round(yDimFlat * SScale * 1000, 2)
        Z = Round(zDimFlat * SScale * 1000, 2)
        myNote = Draw.CreateText2(Partname & vbNewLine & "Qty - " & FanNos & vbNewLine & X & "mm x " & Y & "mm x " & Z & "mm", xRightSec, yTopSec, 0, 0.004, 0)

        boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

        XValueList.Add(X)
        YValueList.Add(Y)
        ZValueList.Add(Z)

        ' Save As
        Part.ViewZoomtofit2()
        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & Partname & ".SLDDRW", 0, 2)
        swApp.CloseAllDocuments(True)

        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & Partname & ".SLDDRW", 3, 0, "", longstatus, longwarnings)
        boolstatus = Part.EditRebuild3()
        longstatus = Part.SaveAs3(SaveFolder & "\PDF Drawings\" & Partname & ".PDF", 0, 2)
        swApp.CloseAllDocuments(True)

    End Sub

    Public Sub VerCDrawings(XFans As Integer)
        'Exit Sub
        swApp.ActivateDoc2(AHUName & "_04A_Vertical Channel", False, longstatus)
        Part = swApp.ActiveDoc

        If Part Is Nothing Then
            MsgBox("Please open an appropriate Part Document")
            Exit Sub
        End If

        'Bounding Box
        Dim BBox As Object = StdFunc.BoundingBox()
        Dim xDim As Decimal = Abs(BBox(0)) + Abs(BBox(3))
        Dim yDim As Decimal = Abs(BBox(1)) + Abs(BBox(4))
        Dim zDim As Decimal = Abs(BBox(2)) + Abs(BBox(5))

        'Bounding Box - Flat
        boolstatus = Part.Extension.SelectByID2("Flat-Pattern23", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditUnsuppress2()

        Dim BBoxFlat As Object = StdFunc.BoundingBox()
        Dim xDimFlat As Decimal = Abs(BBoxFlat(0)) + Abs(BBoxFlat(3))
        Dim yDimFlat As Decimal = Abs(BBoxFlat(1)) + Abs(BBoxFlat(4))
        Dim zDimFlat As Decimal = Abs(BBoxFlat(2)) + Abs(BBoxFlat(5))

        boolstatus = Part.Extension.SelectByID2("Flat-Pattern23", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditSuppress2()

        'swApp.CloseAllDocuments(True)

        ' Sheet Scale
        Dim SScaleX As Integer = Ceiling((xDimFlat + zDim + xDim) / (0.297 - (0.03 + 0.03 + 0.03)))                        '(0.297 - (0.03 + 0.04 + 0.04 + 0.03)))
        Dim SScaleY As Integer = Ceiling((zDim + yDimFlat) / (0.21 - (0.03 + 0.02 + 0.03)))                               ' (0.21 - (0.03 + 0.04 + 0.03)))
        Dim SScale As Integer = SScaleX
        If SScaleX < SScaleY Then
            SScale = SScaleY
        End If

        ' Adjust Values for Scale
        xDim /= SScale
        yDim /= SScale
        zDim /= SScale

        xDimFlat /= SScale
        yDimFlat /= SScale
        zDimFlat /= SScale

        ' Get Margins
        Dim ClearX As Decimal = (0.297 - (xDimFlat + 0.04 + zDim + 0.04 + xDim)) / 2
        If ClearX < 0.03 Then ClearX = 0.03
        Dim ClearY As Decimal = (0.21 - (yDimFlat + 0.04 + zDim)) / 2
        If ClearY < 0.03 Then ClearY = 0.03

        ' Calculate View Placements
        Dim xFrontFlat As Decimal = ClearX + xDimFlat / 2
        Dim yTopSec As Decimal = ClearY + zDim / 2
        Dim yFrontFlat As Decimal = yTopSec + zDim / 2 + 0.04 + yDimFlat / 2
        Dim xRightSec As Decimal = xFrontFlat + xDimFlat / 2 + 0.04 + zDim / 2
        Dim xIso As Decimal = xRightSec + zDim / 2 + 0.04 + xDim / 2

        Dim X, Y, Z As Decimal
        X = Round(xDimFlat * SScale * 1000, 2)
        Y = Round(yDimFlat * SScale * 1000, 2)
        Z = Round(zDimFlat * SScale * 1000, 2)

        ' New Drawing
        Part = swApp.NewDocument(DrawSheet, 2, 0.297, 0.21)
        Draw = Part

        swSheet = Draw.GetCurrentSheet()
        swSheet.SetProperties2(12, 12, 1, SScale, False, 0.297, 0.21, True)
        swSheet.SetTemplateName(DrawTemp)
        swSheet.ReloadTemplate(True)
        swApp.ActivateDoc2("Draw1 - Sheet1", False, longstatus)
        Part = swApp.ActiveDoc

        PartName = AHUName & "_04A_Vertical Channel"

        ' Views
        boolstatus = Draw.CreateFlatPatternViewFromModelView(PartName, "Default", xFrontFlat, yFrontFlat - 0.005, 0)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Draw.ActivateView("Drawing View1")
        Part.ClearSelection2(True)

        'Front - Outside
        myView = Part.CreateDrawViewFromModelView3(PartName, "*Front", -xFrontFlat, yFrontFlat - 0.005, 0)
        boolstatus = Draw.ActivateView("Drawing View2")
        Part.ClearSelection2(True)

        'Top - Section
        boolstatus = Draw.ActivateView("Drawing View2")
        skSegment = Part.SketchManager.CreateLine((xDim * SScale / 2) + 0.15, 0, 0, -(xDim * SScale / 2) - 0.15, 0, 0)
        boolstatus = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xFrontFlat, yTopSec, 0, "A", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View3")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View3", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewVerticalCenter, BaseView)
        Part.ClearSelection2(True)

        'Right - Section
        boolstatus = Draw.ActivateView("Drawing View2")
        skSegment = Part.SketchManager.CreateLine(0, (yDim * SScale / 2) + 0.15, 0, 0, -(yDim * SScale / 2) - 0.15, 0)
        boolstatus = Part.Extension.SelectByID2("Line2", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xRightSec, yFrontFlat, 0, "B", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View4")
        Part.ClearSelection2(True)

        'Isometric
        myView = Part.CreateDrawViewFromModelView3(PartName, "*Dimetric", xIso + 0.05, yFrontFlat - 0.005, 0)
        boolstatus = Draw.ActivateView("Drawing View2")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View5", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject6(1, -1)
        boolstatus = DrawView.SetDisplayMode4(False, swDisplayMode_e.swSHADED, True, True, False)
        Part.ClearSelection2(True)

        For a = 1 To 50
            boolstatus = Part.ActivateView("Drawing View3")
            boolstatus = Part.Extension.SelectByID2("DetailItem" & a & "@Drawing View3", "NOTE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
            boolstatus = Part.ActivateView("Drawing View4")
            boolstatus = Part.Extension.SelectByID2("DetailItem" & a & "@Drawing View4", "NOTE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
        Next

        'Dimensions--------------------------------------------------------------------------------------------------------------xFrontFlat, yFrontFlat - 0.005,
        boolstatus = Draw.ActivateSheet("Sheet1")

        'Width - Flat
        boolstatus = Part.ActivateView("Drawing View1")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat + (0.08281 / SScale), yFrontFlat - 0.005, -5500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat - (0.08281 / SScale), yFrontFlat - 0.005, -5500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddHorizontalDimension2(xFrontFlat + 0.02, yFrontFlat + 0.06, 0)
        Part.ClearSelection2(True)

        'Height - Flat
        boolstatus = Part.ActivateView("Drawing View1")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, -((Y / (SScale * 2000)) - (yFrontFlat - 0.005)), -5500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, (Y / (SScale * 2000)) + (yFrontFlat - 0.005), -5500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddVerticalDimension2(xFrontFlat - 0.035, yFrontFlat, 0)
        Part.ClearSelection2(True)

        Dim NewMidPt As Decimal
        NewMidPt = ((0.035 / SScale) + yTopSec)

        'Width - A-A
        boolstatus = Part.ActivateView("Drawing View3")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat - (0.05 / SScale), NewMidPt, -5000, 0, 0, -1, 0.001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat + (0.05 / SScale), NewMidPt, -5000, 0, 0, -1, 0.001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat + 0.02, NewMidPt + 0.007, 0)
        Part.ClearSelection2(True)

        'Width 2 - A-A
        boolstatus = Part.ActivateView("Drawing View3")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat - (0.05 / SScale), NewMidPt, -5000, 0, 0, -1, 0.001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat - (0.04 / SScale), NewMidPt - (0.014 / SScale), -5000, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat - (0.04 / SScale), NewMidPt - 0.01, 0)
        Part.ClearSelection2(True)

        'Height - A-A
        boolstatus = Part.ActivateView("Drawing View3")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, NewMidPt + (0.015 / SScale), -5500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat - (0.0435 / SScale), NewMidPt - (0.015 / SScale), -5500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddVerticalDimension2(xFrontFlat - 0.01, NewMidPt - 0.0001, 0)
        Part.ClearSelection2(True)


        NewMidPt = yFrontFlat - 0.005

        'Width - B-B 
        boolstatus = Part.ActivateView("Drawing View4")
        boolstatus = Part.Extension.SelectByRay(xRightSec + (0.05 / SScale), NewMidPt + ((WallHt - 4) / (2000 * SScale)), -5500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xRightSec - (0.05 / SScale), NewMidPt, -5500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xRightSec, yFrontFlat + 0.057, 0)
        Part.ClearSelection2(True)

        'Height - B-B
        boolstatus = Part.ActivateView("Drawing View4")
        boolstatus = Part.Extension.SelectByRay(xRightSec, -(((WallHt - 4) / (SScale * 2000)) - NewMidPt), -5500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xRightSec, ((WallHt - 4) / (SScale * 2000) + NewMidPt), -5500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xRightSec - 0.01, NewMidPt, 0)
        Part.ClearSelection2(True)



        ' Note
        boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingBalloonStyle, 0, swBalloonStyle_e.swBS_None)
        myNote = Draw.CreateText2(AHUName & "_04A_Vertical Channel" & vbNewLine & "Qty - " & XFans + 1 & vbNewLine & Y & "mm x " & X & "mm x " & "2.00mm", xRightSec, yTopSec, 0, 0.004, 0)
        boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

        XValueList.Add(X)
        YValueList.Add(Y)
        ZValueList.Add(Z)

        ' Save As
        Part.ViewZoomtofit2()
        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_04A_Vertical Channel.SLDDRW", 0, 2)
        swApp.CloseAllDocuments(True)

        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_04A_Vertical Channel.SLDDRW", 3, 0, "", longstatus, longwarnings)
        boolstatus = Part.EditRebuild3()
        longstatus = Part.SaveAs3(SaveFolder & "\PDF Drawings\" & AHUName & "_04A_Vertical Channel.PDF", 0, 2)
        swApp.CloseAllDocuments(True)

    End Sub

    Public Sub SideLDrawings()
        'Exit Sub

        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_03A_Side_L_Left.SLDPRT", 1, 32, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_03A_Side_L_Right.SLDPRT", 1, 32, "", longstatus, longwarnings)

        swApp.ActivateDoc2(AHUName & "_03A_Side_L_Left", False, longstatus)
        Part = swApp.ActiveDoc

        If Part Is Nothing Then
            MsgBox("Please open an appropriate Part Document")
            Exit Sub
        End If

        'Bounding Box
        Dim BBox As Object = StdFunc.BoundingBox()
        Dim xDim As Decimal = Abs(BBox(0)) + Abs(BBox(3))
        Dim yDim As Decimal = Abs(BBox(1)) + Abs(BBox(4))
        Dim zDim As Decimal = Abs(BBox(2)) + Abs(BBox(5))

        'Bounding Box - Flat
        boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditUnsuppress2()

        Dim BBoxFlat As Object = StdFunc.BoundingBox()
        Dim xDimFlat As Decimal = Abs(BBoxFlat(0)) + Abs(BBoxFlat(3))
        Dim yDimFlat As Decimal = Abs(BBoxFlat(1)) + Abs(BBoxFlat(4))
        Dim zDimFlat As Decimal = Abs(BBoxFlat(2)) + Abs(BBoxFlat(5))

        boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditSuppress2()

        'swApp.CloseAllDocuments(True)

        ' Sheet Scale
        Dim SScaleX As Integer = Ceiling((xDimFlat + zDim + xDim) / (0.297 - (0.03 + 0.03 + 0.03)))                  ' (0.297 - (0.03 + 0.04 + 0.04 + 0.03)))
        Dim SScaleY As Integer = Ceiling((zDim + yDimFlat) / (0.21 - (0.03 + 0.04 + 0.03)))                         ' (0.21 - (0.03 + 0.04 + 0.03)))
        Dim SScale As Integer = SScaleX
        If SScaleX < SScaleY Then
            SScale = SScaleY
        End If

        ' Adjust Values for Scale
        xDim /= SScale
        yDim /= SScale
        zDim /= SScale

        xDimFlat /= SScale
        yDimFlat /= SScale
        zDimFlat /= SScale

        ' Get Margins
        Dim ClearX As Decimal = (0.297 - (xDimFlat + 0.04 + zDim + 0.04 + xDim)) / 2
        If ClearX < 0.03 Then ClearX = 0.03
        Dim ClearY As Decimal = (0.21 - (yDimFlat + 0.045 + zDim)) / 2
        If ClearY < 0.03 Then ClearY = 0.03

        ' Calculate View Placements
        Dim xFrontFlat As Decimal = (ClearX / 2) - xDimFlat - 0.005
        Dim yTopSec As Decimal = ClearY + zDim / 2
        Dim yFrontFlat As Decimal = yTopSec + zDim / 2 + 0.055 + yDimFlat / 2
        Dim xRightSec As Decimal = xFrontFlat + xDimFlat / 2 + 0.04 + zDim / 2
        Dim xIso As Decimal = xRightSec + zDim / 2 + 0.06 + xDim / 2

        ' New Drawing
        Part = swApp.NewDocument(DrawSheet, 2, 0.297, 0.21)
        Draw = Part

        swSheet = Draw.GetCurrentSheet()
        swSheet.SetProperties2(12, 12, 1, SScale, False, 0.297, 0.21, True)
        swSheet.SetTemplateName(DrawTemp)
        swSheet.ReloadTemplate(True)
        swApp.ActivateDoc2("Draw1 - Sheet1", False, longstatus)
        Part = swApp.ActiveDoc
        Part = Draw

        PartName = AHUName & "_03A_Side_L_Left"

        'Flat
        boolstatus = Draw.CreateFlatPatternViewFromModelView(PartName, "Default", xFrontFlat, yFrontFlat, 0)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Draw.ActivateView("Drawing View1")
        Part.ClearSelection2(True)


        'Front - Outside
        myView = Part.CreateDrawViewFromModelView3(PartName, "*Front", -xFrontFlat, yFrontFlat, 0)
        boolstatus = Draw.ActivateView("Drawing View2")
        Part.ClearSelection2(True)

        'Top - Section
        boolstatus = Draw.ActivateView("Drawing View2")
        skSegment = Part.SketchManager.CreateLine((xDim * SScale / 2) + 0.15, 0, 0, -(xDim * SScale / 2) - 0.15, 0, 0)
        boolstatus = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xFrontFlat, yTopSec + ClearY - 0.005, 0, "A", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View3")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View3", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewVerticalCenter, BaseView)
        Part.ClearSelection2(True)
        boolstatus = Part.ActivateSheet("Sheet1")

        'Isometric
        myView = Part.CreateDrawViewFromModelView3(AHUName & "_03A_Side_L_Left", "*Isometric", yDimFlat, yFrontFlat, 0)
        boolstatus = Draw.ActivateView("Drawing View4")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View4", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject6(1, -1)
        boolstatus = DrawView.SetDisplayMode4(False, swDisplayMode_e.swSHADED, True, True, False)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View3", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewVerticalCenter, BaseView)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View4", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewHorizontalCenter, BaseView)
        Part.ClearSelection2(True)

        'Dimensions--------------------------------------------------------------------------------------------------------------

        'Width - Flat
        boolstatus = Part.ActivateView("Drawing View1")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat + (0.0482 / SScale), yFrontFlat, -5000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat - (0.0482 / SScale), yFrontFlat, -5000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat + 0.02, yFrontFlat + 0.055, 0)
        Part.ClearSelection2(True)

        'Height - Flat
        boolstatus = Part.ActivateView("Drawing View1")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, (yFrontFlat - ((WallHt - 100) / (SScale * 2000))), -5000, 0, 0, -1, 0.001, 1, True, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, (yFrontFlat + ((WallHt - 100) / (SScale * 2000))), -5000, 0, 0, -1, 0.001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat - 0.01, yFrontFlat, 0)
        Part.ClearSelection2(True)

        'Width - Section   
        boolstatus = Part.ActivateView("Drawing View3")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat + (0.025 / SScale), (yTopSec + ClearY - 0.005) + (0.0245 / SScale), -8500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat - (0.025 / SScale), (yTopSec + ClearY - 0.005), -8500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat, (yTopSec + ClearY - 0.005) + 0.01, 0)
        Part.ClearSelection2(True)

        'Height - Section
        boolstatus = Part.ActivateView("Drawing View3")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat - (0.024 / SScale), (yTopSec + ClearY - 0.005) - (0.025 / SScale), -5000, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat + (0.014 / SScale), (yTopSec + ClearY - 0.005) + (0.025 / SScale), -5000, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat - 0.01, (yTopSec + ClearY - 0.005), 0)
        Part.ClearSelection2(True)

        ' Note
        Dim X, Y, Z As Decimal
        boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingBalloonStyle, 0, swBalloonStyle_e.swBS_None)

        X = Round(xDimFlat * SScale * 1000, 2)
        Y = Round(yDimFlat * SScale * 1000, 2)
        Z = Round(zDimFlat * SScale * 1000, 2)

        myNote = Draw.CreateText2(AHUName & "_03A_Side_L_Left" & vbNewLine & "Qty - 1 " & vbNewLine & Y & "mm x " & X & "mm x " & "2.00mm", ClearY, ClearY + xDim, 0, 0.004, 0)

        boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

        XValueList.Add(X)
        YValueList.Add(Y)
        ZValueList.Add(Z)

        '======================= Side L Right =============================
        swApp.ActivateDoc2(AHUName & "_03A_Side_L_Right", False, longstatus)
        Part = swApp.ActiveDoc

        If Part Is Nothing Then
            MsgBox("Please open an appropriate Part Document")
            Exit Sub
        End If

        swSheet = Draw.GetCurrentSheet()
        swSheet.SetProperties2(12, 12, 1, SScale, False, 0.297, 0.21, True)
        swSheet.SetTemplateName(DrawTemp)
        swSheet.ReloadTemplate(True)
        swApp.ActivateDoc2("Draw1 - Sheet1", False, longstatus)
        Part = swApp.ActiveDoc
        Part = Draw

        PartName = AHUName & "_03A_Side_L_Right"

        'Flat
        boolstatus = Draw.CreateFlatPatternViewFromModelView(PartName, "Default", xIso + xFrontFlat, yFrontFlat, 0)
        boolstatus = Part.Extension.SelectByID2("Drawing View5", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Draw.ActivateView("Drawing View5")
        Part.ClearSelection2(True)

        'Front - Outside
        myView = Part.CreateDrawViewFromModelView3(PartName, "*Front", (xIso + xFrontFlat + xIso), yFrontFlat, 0)
        boolstatus = Draw.ActivateView("Drawing View6")
        Part.ClearSelection2(True)

        'Top - Section
        boolstatus = Draw.ActivateView("Drawing View6")
        skSegment = Part.SketchManager.CreateLine((xDim * SScale / 2) + 0.15, 0, 0, -(xDim * SScale / 2) - 0.15, 0, 0)
        boolstatus = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xIso + xFrontFlat, yTopSec + ClearY - 0.005, 0, "A", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View7")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View7", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View5", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewVerticalCenter, BaseView)
        Part.ClearSelection2(True)
        boolstatus = Part.ActivateSheet("Sheet1")

        'Isometric
        myView = Part.CreateDrawViewFromModelView3(AHUName & "_03A_Side_L_Right", "*Isometric", xIso + yDimFlat, yFrontFlat, 0)
        boolstatus = Draw.ActivateView("Drawing View8")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View8", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject6(1, -1)
        boolstatus = DrawView.SetDisplayMode4(False, swDisplayMode_e.swSHADED, True, True, False)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View8", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View5", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewHorizontalCenter, BaseView)
        Part.ClearSelection2(True)

        boolstatus = Part.ActivateView("Drawing View3")
        boolstatus = Part.Extension.SelectByID2("DetailItem3@Drawing View3", "NOTE", 0.0550520001073925, 0.0472070823504804, 0, False, 0, Nothing, 0)
        Part.EditDelete()
        boolstatus = Part.ActivateView("Drawing View7")
        boolstatus = Part.Extension.SelectByID2("DetailItem7@Drawing View7", "NOTE", 0.189095899523416, 0.0479559309505699, 0, False, 0, Nothing, 0)
        Part.EditDelete()

        'Dimensions--------------------------------------------------------------------------------------------------------------

        'Width - Flat 
        boolstatus = Part.ActivateView("Drawing View5")
        boolstatus = Part.Extension.SelectByRay((xIso + xFrontFlat) + (0.0482 / SScale), yFrontFlat, -5000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay((xIso + xFrontFlat) - (0.0482 / SScale), yFrontFlat, -5000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2((xIso + xFrontFlat) + 0.02, yFrontFlat + 0.055, 0)
        Part.ClearSelection2(True)

        'Height - Flat
        boolstatus = Part.ActivateView("Drawing View5")
        boolstatus = Part.Extension.SelectByRay((xIso + xFrontFlat), (yFrontFlat - ((WallHt - 100) / (SScale * 2000))), -5000, 0, 0, -1, 0.001, 1, True, 0, 0)
        boolstatus = Part.Extension.SelectByRay((xIso + xFrontFlat), (yFrontFlat + ((WallHt - 100) / (SScale * 2000))), -5000, 0, 0, -1, 0.001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2((xIso + xFrontFlat) - 0.01, yFrontFlat, 0)
        Part.ClearSelection2(True)


        'Width - Section   
        boolstatus = Part.ActivateView("Drawing View7")
        boolstatus = Part.Extension.SelectByRay((xIso + xFrontFlat) + (0.025 / SScale), (yTopSec + ClearY - 0.005) - (0.024 / SScale), -8500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay((xIso + xFrontFlat) - (0.025 / SScale), (yTopSec + ClearY - 0.005), -8500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2((xIso + xFrontFlat) + 0.015, (yTopSec + ClearY - 0.005) - 0.007, 0)
        Part.ClearSelection2(True)

        'Height - Section
        boolstatus = Part.ActivateView("Drawing View7")
        boolstatus = Part.Extension.SelectByRay((xIso + xFrontFlat) - (0.024 / SScale), (yTopSec + ClearY - 0.005) + (0.025 / SScale), -5000, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay((xIso + xFrontFlat) + (0.014 / SScale), (yTopSec + ClearY - 0.005) - (0.025 / SScale), -5000, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2((xIso + xFrontFlat) - 0.01, (yTopSec + ClearY - 0.005), 0)
        Part.ClearSelection2(True)

        ' Note
        boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingBalloonStyle, 0, swBalloonStyle_e.swBS_None)

        X = Round(xDimFlat * SScale * 1000, 2)
        Y = Round(yDimFlat * SScale * 1000, 2)
        Z = Round(zDimFlat * SScale * 1000, 2)

        myNote = Draw.CreateText2(AHUName & "_03A_Side_L_Right" & vbNewLine & "Qty - 1 " & vbNewLine & Y & "mm x " & X & "mm x " & "2.00mm", xIso + xFrontFlat, ClearY, 0, 0.004, 0)

        boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

        XValueList.Add(X)
        YValueList.Add(Y)
        ZValueList.Add(Z)

        ' Save As
        Part.ViewZoomtofit2()
        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_03A_Side_L_Left & Right.SLDDRW", 0, 2)
        swApp.CloseAllDocuments(True)

        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_03A_Side_L_Left & Right.SLDDRW", 3, 0, "", longstatus, longwarnings)
        boolstatus = Part.EditRebuild3()
        longstatus = Part.SaveAs3(SaveFolder & "\PDF Drawings\" & AHUName & "_03A_Side_L_Left & Right.PDF", 0, 2)

        swApp.CloseAllDocuments(True)

    End Sub

    Public Sub TopBotLDrawings()
        'Exit Sub

        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_02A_Bot_L.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_02A_Top_L.SLDPRT", 1, 0, "", longstatus, longwarnings)

        swApp.ActivateDoc2(AHUName & "_02A_Bot_L", False, longstatus)
        Part = swApp.ActiveDoc

        If Part Is Nothing Then
            MsgBox("Please open an appropriate Part Document")
            Exit Sub
        End If

        'Bounding Box
        Dim BBox As Object = StdFunc.BoundingBox()
        Dim xDim As Decimal = Abs(BBox(0)) + Abs(BBox(3))
        Dim yDim As Decimal = Abs(BBox(1)) + Abs(BBox(4))
        Dim zDim As Decimal = Abs(BBox(2)) + Abs(BBox(5))

        'Bounding Box - Flat
        boolstatus = Part.Extension.SelectByID2("Flat-Pattern16", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditUnsuppress2()

        Dim BBoxFlat As Object = StdFunc.BoundingBox()
        Dim xDimFlat As Decimal = Abs(BBoxFlat(0)) + Abs(BBoxFlat(3))
        Dim yDimFlat As Decimal = Abs(BBoxFlat(1)) + Abs(BBoxFlat(4))
        Dim zDimFlat As Decimal = Abs(BBoxFlat(2)) + Abs(BBoxFlat(5))

        boolstatus = Part.Extension.SelectByID2("Flat-Pattern16", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditSuppress2()

        'swApp.CloseAllDocuments(True)

        ' Sheet Scale
        Dim SScaleX As Integer = Ceiling((xDimFlat + zDim + xDim) / (0.297 - (0.03 + 0.03 + 0.03)))                 '(0.297 - (0.03 + 0.04 + 0.04 + 0.03)))
        Dim SScaleY As Integer = Ceiling((zDim + yDimFlat) / (0.21 - (0.03 + 0.02 + 0.03)))                         ' (0.21 - (0.03 + 0.04 + 0.03)))
        Dim SScale As Integer = SScaleX
        If SScaleX < SScaleY Then
            SScale = SScaleY
        End If

        ' Adjust Values for Scale
        xDim /= SScale
        yDim /= SScale
        zDim /= SScale

        xDimFlat /= SScale
        yDimFlat /= SScale
        zDimFlat /= SScale

        ' Get Margins
        Dim ClearX As Decimal = (0.297 - (xDimFlat + 0.03 + zDim + 0.03 + xDim)) / 2
        If ClearX < 0.03 Then ClearX = 0.03
        Dim ClearY As Decimal = (0.21 - (yDimFlat + 0.04 + zDim)) / 2
        If ClearY < 0.03 Then ClearY = 0.03


        ' Calculate View Placements
        Dim xFrontFlat As Decimal = ClearX + xDimFlat / 2
        Dim yTopSec As Decimal = ClearY + zDim / 2
        Dim yFrontFlat As Decimal = yTopSec + ClearY + 0.01     '---Up
        Dim xRightSec As Decimal = xFrontFlat + xDimFlat + 0.05
        Dim xIso As Decimal = ClearY + 0.05

        ' New Drawing
        Part = swApp.NewDocument(DrawSheet, 2, 0.297, 0.21)
        Draw = Part

        swSheet = Draw.GetCurrentSheet()
        swSheet.SetProperties2(12, 12, 1, SScale, False, 0.297, 0.21, True)
        swSheet.SetTemplateName(DrawTemp)
        swSheet.ReloadTemplate(True)
        swApp.ActivateDoc2("Draw1 - Sheet1", False, longstatus)
        Part = swApp.ActiveDoc

        ' Views
        boolstatus = Draw.CreateFlatPatternViewFromModelView(AHUName & "_02A_Bot_L", "Default", xFrontFlat, yFrontFlat, 0)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Draw.ActivateView("Drawing View1")
        Part.ClearSelection2(True)

        'Front - Outside
        myView = Part.CreateDrawViewFromModelView3(AHUName & "_02A_Bot_L", "*Front", -xFrontFlat, yFrontFlat, 0)
        boolstatus = Draw.ActivateView("Drawing View2")
        Part.ClearSelection2(True)


        'Top - Section
        boolstatus = Draw.ActivateView("Drawing View2")
        skSegment = Part.SketchManager.CreateLine((xDim * SScale / 2) + 0.15, 0.015, 0, -(xDim * SScale / 2) - 0.15, 0.015, 0)
        boolstatus = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xFrontFlat + 0.04, xIso + 0.01, 0, "A", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View3")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View3", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewVerticalCenter, BaseView)
        Part.ClearSelection2(True)
        boolstatus = Part.ActivateSheet("Sheet1")

        'Right - Section
        boolstatus = Draw.ActivateView("Drawing View2")
        skSegment = Part.SketchManager.CreateLine(0, (yDim * SScale / 2) + 0.15, 0, 0, -(yDim * SScale / 2) - 0.15, 0)
        boolstatus = Part.Extension.SelectByID2("Line2", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xIso + 0.03, yFrontFlat, 0, "B", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View4")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View4", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewHorizontalCenter, BaseView)
        Part.ClearSelection2(True)

        'Isometric
        myView = Part.CreateDrawViewFromModelView3(AHUName & "_02A_Bot_L", "*Dimetric", xRightSec, yFrontFlat, 0)
        boolstatus = Draw.ActivateView("Drawing View5")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View5", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject6(1, -1)
        boolstatus = DrawView.SetDisplayMode4(False, swDisplayMode_e.swSHADED, True, True, False)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View5", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewHorizontalCenter, BaseView)
        Part.ClearSelection2(True)

        For a = 1 To 50
            boolstatus = Part.ActivateView("Drawing View3")
            boolstatus = Part.Extension.SelectByID2("DetailItem" & a & "@Drawing View3", "NOTE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
            boolstatus = Part.ActivateView("Drawing View4")
            boolstatus = Part.Extension.SelectByID2("DetailItem" & a & "@Drawing View4", "NOTE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
        Next


        'Width - Flat View  
        boolstatus = Part.ActivateView("Drawing View1")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, yFrontFlat + (0.0482 / SScale), -10500, 0, 0, -1, 0.0001, 1, False, 0, 0)
        'boolstatus = Part.Extension.SelectByRay(xFrontFlat, yFrontFlat - (0.0482 / SScale), -10500, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat, yFrontFlat + 0.017, 0)
        Part.ClearSelection2(True)

        'Height - Flat View   
        boolstatus = Part.ActivateView("Drawing View1")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, yFrontFlat + (0.0482 / SScale), -10500, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, yFrontFlat - (0.0482 / SScale), -10500, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat - 0.06, yFrontFlat, 0)
        Part.ClearSelection2(True)



        'Height - A View 
        boolstatus = Part.ActivateView("Drawing View3")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, (xIso + 0.01) - (0.025 / SScale), -10500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat - ((WallWth - 2) / (SScale * 2000)), (xIso + 0.01) + (0.025 / SScale), -10500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat - 0.06, (xIso + 0.01), 0)
        Part.ClearSelection2(True)

        'Height - A View 
        boolstatus = Part.ActivateView("Drawing View3")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, (xIso + 0.01) - (0.025 / SScale), -10500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat + ((WallWth - 2) / (SScale * 2000)), (xIso + 0.01) + (0.025 / SScale), -10500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat + 0.06, (xIso + 0.01) - 0.01, 0)
        Part.ClearSelection2(True)

        'Width - A View    + (0.025 / SScale)
        boolstatus = Part.ActivateView("Drawing View3")
        boolstatus = Part.Extension.SelectByRay((xFrontFlat + ((WallWth) / (SScale * 2000))), (xIso + 0.01), -13000, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay((xFrontFlat - ((WallWth) / (SScale * 2000))), (xIso + 0.01), -13000, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddHorizontalDimension2(xFrontFlat, (xIso + 0.01) - 0.011, 0)
        Part.ClearSelection2(True)



        'Width - B View 
        boolstatus = Part.ActivateView("Drawing View4")
        boolstatus = Part.Extension.SelectByRay((xIso + 0.03) + (0.025 / SScale), yFrontFlat + (0.0137 / SScale), -10500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        'boolstatus = Part.Extension.SelectByRay((xIso + 0.03) + (0.025 / SScale), yFrontFlat, -10500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay((xIso + 0.03) - (0.025 / SScale), yFrontFlat - (0.024 / SScale), -10500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2((xIso + 0.03) - 0.01, yFrontFlat - 0.012, 0)
        Part.ClearSelection2(True)

        'Height - B View
        boolstatus = Part.ActivateView("Drawing View4")
        boolstatus = Part.Extension.SelectByRay((xIso + 0.03) + (0.024 / SScale), yFrontFlat + (0.025 / SScale), -13000, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay((xIso + 0.03) - (0.014 / SScale), yFrontFlat - (0.025 / SScale), -13000, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2((xIso + 0.03) + 0.012, yFrontFlat + 0.01, 0)
        Part.ClearSelection2(True)



        ' Note
        Dim X, Y, Z As Decimal
        boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingBalloonStyle, 0, swBalloonStyle_e.swBS_None)

        X = Round(xDimFlat * SScale * 1000, 2)
        Y = Round(yDimFlat * SScale * 1000, 2)
        Z = Round(zDimFlat * SScale * 1000, 2)
        myNote = Draw.CreateText2(AHUName & "_02A_Bot_L" & vbNewLine & "Qty - 1" & vbNewLine & X & "mm x " & Y & "mm x " & "2.00mm", xRightSec - 0.05, xIso + 0.007, 0, 0.004, 0)

        boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

        XValueList.Add(X)
        YValueList.Add(Y)
        ZValueList.Add(Z)


        '==================================================================================== Top L ====================================================================================================================

        swApp.ActivateDoc2(AHUName & "_02A_Top_L", False, longstatus)
        Part = swApp.ActiveDoc

        If Part Is Nothing Then
            MsgBox("Please open an appropriate Part Document")
            Exit Sub
        End If

        swSheet = Draw.GetCurrentSheet()
        swSheet.SetProperties2(12, 12, 1, SScale, False, 0.297, 0.21, True)
        swSheet.SetTemplateName(DrawTemp)
        swSheet.ReloadTemplate(True)
        swApp.ActivateDoc2("Draw1 - Sheet1", False, longstatus)
        Part = swApp.ActiveDoc
        Part = Draw

        ' Flat Views
        boolstatus = Draw.CreateFlatPatternViewFromModelView(AHUName & "_02A_Top_L", "Default", xFrontFlat, yTopSec, 0)
        boolstatus = Draw.ActivateView("Drawing View6")
        Part.ClearSelection2(True)

        'Front 
        myView = Draw.CreateDrawViewFromModelView3(AHUName & "_02A_Top_L", "*Front", -xFrontFlat, yTopSec, 0)
        boolstatus = Draw.ActivateView("Drawing View7")
        Part.ClearSelection2(True)

        'Top - Section
        boolstatus = Draw.ActivateView("Drawing View7")
        skSegment = Part.SketchManager.CreateLine((xDim * SScale / 2) + 0.15, 0.015, 0, -(xDim * SScale / 2) - 0.15, 0.015, 0)
        boolstatus = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xFrontFlat, yTopSec - 0.05, 0, "A", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View8")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View8", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View6", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewVerticalCenter, BaseView)
        Part.ClearSelection2(True)
        boolstatus = Part.ActivateSheet("Sheet1")

        'Right - Section
        boolstatus = Draw.ActivateView("Drawing View7")
        skSegment = Part.SketchManager.CreateLine(0, (yDim * SScale / 2) + 0.15, 0, 0, -(yDim * SScale / 2) - 0.15, 0)
        boolstatus = Part.Extension.SelectByID2("Line2", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xRightSec, yTopSec - 0.05, 0, "B", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View9")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View9", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View6", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewHorizontalCenter, BaseView)
        Part.ClearSelection2(True)

        'Isometric
        myView = Part.CreateDrawViewFromModelView3(AHUName & "_02A_Top_L", "*Dimetric", xIso + ClearX + 0.07, yTopSec, 0)
        boolstatus = Draw.ActivateView("Drawing View10")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View10", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject6(1, -1)
        boolstatus = DrawView.SetDisplayMode4(False, swDisplayMode_e.swSHADED, True, True, False)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View10", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View6", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewHorizontalCenter, BaseView)
        Part.ClearSelection2(True)

        boolstatus = Part.ActivateView("Drawing View9")
        boolstatus = Part.Extension.SelectByID2("Drawing View9", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.ActivateView("Drawing View4")
        boolstatus = Part.Extension.SelectByID2("Drawing View4", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewVerticalCenter, BaseView)
        Part.ClearSelection2(True)

        For a = 1 To 50
            boolstatus = Part.ActivateView("Drawing View9")
            boolstatus = Part.Extension.SelectByID2("DetailItem" & a & "@Drawing View9", "NOTE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
            boolstatus = Part.ActivateView("Drawing View8")
            boolstatus = Part.Extension.SelectByID2("DetailItem" & a & "@Drawing View8", "NOTE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
        Next

        'Dimensions-------------------
        'Width - Flat View   
        boolstatus = Part.ActivateView("Drawing View6")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, yTopSec + (0.0482 / SScale), -10500, 0, 0, -1, 0.0001, 1, False, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat, yTopSec + 0.017, 0)
        Part.ClearSelection2(True)

        'Height - Flat View   
        boolstatus = Part.ActivateView("Drawing View6")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, yTopSec + (0.0482 / SScale), -10500, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, yTopSec - (0.0482 / SScale), -10500, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat - 0.06, yTopSec, 0)
        Part.ClearSelection2(True)


        'Height - A View   
        boolstatus = Part.ActivateView("Drawing View8")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, (yTopSec - 0.05) - (0.025 / SScale), -10500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat - ((WallWth - 2) / (SScale * 2000)), (yTopSec - 0.05) + (0.025 / SScale), -10500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat - 0.06, (yTopSec - 0.05), 0)
        Part.ClearSelection2(True)

        'Height - A View   
        boolstatus = Part.ActivateView("Drawing View8")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, (yTopSec - 0.05) - (0.025 / SScale), -10500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat + ((WallWth - 2) / (SScale * 2000)), (yTopSec - 0.05) + (0.025 / SScale), -10500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat + 0.06, (yTopSec - 0.05) - 0.01, 0)
        Part.ClearSelection2(True)

        'Width - A View        + (0.025 / SScale)
        boolstatus = Part.ActivateView("Drawing View8")
        boolstatus = Part.Extension.SelectByRay((xFrontFlat + ((WallWth) / (SScale * 2000))), (yTopSec - 0.05), -13000, 0, 0, -1, 0.0005 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay((xFrontFlat - ((WallWth) / (SScale * 2000))), (yTopSec - 0.05), -13000, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat, (yTopSec - 0.05) - 0.011, 0)
        Part.ClearSelection2(True)


        'Width - B View  
        boolstatus = Part.ActivateView("Drawing View9")
        boolstatus = Part.Extension.SelectByRay((xIso + 0.03) + (0.025 / SScale), (yTopSec) + (0.0137 / SScale), -5500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        'boolstatus = Part.Extension.SelectByRay((xIso + 0.03) + (0.025 / SScale), (yTopSec), -5500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay((xIso + 0.03) - (0.025 / SScale), (yTopSec) - (0.024 / SScale), -5500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2((xIso + 0.03) - 0.01, (yTopSec) - 0.012, 0)
        Part.ClearSelection2(True)

        'Height - B View
        boolstatus = Part.ActivateView("Drawing View9")
        boolstatus = Part.Extension.SelectByRay((xIso + 0.03) + (0.024 / SScale), (yTopSec) + (0.025 / SScale), -13000, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay((xIso + 0.03) - (0.014 / SScale), (yTopSec) - (0.025 / SScale), -13000, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2((xIso + 0.03) + 0.015, (yTopSec) + 0.01, 0)
        Part.ClearSelection2(True)


        ''Width - B View  
        'boolstatus = Part.ActivateView("Drawing View9")
        'boolstatus = Part.Extension.SelectByRay((xIso + 0.03) + (0.025 / SScale), (yTopSec), -5500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        'boolstatus = Part.Extension.SelectByRay((xIso + 0.03) - (0.025 / SScale), (yTopSec) + (0.024 / SScale), -5500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        'myDisplayDim = Part.AddDimension2((xIso + 0.03) - 0.01, (yTopSec) + 0.012, 0)
        'Part.ClearSelection2(True)

        ''Height - B View
        'boolstatus = Part.ActivateView("Drawing View9")
        'boolstatus = Part.Extension.SelectByRay((xIso + 0.03) + (0.024 / SScale), (yTopSec) - (0.025 / SScale), -13000, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        'boolstatus = Part.Extension.SelectByRay((xIso + 0.03) - (0.014 / SScale), (yTopSec) + (0.025 / SScale), -13000, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        'myDisplayDim = Part.AddDimension2((xIso + 0.03) + 0.015, (yTopSec) + 0.01, 0)
        'Part.ClearSelection2(True)

        'Note
        myNote = Draw.CreateText2(AHUName & "_02A_Top_L" & vbNewLine & "Qty - 1" & vbNewLine & X & "mm x " & Y & "mm x " & "2.00mm", xIso + ClearX + 0.02, yTopSec - 0.04, 0, 0.004, 0)

        ' Save As
        Part.ViewZoomtofit2()
        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02A_Bot & Top_L.SLDDRW", 0, 2)
        swApp.CloseAllDocuments(True)

        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_02A_Bot & Top_L.SLDDRW", 3, 0, "", longstatus, longwarnings)
        boolstatus = Part.EditRebuild3()
        longstatus = Part.SaveAs3(SaveFolder & "\PDF Drawings\" & AHUName & "_02A_Bot & Top_L.PDF", 0, 2)

        swApp.CloseAllDocuments(True)
    End Sub

    Public Sub TopBot2BLDrawings()
        'Exit Sub

        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_02B_Bot_L2.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_02B_Top_L2.SLDPRT", 1, 0, "", longstatus, longwarnings)

        swApp.ActivateDoc2(AHUName & "_02B_Bot_L2", False, longstatus)
        Part = swApp.ActiveDoc

        If Part Is Nothing Then
            MsgBox("Please open an appropriate Part Document")
            Exit Sub
        End If

        'Bounding Box
        Dim BBox As Object = StdFunc.BoundingBox()
        Dim xDim As Decimal = Abs(BBox(0)) + Abs(BBox(3))
        Dim yDim As Decimal = Abs(BBox(1)) + Abs(BBox(4))
        Dim zDim As Decimal = Abs(BBox(2)) + Abs(BBox(5))

        'Bounding Box - Flat
        boolstatus = Part.Extension.SelectByID2("Flat-Pattern16", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditUnsuppress2()

        Dim BBoxFlat As Object = StdFunc.BoundingBox()
        Dim xDimFlat As Decimal = Abs(BBoxFlat(0)) + Abs(BBoxFlat(3))
        Dim yDimFlat As Decimal = Abs(BBoxFlat(1)) + Abs(BBoxFlat(4))
        Dim zDimFlat As Decimal = Abs(BBoxFlat(2)) + Abs(BBoxFlat(5))

        boolstatus = Part.Extension.SelectByID2("Flat-Pattern16", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditSuppress2()

        'swApp.CloseAllDocuments(True)

        ' Sheet Scale
        Dim SScaleX As Integer = Ceiling((xDimFlat + zDim + xDim) / (0.297 - (0.03 + 0.03 + 0.03)))                 '(0.297 - (0.03 + 0.04 + 0.04 + 0.03)))
        Dim SScaleY As Integer = Ceiling((zDim + yDimFlat) / (0.21 - (0.03 + 0.02 + 0.03)))                         ' (0.21 - (0.03 + 0.04 + 0.03)))
        Dim SScale As Integer = SScaleX
        If SScaleX < SScaleY Then
            SScale = SScaleY
        End If

        ' Adjust Values for Scale
        xDim /= SScale
        yDim /= SScale
        zDim /= SScale

        xDimFlat /= SScale
        yDimFlat /= SScale
        zDimFlat /= SScale

        ' Get Margins
        Dim ClearX As Decimal = (0.297 - (xDimFlat + 0.03 + zDim + 0.03 + xDim)) / 2
        If ClearX < 0.03 Then ClearX = 0.03
        Dim ClearY As Decimal = (0.21 - (yDimFlat + 0.04 + zDim)) / 2
        If ClearY < 0.03 Then ClearY = 0.03


        ' Calculate View Placements
        Dim xFrontFlat As Decimal = ClearX + xDimFlat / 2
        Dim yTopSec As Decimal = ClearY + zDim / 2
        Dim yFrontFlat As Decimal = yTopSec + ClearY + 0.01     '---Up
        Dim xRightSec As Decimal = xFrontFlat + xDimFlat + 0.05
        Dim xIso As Decimal = ClearY + 0.05

        ' New Drawing
        Part = swApp.NewDocument(DrawSheet, 2, 0.297, 0.21)
        Draw = Part

        swSheet = Draw.GetCurrentSheet()
        swSheet.SetProperties2(12, 12, 1, SScale, False, 0.297, 0.21, True)
        swSheet.SetTemplateName(DrawTemp)
        swSheet.ReloadTemplate(True)
        swApp.ActivateDoc2("Draw1 - Sheet1", False, longstatus)
        Part = swApp.ActiveDoc

        ' Views
        boolstatus = Draw.CreateFlatPatternViewFromModelView(AHUName & "_02B_Bot_L2", "Default", xFrontFlat, yFrontFlat, 0)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Draw.ActivateView("Drawing View1")
        Part.ClearSelection2(True)

        'Front - Outside
        myView = Part.CreateDrawViewFromModelView3(AHUName & "_02B_Bot_L2", "*Front", -xFrontFlat, yFrontFlat, 0)
        boolstatus = Draw.ActivateView("Drawing View2")
        Part.ClearSelection2(True)


        'Top - Section
        boolstatus = Draw.ActivateView("Drawing View2")
        skSegment = Part.SketchManager.CreateLine((xDim * SScale / 2) + 0.15, 0.015, 0, -(xDim * SScale / 2) - 0.15, 0.015, 0)
        boolstatus = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xFrontFlat + 0.04, xIso + 0.01, 0, "A", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View3")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View3", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewVerticalCenter, BaseView)
        Part.ClearSelection2(True)
        boolstatus = Part.ActivateSheet("Sheet1")

        'Right - Section
        boolstatus = Draw.ActivateView("Drawing View2")
        skSegment = Part.SketchManager.CreateLine(0, (yDim * SScale / 2) + 0.15, 0, 0, -(yDim * SScale / 2) - 0.15, 0)
        boolstatus = Part.Extension.SelectByID2("Line2", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xIso + 0.03, yFrontFlat, 0, "B", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View4")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View4", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewHorizontalCenter, BaseView)
        Part.ClearSelection2(True)

        'Isometric
        myView = Part.CreateDrawViewFromModelView3(AHUName & "_02B_Bot_L2", "*Dimetric", xRightSec, yFrontFlat, 0)
        boolstatus = Draw.ActivateView("Drawing View5")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View5", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject6(1, -1)
        boolstatus = DrawView.SetDisplayMode4(False, swDisplayMode_e.swSHADED, True, True, False)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View5", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewHorizontalCenter, BaseView)
        Part.ClearSelection2(True)

        For a = 1 To 50
            boolstatus = Part.ActivateView("Drawing View3")
            boolstatus = Part.Extension.SelectByID2("DetailItem" & a & "@Drawing View3", "NOTE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
            boolstatus = Part.ActivateView("Drawing View4")
            boolstatus = Part.Extension.SelectByID2("DetailItem" & a & "@Drawing View4", "NOTE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
        Next

        'DIMENSIONS - Width - Flat View  
        boolstatus = Part.ActivateView("Drawing View1")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, yFrontFlat + (0.0482 / SScale), -10500, 0, 0, -1, 0.0001, 1, False, 0, 0)
        'boolstatus = Part.Extension.SelectByRay(xFrontFlat, yFrontFlat - (0.0482 / SScale), -10500, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat, yFrontFlat + 0.017, 0)
        Part.ClearSelection2(True)

        'Height - Flat View   
        boolstatus = Part.ActivateView("Drawing View1")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, yFrontFlat + (0.0482 / SScale), -10500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, yFrontFlat - (0.0482 / SScale), -10500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat - 0.06, yFrontFlat, 0)
        Part.ClearSelection2(True)

        'Height - A View 
        boolstatus = Part.ActivateView("Drawing View3")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, (xIso + 0.01) - (0.025 / SScale), -10500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat - ((WallWth - 2) / (SScale * 2000)), (xIso + 0.01) + (0.025 / SScale), -10500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat - 0.06, (xIso + 0.01), 0)
        Part.ClearSelection2(True)

        'Height - A View 
        boolstatus = Part.ActivateView("Drawing View3")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, (xIso + 0.01) - (0.025 / SScale), -10500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat + ((WallWth - 2) / (SScale * 2000)), (xIso + 0.01) + (0.025 / SScale), -10500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat + 0.06, (xIso + 0.01) - 0.01, 0)
        Part.ClearSelection2(True)

        'Width - A View 
        boolstatus = Part.ActivateView("Drawing View3")
        boolstatus = Part.Extension.SelectByRay((xFrontFlat + ((WallWth) / (SScale * 2000))), (xIso + 0.01) + (0.025 / SScale), -13000, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay((xFrontFlat - ((WallWth) / (SScale * 2000))), (xIso + 0.01) + (0.025 / SScale), -13000, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddHorizontalDimension2(xFrontFlat, (xIso + 0.01) + 0.011, 0)
        Part.ClearSelection2(True)

        'Width - B View 
        boolstatus = Part.ActivateView("Drawing View4")
        boolstatus = Part.Extension.SelectByRay((xIso + 0.03) + (0.025 / SScale), yFrontFlat, -10500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay((xIso + 0.03) - (0.025 / SScale), yFrontFlat - (0.024 / SScale), -10500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2((xIso + 0.03) - 0.01, yFrontFlat - 0.012, 0)
        Part.ClearSelection2(True)

        'Height - B View
        boolstatus = Part.ActivateView("Drawing View4")
        boolstatus = Part.Extension.SelectByRay((xIso + 0.03) + (0.024 / SScale), yFrontFlat + (0.025 / SScale), -13000, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay((xIso + 0.03) - (0.014 / SScale), yFrontFlat - (0.025 / SScale), -13000, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2((xIso + 0.03) + 0.012, yFrontFlat + 0.01, 0)
        Part.ClearSelection2(True)



        ''Height - B View
        'boolstatus = Part.ActivateView("Drawing View4")
        'boolstatus = Part.Extension.SelectByRay((xIso + 0.03) - (0.024 / SScale), yFrontFlat + (0.025 / SScale), -5000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        'boolstatus = Part.Extension.SelectByRay((xIso + 0.03), yFrontFlat - (0.025 / SScale), -5000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        'myDisplayDim = Part.AddDimension2((xIso + 0.03) + 0.012, yFrontFlat + 0.01, 0)
        'Part.ClearSelection2(True)

        ' Note
        Dim X, Y, Z As Decimal
        boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingBalloonStyle, 0, swBalloonStyle_e.swBS_None)

        X = Round(xDimFlat * SScale * 1000, 2)
        Y = Round(yDimFlat * SScale * 1000, 2)
        Z = Round(zDimFlat * SScale * 1000, 2)
        myNote = Draw.CreateText2(AHUName & "_02B_Bot_L2" & vbNewLine & "Qty - 1" & vbNewLine & X & "mm x " & Y & "mm x " & "2.00mm", xRightSec - 0.05, xIso + 0.007, 0, 0.004, 0)

        boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

        XValueList.Add(X)
        YValueList.Add(Y)
        ZValueList.Add(Z)


        '==================================================================================== 2B Top L2 ====================================================================================================================

        swApp.ActivateDoc2(AHUName & "_02B_Top_L2", False, longstatus)
        Part = swApp.ActiveDoc

        If Part Is Nothing Then
            MsgBox("Please open an appropriate Part Document")
            Exit Sub
        End If

        swSheet = Draw.GetCurrentSheet()
        swSheet.SetProperties2(12, 12, 1, SScale, False, 0.297, 0.21, True)
        swSheet.SetTemplateName(DrawTemp)
        swSheet.ReloadTemplate(True)
        swApp.ActivateDoc2("Draw1 - Sheet1", False, longstatus)
        Part = swApp.ActiveDoc
        Part = Draw

        ' Flat Views
        boolstatus = Draw.CreateFlatPatternViewFromModelView(AHUName & "_02B_Top_L2", "Default", xFrontFlat, yTopSec, 0)
        boolstatus = Draw.ActivateView("Drawing View6")
        Part.ClearSelection2(True)

        'Front 
        myView = Draw.CreateDrawViewFromModelView3(AHUName & "_02B_Top_L2", "*Front", -xFrontFlat, yTopSec, 0)
        boolstatus = Draw.ActivateView("Drawing View7")
        Part.ClearSelection2(True)

        'Top - Section
        boolstatus = Draw.ActivateView("Drawing View7")
        skSegment = Part.SketchManager.CreateLine((xDim * SScale / 2) + 0.15, 0.015, 0, -(xDim * SScale / 2) - 0.15, 0.015, 0)
        boolstatus = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xFrontFlat, yTopSec - 0.05, 0, "A", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View8")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View8", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View6", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewVerticalCenter, BaseView)
        Part.ClearSelection2(True)
        boolstatus = Part.ActivateSheet("Sheet1")

        'Right - Section
        boolstatus = Draw.ActivateView("Drawing View7")
        skSegment = Part.SketchManager.CreateLine(0, (yDim * SScale / 2) + 0.15, 0, 0, -(yDim * SScale / 2) - 0.15, 0)
        boolstatus = Part.Extension.SelectByID2("Line2", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xRightSec, yTopSec - 0.05, 0, "B", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View9")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View9", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View6", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewHorizontalCenter, BaseView)
        Part.ClearSelection2(True)

        'Isometric
        myView = Part.CreateDrawViewFromModelView3(AHUName & "_02B_Top_L2", "*Dimetric", xIso + ClearX + 0.07, yTopSec, 0)
        boolstatus = Draw.ActivateView("Drawing View10")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View10", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject6(1, -1)
        boolstatus = DrawView.SetDisplayMode4(False, swDisplayMode_e.swSHADED, True, True, False)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View10", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View6", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewHorizontalCenter, BaseView)
        Part.ClearSelection2(True)

        boolstatus = Part.ActivateView("Drawing View9")
        boolstatus = Part.Extension.SelectByID2("Drawing View9", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.ActivateView("Drawing View4")
        boolstatus = Part.Extension.SelectByID2("Drawing View4", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewVerticalCenter, BaseView)
        Part.ClearSelection2(True)


        For a = 1 To 50
            boolstatus = Part.ActivateView("Drawing View9")
            boolstatus = Part.Extension.SelectByID2("DetailItem" & a & "@Drawing View9", "NOTE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
            boolstatus = Part.ActivateView("Drawing View8")
            boolstatus = Part.Extension.SelectByID2("DetailItem" & a & "@Drawing View8", "NOTE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
        Next

        'Dimensions-------------------
        'Width - Flat View   xFrontFlat, yTopSec
        boolstatus = Part.ActivateView("Drawing View6")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, yTopSec + (0.0482 / SScale), -10500, 0, 0, -1, 0.0001, 1, False, 0, 0)
        'boolstatus = Part.Extension.SelectByRay(xFrontFlat, yFrontFlat - (0.0482 / SScale), -10500, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat, yTopSec + 0.017, 0)
        Part.ClearSelection2(True)

        'Height - Flat View   
        boolstatus = Part.ActivateView("Drawing View6")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, yTopSec + (0.0482 / SScale), -10500, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, yTopSec - (0.0482 / SScale), -10500, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat - 0.06, yTopSec, 0)
        Part.ClearSelection2(True)


        'Height - A View   
        boolstatus = Part.ActivateView("Drawing View8")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, (yTopSec - 0.05) - (0.025 / SScale), -10500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat - ((WallWth - 2) / (SScale * 2000)), (yTopSec - 0.05) + (0.025 / SScale), -10500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat - 0.06, (yTopSec - 0.05), 0)
        Part.ClearSelection2(True)

        'Height - A View   
        boolstatus = Part.ActivateView("Drawing View8")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, (yTopSec - 0.05) - (0.025 / SScale), -10500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat + ((WallWth - 2) / (SScale * 2000)), (yTopSec - 0.05) + (0.025 / SScale), -10500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat + 0.06, (yTopSec - 0.05) - 0.01, 0)
        Part.ClearSelection2(True)

        'Width - A View
        boolstatus = Part.ActivateView("Drawing View8")
        boolstatus = Part.Extension.SelectByRay((xFrontFlat + ((WallWth) / (SScale * 2000))), (yTopSec - 0.05) + (0.025 / SScale), -13000, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay((xFrontFlat - ((WallWth) / (SScale * 2000))), (yTopSec - 0.05) + (0.025 / SScale), -13000, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat, (yTopSec - 0.05) + 0.011, 0)
        Part.ClearSelection2(True)

        'Width - B View  
        boolstatus = Part.ActivateView("Drawing View9")
        boolstatus = Part.Extension.SelectByRay((xIso + 0.03) + (0.025 / SScale), (yTopSec), -5500, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay((xIso + 0.03) - (0.025 / SScale), (yTopSec) - (0.024 / SScale), -5500, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2((xIso + 0.03) - 0.01, (yTopSec) - 0.012, 0)
        Part.ClearSelection2(True)

        'Height - B View
        boolstatus = Part.ActivateView("Drawing View9")
        boolstatus = Part.Extension.SelectByRay((xIso + 0.03) + (0.024 / SScale), (yTopSec) + (0.025 / SScale), -13000, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        boolstatus = Part.Extension.SelectByRay((xIso + 0.03) - (0.014 / SScale), (yTopSec) - (0.025 / SScale), -13000, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2((xIso + 0.03) + 0.015, (yTopSec) + 0.01, 0)
        Part.ClearSelection2(True)

        'Note
        myNote = Draw.CreateText2(AHUName & "_02B_Top_L2" & vbNewLine & "Qty - 1" & vbNewLine & X & "mm x " & Y & "mm x " & "2.00mm", xIso + ClearX + 0.02, yTopSec - 0.04, 0, 0.004, 0)

        ' Save As
        Part.ViewZoomtofit2()
        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_02B_Bot & Top_L2.SLDDRW", 0, 2)
        swApp.CloseAllDocuments(True)

        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_02B_Bot & Top_L2.SLDDRW", 3, 0, "", longstatus, longwarnings)
        boolstatus = Part.EditRebuild3()
        longstatus = Part.SaveAs3(SaveFolder & "\PDF Drawings\" & AHUName & "_02B_Bot & Top_L2.PDF", 0, 2)

        swApp.CloseAllDocuments(True)
    End Sub

    Public Sub BaseStandDrawings(PartName As String, BaseQty As Integer, BaseSize As Decimal)
        ' Exit Sub
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_10_Base_Stand.SLDPRT", 2, 32, "", longstatus, longwarnings)

        swApp.ActivateDoc2(AHUName & "_10_Base_Stand", False, longstatus)
        Part = swApp.ActiveDoc

        If Part Is Nothing Then
            MsgBox("Please open an appropriate Part Document")
            Exit Sub
        End If

        'Bounding Box
        Dim BBox As Object = StdFunc.BoundingBox()
        Dim xDim As Decimal = Abs(BBox(0)) + Abs(BBox(3))
        Dim yDim As Decimal = Abs(BBox(1)) + Abs(BBox(4))
        Dim zDim As Decimal = Abs(BBox(2)) + Abs(BBox(5))

        'Bounding Box - Flat
        boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditUnsuppress2()

        Dim BBoxFlat As Object = StdFunc.BoundingBox()
        Dim xDimFlat As Decimal = Abs(BBoxFlat(0)) + Abs(BBoxFlat(3))
        Dim yDimFlat As Decimal = Abs(BBoxFlat(1)) + Abs(BBoxFlat(4))
        Dim zDimFlat As Decimal = Abs(BBoxFlat(2)) + Abs(BBoxFlat(5))

        boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditSuppress2()

        'swApp.CloseAllDocuments(True)

        ' Sheet Scale
        Dim SScaleX As Integer = Ceiling((xDimFlat + zDim + xDim) / (0.297 - (0.03 + 0.03 + 0.03)))                  ' (0.297 - (0.03 + 0.04 + 0.04 + 0.03)))
        Dim SScaleY As Integer = Ceiling((zDim + yDimFlat) / (0.21 - (0.03 + 0.04 + 0.03)))                         ' (0.21 - (0.03 + 0.04 + 0.03)))
        Dim SScale As Integer = SScaleX
        If SScaleX < SScaleY Then
            SScale = SScaleY
        End If

        ' Adjust Values for Scale
        xDim /= SScale
        yDim /= SScale
        zDim /= SScale

        xDimFlat /= SScale
        yDimFlat /= SScale
        zDimFlat /= SScale

        ' Get Margins
        Dim ClearX As Decimal = (0.297 - (xDimFlat + 0.04 + zDim + 0.04 + xDim)) / 2
        If ClearX < 0.03 Then ClearX = 0.03
        Dim ClearY As Decimal = (0.21 - (yDimFlat + 0.045 + zDim)) / 2
        If ClearY < 0.03 Then ClearY = 0.03

        ' Calculate View Placements
        Dim xFrontFlat As Decimal = ((ClearX + xDimFlat) * 2) - 0.02
        Dim yTopSec As Decimal = ClearY + zDim / 2
        Dim yFrontFlat As Decimal = yTopSec + zDim / 2 + 0.01
        Dim xRightSec As Decimal = xFrontFlat + xDimFlat / 2 + 0.04 + zDim / 2
        Dim xIso As Decimal = xRightSec + zDim / 2 + 0.06 + xDim / 2

        ' New Drawing
        Part = swApp.NewDocument(DrawSheet, 2, 0.297, 0.21)
        Draw = Part

        swSheet = Draw.GetCurrentSheet()
        swSheet.SetProperties2(12, 12, 1, SScale, False, 0.297, 0.21, True)
        swSheet.SetTemplateName(DrawTemp)
        swSheet.ReloadTemplate(True)
        swApp.ActivateDoc2("Draw1 - Sheet1", False, longstatus)
        Part = swApp.ActiveDoc
        Part = Draw


        PartName = AHUName & "_10_Base_Stand"

        'Flat
        boolstatus = Draw.CreateFlatPatternViewFromModelView(PartName, "Default", xFrontFlat, yFrontFlat, 0)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Draw.ActivateView("Drawing View1")
        Part.ClearSelection2(True)


        'Right
        myView = Part.CreateDrawViewFromModelView3(PartName, "*Front", xRightSec, yFrontFlat, 0)
        boolstatus = Draw.ActivateView("Drawing View2")
        Part.ClearSelection2(True)

        'Isometric
        myView = Part.CreateDrawViewFromModelView3(PartName, "*Isometric", xFrontFlat, yTopSec, 0)
        boolstatus = Draw.ActivateView("Drawing View3")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View3", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject6(1, -1)
        boolstatus = DrawView.SetDisplayMode4(False, swDisplayMode_e.swSHADED, True, True, False)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View3", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewVerticalCenter, BaseView)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View2", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewHorizontalCenter, BaseView)
        Part.ClearSelection2(True)

        Dim X, Y, Z As Decimal
        X = Round(xDimFlat * SScale * 1000, 2)
        Y = Round(yDimFlat * SScale * 1000, 2)
        Z = Round(zDimFlat * SScale * 1000, 2)


        'Dimensions--------------------------------------------------------------------------------------------------------------

        'Height - Flat
        boolstatus = Part.ActivateView("Drawing View1")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat + ((Z / 2000) / SScale), yFrontFlat, -5000, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat - ((Z / 2000) / SScale) - 0.005, yFrontFlat, 0)
        Part.ClearSelection2(True)

        'Width - Flat
        boolstatus = Part.ActivateView("Drawing View1")
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, (yFrontFlat + (0.07304 / SScale)), -5000, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat, yFrontFlat + 0.02, 0)
        Part.ClearSelection2(True)

        'Width - Section   
        boolstatus = Part.ActivateView("Drawing View2")
        boolstatus = Part.Extension.SelectByRay(xRightSec + (0.04 / SScale), yFrontFlat, -5000, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xRightSec - (0.04 / SScale), yFrontFlat, -5000, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xRightSec, yFrontFlat + 0.01, 0)
        Part.ClearSelection2(True)

        'Height - Section
        boolstatus = Part.ActivateView("Drawing View2")
        boolstatus = Part.Extension.SelectByRay(xRightSec, yFrontFlat + (0.01 / SScale), -5000, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xRightSec + (0.05 / SScale), yFrontFlat - (0.01 / SScale), -5000, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xRightSec + 0.015, yFrontFlat + 0.007, 0)
        Part.ClearSelection2(True)

        'Width 2 - Section   
        boolstatus = Part.ActivateView("Drawing View2")
        boolstatus = Part.Extension.SelectByRay(xRightSec - (0.038 / SScale), yFrontFlat, -5000, 0, 0, -1, 0.0001 / SScale, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xRightSec - (0.058 / SScale), yFrontFlat - (0.009 / SScale), -5000, 0, 0, -1, 0.0001 / SScale, 1, True, 0, 0)
        myDisplayDim = Part.AddHorizontalDimension2(xRightSec - 0.01, yFrontFlat - 0.01, 0)
        Part.ClearSelection2(True)

        ' Note
        boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingBalloonStyle, 0, swBalloonStyle_e.swBS_None)
        myNote = Draw.CreateText2(AHUName & "_10_Base_Stand" & vbNewLine & "Qty - " & BaseQty & vbNewLine & Z & "mm x " & X & "mm x " & Y & "mm", xRightSec, yTopSec, 0, 0.004, 0)
        boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

        XValueList.Add(X)
        YValueList.Add(Y)
        ZValueList.Add(Z)


        ' Save As
        Part.ViewZoomtofit2()
        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_10_Base_Stand.SLDDRW", 0, 2)
        swApp.CloseAllDocuments(True)

        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_10_Base_Stand.SLDDRW", 3, 0, "", longstatus, longwarnings)
        boolstatus = Part.EditRebuild3()
        longstatus = Part.SaveAs3(SaveFolder & "\PDF Drawings\" & AHUName & "_10_Base_Stand.PDF", 0, 2)

        swApp.CloseAllDocuments(True)

    End Sub

    Public Sub AssemDrawing()
        ' Exit Sub
        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_AHU Final Assembly.SLDASM", 2, 32, "", longstatus, longwarnings)
        Assy = Part
        Assy.ViewZoomtofit2()

        'Bounding Box of assembly
        Dim BBox As Object
        BBox = StdFunc.BoundingBoxOfAssembly
        Dim dimX As Decimal = Abs(BBox(0)) + Abs(BBox(3))
        Dim dimY As Decimal = Abs(BBox(1)) + Abs(BBox(4))
        Dim dimZ As Decimal = Abs(BBox(2)) + Abs(BBox(5))

        ' Sheet Scale
        Dim SScale As Integer = StdFunc.ShtScale(dimX, dimZ, 0.297, 0.21)

        dimX /= SScale
        dimY /= SScale
        dimZ /= SScale


        ' New Drawing
        Part = swApp.NewDocument(DrawSheet, 2, 0.297, 0.21)
        Draw = Part

        swSheet = Draw.GetCurrentSheet()
        swSheet.SetProperties2(12, 12, 1, SScale, False, 0.297, 0.21, True)
        swSheet.SetTemplateName(DrawTemp)
        swSheet.ReloadTemplate(True)
        swApp.ActivateDoc2("Draw1 - Sheet1", False, longstatus)
        Part = swApp.ActiveDoc

        ' Views
        myView = Part.CreateDrawViewFromModelView3(AHUName & "_AHU Final Assembly.SLDASM", "*Back", 0.297 / 3.34, 0.21 / 1.84, 0)
        boolstatus = Draw.ActivateView("Drawing View1")
        Part.ClearSelection2(True)

        ' BOM Table
        boolstatus = Part.ActivateView("Drawing View1")
        myView = Part.ActiveDrawingView
        swBOMTable = myView.InsertBomTable2(False, 0.195, 0.205, 1, 2, "Default", BOMTemp)
        boolstatus = Part.EditRebuild3()

        ' Ballooning
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Draw.ActivateView("Drawing View1")
        autoballoonParams = Part.CreateAutoBalloonOptions()
        autoballoonParams.Layout = swBalloonLayoutType_e.swDetailingBalloonLayout_Square
        autoballoonParams.ReverseDirection = False
        autoballoonParams.IgnoreMultiple = True
        autoballoonParams.InsertMagneticLine = True
        autoballoonParams.LeaderAttachmentToFaces = True
        autoballoonParams.Style = swBalloonStyle_e.swBS_Circular
        autoballoonParams.Size = swBalloonFit_e.swBF_1Char
        autoballoonParams.EditBalloonOption = 1
        autoballoonParams.EditBalloons = 1
        autoballoonParams.UpperTextContent = 1
        autoballoonParams.UpperText = """"
        autoballoonParams.Layername = "0"
        autoballoonParams.ItemNumberStart = 1
        autoballoonParams.ItemNumberIncrement = 1
        autoballoonParams.ItemOrder = 0
        vBaloon = Draw.AutoBalloon5(autoballoonParams)
        Part.ClearSelection2(True)


        ' Note
        Dim X, Y As Decimal
        boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingBalloonStyle, 0, swBalloonStyle_e.swBS_None)

        X = WallWth
        Y = WallHt
        myNote = Draw.CreateText2(AHUName & "_AHU Final Assembly" & vbNewLine & X & "mm x " & Y & "mm", 0.297 / 5.94, 0.21 / 6, 0, 0.004, 0)

        boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDisplayHideAllTypes, True)

        ' Save As
        Part.ViewZoomtofit2()
        longstatus = Part.SaveAs3(SaveFolder & "\AHU\" & AHUName & "_AHU Final Assembly" & ".SLDDRW", 0, 2)
        swApp.CloseAllDocuments(True)

        Part = swApp.OpenDoc6(SaveFolder & "\AHU\" & AHUName & "_AHU Final Assembly" & ".SLDDRW", 3, 0, "", longstatus, longwarnings)
        boolstatus = Part.EditRebuild3()
        longstatus = Part.SaveAs3(SaveFolder & "\PDF Drawings\" & AHUName & "_AHU Final Assembly.PDF", 0, 2)
        swApp.CloseAllDocuments(True)

    End Sub

#End Region

#Region "CNC Partlist"
    Public Sub CreateCNCPartList(ClientName As String, JobNum As String)
        'Exit Sub
        Part = swApp.ActiveDoc
        Assy = Part

        Dim TopLevel As Boolean = False
        Dim fullprtname As New List(Of String)()

        selCount = Assy.GetComponentCount(False)
        AllComp = Assy.GetComponents(False)

        Dim PartNames(selCount - 1) As String
        Dim PartNo(selCount - 1) As String


        For i = 0 To selCount - 1
nextcount:

            If i = selCount Then
                GoTo Jump
            End If

            swComp = AllComp(i)
            PartNames(i) = swComp.Name2

            Dim CompType As ModelDoc2
            CompType = swComp.GetModelDoc2

            If CompType.GetType = 2 Then
                i = i + 1
                GoTo nextcount
            End If

            Dim TempName() As String = Split(PartNames(i), "/")
            If TempName.Length > 1 Then
                PartNames(i) = TempName(1)
            End If

            Dim TempName1() As String = Split(PartNames(i), "-")
            Dim TempName2() As String = Split(TempName1(0), "_")

            '================ skip motor parts =====================
            Dim secondElementLength As Integer = TempName2(2).Length
            If secondElementLength > 3 Then
                i = i + 1

                GoTo nextcount
            End If
            '=======================================================

            If QuantityDictionary.ContainsKey(TempName1(0)) Then
                QuantityDictionary(TempName1(0)) += 1           ' Increment the existing quantity by 1
                i = i + 1
                GoTo nextcount
            Else
                QuantityDictionary.Add(TempName1(0), 1)                  ' Add the part with quantity 1
                i = i + 1
                GoTo nextcount
            End If

Jump:
            For Each kvp1 As KeyValuePair(Of String, Integer) In QuantityDictionary
                fullprtname.Add(kvp1.Key)
                QtyList.Add(kvp1.Value)
            Next

            For Each FullPartName As String In fullprtname
                Dim splitString() As String = FullPartName.Split("_")

                ' Add part numbers (index 0, 1, and 2) to PartNoList                                                         'AAD1235_A_01 --- example
                PartNoList.Add(splitString(0) & "_" & splitString(1) & "_" & splitString(2))

                ' Add part names (index 3 and 4) to PartNameList
                Dim TmpNm As String = ""
                For a = 3 To UBound(splitString)
                    TmpNm = TmpNm & "_" & splitString(a)
                Next

                Dim trimChars As Char() = {"_"}                                                                          ' Array of characters to remove
                TmpNm = TmpNm.TrimStart(trimChars)
                NameList.Add(TmpNm)                                                                                     'vertical channel 01 --- example

            Next

        Next

        My.Computer.FileSystem.CopyFile(LibPath & "\Panel CNC Partlist.xlsx", SaveFolder & "\" & AHUName & "_CNC Part List.xlsx", True)

        Dim oExcel As Excel.Application
        oExcel = CreateObject("excel.application")
        oExcel.Visible = True

        Dim oBook As Excel.Workbook
        oBook = oExcel.Workbooks.Open(SaveFolder & "\" & AHUName & "_CNC Part List.xlsx")

        Dim oSheet As Excel.Worksheet
        oSheet = oBook.Worksheets("TestSheet")

        oSheet.Range("C4").Value = "Job No:-  " & JobNum                                                      'job no.
        oSheet.Range("A3").Value = "Client Name:-  " & ClientName                                             'Client Name


        'Dim CellRange As Range
        Dim CellNo As Integer = 9

        For i = 0 To PartNoList.Count - 1

            oSheet.Rows(CellNo).Resize(1).insert()

            oSheet.Range("A" & CellNo).Value = PartNoList(i)
            oSheet.Range("B" & CellNo).Value = NameList(i)
            oSheet.Range("C" & CellNo).Value = QtyList(i)

            CellNo += 1
        Next

        ' Get the range of cells to sort
        Dim range As Excel.Range = oSheet.Range("A9:C" & CellNo - 1)                                          ' Update with your desired range

        '' Sort the range in ascending order (A to Z)
        'range.Sort(Key1:=range.Columns(1), Order1:=Excel.XlSortOrder.xlAscending, Header:=Excel.XlYesNoGuess.xlYes)


        ' Add the formula to cell C19
        oSheet.Range("C" & CellNo).Formula = "=SUM(C9:C" & CellNo - 1 & ")"


        oBook.Save()
        oBook.Close()
        oBook = Nothing
        oExcel.Quit()
        oExcel = Nothing

        swApp.CloseAllDocuments(True)

    End Sub

#End Region

#Region "Functions ----------------------------------------------------------------------------------------------"

    '8.5mm Holes ------------------------------------------------------------------------------------------------
    Function InterBoltingNumber(Length As Decimal, HoleClearence As Decimal) As Decimal

        Dim MinDis As Decimal = 250
        Dim MaxDis As Decimal = 300

        Dim ActualLength As Decimal = Length - (2 * HoleClearence)
        Dim ActualBoltDis As Decimal

        If ActualLength < MinDis Then
            LastLgt = ActualLength
        End If

        Dim BoltNoMax As Integer = Truncate(ActualLength / MaxDis)
        Dim BoltNoMin As Integer = Truncate(ActualLength / MinDis)

        If BoltNoMax > 1 Or BoltNoMin > 1 Then
            If BoltNoMax > 1 Then
                ActualBoltDis = (ActualLength / (BoltNoMax - 1))
            Else
                ActualBoltDis = (ActualLength / (BoltNoMin - 1))
            End If

            If ActualBoltDis > MaxDis Then
                While ActualBoltDis > MaxDis
                    BoltNoMax += 1
                    ActualBoltDis = (ActualLength / (BoltNoMax - 1))
                    LastLgt = (ActualLength / (BoltNoMax - 1))
                End While

            ElseIf ActualBoltDis < MinDis Then
                While ActualBoltDis < MinDis
                    BoltNoMax -= 1
                    ActualBoltDis = (ActualLength / (BoltNoMax - 1))
                End While
            End If
        Else
            BoltNoMax = 1
        End If

        If ActualLength > 200 And ActualLength <= 400 Then
            BoltNoMax = 2
        ElseIf BoltNoMax < 3 And ActualLength > 400 Then
            BoltNoMax = 3
        End If

        'If ActualLength > MaxDis Then
        '    BoltNoMax = BoltNoMax + 1
        'End If

        Return BoltNoMax
        'Return LastLgt / 1000

    End Function

    Function BoltDistance(BoltNumber As Decimal, Length As Decimal, HoleClearence As Decimal) As Decimal

        Dim BoltDis As Decimal

        Dim ActualLength As Decimal = Length - (2 * HoleClearence)

        If BoltNumber >= 2 Then
            BoltDis = ActualLength / (BoltNumber - 1)

        ElseIf BoltNumber = 1 Then
            BoltDis = ActualLength / 1
        Else
            BoltDis = 0
        End If

        If ActualLength > 200 And ActualLength <= 400 Then
            BoltDis = ActualLength
        ElseIf ActualLength > 400 And BoltDis = 0 Then
            BoltDis = ActualLength
        End If

        Return BoltDis / 1000

    End Function

    '310mm to 450mm 9.2mm Holes ------------------------------------------------------------------------------------------------

    Function SmolInterBoltingNumber(Length As Decimal, HoleClearence As Decimal) As Decimal

        Dim MinDis As Decimal = 150
        Dim MaxDis As Decimal = 200

        Dim ActualLength As Decimal = Length - (2 * HoleClearence)
        Dim ActualBoltDis As Decimal

        If ActualLength < MinDis Then
            LastLgt = ActualLength
        End If

        Dim BoltNoMax As Integer = Truncate(ActualLength / MaxDis)
        Dim BoltNoMin As Integer = Truncate(ActualLength / MinDis)

        If BoltNoMax > 1 Or BoltNoMin > 1 Then
            If BoltNoMax > 1 Then
                ActualBoltDis = (ActualLength / (BoltNoMax - 1))
            Else
                ActualBoltDis = (ActualLength / (BoltNoMin - 1))
            End If

            If ActualBoltDis > MaxDis Then
                While ActualBoltDis > MaxDis
                    BoltNoMax += 1
                    ActualBoltDis = (ActualLength / (BoltNoMax - 1))
                    LastLgt = (ActualLength / (BoltNoMax - 1))
                End While

            ElseIf ActualBoltDis < MinDis Then
                While ActualBoltDis < MinDis
                    BoltNoMax -= 1
                    ActualBoltDis = (ActualLength / (BoltNoMax - 1))
                End While
            End If
        Else
            BoltNoMax = 1
        End If

        If ActualLength > 200 And ActualLength <= 400 Then
            BoltNoMax = 2
        ElseIf BoltNoMax < 3 And ActualLength > 400 Then
            BoltNoMax = 3
        End If

        'If ActualLength > MaxDis Then
        '    BoltNoMax = BoltNoMax + 1
        'End If

        Return BoltNoMax
        'Return LastLgt / 1000

    End Function

    Function SmolBoltDistance(BoltNumber As Decimal, Length As Decimal, HoleClearence As Decimal) As Decimal

        Dim BoltDis As Decimal

        If BoltNumber > 1 Then
            BoltDis = (Length - (2 * HoleClearence)) / (BoltNumber - 1)
        Else
            BoltDis = 0
        End If

        Return BoltDis / 1000

    End Function

    'Slots ------------------------------------------------------------------------------------------------
    Function SlotsInterBoltingNumber(Length As Decimal, HoleClearence As Decimal) As Decimal

        Dim MinDis As Decimal = 250
        Dim MaxDis As Decimal = 300

        Dim ActualLength As Decimal = Length - (2 * HoleClearence)
        Dim ActualBoltDis As Decimal

        If ActualLength < MinDis Then
            LastLgt = ActualLength
        End If

        Dim BoltNoMax As Integer = Truncate(ActualLength / MaxDis)
        Dim BoltNoMin As Integer = Truncate(ActualLength / MinDis)

        If BoltNoMax > 1 Or BoltNoMin > 1 Then
            If BoltNoMax > 1 Then
                ActualBoltDis = (ActualLength / (BoltNoMax - 1))
            Else
                ActualBoltDis = (ActualLength / (BoltNoMin - 1))
            End If

            If ActualBoltDis > MaxDis Then
                While ActualBoltDis > MaxDis
                    BoltNoMax += 1
                    ActualBoltDis = (ActualLength / (BoltNoMax - 1))
                    LastLgt = (ActualLength / (BoltNoMax - 1))
                End While

            ElseIf ActualBoltDis < MinDis Then
                While ActualBoltDis < MinDis
                    BoltNoMax -= 1
                    ActualBoltDis = (ActualLength / (BoltNoMax - 1))
                End While
            End If
        Else
            BoltNoMax = 1
        End If

        If ActualLength > 200 And ActualLength <= 400 Then
            BoltNoMax = 2
        ElseIf BoltNoMax < 3 And ActualLength > 400 Then
            BoltNoMax = 3
        End If

        'If ActualLength > MaxDis Then
        '    BoltNoMax = BoltNoMax + 1
        'End If

        Return BoltNoMax
        'Return LastLgt / 1000

    End Function

    Function SlotsBoltDistance(BoltNumber As Decimal, Length As Decimal, HoleClearence As Decimal) As Decimal

        Dim BoltDis As Decimal

        Dim ActualLength As Decimal = Length - (2 * HoleClearence)

        If BoltNumber >= 2 Then
            BoltDis = ActualLength / (BoltNumber - 1)
        Else
            BoltDis = 0
        End If

        If ActualLength > 200 And ActualLength <= 400 Then
            BoltDis = ActualLength
        ElseIf ActualLength > 400 And BoltDis = 0 Then
            BoltDis = ActualLength
        End If

        Return BoltDis / 1000

    End Function

    '4.2mm Holes -----------------------------------------------------------------------------------------------
    Function SideHoleNumber(Length As Decimal, HoleClearence As Decimal) As Decimal

        Dim MinDis As Decimal = 0.125
        Dim MaxDis As Decimal = 0.15

        Dim ActualLength As Decimal = Length - (2 * HoleClearence)

        Dim HoleNo As Integer = Truncate(ActualLength / MaxDis)

        If HoleNo > 1 Then
            Dim ActualHoleDis As Decimal = (ActualLength / HoleNo)
            If ActualHoleDis > MaxDis Then
                While ActualHoleDis > MaxDis
                    HoleNo += 1
                    ActualHoleDis = (ActualLength / HoleNo)
                End While
            ElseIf ActualHoleDis < MinDis Then
                While ActualHoleDis < MinDis
                    HoleNo -= 1
                    ActualHoleDis = (ActualLength / HoleNo)
                End While
            End If
        Else
            HoleNo = 1
        End If

        Return HoleNo

    End Function

    Function SideHoleDist(HoleNumber As Decimal, Length As Decimal, HoleClearence As Decimal) As Decimal

        Dim HoleDis As Decimal
        If HoleNumber > 1 Then
            HoleDis = (Length - (2 * HoleClearence)) / HoleNumber
        Else
            HoleDis = 0
        End If

        Return HoleDis

    End Function

    Function SideHoleNumberTopBotL(Length As Decimal, HoleClearence As Decimal) As Decimal

        Dim MinDis As Decimal = 125
        Dim MaxDis As Decimal = 150

        Dim ActualLength As Decimal = Length - (2 * HoleClearence)

        Dim HoleNo As Integer = Truncate(ActualLength / MaxDis)

        If HoleNo > 1 Then
            Dim ActualHoleDis As Decimal = (ActualLength / HoleNo)
            If ActualHoleDis > MaxDis Then
                While ActualHoleDis > MaxDis
                    HoleNo += 1
                    ActualHoleDis = (ActualLength / HoleNo)
                End While
            ElseIf ActualHoleDis < MinDis Then
                While ActualHoleDis < MinDis
                    HoleNo -= 1
                    ActualHoleDis = (ActualLength / HoleNo)
                End While
            End If
        Else
            HoleNo = 1
        End If

        Return HoleNo

    End Function

    Function SideHoleDistTopBotL(HoleNumber As Decimal, Length As Decimal, HoleClearence As Decimal) As Decimal

        Dim HoleDis As Decimal
        If HoleNumber > 1 Then
            HoleDis = (Length - (2 * HoleClearence)) / HoleNumber
        Else
            HoleDis = 0
        End If

        Return HoleDis / 1000

    End Function


    Public Sub ListofAllParts()

        Dim Directory As String = SaveFolder & "\AHU\"
        Dim Files() As FileInfo
        Dim dirinfo As New DirectoryInfo(Directory)

        Files = dirinfo.GetFiles("*", SearchOption.AllDirectories)
        For Each File In Files
            AssyComponents.Add(File.ToString)
        Next
        'Dim x As Integer = PartNameList.Count
        'For i = 0 To x - 1
        '    Part = swApp.OpenDoc6(PartNameList.Item(i), 1, 0, "", longstatus, longwarnings)
        'Next

    End Sub

#End Region

End Class
