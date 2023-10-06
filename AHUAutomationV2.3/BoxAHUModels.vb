Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst

Imports System.Math

Public Class BoxAHUModels

    Dim swApp As New SldWorks

    Dim Part As ModelDoc2
    Dim Draw As DrawingDoc
    Dim Assy As AssemblyDoc
    Dim boolstatus As Boolean
    Dim longstatus As Long, longwarnings As Long
    Dim myDimension As Dimension
    Dim BlockUtil As MathUtility
    Dim Blockpoint As MathPoint
    Dim myBlockDefinition As SketchBlockDefinition
    Dim excludedComponents As Object
    Dim myFeature As Object
    Dim myMate As Object
    Dim myView As View
    Dim DrawView As View
    Dim BaseView As View
    Dim skSegment As Object
    Dim myDisplayDim As Object
    Dim swBOMTable As BomTableAnnotation
    Dim swTable As TableAnnotation
    Dim autoballoonParams As AutoBalloonOptions
    Dim vBaloon As Object
    Dim myNote As Note

    ReadOnly DrawTemp As String = "C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Drawing\AADTech Drawing Template.DRWDOT"
    ReadOnly DrawSheet As String = "C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Drawing\AADTech Sheet Format - A4 - Landscape.slddrt"
    ReadOnly BOMTemp As String = "C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Drawing\BOM Template.sldbomtbt"

    Dim StdFunc As New Standard_Functions
    Dim predictivedb As New PredictiveDBInput
    Dim bomData As New BOMExcel

    Public Client As String
    Public AHUName As String
    Public JobNo As String
    Public ArticleNoFan As String

#Region "BOX Models"

    Public Sub BackSheet(Wth As Decimal, Ht As Decimal, Dpth As Decimal, FanDia As Decimal, HoleCD As Decimal)

        ' Start 3D Model
        'Open File
        Part = swApp.OpenDoc6("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Box - Sample\01_back bot sheet.sldprt", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("01_back bot sheet", False, longstatus)
        Part = swApp.ActiveDoc

        Part.ViewZoomtofit2()

        'Re Dim
        boolstatus = Part.Extension.SelectByID2("BoxWidth@BaseFlange@01_back bot sheet.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 0)
        myDimension = Part.Parameter("BoxWidth@BaseFlange")
        myDimension.SystemValue = Wth - 0.0012
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("BoxHeight@BaseFlange@01_back bot sheet.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 0)
        myDimension = Part.Parameter("BoxHeight@BaseFlange")
        myDimension.SystemValue = Ht
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("BoxDepth@BottomFlange@01_back bot sheet.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 0)
        myDimension = Part.Parameter("BoxDepth@BottomFlange")
        myDimension.SystemValue = Dpth
        Part.ClearSelection2(True)

        boolstatus = Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Fan Duct
        boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        Part.SketchManager.InsertSketch(True)
        Part.ClearSelection2(True)
        'skSegment = Part.SketchManager.CreateCircle(0, HoleCD, 0, FanDuctDia / 2, HoleCD, 0)
        'Part.SketchAddConstraints("sgFIXED")
        'Part.ClearSelection2(True)

        Dim ArData(2) As Double
        ArData(0) = 0
        ArData(1) = HoleCD
        ArData(2) = 0

        BlockUtil = swApp.GetMathUtility
        Blockpoint = BlockUtil.CreatePoint(ArData)
        myBlockDefinition = Part.SketchManager.MakeSketchBlockFromFile(Blockpoint, "C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Motor Cutout\" & FanDia & "_motor cutout.SLDBLK", False, 1, 0)

        Part.ViewOrientationUndo()
        Part.SketchManager.InsertSketch(True)

        myFeature = Part.FeatureManager.FeatureCut3(True, False, True, 0, 1, 0.002, 0.01, False, False, False, False, 0, 0, False, False, False, False, True, True, True, True, True, False, 0, 0, False)
        Part.SelectionManager.EnableContourSelection = False
        Part.ClearSelection2(True)

        'Save File
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_01_back bot sheet.SLDPRT", 0, 2)
        Part = Nothing

        StdFunc.CloseActiveDoc()

        'BOM Entries
        bomData.EnterValuesInCNC(JobNo & "_01_back bot sheet", (Wth * 1000) + 57.81, ((Ht + Dpth) * 1000) + 112.02, "2.0", 1, Client, AHUName, JobNo)

    End Sub

    Public Sub BackSheetDrawing()
        Exit Sub
        ' Variables
        Dim FilePath As String = "C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_01_back bot sheet.SLDPRT"
        Dim FileName As String = JobNo & "_01_back bot sheet"

        ' Open File
        Part = swApp.OpenDoc6(FilePath, 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2(FileName, False, longstatus)
        Part = swApp.ActiveDoc

        Part.ViewZoomtofit2()

        'Bounding Box
        Dim BBox As Object = StdFunc.BoundingBox()
        Dim xDim As Decimal = Abs(BBox(0)) + Abs(BBox(3))
        Dim yDim As Decimal = Abs(BBox(1)) + Abs(BBox(4))
        Dim zDim As Decimal = Abs(BBox(2)) + Abs(BBox(5))

        'Bounding Box - Flat
        boolstatus = Part.Extension.SelectByID2("Flat-Pattern46", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditUnsuppress2()

        Dim BBoxFlat As Object = StdFunc.BoundingBox()
        Dim xDimFlat As Decimal = Abs(BBoxFlat(0)) + Abs(BBoxFlat(3))
        Dim yDimFlat As Decimal = Abs(BBoxFlat(1)) + Abs(BBoxFlat(4))
        Dim zDimFlat As Decimal = Abs(BBoxFlat(2)) + Abs(BBoxFlat(5))

        boolstatus = Part.Extension.SelectByID2("Flat-Pattern46", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditSuppress2()

        swApp.CloseDoc(FileName & ".SLDPRT")

        ' Sheet Scale
        Dim SScale As Integer = Ceiling((xDimFlat + zDim + xDim) / (0.297 - (0.03 + 0.04 + 0.04 + 0.03)))

        ' Adjust Values for Scale
        xDim /= SScale
        yDim /= SScale
        zDim /= SScale

        xDimFlat /= SScale
        yDimFlat /= SScale
        zDimFlat /= SScale

        ' Get Margins
        Dim marginX As Decimal = (0.297 - (xDimFlat + 0.04 + zDim + 0.04 + xDim)) / 2
        If marginX > 0.03 Then marginX = 0.03
        Dim marginY As Decimal = (0.21 - (zDim + 0.04 + yDimFlat)) / 2

        ' Calculate View Placements
        Dim xTop As Decimal = marginX + xDimFlat / 2
        Dim yTop As Decimal = marginY + zDim / 2
        Dim yFrontFlat As Decimal = yTop + zDim / 2 + 0.04 + yDimFlat / 2
        Dim xRight As Decimal = xTop + xDimFlat / 2 + 0.04 + zDim / 2
        Dim xIso As Decimal = xRight + zDim / 2 + 0.04 + xDim / 2

        ' Start Drawing
        Part = swApp.NewDocument(DrawTemp, 2, 0.297, 0.21)
        Draw = Part
        boolstatus = Draw.SetupSheet5("Sheet1", 11, 12, 1, SScale, False, DrawSheet, 0.297, 0.21, "Default", False)

        swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swAutomaticScaling3ViewDrawings, False)

        ' Views
        boolstatus = Draw.GenerateViewPaletteViews(FilePath)

        'Front Flat
        boolstatus = Draw.CreateFlatPatternViewFromModelView(FilePath, "Default", xTop, yFrontFlat, 0)
        boolstatus = Part.ActivateView("Drawing View12")
        Part.ClearSelection2(True)

        'Front - Outside Sheet
        myView = Draw.CreateDrawViewFromModelView3(FilePath, "*Front", -xTop, yFrontFlat, 0)
        boolstatus = Draw.ActivateView("Drawing View13")
        Part.ClearSelection2(True)

        'Right - Section
        boolstatus = Draw.ActivateView("Drawing View13")
        skSegment = Part.SketchManager.CreateLine(0, (yDim * SScale / 2) + 0.15, 0, 0, -(yDim * SScale / 2) - 0.15, 0)
        boolstatus = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xRight, yFrontFlat, 0, "A", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View14")
        Part.ClearSelection2(True)

        'Top - Section
        boolstatus = Draw.ActivateView("Drawing View13")
        skSegment = Part.SketchManager.CreateLine(-(xDim * SScale / 2) - 0.15, 0, 0, (xDim * SScale / 2) + 0.15, 0, 0)
        boolstatus = Part.Extension.SelectByID2("Line2", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xTop, yTop, 0, "B", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View15")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View15", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swNoViewAlignment, Nothing)
        Part.ClearSelection2(True)

        boolstatus = Draw.ActivateView("Drawing View15")
        boolstatus = Part.Extension.SelectByID2("Drawing View12", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewVerticalCenter, BaseView)
        Part.ClearSelection2(True)

        'Isometric
        myView = Draw.CreateDrawViewFromModelView3(FilePath, "*Isometric", xIso, yFrontFlat, 0)
        boolstatus = Draw.ActivateView("Drawing View16")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View16", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Draw.SelectionManager.GetSelectedObject6(1, -1)
        boolstatus = DrawView.SetDisplayMode4(False, swDisplayMode_e.swSHADED, True, True, False)

        boolstatus = Part.ActivateSheet("Sheet1")

        ' Dimentions
        'Height - Flat
        boolstatus = Part.Extension.SelectByRay(xTop, yFrontFlat + (yDimFlat / 2), -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xTop, yFrontFlat - (yDimFlat / 2), -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xTop + (xDimFlat / 2) + 0.015, yFrontFlat, 0)
        Part.ClearSelection2(True)

        'Width - Flat
        boolstatus = Part.Extension.SelectByRay(xTop + (xDimFlat / 2) - (0.0136 / SScale), yFrontFlat, -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xTop - (xDimFlat / 2), yFrontFlat, -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xTop, yFrontFlat - (yDimFlat / 2) - 0.015, 0)
        Part.ClearSelection2(True)

        'Width - Top Section
        boolstatus = Part.Extension.SelectByRay(xTop + (xDim / 2), yTop + (zDim / 2) - (0.0125 / SScale), -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xTop - (xDim / 2), yTop + (zDim / 2) - (0.0125 / SScale), -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xTop, yTop + (zDim / 2) + 0.01, 0)
        Part.ClearSelection2(True)

        'Depth - Top Section
        boolstatus = Part.Extension.SelectByRay(xTop + (xDim / 2) - (0.0125 / SScale), yTop + (zDim / 2), -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xTop + (xDim / 2) - (0.0075 / SScale), yTop + (zDim / 2) - (0.05 / SScale), -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xTop + (xDim / 2) + 0.01, yTop + (zDim / 2) + 0.01, 0)
        Part.ClearSelection2(True)

        'Height - Right Section
        boolstatus = Part.Extension.SelectByRay(xRight + (zDim / 2) - (0.0125 / SScale), yFrontFlat + (yDim / 2), -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xRight, yFrontFlat - (yDim / 2), -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xRight + (zDim / 2) + 0.015, yFrontFlat, 0)
        Part.ClearSelection2(True)

        'Width - Right Section
        boolstatus = Part.Extension.SelectByRay(xRight + (zDim / 2), yFrontFlat - (yDim / 2) + (0.05 / SScale), -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xRight - (zDim / 2), yFrontFlat - (yDim / 2) + (0.0125 / SScale), -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xRight, yFrontFlat - (yDim / 2) - 0.015, 0)
        Part.ClearSelection2(True)

        boolstatus = Part.ActivateSheet("Sheet1")

        ' Note
        Dim X, Y, Z As Decimal
        boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingBalloonStyle, 0, swBalloonStyle_e.swBS_None)

        X = Round(xDimFlat * SScale * 1000, 2)
        Y = Round(yDimFlat * SScale * 1000, 2)
        Z = Round(zDimFlat * SScale * 1000, 2)
        myNote = Draw.CreateText2(FileName & vbNewLine & "Qty - " & vbNewLine & X & "mm x " & Y & "mm x " & Z & "mm", xRight - (zDim / 2) - 0.015, yTop, 0, 0.004, 0)

        ' Save As
        Part.ViewZoomtofit2()
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Drawings\Motor Box\" & JobNo & "_01_back bot sheet.SLDDRW", 0, 2)
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Drawings\Motor Box\" & JobNo & "_01_back bot sheet.PDF", 0, 2)
        swApp.CloseAllDocuments(True)

    End Sub

    Public Sub LHSSheet(Ht As Decimal, Dpth As Decimal, HoleCD As Decimal)

        Part = swApp.OpenDoc6("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Box - Sample\02B_lhs sheet.sldprt", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("02B_lhs sheet", False, longstatus)
        Part = swApp.ActiveDoc

        Part.ViewZoomtofit2()

        swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchInference, False)    'Snaping OFF

        'Resize Dimentions
        boolstatus = Part.Extension.SelectByID2("BoxDepth@BaseFlange@02B_lhs sheet.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 0)
        myDimension = Part.Parameter("Depth@BaseFlange")
        myDimension.SystemValue = Dpth - 0.0015
        boolstatus = Part.Extension.SelectByID2("BoxHeight@BaseFlange@02B_lhs sheet.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 0)
        myDimension = Part.Parameter("Height@BaseFlange")
        myDimension.SystemValue = Ht - 0.0015
        Part.ClearSelection2(True)

        boolstatus = Part.EditRebuild3()

        Part.ClearSelection2(True)
        Part.ViewZoomtofit2()

        'Right Bolt Holes
        boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        Part.SketchManager.InsertSketch(True)
        Part.ClearSelection2(True)

        skSegment = Part.SketchManager.CreateCircle(Dpth - 0.025, Ht - 0.05 - 0.0015, 0#, Dpth - 0.025, Ht - 0.05 - 0.0015 + 0.0045, 0#)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)
        skSegment = Part.SketchManager.CreateCircle(Dpth - 0.025, (Ht / 2) - 0.0015, 0#, Dpth - 0.025, (Ht / 2) - 0.0015 + 0.0045, 0#)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)
        skSegment = Part.SketchManager.CreateCircle(Dpth - 0.025, 0.0235, 0#, Dpth - 0.025, 0.0235 + 0.0045, 0#)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)

        skSegment = Part.SketchManager.CreateCircle(0.025, Ht - 0.05 - 0.0015, 0#, 0.025, Ht - 0.05 - 0.0015 + 0.0045, 0#)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)
        skSegment = Part.SketchManager.CreateCircle(0.025, (Ht / 2) - 0.0015, 0#, 0.025, (Ht / 2) - 0.0015 + 0.0045, 0#)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)
        skSegment = Part.SketchManager.CreateCircle(0.05, 0.0235, 0#, 0.05, 0.0235 + 0.0045, 0#)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)

        Part.ViewOrientationUndo()
        Part.ClearSelection2(True)

        myFeature = Part.FeatureManager.FeatureCut3(True, False, True, 1, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False)
        Part.SelectionManager.EnableContourSelection = False
        Part.ClearSelection2(True)
        Part.ViewZoomtofit2()

        swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchInference, True)    'Snaping ON

        'Save File      
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_02B_lhs sheet.SLDPRT", 0, 2)
        Part = Nothing
        StdFunc.CloseActiveDoc()

        'BOM Entries
        bomData.EnterValuesInCNC(JobNo & "_02B_lhs sheet", (Dpth * 1000) + 67.71, (Ht * 1000) + 67.71, "2.0", 1, Client, AHUName, JobNo)

    End Sub

    Public Sub LHSBeam(Ht As Decimal, Dpth As Decimal)

        'Variables
        Ht -= 0.05

        'Open File
        Part = swApp.OpenDoc6("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Box - Sample\02_lhs-l.sldprt", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("04_rhs -l", False, longstatus)
        Part = swApp.ActiveDoc

        Part.ViewZoomtofit2()

        'Resize Dimentions
        boolstatus = Part.Extension.SelectByID2("Height@BaseFlange@02_lhs-l.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 0)
        myDimension = Part.Parameter("Height@BaseFlange")
        myDimension.SystemValue = Ht
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Depth@Edge-Flange1@02_lhs-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        myDimension = Part.Parameter("Depth@Edge-Flange1")
        myDimension.SystemValue = Dpth - 0.05
        Part.ClearSelection2(True)
        boolstatus = Part.EditRebuild3()

        Part.ViewZoomtofit2()

        swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchInference, False)    'Snaping OFF

        'Holes - Right
        boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        Part.SketchManager.InsertSketch(True)
        Part.ClearSelection2(True)

        skSegment = Part.SketchManager.CreateCircle(0.025, Ht - 0.05, 0, 0.025, Ht - 0.05 + 0.0045, 0#)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)
        skSegment = Part.SketchManager.CreateCircle(0.025, Ht / 2, 0, 0.025, (Ht / 2) + 0.0045, 0#)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)
        skSegment = Part.SketchManager.CreateCircle(0.025, 0.025, 0, 0.025, 0.025 + 0.0045, 0#)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)

        Part.ViewOrientationUndo()
        Part.SketchManager.InsertSketch(True)

        myFeature = Part.FeatureManager.FeatureCut3(True, False, False, 1, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False)
        Part.SelectionManager.EnableContourSelection = False
        Part.ClearSelection2(True)

        'Holes - Front
        boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        Part.SketchManager.InsertSketch(True)
        Part.ClearSelection2(True)

        skSegment = Part.SketchManager.CreateCircle(-0.015, Ht - 0.05, 0, -0.015, Ht - 0.05 - 0.0015 + 0.0045, 0)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)
        skSegment = Part.SketchManager.CreateCircle(-0.015, Ht / 2, 0, -0.015, (Ht / 2) - 0.0015 + 0.0045, 0)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)
        skSegment = Part.SketchManager.CreateCircle(-0.015, 0.025, 0, -0.015, 0.0235 + 0.0045, 0)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)

        skSegment = Part.SketchManager.CreateCircle(0.015, Ht - 0.05, 0, 0.015, Ht - 0.05 + 0.0045, 0)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)
        skSegment = Part.SketchManager.CreateCircle(0.015, Ht / 2, 0, 0.015, (Ht / 2) + 0.0045, 0)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)
        skSegment = Part.SketchManager.CreateCircle(0.015, 0.025, 0, 0.015, 0.025 + 0.0045, 0)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)

        Part.ViewOrientationUndo()
        Part.SketchManager.InsertSketch(True)

        myFeature = Part.FeatureManager.FeatureCut3(True, False, False, 11, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False)
        Part.SelectionManager.EnableContourSelection = False
        Part.ClearSelection2(True)

        swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchInference, True)    'Snaping ON

        'Save File
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_02_lhs-l.SLDPRT", 0, 2)
        Part = Nothing
        StdFunc.CloseActiveDoc()

        'BOM Entries
        bomData.EnterValuesInCNC(JobNo & "_02_lhs-l", ((Dpth + Ht) * 1000) - 3.6, "108.81", "2.0", 1, Client, AHUName, JobNo)

    End Sub

    Public Sub LHSDrawing()
        Exit Sub
        ' Variables
        Dim FilePathLHS As String = "C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_02_lhs-l.SLDPRT"
        Dim FileNameLHS As String = JobNo & "_02_lhs-l"
        Dim FilePathLHSSheet As String = "C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_02B_lhs sheet.SLDPRT"
        Dim FileNameLHSSheet As String = JobNo & "_02B_lhs sheetl"

        ' LHS
        'Open File
        Part = swApp.OpenDoc6(FilePathLHS, 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2(FileNameLHS, False, longstatus)
        Part = swApp.ActiveDoc

        Part.ViewZoomtofit2()

        'Bounding Box
        Dim BBox_LHS As Object = StdFunc.BoundingBox()
        Dim xDimLHS As Decimal = Abs(BBox_LHS(0)) + Abs(BBox_LHS(3))
        Dim yDimLHS As Decimal = Abs(BBox_LHS(1)) + Abs(BBox_LHS(4))
        Dim zDimLHS As Decimal = Abs(BBox_LHS(2)) + Abs(BBox_LHS(5))

        'Bounding Box - Flat
        boolstatus = Part.Extension.SelectByID2("Flat-Pattern15", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditUnsuppress2()

        Dim BBoxFlatLHS As Object = StdFunc.BoundingBox()
        Dim xDimFlatLHS As Decimal = Abs(BBoxFlatLHS(0)) + Abs(BBoxFlatLHS(3))
        Dim yDimFlatLHS As Decimal = Abs(BBoxFlatLHS(1)) + Abs(BBoxFlatLHS(4))
        Dim zDimFlatLHS As Decimal = Abs(BBoxFlatLHS(2)) + Abs(BBoxFlatLHS(5))

        boolstatus = Part.Extension.SelectByID2("Flat-Pattern15", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditSuppress2()

        swApp.CloseAllDocuments(True)

        ' LHS Sheet
        'Open File
        Part = swApp.OpenDoc6(FilePathLHSSheet, 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2(FileNameLHSSheet, False, longstatus)
        Part = swApp.ActiveDoc

        Part.ViewZoomtofit2()

        'Bounding Box
        Dim BBox_LHSSheet As Object = StdFunc.BoundingBox()
        Dim xDimLHSSheet As Decimal = Abs(BBox_LHSSheet(0)) + Abs(BBox_LHSSheet(3))
        Dim yDimLHSSheet As Decimal = Abs(BBox_LHSSheet(1)) + Abs(BBox_LHSSheet(4))
        Dim zDimLHSSheet As Decimal = Abs(BBox_LHSSheet(2)) + Abs(BBox_LHSSheet(5))

        'Bounding Box - Flat
        boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditUnsuppress2()

        Dim BBoxFlatLHSSheet As Object = StdFunc.BoundingBox()
        Dim xDimFlatLHSSheet As Decimal = Abs(BBoxFlatLHSSheet(0)) + Abs(BBoxFlatLHSSheet(3))
        Dim yDimFlatLHSSheet As Decimal = Abs(BBoxFlatLHSSheet(1)) + Abs(BBoxFlatLHSSheet(4))
        Dim zDimFlatLHSSheet As Decimal = Abs(BBoxFlatLHSSheet(2)) + Abs(BBoxFlatLHSSheet(5))

        boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditSuppress2()

        swApp.CloseAllDocuments(True)

        ' Sheet Scale
        Dim SScaleX As Integer = Ceiling((xDimFlatLHS + zDimLHS + zDimLHSSheet) / (0.297 - (0.03 + 0.04 + 0.06 + 0.03)))
        Dim SScaleY As Integer = Ceiling((0.05 + yDimFlatLHS) / (0.21 - (0.03 + 0.03 + 0.03)))
        Dim SScale As Integer = SScaleX
        If SScaleX < SScaleY Then
            SScale = SScaleY
        End If

        ' Adjust Values for Scale
        xDimLHS /= SScale
        yDimLHS /= SScale
        zDimLHS /= SScale

        xDimFlatLHS /= SScale
        yDimFlatLHS /= SScale
        zDimFlatLHS /= SScale

        xDimLHSSheet /= SScale
        yDimLHSSheet /= SScale
        zDimLHSSheet /= SScale

        xDimFlatLHSSheet /= SScale
        yDimFlatLHSSheet /= SScale
        zDimFlatLHSSheet /= SScale

        ' Get Margins
        Dim marginX As Decimal = (0.297 - (xDimFlatLHS + 0.04 + zDimLHS + 0.04 + zDimLHSSheet)) / 2
        If marginX < 0.03 Then marginX = 0.03

        Dim marginY As Decimal = (0.21 - (yDimFlatLHS + 0.03 + (0.05 / SScale))) / 2
        If marginY < 0.03 Then marginY = 0.03

        ' Calculate View Placements
        Dim xFrontFlatLHS As Decimal = marginX + xDimFlatLHS / 2
        Dim yFrontSecLHS As Decimal = marginY + 0.05 / SScale / 2
        Dim yFrontFlatLHS As Decimal = yFrontSecLHS + 0.05 / SScale / 2 + 0.03 + yDimFlatLHS / 2
        Dim xRightSecLHS As Decimal = xFrontFlatLHS + xDimFlatLHS / 2 + 0.04 + zDimLHS / 2
        Dim xFrontLHSSheet As Decimal = xRightSecLHS + zDimLHS / 2 + 0.04 + zDimLHSSheet / 2
        Dim yFrontLHSSheet As Decimal = marginY + yDimLHSSheet / 2

        ' Open Parts
        Part = swApp.OpenDoc6(FilePathLHS, 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(FilePathLHSSheet, 1, 0, "", longstatus, longwarnings)

        ' Start Drawing
        Part = swApp.NewDocument(DrawTemp, 2, 0.297, 0.21)
        Draw = Part
        boolstatus = Draw.SetupSheet5("Sheet1", 11, 12, 1, SScale, False, DrawSheet, 0.297, 0.21, "Default", False)

        swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swAutomaticScaling3ViewDrawings, False)

        ' Views - LHS
        'Front Flat
        boolstatus = Draw.CreateFlatPatternViewFromModelView(FilePathLHS, "Default", xFrontFlatLHS, yFrontFlatLHS, 0)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Draw.ActivateView("Drawing View1")
        Part.ClearSelection2(True)

        'Front - Outside
        myView = Draw.CreateDrawViewFromModelView3(FilePathLHS, "*Front", -xFrontFlatLHS, yFrontFlatLHS, 0)
        boolstatus = Draw.ActivateView("Drawing View2")
        Part.ClearSelection2(True)

        'Side - Outside
        myView = Draw.CreateDrawViewFromModelView3(FilePathLHS, "*Right", -xFrontFlatLHS, yFrontSecLHS - (yDimLHS / 2) + 0.05 / SScale / 2, 0)
        boolstatus = Draw.ActivateView("Drawing View3")
        Part.ClearSelection2(True)

        'Right - Section
        boolstatus = Draw.ActivateView("Drawing View2")
        skSegment = Part.SketchManager.CreateLine(0, (yDimLHS * SScale / 2) + 0.15, 0, 0, -(yDimLHS * SScale / 2) - 0.15, 0)
        boolstatus = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xRightSecLHS, yFrontFlatLHS, 0, "A", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View4")
        Part.ClearSelection2(True)

        'Front - Section
        boolstatus = Draw.ActivateView("Drawing View3")
        skSegment = Part.SketchManager.CreateLine(0, (yDimLHS * SScale / 2) + 0.1, 0, 0, (yDimLHS * SScale / 2) - 0.15, 0)
        boolstatus = Part.Extension.SelectByID2("Line2", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xFrontFlatLHS, yFrontSecLHS, 0, "B", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View5")
        Part.ClearSelection2(True)

        ' Views - LHS Sheet
        'Front
        myView = Part.CreateDrawViewFromModelView3(FilePathLHSSheet, "*Right", xFrontLHSSheet, yFrontLHSSheet, 0)
        boolstatus = Draw.ActivateView("Drawing View7")
        Part.ClearSelection2(True)

        boolstatus = Part.ActivateSheet("Sheet1")

        ' Dimentions
        'Height - LHS Flat
        boolstatus = Part.Extension.SelectByRay(xFrontFlatLHS, yFrontFlatLHS + (yDimFlatLHS / 2), -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlatLHS, yFrontFlatLHS - (yDimFlatLHS / 2), -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlatLHS + (xDimFlatLHS / 2) + 0.015, yFrontFlatLHS, 0)
        Part.ClearSelection2(True)

        'Width - LHS Flat
        boolstatus = Part.Extension.SelectByRay(xFrontFlatLHS + (xDimFlatLHS / 2), yFrontFlatLHS + (yDimFlatLHS / 2) - (zDimLHS / 2), -7000, 0, 0, -1, 0.001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlatLHS - (xDimFlatLHS / 2), yFrontFlatLHS + (yDimFlatLHS / 2) - (0.1 / SScale), -7000, 0, 0, -1, 0.001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlatLHS + (xDimFlatLHS / 2) + 0.015, yFrontFlatLHS + (yDimFlatLHS / 2) + 0.015, 0)
        Part.ClearSelection2(True)

        'Height - LHS Right Section
        boolstatus = Part.Extension.SelectByRay(xRightSecLHS, yFrontFlatLHS + (yDimLHS / 2), -7000, 0, 0, -1, 0.00001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xRightSecLHS - (zDimLHS / 2) + (0.001 / SScale), yFrontFlatLHS - (yDimLHS / 2), -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xRightSecLHS - (zDimLHS / 2) - 0.015, yFrontFlatLHS, 0)
        Part.ClearSelection2(True)

        'Width - LHS Right Section
        boolstatus = Part.Extension.SelectByRay(xRightSecLHS + (zDimLHS / 2), yFrontFlatLHS + (yDimLHS / 2) - (0.001 / SScale), -7000, 0, 0, -1, 0.00001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xRightSecLHS - (zDimLHS / 2), yFrontFlatLHS, -7000, 0, 0, -1, 0.00001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xRightSecLHS, yFrontFlatLHS + (yDimLHS / 2) + 0.015, 0)
        Part.ClearSelection2(True)

        'Hight - LHS Front Section
        boolstatus = Part.Extension.SelectByRay(xFrontFlatLHS, yFrontSecLHS + 0.05 / SScale / 2, -7000, 0, 0, -1, 0.00001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlatLHS + (xDimLHS / 2) - (0.016 / SScale / 2), yFrontSecLHS - 0.05 / SScale / 2, -7000, 0, 0, -1, 0.00001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlatLHS + (xDimLHSSheet / 2) + 0.015, yFrontSecLHS - (0.05 / SScale / 2) - 0.015, 0)
        Part.ClearSelection2(True)

        'Width - LHS Front Section
        boolstatus = Part.Extension.SelectByRay(xFrontFlatLHS - (xDimLHS / 2), yFrontSecLHS + (0.05 - 0.016) / SScale / 2, -7000, 0, 0, -1, 0.00001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlatLHS + (xDimLHS / 2), yFrontSecLHS, -7000, 0, 0, -1, 0.00001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlatLHS + (xDimLHS / 2) + 0.015, yFrontSecLHS + (0.05 / SScale / 2) + 0.015, 0)
        Part.ClearSelection2(True)

        'Height - LHS Sheet
        boolstatus = Part.Extension.SelectByRay(xFrontLHSSheet, yFrontLHSSheet + (yDimLHSSheet / 2), -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontLHSSheet, yFrontLHSSheet - (yDimLHSSheet / 2), -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontLHSSheet + (zDimLHSSheet / 2) + 0.015, yFrontLHSSheet, 0)
        Part.ClearSelection2(True)

        'Width - LHS Sheet
        boolstatus = Part.Extension.SelectByRay(xFrontLHSSheet - (zDimLHSSheet / 2), yFrontLHSSheet, -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontLHSSheet + (zDimLHSSheet / 2), yFrontLHSSheet, -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontLHSSheet, yFrontLHSSheet + yDimLHSSheet / 2 + 0.015, 0)
        Part.ClearSelection2(True)

        ' Note
        Dim X, Y, Z As Decimal
        boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingBalloonStyle, 0, swBalloonStyle_e.swBS_None)

        'LHS
        X = Round(xDimFlatLHS * SScale * 1000, 2)
        Y = Round(yDimFlatLHS * SScale * 1000, 2)
        Z = Round(zDimFlatLHS * SScale * 1000, 2)
        myNote = Draw.CreateText2(FileNameLHS & vbNewLine & "Qty - " & vbNewLine & X & "mm x " & Y & "mm x " & Z & "mm", xRightSecLHS - (zDimLHS / 2) - 0.015, yFrontFlatLHS - (yDimLHS / 2) - 0.025, 0, 0.004, 0)

        'LHS Sheet
        X = Round(xDimFlatLHSSheet * SScale * 1000, 2)
        Y = Round(yDimFlatLHSSheet * SScale * 1000, 2)
        Z = Round(zDimFlatLHSSheet * SScale * 1000, 2)
        myNote = Draw.CreateText2(FileNameLHSSheet & vbNewLine & "Qty - " & vbNewLine & Z & "mm x " & Y & "mm x " & X & "mm", xFrontLHSSheet - (zDimLHSSheet / 2) - 0.015, yFrontLHSSheet - (yDimLHSSheet / 2) - 0.01, 0, 0.004, 0)

        ' Save As
        Part.ViewZoomtofit2()
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Drawings\Motor Box\" & JobNo & "_02 & 02B.SLDDRW", 0, 2)
        swApp.CloseAllDocuments(True)

        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Drawings\Motor Box\" & JobNo & "_02 & 02B.SLDDRW", 3, 0, "", longstatus, longwarnings)
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Drawings\Motor Box\" & JobNo & "_02 & 02B.SLDDRW", 0, 2)
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Drawings\Motor Box\" & JobNo & "_02 & 02B.PDF", 0, 2)
        swApp.CloseAllDocuments(True)

    End Sub

    Public Sub BottomSupport(Dpth As Decimal)

        Part = swApp.OpenDoc6("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Box - Sample\03_bot support.sldprt", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("03_bot support", False, longstatus)
        Part = swApp.ActiveDoc

        Part.ViewZoomtofit2()

        'Resize Dimentions
        boolstatus = Part.Extension.SelectByID2("ChannelLength@BaseFlange@03_bot support.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 0)
        myDimension = Part.Parameter("ChannelLength@BaseFlange")
        myDimension.SystemValue = Dpth - 0.017
        Part.ClearSelection2(True)
        boolstatus = Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Save File      
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_03_bot support.SLDPRT", 0, 2)
        Part = Nothing
        StdFunc.CloseActiveDoc()

        'BOM Entries
        bomData.EnterValuesInCNC(JobNo & "_03_bot support", (Dpth * 1000) - 16.5, "171.61", "2.0", 1, Client, AHUName, JobNo)

    End Sub

    Public Sub RHSBeam(Ht As Decimal, Dpth As Decimal)

        'Variables
        Ht -= 0.05

        'Open File
        Part = swApp.OpenDoc6("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Box - Sample\04_rhs -l.sldprt", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("04_rhs -l", False, longstatus)
        Part = swApp.ActiveDoc

        Part.ViewZoomtofit2()

        'Resize Dimentions
        boolstatus = Part.Extension.SelectByID2("BoxHeight@BaseFlange@04_rhs -l.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 0)
        myDimension = Part.Parameter("BoxHeight@BaseFlange")
        myDimension.SystemValue = Ht
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("BoxDepth@Edge-Flange1@04_rhs -l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        myDimension = Part.Parameter("BoxDepth@Edge-Flange1")
        myDimension.SystemValue = Dpth - 0.05
        Part.ClearSelection2(True)
        boolstatus = Part.EditRebuild3()

        Part.ViewZoomtofit2()

        swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchInference, False)    'Snaping OFF

        'Holes - Left
        boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        Part.SketchManager.InsertSketch(True)
        Part.ClearSelection2(True)

        skSegment = Part.SketchManager.CreateCircle(0.025, Ht - 0.05 - 0.05, 0, 0.0295, Ht - 0.05 - 0.05, 0)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)
        skSegment = Part.SketchManager.CreateCircle(0.025, (Ht / 2) - 0.05, 0, 0.0295, (Ht / 2) - 0.05, 0)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)

        Part.ViewOrientationUndo()
        Part.SketchManager.InsertSketch(True)

        myFeature = Part.FeatureManager.FeatureCut3(True, False, False, 1, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False)
        Part.SelectionManager.EnableContourSelection = False
        Part.ClearSelection2(True)

        'Holes - Front
        boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        Part.SketchManager.InsertSketch(True)
        Part.ClearSelection2(True)

        skSegment = Part.SketchManager.CreateCircle(-0.01, Ht - 0.05, 0, -0.01, Ht - 0.05 + 0.0045, 0)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)
        skSegment = Part.SketchManager.CreateCircle(-0.01, Ht / 2, 0, -0.01, (Ht / 2) + 0.0045, 0)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)
        skSegment = Part.SketchManager.CreateCircle(-0.01, 0.025, 0, -0.01, 0.025 + 0.0045, 0)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)

        skSegment = Part.SketchManager.CreateCircle(-0.04, Ht - 0.05, 0, -0.04, Ht - 0.05 + 0.0045, 0)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)
        skSegment = Part.SketchManager.CreateCircle(-0.04, Ht / 2, 0, -0.04, (Ht / 2) + 0.0045, 0)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)
        skSegment = Part.SketchManager.CreateCircle(-0.04, 0.025, 0, -0.04, 0.025 + 0.0045, 0)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)

        Part.ViewOrientationUndo()
        Part.SketchManager.InsertSketch(True)

        myFeature = Part.FeatureManager.FeatureCut3(True, False, False, 11, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False)
        Part.SelectionManager.EnableContourSelection = False
        Part.ClearSelection2(True)

        swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchInference, True)    'Snaping ON

        'Save File
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_04_rhs -l.SLDPRT", 0, 2)
        Part = Nothing
        StdFunc.CloseActiveDoc()

        'BOM Entries
        bomData.EnterValuesInCNC(JobNo & "_04_rhs -l", ((Dpth + Ht) * 1000) - 3.6, "108.81", "2.0", 1, Client, AHUName, JobNo)

    End Sub

    Public Sub FrontTopBeam(Wth As Decimal, Ht As Decimal, Dpth As Decimal, HoleCD As Decimal)

        Part = swApp.OpenDoc6("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Box - Sample\05_front top l.sldprt", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("05_front top l", False, longstatus)
        Part = swApp.ActiveDoc

        Part.ViewZoomtofit2()

        'Resize Dimentions
        boolstatus = Part.Extension.SelectByID2("BoxWidth@BaseFlange@05_front top l.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 0)
        myDimension = Part.Parameter("BoxWidth@BaseFlange")
        myDimension.SystemValue = Wth - 0.1 - 0.001
        Part.ClearSelection2(True)
        boolstatus = Part.EditRebuild3()
        Part.ViewZoomtofit2()

        swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchInference, False)    'Snaping OFF

        'Holes
        boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        Part.SketchManager.InsertSketch(True)
        Part.ClearSelection2(True)

        skSegment = Part.SketchManager.CreateCircle(-1 * ((Wth / 2) - 0.075), -0.025, 0#, -1 * ((Wth / 2) - 0.075), -0.0295, 0#)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)
        skSegment = Part.SketchManager.CreateCircle(0, -0.025, 0#, 0, -0.0295, 0#)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)
        skSegment = Part.SketchManager.CreateCircle((Wth / 2) - 0.075, -0.025, 0#, (Wth / 2) - 0.075, -0.0295, 0#)
        Part.SketchAddConstraints("sgFIXED")
        Part.ClearSelection2(True)

        Part.ViewOrientationUndo()
        Part.SketchManager.InsertSketch(True)

        myFeature = Part.FeatureManager.FeatureCut3(True, False, False, 1, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False)
        Part.SelectionManager.EnableContourSelection = False
        Part.ClearSelection2(True)

        swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchInference, True)    'Snaping ON

        'Save File
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_05_front top l.SLDPRT", 0, 2)
        Part = Nothing
        StdFunc.CloseActiveDoc()

        'BOM Entries
        bomData.EnterValuesInCNC(JobNo & "_05_front top l", (Wth * 1000) - 100, "96.40", "2.0", 1, Client, AHUName, JobNo)

    End Sub

    Public Sub BotSupp_RHSBeam_FrontTopBeam_Drawing()

        ' Variables
        Dim FilePathBotSupp As String = "C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_03_bot support.SLDPRT"
        Dim FileNameBotSupp As String = JobNo & "_03_bot support"
        Dim FilePathRHSBeam As String = "C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_04_rhs -l.SLDPRT"
        Dim FileNameRHSBeam As String = JobNo & "_04_rhs -l"
        Dim FilePathFntTopBeam As String = "C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_05_front top l.SLDPRT"
        Dim FileNameFntTopBeam As String = JobNo & "_05_front top l"

        ' Bottom Support
        ' Open File
        Part = swApp.OpenDoc6(FilePathBotSupp, 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2(FileNameBotSupp, False, longstatus)
        Part = swApp.ActiveDoc

        Part.ViewZoomtofit2()

        'Bounding Box
        Dim BBox_BotSupp As Object = StdFunc.BoundingBox()
        Dim xDimBotSupp As Decimal = Abs(BBox_BotSupp(0)) + Abs(BBox_BotSupp(3))
        Dim yDimBotSupp As Decimal = Abs(BBox_BotSupp(1)) + Abs(BBox_BotSupp(4))
        Dim zDimBotSupp As Decimal = Abs(BBox_BotSupp(2)) + Abs(BBox_BotSupp(5))

        'Bounding Box - Flat
        boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditUnsuppress2()

        Dim BBoxFlatBotSupp As Object = StdFunc.BoundingBox()
        Dim xDimFlatBotSupp As Decimal = Abs(BBoxFlatBotSupp(0)) + Abs(BBoxFlatBotSupp(3))
        Dim yDimFlatBotSupp As Decimal = Abs(BBoxFlatBotSupp(1)) + Abs(BBoxFlatBotSupp(4))
        Dim zDimFlatBotSupp As Decimal = Abs(BBoxFlatBotSupp(2)) + Abs(BBoxFlatBotSupp(5))

        boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditSuppress2()

        swApp.CloseDoc(FileNameBotSupp & ".SLDPRT")

        ' RHS Beam
        ' Open File
        Part = swApp.OpenDoc6(FilePathRHSBeam, 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2(FileNameRHSBeam, False, longstatus)
        Part = swApp.ActiveDoc

        Part.ViewZoomtofit2()

        'Bounding Box
        Dim BBox_RHS As Object = StdFunc.BoundingBox()
        Dim xDimRHS As Decimal = Abs(BBox_RHS(0)) + Abs(BBox_RHS(3))
        Dim yDimRHS As Decimal = Abs(BBox_RHS(1)) + Abs(BBox_RHS(4))
        Dim zDimRHS As Decimal = Abs(BBox_RHS(2)) + Abs(BBox_RHS(5))

        'Bounding Box - Flat
        boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditUnsuppress2()

        Dim BBoxFlatRHS As Object = StdFunc.BoundingBox()
        Dim xDimFlatRHS As Decimal = Abs(BBoxFlatRHS(0)) + Abs(BBoxFlatRHS(3))
        Dim yDimFlatRHS As Decimal = Abs(BBoxFlatRHS(1)) + Abs(BBoxFlatRHS(4))
        Dim zDimFlatRHS As Decimal = Abs(BBoxFlatRHS(2)) + Abs(BBoxFlatRHS(5))

        boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditSuppress2()

        swApp.CloseDoc(FileNameRHSBeam & ".SLDPRT")

        ' Front Top Beam
        ' Open File
        Part = swApp.OpenDoc6(FilePathFntTopBeam, 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2(FileNameFntTopBeam, False, longstatus)
        Part = swApp.ActiveDoc

        Part.ViewZoomtofit2()

        'Bounding Box
        Dim BBox_FntTop As Object = StdFunc.BoundingBox()
        Dim xDimFntTop As Decimal = Abs(BBox_FntTop(0)) + Abs(BBox_FntTop(3))
        Dim yDimFntTop As Decimal = Abs(BBox_FntTop(1)) + Abs(BBox_FntTop(4))
        Dim zDimFntTop As Decimal = Abs(BBox_FntTop(2)) + Abs(BBox_FntTop(5))

        'Bounding Box - Flat
        boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditUnsuppress2()

        Dim BBoxFlatFntTop As Object = StdFunc.BoundingBox()
        Dim xDimFlatFntTop As Decimal = Abs(BBoxFlatFntTop(0)) + Abs(BBoxFlatFntTop(3))
        Dim yDimFlatFntTop As Decimal = Abs(BBoxFlatFntTop(1)) + Abs(BBoxFlatFntTop(4))
        Dim zDimFlatFntTop As Decimal = Abs(BBoxFlatFntTop(2)) + Abs(BBoxFlatFntTop(5))

        boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditSuppress2()

        swApp.CloseDoc(FileNameFntTopBeam & ".SLDPRT")

        ' Sheet Scale
        Dim SScaleX As Integer = Ceiling((xDimFlatRHS + zDimRHS + 10 * xDimBotSupp) / (0.297 - (0.03 + 0.04 + 0.06 + 0.03)))
        Dim SScaleY As Integer = Ceiling((0.05 + yDimFlatRHS) / (0.21 - (0.03 + 0.03 + 0.03)))
        Dim SScale As Integer = SScaleX
        If SScaleX < SScaleY Then
            SScale = SScaleY
        End If

        ' Adjust Values for Scale
        xDimBotSupp /= SScale
        yDimBotSupp /= SScale
        zDimBotSupp /= SScale

        xDimFlatBotSupp /= SScale
        yDimFlatBotSupp /= SScale
        zDimFlatBotSupp /= SScale

        xDimRHS /= SScale
        yDimRHS /= SScale
        zDimRHS /= SScale

        xDimFlatRHS /= SScale
        yDimFlatRHS /= SScale
        zDimFlatRHS /= SScale

        xDimFntTop /= SScale
        yDimFntTop /= SScale
        zDimFntTop /= SScale

        xDimFlatFntTop /= SScale
        yDimFlatFntTop /= SScale
        zDimFlatFntTop /= SScale

        ' Get Margins
        Dim marginX As Decimal = (0.297 - (xDimFlatRHS + 0.04 + zDimRHS + 0.06 + 10 * xDimBotSupp)) / 2
        If marginX < 0.03 Then marginX = 0.03
        Dim marginY As Decimal = (0.21 - (yDimFlatRHS + 0.03 + (0.05 / SScale))) / 2
        If marginY < 0.03 Then marginY = 0.03

        ' Calculate View Placements
        Dim xFrontFlatRHS As Decimal = marginX + xDimFlatRHS / 2
        Dim yFrontSecRHS As Decimal = marginY + 0.05 / SScale / 2
        Dim yFrontFlatRHS As Decimal = yFrontSecRHS + 0.05 / SScale / 2 + 0.03 + yDimFlatRHS / 2
        Dim xRightSecRHS As Decimal = xFrontFlatRHS + xDimFlatRHS / 2 + 0.04 + zDimRHS / 2
        Dim xFrontFnrTop As Decimal = xRightSecRHS + zDimRHS / 2 + 0.06 + 10 * xDimBotSupp / 2
        Dim yFrontFnrTop As Decimal = 0.21 - (marginY + 10 * yDimFntTop / 2)
        Dim yFrontBotSup As Decimal = marginY + 10 * yDimBotSupp

        ' Open Parts
        Part = swApp.OpenDoc6(FilePathBotSupp, 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(FilePathRHSBeam, 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6(FilePathFntTopBeam, 1, 0, "", longstatus, longwarnings)

        ' Start Drawing
        Part = swApp.NewDocument(DrawTemp, 2, 0.297, 0.21)
        Draw = Part
        boolstatus = Draw.SetupSheet5("Sheet1", 11, 12, 1, SScale, False, DrawSheet, 0.297, 0.21, "Default", False)

        swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swAutomaticScaling3ViewDrawings, False)

        ' Views - RHS
        'Front Flat
        boolstatus = Draw.CreateFlatPatternViewFromModelView(FilePathRHSBeam, "Default", xFrontFlatRHS, yFrontFlatRHS, 0)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Draw.ActivateView("Drawing View1")
        Part.ClearSelection2(True)

        'Front - Outside
        myView = Draw.CreateDrawViewFromModelView3(FilePathRHSBeam, "*Front", -xFrontFlatRHS, yFrontFlatRHS, 0)
        boolstatus = Draw.ActivateView("Drawing View2")
        Part.ClearSelection2(True)

        'Side - Outside
        myView = Part.CreateDrawViewFromModelView3(FilePathRHSBeam, "*Right", -xFrontFlatRHS, yFrontSecRHS - (yDimRHS / 2) + 0.05 / SScale / 2, 0)
        boolstatus = Draw.ActivateView("Drawing View3")
        Part.ClearSelection2(True)

        'Right - Section
        boolstatus = Draw.ActivateView("Drawing View2")
        skSegment = Part.SketchManager.CreateLine(0, (yDimRHS * SScale / 2) + 0.15, 0, 0, -(yDimRHS * SScale / 2) - 0.15, 0)
        boolstatus = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xRightSecRHS, yFrontFlatRHS, 0, "A", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View4")
        Part.ClearSelection2(True)

        'Front - Section
        boolstatus = Draw.ActivateView("Drawing View3")
        skSegment = Part.SketchManager.CreateLine(0, (yDimRHS * SScale / 2) + 0.1, 0, 0, (yDimRHS * SScale / 2) - 0.15, 0)
        boolstatus = Part.Extension.SelectByID2("Line2", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xFrontFlatRHS, yFrontSecRHS, 0, "B", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View5")
        Part.ClearSelection2(True)

        ' Views - Bottom Support & Front Top Beam
        'Right - Front Top Beam
        myView = Draw.CreateDrawViewFromModelView3(FilePathFntTopBeam, "*Right", xFrontFnrTop, yFrontFnrTop, 0)
        boolstatus = Draw.ActivateView("Drawing View6")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View6", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        myView = Part.SelectionManager.GetSelectedObject6(1, -1)
        myView.UseParentScale = False
        myView.ScaleDecimal *= 10

        'Front - Bottom Support
        myView = Draw.CreateDrawViewFromModelView3(FilePathBotSupp, "*Front", xFrontFnrTop, yFrontBotSup, 0)
        boolstatus = Draw.ActivateView("Drawing View7")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View7", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        myView = Part.SelectionManager.GetSelectedObject6(1, -1)
        myView.UseParentScale = False
        myView.ScaleDecimal *= 10

        boolstatus = Part.ActivateSheet("Sheet1")

        ' Dimentions
        'Height - RHS Flat
        boolstatus = Part.Extension.SelectByRay(xFrontFlatRHS, yFrontFlatRHS + (yDimFlatRHS / 2), -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlatRHS, yFrontFlatRHS - (yDimFlatRHS / 2), -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlatRHS + (xDimFlatRHS / 2) + 0.015, yFrontFlatRHS, 0)
        Part.ClearSelection2(True)

        'Width - RHS Flat
        boolstatus = Part.Extension.SelectByRay(xFrontFlatRHS + (xDimFlatRHS / 2), yFrontFlatRHS, -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlatRHS - (xDimFlatRHS / 2), yFrontFlatRHS, -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlatRHS + (xDimFlatRHS / 2) + 0.015, yFrontFlatRHS + (yDimFlatRHS / 2) + 0.015, 0)
        Part.ClearSelection2(True)

        'Height - RHS Right Section
        boolstatus = Part.Extension.SelectByRay(xRightSecRHS, yFrontFlatRHS + (yDimRHS / 2), -7000, 0, 0, -1, 0.00001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xRightSecRHS - (zDimRHS / 2) + (0.001 / SScale), yFrontFlatRHS - (yDimRHS / 2), -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xRightSecRHS - (zDimRHS / 2) - 0.015, yFrontFlatRHS, 0)
        Part.ClearSelection2(True)

        'Width - RHS Right Section
        boolstatus = Part.Extension.SelectByRay(xRightSecRHS + (zDimRHS / 2), yFrontFlatRHS + (yDimRHS / 2) - (0.001 / SScale), -7000, 0, 0, -1, 0.00001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xRightSecRHS - (zDimRHS / 2), yFrontFlatRHS, -7000, 0, 0, -1, 0.00001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xRightSecRHS, yFrontFlatRHS + (yDimRHS / 2) + 0.015, 0)
        Part.ClearSelection2(True)

        'Width - RHS Front Section
        boolstatus = Part.Extension.SelectByRay(xFrontFlatRHS - (xDimRHS / 2), yFrontSecRHS, -7000, 0, 0, -1, 0.00001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlatRHS + (xDimRHS / 2), yFrontSecRHS + (0.05 - 0.016) / 2 / SScale, -7000, 0, 0, -1, 0.00001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlatRHS + (xDimRHS / 2) + 0.015, yFrontSecRHS + (0.05 / SScale / 2) + 0.015, 0)
        Part.ClearSelection2(True)

        'Hight - RHS Front Section
        boolstatus = Part.Extension.SelectByRay(xFrontFlatRHS, yFrontSecRHS + 0.05 / 2 / SScale, -7000, 0, 0, -1, 0.00001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlatRHS + (xDimRHS / 2) - (0.001 / SScale), yFrontSecRHS + 0.05 / 2 / SScale - 0.016 / SScale, -7000, 0, 0, -1, 0.00001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlatRHS + (xDimRHS / 2) + 0.015, yFrontSecRHS - (0.05 / SScale / 2) - 0.015, 0)
        Part.ClearSelection2(True)

        'Height - Front Top Beam
        boolstatus = Part.Extension.SelectByRay(xFrontFnrTop, yFrontFnrTop + (10 * yDimFntTop / 2), -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFnrTop + (10 * zDimFntTop / 2) - (0.01 / SScale), yFrontFnrTop - (10 * yDimFntTop / 2), -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFnrTop + (10 * zDimFntTop / 2) + 0.015, yFrontFnrTop, 0)
        Part.ClearSelection2(True)

        'Width - Front Top Beam
        boolstatus = Part.Extension.SelectByRay(xFrontFnrTop - (10 * zDimFntTop / 2), yFrontFnrTop + (10 * yDimFntTop / 2) - (0.01 / SScale), -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFnrTop + (10 * zDimFntTop / 2), yFrontFnrTop, -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFnrTop, yFrontFnrTop + (10 * yDimBotSupp / 2) + 0.015, 0)
        Part.ClearSelection2(True)

        'Height - Bottom Support
        boolstatus = Part.Extension.SelectByRay(xFrontFnrTop, yFrontBotSup + (10 * yDimBotSupp / 2), -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFnrTop + (10 * xDimBotSupp / 2) - (0.05 / SScale), yFrontBotSup - (10 * yDimBotSupp / 2), -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFnrTop + (10 * xDimBotSupp / 2) + 0.015, yFrontBotSup, 0)
        Part.ClearSelection2(True)

        'Width - Bottom Support
        boolstatus = Part.Extension.SelectByRay(xFrontFnrTop - (10 * xDimBotSupp / 2), yFrontBotSup, -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFnrTop + (10 * xDimBotSupp / 2), yFrontBotSup, -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFnrTop, yFrontBotSup + (10 * yDimBotSupp / 2) + 0.015, 0)
        Part.ClearSelection2(True)

        'Width - Bottom Support - Lower Bend
        boolstatus = Part.Extension.SelectByRay(xFrontFnrTop + (10 * xDimBotSupp / 2), yFrontBotSup, -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFnrTop + (10 * xDimBotSupp / 2) - (0.1 / SScale), yFrontBotSup - (10 * yDimBotSupp / 2) + (0.01 / SScale), -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFnrTop + (10 * xDimBotSupp / 2) + 0.015, yFrontBotSup - (10 * yDimBotSupp / 2) - 0.015, 0)
        Part.ClearSelection2(True)

        ' Note
        Dim X, Y, Z As Decimal
        boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingBalloonStyle, 0, swBalloonStyle_e.swBS_None)

        'RHS Beam
        X = Round(xDimFlatRHS * SScale * 1000, 2)
        Y = Round(yDimFlatRHS * SScale * 1000, 2)
        Z = Round(zDimFlatRHS * SScale * 1000, 2)
        myNote = Draw.CreateText2(FileNameRHSBeam & vbNewLine & "Qty - " & vbNewLine & X & "mm x " & Y & "mm x " & Z & "mm", xRightSecRHS - (zDimRHS / 2) - 0.015, yFrontSecRHS + (0.05 / SScale / 2) + 0.015, 0, 0.004, 0)

        'Front Top Beam
        X = Round(xDimFlatFntTop * SScale * 1000, 2)
        Y = Round(yDimFlatFntTop * SScale * 1000, 2)
        Z = Round(zDimFlatFntTop * SScale * 1000, 2)
        myNote = Draw.CreateText2(FileNameRHSBeam & vbNewLine & "Qty - " & vbNewLine & X & "mm x " & Y & "mm x " & Z & "mm", xFrontFnrTop - (zDimBotSupp / 2) - 0.015, yFrontFnrTop - (yDimFntTop / 2) - 0.025, 0, 0.004, 0)

        'Bottom Support
        X = Round(xDimFlatBotSupp * SScale * 1000, 2)
        Y = Round(zDimFlatBotSupp * SScale * 1000, 2)
        Z = Round(yDimFlatBotSupp * SScale * 1000, 2)
        myNote = Draw.CreateText2(FileNameRHSBeam & vbNewLine & "Qty - " & vbNewLine & X & "mm x " & Y & "mm x " & Z & "mm", xFrontFnrTop - (zDimBotSupp / 2) - 0.015, yFrontBotSup - (10 * yDimBotSupp / 2) - 0.025, 0, 0.004, 0)

        ' Save As
        Part.ViewZoomtofit2()
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Drawings\Motor Box\" & JobNo & "_03 to 05.SLDDRW", 0, 2)
        swApp.CloseAllDocuments(True)

        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Drawings\Motor Box\" & JobNo & "_03 to 05.SLDDRW", 3, 0, "", longstatus, longwarnings)
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Drawings\Motor Box\" & JobNo & "_03 to 05.SLDDRW", 0, 2)
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Drawings\Motor Box\" & JobNo & "_03 to 05.PDF", 0, 2)
        swApp.CloseAllDocuments(True)

    End Sub

    Public Sub FanVerticalStand(HoleCD As Decimal, FanBoltHole As Decimal)

        'Variables
        Dim StandHt As Decimal = HoleCD - 0.05 - FanBoltHole

        'Open File
        Part = swApp.OpenDoc6("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Box - Sample\05_lrhs base.SLDPRT", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("05_lrhs base", False, longstatus)
        Part = swApp.ActiveDoc

        Part.ViewZoomtofit2()

        'Re-Dim
        boolstatus = Part.Extension.SelectByID2("Height@BaseFlange@05_lrhs base.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 0)
        myDimension = Part.Parameter("Height@BaseFlange")
        myDimension.SystemValue = StandHt
        Part.ClearSelection2(True)
        boolstatus = Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Save
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor\" & ArticleNoFan & "_05_lrhs base.SLDPRT", 0, 2)
        Part = Nothing
        StdFunc.CloseActiveDoc()

    End Sub

    Public Sub FanHorizontalStand()

        'Variables
        Dim BottomSupportDis As Decimal = StdFunc.GetFromTable("Stnd_Dist", "article_no_table", "article_no", ArticleNoFan) / 1000

        'Open File
        Part = swApp.OpenDoc6("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Box - Sample\02_base support.SLDPRT", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("02_base support", False, longstatus)
        Part = swApp.ActiveDoc

        'Re-Dim
        boolstatus = Part.Extension.SelectByID2("Length@BaseFlange@02_base support.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 0)
        myDimension = Part.Parameter("Length@BaseFlange")
        myDimension.SystemValue = BottomSupportDis
        Part.ClearSelection2(True)
        boolstatus = Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Save
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor\" & ArticleNoFan & "_02_base support.SLDPRT", 0, 2)
        Part = Nothing
        StdFunc.CloseActiveDoc()

    End Sub

    Public Sub BoxAssembly(Wth As Decimal, Ht As Decimal, Dpth As Decimal, HoleCD As Decimal)

        'Variables
        Dim BottomSupportDis As Decimal = (StdFunc.GetFromTable("Stnd_Dist", "article_no_table", "article_no", ArticleNoFan) / 1000) + 0.07

        'Open Part Files
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_01_back bot sheet.sldprt", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_02_lhs-l.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_02B_lhs sheet.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_03_bot support.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_04_rhs -l.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_05_front top l.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor\" & ArticleNoFan & "_02_base support.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor\" & ArticleNoFan & "_05_lrhs base.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Box - MOTOR\" & ArticleNoFan & "\" & ArticleNoFan & ".sldasm", 2, 0, "", longstatus, longwarnings)

        'New Assembly Document
        Dim assyTemp As String = swApp.GetUserPreferenceStringValue(9)
        Part = swApp.NewDocument(assyTemp, 0, 0, 0)
        swApp.ActivateDoc2("Assem1", False, longstatus)
        Part = swApp.ActiveDoc
        Assy = Part

        'Incert Parts
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_01_back bot sheet.SLDPRT", 0, 0, Dpth / 2)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_02_lhs-l.SLDPRT", -1, 0.5, 0.5)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_02B_lhs sheet.SLDPRT", -1, 0.5, 0.5)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_03_bot support.SLDPRT", -0.5, 0.1, 1)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_03_bot support.SLDPRT", 0.5, 0.1, 1)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_04_rhs -l.SLDPRT", 1, 0.5, 0.5)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_05_front top l.SLDPRT", 0.2, 0.7, 0.7)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor\" & ArticleNoFan & "_02_base support.SLDPRT", 1, -0.5, 0.5)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor\" & ArticleNoFan & "_02_base support.SLDPRT", 1.1, -0.5, 0.8)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor\" & ArticleNoFan & "_05_lrhs base.SLDPRT", 0.8, 0.7, 0.4)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor\" & ArticleNoFan & "_05_lrhs base.SLDPRT", 0.8, 0.7, 0.8)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor\" & ArticleNoFan & "_05_lrhs base.SLDPRT", -0.8, 0.7, 0.4)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor\" & ArticleNoFan & "_05_lrhs base.SLDPRT", -0.8, 0.7, 0.8)
        boolstatus = Assy.AddComponent("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Box - MOTOR\" & ArticleNoFan & "\" & ArticleNoFan & ".sldasm", 0, 0.1, 1.5)

        'Close Part Files
        swApp.CloseAllDocuments(False)

        'Save Assembly Document
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_Box Assembly.SLDASM", 0, 2)
        Part = Nothing
        StdFunc.CloseActiveDoc()

        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_Box Assembly.SLDASM", 2, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2(JobNo & "_Box Assembly", False, longstatus)
        Part = swApp.ActiveDoc
        Assy = Part

        Part.ViewZoomtofit2()

        'Mates
        'LHS
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_01_back bot sheet-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_02_lhs-l-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 0, False, Dpth, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_01_back bot sheet-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_02_lhs-l-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 0, False, 0.05, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_01_back bot sheet-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_02_lhs-l-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 0, True, Wth / 2 - 0.025, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        Part.ViewZoomtofit2()

        'LHS Sheet
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_01_back bot sheet-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_02B_lhs sheet-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 0, False, Dpth, Dpth, Dpth, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_01_back bot sheet-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_02B_lhs sheet-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 0, False, 0.0015, 0.0015, 0.0015, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_01_back bot sheet-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_02B_lhs sheet-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 0, True, Wth / 2, Wth / 2, Wth / 2, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        Part.ViewZoomtofit2()

        'Bottom Support - Left
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_01_back bot sheet-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_03_bot support-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 0, False, 0.0015, 0.0015, 0.0015, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_01_back bot sheet-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_03_bot support-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 0, True, BottomSupportDis / 2, BottomSupportDis / 2, BottomSupportDis / 2, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_01_back bot sheet-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_03_bot support-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 0, False, 0.05, 0.05, 0.05, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        Part.ViewZoomtofit2()

        'Bottom Support - Right
        Part.ViewZoomtofit2()
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_03_bot support-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_03_bot support-2@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0.455, 0.001, 0.001, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_03_bot support-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_03_bot support-2@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0.1, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_03_bot support-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_03_bot support-2@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 0, False, BottomSupportDis, BottomSupportDis, BottomSupportDis, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        Part.ViewZoomtofit2()

        'RHS Beam
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_01_back bot sheet-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_04_rhs -l-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 0, False, Dpth, Dpth, Dpth, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_01_back bot sheet-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_04_rhs -l-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 0, False, Wth / 2, Wth / 2, Wth / 2, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_01_back bot sheet-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_04_rhs -l-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 0, False, 0.05, 0.05, 0.05, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        Part.ViewZoomtofit2()

        'Front Top Support
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_01_back bot sheet-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_05_front top l-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 1, False, Dpth, Dpth, Dpth, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_01_back bot sheet-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_05_front top l-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 0, False, Ht, Ht, Ht, 0.001, 0.001, 0.39309481042652, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_01_back bot sheet-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_05_front top l-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 1, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        'Fan Motor
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_01_back bot sheet-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Plane3@" & ArticleNoFan & "-1@" & JobNo & "_Box Assembly/" & ArticleNoFan & "_motor-1@" & ArticleNoFan, "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 1, True, 0.002, 0.002, 0.002, 0.001, 0.001, 1.5707963267949, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_01_back bot sheet-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & ArticleNoFan & "-1@" & JobNo & "_Box Assembly/" & ArticleNoFan & "_motor-1@" & ArticleNoFan, "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 1, False, HoleCD, HoleCD, HoleCD, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Axis1@" & ArticleNoFan & "-1@" & JobNo & "_Box Assembly/" & ArticleNoFan & "_motor-1@" & ArticleNoFan, "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_01_back bot sheet-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, -1, False, 0.2325, 0, 0, 0.001, 0.001, 0.5235987755983, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        Part.ViewZoomtofit2()

        'FAN STAND
        'Vertical Stand 1
        boolstatus = Part.Extension.SelectByID2("Axis1@" & ArticleNoFan & "_05_lrhs base-1@" & JobNo & "_Box Assembly", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Axis1@" & ArticleNoFan & "-1@" & JobNo & "_Box Assembly", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 1, False, 0.571899685259571, 0.001, 0.001, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_03_bot support-2@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & ArticleNoFan & "_05_lrhs base-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_03_bot support-2@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & ArticleNoFan & "_05_lrhs base-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(3, 0, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        Part.ViewZoomtofit2()

        'Vertical Stand 2
        boolstatus = Part.Extension.SelectByID2("Axis1@" & ArticleNoFan & "_05_lrhs base-2@" & JobNo & "_Box Assembly", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Axis2@" & ArticleNoFan & "-1@" & JobNo & "_Box Assembly", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 1, False, 0.571899685259571, 0.001, 0.001, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_03_bot support-2@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & ArticleNoFan & "_05_lrhs base-2@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_03_bot support-2@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & ArticleNoFan & "_05_lrhs base-2@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(3, 0, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        Part.ViewZoomtofit2()

        'Vertical Stand 3
        boolstatus = Part.Extension.SelectByID2("Axis1@" & ArticleNoFan & "_05_lrhs base-3@" & JobNo & "_Box Assembly", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Axis3@" & ArticleNoFan & "-1@" & JobNo & "_Box Assembly", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 1, False, 0.571899685259571, 0.001, 0.001, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_03_bot support-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & ArticleNoFan & "_05_lrhs base-3@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_03_bot support-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & ArticleNoFan & "_05_lrhs base-3@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(3, 1, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        Part.ViewZoomtofit2()

        'Vertical Stand 4
        boolstatus = Part.Extension.SelectByID2("Axis1@" & ArticleNoFan & "_05_lrhs base-4@" & JobNo & "_Box Assembly", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Axis4@" & ArticleNoFan & "-1@" & JobNo & "_Box Assembly", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 1, False, 0.571899685259571, 0.001, 0.001, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_03_bot support-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & ArticleNoFan & "_05_lrhs base-4@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_03_bot support-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & ArticleNoFan & "_05_lrhs base-4@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(3, 1, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        Part.ViewZoomtofit2()

        'Horisontal Stand 1
        boolstatus = Part.Extension.SelectByID2("Axis1@" & ArticleNoFan & "_02_base support-1@" & JobNo & "_Box Assembly", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Axis2@" & ArticleNoFan & "_05_lrhs base-1@" & JobNo & "_Box Assembly", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 1, False, 0.0565, 0.001, 0.001, 0.001, 0.001, 1.5707963267949, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & ArticleNoFan & "_02_base support-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & ArticleNoFan & "_05_lrhs base-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Front Plane@" & ArticleNoFan & "_02_base support-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_01_back bot sheet-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0.68, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        Part.ViewZoomtofit2()

        'Horisontal Stand 2
        boolstatus = Part.Extension.SelectByID2("Axis1@" & ArticleNoFan & "_02_base support-2@" & JobNo & "_Box Assembly", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Axis2@" & ArticleNoFan & "_05_lrhs base-2@" & JobNo & "_Box Assembly", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 1, False, 0.0565, 0.001, 0.001, 0.001, 0.001, 1.5707963267949, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & ArticleNoFan & "_02_base support-2@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & ArticleNoFan & "_05_lrhs base-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Front Plane@" & ArticleNoFan & "_02_base support-2@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_01_back bot sheet-1@" & JobNo & "_Box Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0.68, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        Part.ViewZoomtofit2()

        'Save Assembly Document
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_Box Assembly.SLDASM", 0, 2)
        Part = Nothing
        StdFunc.CloseActiveDoc()

    End Sub

    Public Sub BoxAssyDrawing()
        Exit Sub
        ' Variables
        Dim FilePath As String = "C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_Box Assembly.SLDASM"
        Dim FileName As String = JobNo & "_Box Assembly"

        'Open File
        Part = swApp.OpenDoc6(FilePath, 2, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2(FileName, False, longstatus)
        Part = swApp.ActiveDoc

        Part.ViewZoomtofit2()

        'Bounding Box
        Dim BBox As Object = StdFunc.BoundingBoxOfAssembly()
        Dim xDim As Decimal = Abs(BBox(0)) + Abs(BBox(3))
        Dim yDim As Decimal = Abs(BBox(1)) + Abs(BBox(4))
        Dim zDim As Decimal = Abs(BBox(2)) + Abs(BBox(5))

        swApp.CloseAllDocuments(True)

        ' Sheet Scale
        Dim SScaleX As Integer = Ceiling((xDim + zDim + xDim) / (0.297 - (0.03 + 0.04 + 0.02 + 0.03)))
        Dim SScaleY As Integer = Ceiling((yDim + zDim) / (0.21 - (0.03 + 0.04 + 0.03)))
        Dim SScale As Integer = SScaleX
        If SScaleX < SScaleY Then
            SScale = SScaleY
        End If

        ' Adjust Values for Scale
        xDim /= SScale
        yDim /= SScale
        zDim /= SScale

        ' Get Margins
        Dim marginX As Decimal = (0.297 - (xDim + 0.04 + zDim + 0.02 + xDim)) / 2
        If marginX < 0.03 Then marginX = 0.03

        Dim marginY As Decimal = (0.21 - (yDim + 0.04 + zDim)) / 2
        If marginY < 0.03 Then marginY = 0.03

        ' Calculate View Placements
        Dim xFrontSec As Decimal = marginX + xDim / 2
        Dim yFrontSec As Decimal = marginY + yDim / 2
        Dim yTopSec As Decimal = yFrontSec + yDim / 2 + 0.04 + zDim / 2
        Dim xRightSec As Decimal = xFrontSec + xDim / 2 + 0.04 + zDim / 2
        Dim xIso As Decimal = xRightSec + zDim / 2 + 0.02 + xDim / 2

        ' Open Parts
        Part = swApp.OpenDoc6(FilePath, 1, 0, "", longstatus, longwarnings)

        ' Start Drawing
        Part = swApp.NewDocument(DrawTemp, 2, 0.297, 0.21)
        Draw = Part
        boolstatus = Draw.SetupSheet5("Sheet1", 11, 12, 1, SScale, False, DrawSheet, 0.297, 0.21, "Default", False)

        swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swAutomaticScaling3ViewDrawings, False)

        ' Views
        'Front - Outside
        myView = Draw.CreateDrawViewFromModelView3(FilePath, "*Front", -xFrontSec, yFrontSec, 0)
        boolstatus = Draw.ActivateView("Drawing View1")
        Part.ClearSelection2(True)

        'Right - Outside
        myView = Draw.CreateDrawViewFromModelView3(FilePath, "*Right", -xRightSec, yFrontSec, 0)
        boolstatus = Draw.ActivateView("Drawing View2")
        Part.ClearSelection2(True)

        'Front - Section
        boolstatus = Draw.ActivateView("Drawing View2")
        skSegment = Part.SketchManager.CreateLine(0, (yDim * SScale / 2) + 0.15, 0, 0, -(yDim * SScale / 2) - 0.15, 0)
        boolstatus = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xFrontSec, yFrontSec, 0, "A", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View3")
        Part.ClearSelection2(True)

        'Right - Section
        boolstatus = Draw.ActivateView("Drawing View1")
        skSegment = Part.SketchManager.CreateLine(0, (yDim * SScale / 2) + 0.1, 0, 0, -(yDim * SScale / 2), 0)
        boolstatus = Part.Extension.SelectByID2("Line2", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xRightSec, yFrontSec, 0, "B", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View4")
        Part.ClearSelection2(True)

        'Top - Section
        boolstatus = Draw.ActivateView("Drawing View1")
        skSegment = Part.SketchManager.CreateLine((xDim * SScale / 2) + 0.1, 0, 0, -(xDim * SScale / 2) - 0.1, 0, 0)
        boolstatus = Part.Extension.SelectByID2("Line3", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        excludedComponents = vbEmpty
        myView = Draw.CreateSectionViewAt5(xFrontSec, yTopSec, 0, "C", swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
        boolstatus = Draw.ActivateView("Drawing View5")

        boolstatus = Part.Extension.SelectByID2("Drawing View5", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = Part.Extension.SelectByID2("Drawing View3", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        BaseView = Part.SelectionManager.GetSelectedObject5(1)
        boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewVerticalOrigin, BaseView)
        Part.ClearSelection2(True)

        'Isometric
        myView = Draw.CreateDrawViewFromModelView3(FilePath, "*Isometric", xIso, yTopSec, 0)
        boolstatus = Draw.ActivateView("Drawing View6")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View6", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Draw.SelectionManager.GetSelectedObject6(1, -1)
        boolstatus = DrawView.SetDisplayMode4(False, swDisplayMode_e.swSHADED, True, True, False)

        boolstatus = Part.ActivateSheet("Sheet1")

        ' Dimentions
        'Height
        boolstatus = Part.Extension.SelectByRay(xFrontSec + (xDim / 2) - (0.025 / SScale), yFrontSec + (yDim / 2), -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontSec, yFrontSec - (yDim / 2), -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontSec + (xDim / 2) + 0.015, yFrontSec, 0)
        Part.ClearSelection2(True)

        'Width
        boolstatus = Part.Extension.SelectByRay(xFrontSec + (xDim / 2), yFrontSec, -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontSec - (xDim / 2), yFrontSec + (yDim / 2) - (0.025 / SScale), -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontSec, yFrontSec + (yDim / 2) + 0.015, 0)
        Part.ClearSelection2(True)

        'Depth
        boolstatus = Part.Extension.SelectByRay(xRightSec + (zDim / 2), yFrontSec + (yDim / 2) - (0.025 / SScale), -7000, 0, 0, -1, 0.0002, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xRightSec - (zDim / 2), yFrontSec + (yDim / 2) - (0.025 / SScale), -7000, 0, 0, -1, 0.0002, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xRightSec, yFrontSec + (yDim / 2) + 0.015, 0)
        Part.ClearSelection2(True)

        ' Note
        boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingBalloonStyle, 0, swBalloonStyle_e.swBS_None)

        myNote = Draw.CreateText2(FileName & vbNewLine & "Qty - ", xIso - (xDim / 2), yFrontSec, 0, 0.004, 0)

        ' Save As
        Part.ViewZoomtofit2()
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Drawings\Motor Box\" & JobNo & "_Motor Box.SLDDRW", 0, 2)
        swApp.CloseAllDocuments(True)

        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Drawings\Motor Box\" & JobNo & "_Motor Box.SLDDRW", 3, 0, "", longstatus, longwarnings)
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Drawings\Motor Box\" & JobNo & "_Motor Box.SLDDRW", 0, 2)
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Drawings\Motor Box\" & JobNo & "_Motor Box.PDF", 0, 2)
        swApp.CloseAllDocuments(True)

    End Sub

#End Region

#Region "Frame"

    Public Sub FrameBotSec(SectionWidth As Decimal, SectionHeight As Decimal, SectionWthSet As Integer(), WallWth As Integer, BoxWth As Integer, FanWthNo As Integer, RHSSecWth As Integer, DoorWth As Integer, Stamp As Integer, lastpart As Integer)

        'Variables
        Dim PartNo As Integer = Convert.ToInt16(Stamp)
        Dim SideBlankWth As Integer = (WallWth - DoorWth - (FanWthNo * BoxWth))

        Dim TempSecWth As Integer = 0
        If PartNo > 1 Then
            For i = 0 To PartNo - 2
                TempSecWth = TempSecWth + (SectionWthSet(i))
            Next
        End If

        Dim TempBoxWth As Integer = SideBlankWth + 50
        Dim NoOfBox As Integer = 0
        If PartNo > 1 Then
            While TempBoxWth < TempSecWth
                NoOfBox += 1
                TempBoxWth = TempBoxWth + BoxWth
            End While
        End If

        Dim TopHoleCenter As Decimal
        If PartNo = 1 Then
            TopHoleCenter = (RHSSecWth - 50) / 1000
        Else
            TopHoleCenter = (TempBoxWth - TempSecWth) / 1000
        End If

        If PartNo = 1 And SideBlankWth >= 200 Then
            BoxWth = SideBlankWth
        End If

        'Open Doc
        Part = swApp.OpenDoc6("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Frame - Sample\_04_bot-l.SLDPRT", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("_04_bot-l", False, longstatus)
        Part = swApp.ActiveDoc

        'Re-Dim Part
        boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_04_bot-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        myDimension = Part.Parameter("Width@BaseFlange")
        myDimension.SystemValue = SectionWidth - 0.001
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Height@BaseFlange@_04_bot-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        myDimension = Part.Parameter("Height@BaseFlange")
        myDimension.SystemValue = SectionHeight
        Part.ClearSelection2(True)

        SideTOPBOTHoles(SectionWidth - 0.001)

        boolstatus = Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Interbolting Holes Top - 1
        boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        Part.SketchManager.InsertSketch(True)
        Part.ClearSelection2(True)

        If PartNo = 1 Then
            skSegment = Part.SketchManager.CreateCircle(-0.025, -0.025, 0, -0.025, -0.025 + 0.0045, 0)
            Part.SketchAddConstraints("sgFIXED")
            Part.ClearSelection2(True)

            skSegment = Part.SketchManager.CreateCircle(-1 * ((RHSSecWth / 1000) - 0.025), -0.025, 0, -1 * ((RHSSecWth / 1000) - 0.025), -0.025 + 0.0045, 0)
            Part.SketchAddConstraints("sgFIXED")
            Part.ClearSelection2(True)
        Else
            If (TopHoleCenter + 0.025) <= SectionWidth Then
                skSegment = Part.SketchManager.CreateCircle(-1 * (TopHoleCenter - 0.025), -0.025, 0, -1 * (TopHoleCenter - 0.025), -0.025 + 0.0045, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)
            End If

            If (TopHoleCenter - 0.025) > 0.01 Then
                skSegment = Part.SketchManager.CreateCircle(-1 * (TopHoleCenter + 0.025), -0.025, 0, -1 * (TopHoleCenter + 0.025), -0.025 + 0.0045, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)
            End If
        End If

        myFeature = Part.FeatureManager.FeatureCut4(True, False, True, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
        Part.SelectionManager.EnableContourSelection = False
        Part.ClearSelection2(True)

        Part.ViewOrientationUndo()

        'Interbolting Holes Top - 2
        If TopHoleCenter + (BoxWth / 1000) <= SectionWidth Then
            boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            Part.SketchManager.InsertSketch(True)
            Part.ClearSelection2(True)

            If (TopHoleCenter + (BoxWth / 1000) - 0.025) <= SectionWidth Then
                skSegment = Part.SketchManager.CreateCircle(-1 * (TopHoleCenter + (BoxWth / 1000) - 0.025), -0.025, 0, -1 * (TopHoleCenter + (BoxWth / 1000) - 0.025), -0.025 + 0.0045, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)
            End If

            If (TopHoleCenter + (BoxWth / 1000) + 0.025) <= SectionWidth Then
                skSegment = Part.SketchManager.CreateCircle(-1 * (TopHoleCenter + (BoxWth / 1000) + 0.025), -0.025, 0, -1 * (TopHoleCenter + (BoxWth / 1000) + 0.025), -0.025 + 0.0045, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)
            End If

            myFeature = Part.FeatureManager.FeatureCut4(True, False, True, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
            Part.SelectionManager.EnableContourSelection = False
            Part.ClearSelection2(True)

            Part.ViewOrientationUndo()
        End If

        'Interbolting Holes Back - 1
        boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        Part.SketchManager.InsertSketch(True)
        Part.ClearSelection2(True)

        If PartNo > 1 Then
            If (TopHoleCenter - 0.1) > 0.01 Then
                skSegment = Part.SketchManager.CreateCircle(-1 * TopHoleCenter + 0.1, -0.025, 0, -1 * TopHoleCenter + 0.1, -0.025 + 0.0045, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)
            End If
        End If

        If (TopHoleCenter + 0.1) <= SectionWidth Then
            skSegment = Part.SketchManager.CreateCircle(-1 * TopHoleCenter - 0.1, -0.025, 0, -1 * TopHoleCenter - 0.1, -0.025 + 0.0045, 0)
            Part.SketchAddConstraints("sgFIXED")
            Part.ClearSelection2(True)
        End If

        myFeature = Part.FeatureManager.FeatureCut4(True, False, True, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
        Part.SelectionManager.EnableContourSelection = False
        Part.ClearSelection2(True)

        Part.ViewOrientationUndo()

        If Stamp = 1 Then
            boolstatus = Part.Extension.SelectByID2("ExtRight@Right@_04_bot-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("ExtRight@Right")
            myDimension.SystemValue = 0.1
            boolstatus = Part.EditRebuild3()
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("DiaRight@CutRight@_04_bot-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("DiaRight@CutRight")
            myDimension.SystemValue = 0.0042
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("DistRight@CutRight@_04_bot-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("DistRight@CutRight")
            myDimension.SystemValue = 0.075
            Part.ClearSelection2(True)

            boolstatus = Part.EditRebuild3()
        End If

        If Stamp = lastpart Then
            boolstatus = Part.Extension.SelectByID2("ExtLeft@Left@_04_bot-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("ExtLeft@Left")
            myDimension.SystemValue = 0.1
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("DiaLeft@CutLeft@_05_top-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("DiaLeft@CutLeft")
            myDimension.SystemValue = 0.0042
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("DistLeft@CutLeft@_05_top-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("DistLeft@CutLeft")
            myDimension.SystemValue = 0.075
            Part.ClearSelection2(True)

            boolstatus = Part.EditRebuild3()
        End If

        'Interbolting Holes Back - 2
        If TopHoleCenter + (BoxWth / 1000) <= SectionWidth Then
            boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            Part.SketchManager.InsertSketch(True)
            Part.ClearSelection2(True)

            If (TopHoleCenter + (BoxWth / 1000) - 0.1) <= SectionWidth Then
                skSegment = Part.SketchManager.CreateCircle(-1 * TopHoleCenter - (BoxWth / 1000) + 0.1, -0.025, 0, -1 * TopHoleCenter - (BoxWth / 1000) + 0.1, -0.025 + 0.0045, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)
            End If

            If PartNo < SectionWthSet.Length Then
                If (TopHoleCenter + (BoxWth / 1000) + 0.1) <= SectionWidth Then
                    skSegment = Part.SketchManager.CreateCircle(-1 * TopHoleCenter - (BoxWth / 1000) - 0.1, -0.025, 0, -1 * TopHoleCenter - (BoxWth / 1000) - 0.1, -0.025 + 0.0045, 0)
                    Part.SketchAddConstraints("sgFIXED")
                    Part.ClearSelection2(True)
                End If
            End If

            myFeature = Part.FeatureManager.FeatureCut4(True, False, True, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
            Part.SelectionManager.EnableContourSelection = False
            Part.ClearSelection2(True)

            Part.ViewOrientationUndo()
        End If

        'Last Holes for door Type
        If PartNo = SectionWthSet.Length And DoorWth > 0 Then
            'Top Holes
            boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            Part.SketchManager.InsertSketch(True)
            Part.ClearSelection2(True)

            skSegment = Part.SketchManager.CreateCircle(-1 * (SectionWidth - 0.075), -0.025, 0, -1 * (SectionWidth - 0.075), -0.025 + 0.0045, 0)
            Part.SketchAddConstraints("sgFIXED")
            Part.ClearSelection2(True)

            skSegment = Part.SketchManager.CreateCircle(-1 * (SectionWidth - 0.025), -0.025, 0, -1 * (SectionWidth - 0.025), -0.025 + 0.0045, 0)
            Part.SketchAddConstraints("sgFIXED")
            Part.ClearSelection2(True)

            myFeature = Part.FeatureManager.FeatureCut4(True, False, True, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
            Part.SelectionManager.EnableContourSelection = False
            Part.ClearSelection2(True)

            Part.ViewOrientationUndo()

            'Back Hole
            boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            Part.SketchManager.InsertSketch(True)
            Part.ClearSelection2(True)

            skSegment = Part.SketchManager.CreateCircle(-1 * (TopHoleCenter - 0.1), -0.025, 0, -1 * (TopHoleCenter - 0.1), -0.025 + 0.0045, 0)
            Part.SketchAddConstraints("sgFIXED")
            Part.ClearSelection2(True)

            myFeature = Part.FeatureManager.FeatureCut4(True, False, True, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
            Part.SelectionManager.EnableContourSelection = False
            Part.ClearSelection2(True)

            Part.ViewOrientationUndo()
        End If

        'Save File
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_06" & Convert.ToChar(Stamp + 64) & "_Bot-L.SLDPRT", 0, 2)
        Part = Nothing
        StdFunc.CloseActiveDoc()

        'BOM Entries
        bomData.EnterValuesInCNC(JobNo & "_06" & Convert.ToChar(Stamp + 64) & "_bot section -l", (SectionWidth * 1000) - 2 + 85.6, (SectionHeight * 1000) - 1 + 157.21, "4.0", 1, Client, AHUName, JobNo)

        'Add entry to predictive Database
        Dim name As String = "FrameBotSec_" & Convert.ToString(Truncate(SectionWidth * 1000)) & "_" & Convert.ToString(Truncate(WallWth)) & "_" & Convert.ToString(Truncate(BoxWth)) & "_" & Stamp
        predictivedb.AHUPartCount(name)

    End Sub

    Public Sub FrameTopSec(SectionWidth As Decimal, SectionHeight As Decimal, SectionWthSet As Integer(), WallWth As Integer, BoxWth As Integer, FanWthNo As Integer, RHSSecWth As Integer, DoorWth As Integer, Stamp As Integer, lastpart As Integer)

        'Variables
        Dim PartNo As Integer = Convert.ToInt16(Stamp)
        Dim SideBlankWth As Integer = (WallWth - DoorWth - (FanWthNo * BoxWth))

        Dim TempSecWth As Integer = 0
        If PartNo > 1 Then
            For i = 0 To PartNo - 2
                TempSecWth = TempSecWth + (SectionWthSet(i))
            Next
        End If

        Dim TempBoxWth As Integer = SideBlankWth + 50
        Dim NoOfBox As Integer = 0
        If PartNo > 1 Then
            While TempBoxWth < TempSecWth
                NoOfBox += 1
                TempBoxWth += BoxWth
            End While
        End If

        Dim TopHoleCenter As Decimal
        If PartNo = 1 Then
            TopHoleCenter = (RHSSecWth - 50) / 1000
        Else
            TopHoleCenter = (TempBoxWth - TempSecWth) / 1000
        End If

        If PartNo = 1 And SideBlankWth >= 200 Then
            BoxWth = SideBlankWth
        End If

        'Open Doc
        Part = swApp.OpenDoc6("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Frame - Sample\_05_top-l.SLDPRT", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("_05_top-l", False, longstatus)
        Part = swApp.ActiveDoc

        'Re-Dim Part
        boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_05_top-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        myDimension = Part.Parameter("Width@BaseFlange")
        myDimension.SystemValue = SectionWidth - 0.001
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Height@BaseFlange@_05_top-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        myDimension = Part.Parameter("Height@BaseFlange")
        myDimension.SystemValue = SectionHeight
        Part.ClearSelection2(True)

        SideTOPBOTHoles(SectionWidth - 0.001)

        boolstatus = Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Interbolting Holes Bottom - 1
        boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        Part.SketchManager.InsertSketch(True)
        Part.ClearSelection2(True)

        If PartNo = 1 Then
            skSegment = Part.SketchManager.CreateCircle(-0.025, -0.025, 0, -0.025, -0.025 + 0.0045, 0)
            Part.SketchAddConstraints("sgFIXED")
            Part.ClearSelection2(True)

            skSegment = Part.SketchManager.CreateCircle(-1 * ((RHSSecWth / 1000) - 0.025), -0.025, 0, -1 * ((RHSSecWth / 1000) - 0.025), -0.025 + 0.0045, 0)
            Part.SketchAddConstraints("sgFIXED")
            Part.ClearSelection2(True)
        Else
            If (TopHoleCenter + 0.025) <= SectionWidth Then
                skSegment = Part.SketchManager.CreateCircle(-1 * (TopHoleCenter - 0.025), -0.025, 0, -1 * (TopHoleCenter - 0.025), -0.025 + 0.0045, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)
            End If

            If (TopHoleCenter - 0.025) > 0.01 Then
                skSegment = Part.SketchManager.CreateCircle(-1 * (TopHoleCenter + 0.025), -0.025, 0, -1 * (TopHoleCenter + 0.025), -0.025 + 0.0045, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)
            End If
        End If

        myFeature = Part.FeatureManager.FeatureCut4(True, False, False, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
        Part.SelectionManager.EnableContourSelection = False
        Part.ClearSelection2(True)

        Part.ViewOrientationUndo()

        'Interbolting Holes Bottom - 2
        If TopHoleCenter + (BoxWth / 1000) <= SectionWidth Then
            boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            Part.SketchManager.InsertSketch(True)
            Part.ClearSelection2(True)

            If (TopHoleCenter + (BoxWth / 1000) - 0.025) <= SectionWidth Then
                skSegment = Part.SketchManager.CreateCircle(-1 * (TopHoleCenter + (BoxWth / 1000) - 0.025), -0.025, 0, -1 * (TopHoleCenter + (BoxWth / 1000) - 0.025), -0.025 + 0.0045, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)
            End If

            If (TopHoleCenter + (BoxWth / 1000) + 0.025) <= SectionWidth Then
                skSegment = Part.SketchManager.CreateCircle(-1 * (TopHoleCenter + (BoxWth / 1000) + 0.025), -0.025, 0, -1 * (TopHoleCenter + (BoxWth / 1000) + 0.025), -0.025 + 0.0045, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)
            End If

            myFeature = Part.FeatureManager.FeatureCut4(True, False, False, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
            Part.SelectionManager.EnableContourSelection = False
            Part.ClearSelection2(True)

            Part.ViewOrientationUndo()
        End If

        'Interbolting Holes Back - 1
        boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        Part.SketchManager.InsertSketch(True)
        Part.ClearSelection2(True)

        If PartNo > 1 Then
            If (TopHoleCenter - 0.1) > 0.01 Then
                skSegment = Part.SketchManager.CreateCircle(-1 * TopHoleCenter + 0.1, -1 * (SectionHeight - 0.075), 0, -1 * TopHoleCenter + 0.1, -1 * (SectionHeight - 0.075) + 0.0045, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)
            End If
        End If

        If (TopHoleCenter + 0.1) <= SectionWidth Then
            skSegment = Part.SketchManager.CreateCircle(-1 * TopHoleCenter - 0.1, -1 * (SectionHeight - 0.075), 0, -1 * TopHoleCenter - 0.1, -1 * (SectionHeight - 0.075) + 0.0045, 0)
            Part.SketchAddConstraints("sgFIXED")
            Part.ClearSelection2(True)
        End If

        myFeature = Part.FeatureManager.FeatureCut4(True, False, True, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
        Part.SelectionManager.EnableContourSelection = False
        Part.ClearSelection2(True)

        Part.ViewOrientationUndo()

        If Stamp = 1 Then
            boolstatus = Part.Extension.SelectByID2("ExtRight@Right@_05_top-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("ExtRight@Right")
            myDimension.SystemValue = 0.1
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("DiaRight@CutRight@_05_top-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("DiaRight@CutRight")
            myDimension.SystemValue = 0.0042
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("DistRight@CutRight@_05_top-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("DistRight@CutRight")
            myDimension.SystemValue = 0.075
            Part.ClearSelection2(True)

            boolstatus = Part.EditRebuild3()
        End If

        If Stamp = lastpart Then
            boolstatus = Part.Extension.SelectByID2("ExtLeft@Left@_05_top-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("ExtLeft@Left")
            myDimension.SystemValue = 0.1
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("DiaLeft@CutLeft@_05_top-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("DiaLeft@CutLeft")
            myDimension.SystemValue = 0.0042
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("DistLeft@CutLeft@_05_top-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = Part.Parameter("DistLeft@CutLeft")
            myDimension.SystemValue = 0.075
            Part.ClearSelection2(True)

            boolstatus = Part.EditRebuild3()
        End If


        'Interbolting Holes Back - 2
        If TopHoleCenter + (BoxWth / 1000) <= SectionWidth Then
            boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            Part.SketchManager.InsertSketch(True)
            Part.ClearSelection2(True)

            If (TopHoleCenter + (BoxWth / 1000) - 0.1) <= SectionWidth Then
                skSegment = Part.SketchManager.CreateCircle(-1 * TopHoleCenter - (BoxWth / 1000) + 0.1, -1 * (SectionHeight - 0.075), 0, -1 * TopHoleCenter - (BoxWth / 1000) + 0.1, -1 * (SectionHeight - 0.075) + 0.0045, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)
            End If

            If PartNo < SectionWthSet.Length Then
                If (TopHoleCenter + (BoxWth / 1000) + 0.1) <= SectionWidth Then
                    skSegment = Part.SketchManager.CreateCircle(-1 * TopHoleCenter - (BoxWth / 1000) - 0.1, -1 * (SectionHeight - 0.075), 0, -1 * TopHoleCenter - (BoxWth / 1000) - 0.1, -1 * (SectionHeight - 0.075) + 0.0045, 0)
                    Part.SketchAddConstraints("sgFIXED")
                    Part.ClearSelection2(True)
                End If
            End If

            myFeature = Part.FeatureManager.FeatureCut4(True, False, True, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
            Part.SelectionManager.EnableContourSelection = False
            Part.ClearSelection2(True)

            Part.ViewOrientationUndo()
        End If

        'Last Holes for door Type
        If PartNo = SectionWthSet.Length And DoorWth > 0 Then
            'Bottom Holes
            boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            Part.SketchManager.InsertSketch(True)
            Part.ClearSelection2(True)

            skSegment = Part.SketchManager.CreateCircle(-1 * (SectionWidth - 0.075), -0.025, 0, -1 * (SectionWidth - 0.075), -0.025 + 0.0045, 0)
            Part.SketchAddConstraints("sgFIXED")
            Part.ClearSelection2(True)

            skSegment = Part.SketchManager.CreateCircle(-1 * (SectionWidth - 0.025), -0.025, 0, -1 * (SectionWidth - 0.025), -0.025 + 0.0045, 0)
            Part.SketchAddConstraints("sgFIXED")
            Part.ClearSelection2(True)

            myFeature = Part.FeatureManager.FeatureCut4(True, False, False, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
            Part.SelectionManager.EnableContourSelection = False
            Part.ClearSelection2(True)

            Part.ViewOrientationUndo()

            'Back Hole
            boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            Part.SketchManager.InsertSketch(True)
            Part.ClearSelection2(True)

            skSegment = Part.SketchManager.CreateCircle(-1 * (TopHoleCenter - 0.1), -0.025, 0, -1 * (TopHoleCenter - 0.1), -0.025 + 0.0045, 0)
            Part.SketchAddConstraints("sgFIXED")
            Part.ClearSelection2(True)

            myFeature = Part.FeatureManager.FeatureCut4(True, False, True, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
            Part.SelectionManager.EnableContourSelection = False
            Part.ClearSelection2(True)

            Part.ViewOrientationUndo()
        End If

        'Save File
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_07" & Convert.ToChar(Stamp + 64) & "_Top-L.SLDPRT", 0, 2)
        Part = Nothing
        StdFunc.CloseActiveDoc()

        'BOM Entries
        bomData.EnterValuesInCNC(JobNo & "_07" & Convert.ToChar(Stamp + 64) & "_Top-L", (SectionWidth * 1000) - 2 + 85.6, (SectionHeight * 1000) - 1 + 157.21, "4.0", 1, Client, AHUName, JobNo)

        'Add entry to predictive Database
        Dim name As String = "FrameTopSec_" & Convert.ToString(Truncate(SectionWidth * 1000)) & "x" & Convert.ToString(Truncate(SectionHeight * 1000)) & "_" & Convert.ToString(Truncate(WallWth)) & "_" & Convert.ToString(Truncate(BoxWth)) & "_" & Stamp
        predictivedb.AHUPartCount(name)

    End Sub

    Public Sub FrameSideSecRHS(SectionWidth As Decimal, SectionHeight As Decimal, SectionHtSet As Integer(), WallHt As Integer, BoxHt As Integer, Stamp As Integer, lastpart As Integer)

        'Variables
        Dim PartNo As Integer = Convert.ToInt16(Stamp)
        Dim SideHoleCenter As Decimal
        Dim LastPartNo As Integer = Convert.ToInt16(Stamp)

        Dim TempSecHt As Integer = 0
        If PartNo > 1 Then
            For i = 0 To PartNo - 2
                TempSecHt = TempSecHt + (SectionHtSet(i))
            Next
        End If

        Dim TempBoxHt As Integer = 0
        Dim NoOfBox As Integer = 0
        If PartNo > 1 Then
            While TempBoxHt < TempSecHt
                NoOfBox += 1
                TempBoxHt = NoOfBox * BoxHt
            End While
        End If

        If PartNo <> lastpart Then
            SideHoleCenter = (BoxHt - 100) / 1000
        Else
            SideHoleCenter = (TempBoxHt - 100 - TempSecHt) / 1000
        End If

        'Open Doc
        Part = swApp.OpenDoc6("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Frame - Sample\_08_side rhs bot-l.SLDPRT", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("_08_side rhs bot-l", False, longstatus)
        Part = swApp.ActiveDoc

        'Re-Dim Part
        boolstatus = Part.Extension.SelectByID2("Height@BaseFlange@_08_side rhs bot-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        myDimension = Part.Parameter("Height@BaseFlange")
        myDimension.SystemValue = SectionHeight - 0.001
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_08_side rhs bot-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        myDimension = Part.Parameter("Width@BaseFlange")
        myDimension.SystemValue = SectionWidth
        Part.ClearSelection2(True)

        SideLHSRHSHoles(SectionHeight - 0.001)

        boolstatus = Part.EditRebuild3()
        Part.ViewZoomtofit2()

        If PartNo = lastpart Then
            If SectionHeight - 0.001 > BoxHt / 1000 Then
                'Interbolting Holes Side - 1
                If SideHoleCenter <= SectionHeight Then

                    If (SideHoleCenter + 0.025) < SectionHeight Then

                        Part = swApp.ActiveDoc
                        boolstatus = Part.Extension.SelectByID2("SideHoles1", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                        Part.EditSketch()

                        boolstatus = Part.Extension.SelectByID2("SideHoleDist@SideHoles1@_08_side rhs bot-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                        myDimension = Part.Parameter("SideHoleDist@SideHoles1")
                        myDimension.SystemValue = SideHoleCenter + 0.024
                        Part.ClearSelection2(True)

                    End If

                    Part.ViewOrientationUndo()
                End If

                'Interbolting Holes Side - 2
                If SideHoleCenter + (BoxHt / 1000) <= SectionHeight Then
                    boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    Part.SketchManager.InsertSketch(True)
                    Part.ClearSelection2(True)

                    skSegment = Part.SketchManager.CreateCircle(-0.025, SideHoleCenter + (BoxHt / 1000) - 0.025, 0, -0.025 + 0.0045, SideHoleCenter + (BoxHt / 1000) - 0.025, 0)
                    Part.SketchAddConstraints("sgFIXED")
                    Part.ClearSelection2(True)

                    If (SideHoleCenter + (BoxHt / 1000) + 0.025) < SectionHeight Then
                        skSegment = Part.SketchManager.CreateCircle(-0.025, SideHoleCenter + (BoxHt / 1000) + 0.025, 0, -0.025 + 0.0045, SideHoleCenter + (BoxHt / 1000) + 0.025, 0)
                        Part.SketchAddConstraints("sgFIXED")
                        Part.ClearSelection2(True)
                    End If

                    myFeature = Part.FeatureManager.FeatureCut4(True, False, False, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
                    Part.SelectionManager.EnableContourSelection = False
                    Part.ClearSelection2(True)

                    Part.ViewOrientationUndo()
                End If
            End If

        Else
            'Interbolting Holes Side - 1
            If SideHoleCenter <= SectionHeight Then

                If (SideHoleCenter + 0.025) < SectionHeight Then

                    Part = swApp.ActiveDoc
                    boolstatus = Part.Extension.SelectByID2("SideHoles1", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                    Part.EditSketch()

                    boolstatus = Part.Extension.SelectByID2("SideHoleDist@SideHoles1@_08_side rhs bot-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                    myDimension = Part.Parameter("SideHoleDist@SideHoles1")
                    myDimension.SystemValue = SideHoleCenter + 0.024
                    Part.ClearSelection2(True)

                End If

                Part.ViewOrientationUndo()
            End If

            'Interbolting Holes Side - 2
            If SideHoleCenter + (BoxHt / 1000) <= SectionHeight Then
                boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                Part.SketchManager.InsertSketch(True)
                Part.ClearSelection2(True)

                skSegment = Part.SketchManager.CreateCircle(-0.025, SideHoleCenter + (BoxHt / 1000) - 0.025, 0, -0.025 + 0.0045, SideHoleCenter + (BoxHt / 1000) - 0.025, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)

                If (SideHoleCenter + (BoxHt / 1000) + 0.025) < SectionHeight Then
                    skSegment = Part.SketchManager.CreateCircle(-0.025, SideHoleCenter + (BoxHt / 1000) + 0.025, 0, -0.025 + 0.0045, SideHoleCenter + (BoxHt / 1000) + 0.025, 0)
                    Part.SketchAddConstraints("sgFIXED")
                    Part.ClearSelection2(True)
                End If

                myFeature = Part.FeatureManager.FeatureCut4(True, False, False, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
                Part.SelectionManager.EnableContourSelection = False
                Part.ClearSelection2(True)

                Part.ViewOrientationUndo()
            End If
        End If


        'Interbolting Holes Back - 1
        If SideHoleCenter - 0.11 <= SectionHeight Then
            boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            Part.SketchManager.InsertSketch(True)
            Part.ClearSelection2(True)

            If SideHoleCenter + 0.1 < SectionHeight Then
                skSegment = Part.SketchManager.CreateCircle(-1 * (SectionWidth - 0.075), SideHoleCenter + 0.1, 0, -1 * (SectionWidth - 0.075) + 0.0045, SideHoleCenter + 0.1, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)
            End If

            If SideHoleCenter - 0.1 > 0.01 Then
                skSegment = Part.SketchManager.CreateCircle(-1 * (SectionWidth - 0.075), SideHoleCenter - 0.1, 0, -1 * (SectionWidth - 0.075) + 0.0045, SideHoleCenter - 0.1, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)
            End If

            myFeature = Part.FeatureManager.FeatureCut4(True, False, True, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
            Part.SelectionManager.EnableContourSelection = False
            Part.ClearSelection2(True)

            Part.ViewOrientationUndo()
        End If

        'Interbolting Holes Back - 2
        If SideHoleCenter + (BoxHt / 1000) - 0.09 <= SectionHeight Then
            boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            Part.SketchManager.InsertSketch(True)
            Part.ClearSelection2(True)

            If (SideHoleCenter + (BoxHt / 1000) + 0.1) < SectionHeight Then
                skSegment = Part.SketchManager.CreateCircle(-1 * (SectionWidth - 0.075), SideHoleCenter + (BoxHt / 1000) + 0.1, 0, -1 * (SectionWidth - 0.075) + 0.0045, SideHoleCenter + (BoxHt / 1000) + 0.1, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)
            End If

            skSegment = Part.SketchManager.CreateCircle(-1 * (SectionWidth - 0.075), SideHoleCenter + (BoxHt / 1000) - 0.1, 0, -1 * (SectionWidth - 0.075) + 0.0045, SideHoleCenter + (BoxHt / 1000) - 0.1, 0)
            Part.SketchAddConstraints("sgFIXED")
            Part.ClearSelection2(True)

            myFeature = Part.FeatureManager.FeatureCut4(True, False, True, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
            Part.SelectionManager.EnableContourSelection = False
            Part.ClearSelection2(True)

            Part.ViewOrientationUndo()
        End If

        'Top Blank Top Hole
        If PartNo = (SectionHtSet.Length) And (WallHt - 50 - (BoxHt * NoOfBox)) >= 200 Then
            boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            Part.SketchManager.InsertSketch(True)
            Part.ClearSelection2(True)

            skSegment = Part.SketchManager.CreateCircle(-1 * (SectionWidth - 0.075), SectionHeight - 0.05, 0, -1 * (SectionWidth - 0.075) + 0.0045, SectionHeight - 0.05, 0)
            Part.SketchAddConstraints("sgFIXED")
            Part.ClearSelection2(True)

            myFeature = Part.FeatureManager.FeatureCut4(True, False, True, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
            Part.SelectionManager.EnableContourSelection = False
            Part.ClearSelection2(True)

            Part.ViewOrientationUndo()
        End If

        'Save File
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_08" & Convert.ToChar(Stamp + 64) & "_RHS-L.SLDPRT", 0, 2)
        Part = Nothing
        StdFunc.CloseActiveDoc()

        'BOM Entries
        bomData.EnterValuesInCNC(JobNo & "_08" & Convert.ToChar(Stamp + 64) & "_RHS-L", (SectionWidth * 1000) - 1 + 157.21, (SectionHeight * 1000) - 2 + 85.6, "4.0", 1, Client, AHUName, JobNo)

        'Add entry to predictive Database
        Dim name As String = "FrameSideSecRHS_" & Convert.ToString(Truncate(SectionWidth * 1000)) & "_" & Convert.ToString(Truncate(SectionHeight * 1000)) & "_" & Convert.ToString(Truncate(WallHt)) & "_" & Convert.ToString(Truncate(BoxHt)) & "_" & Stamp
        predictivedb.AHUPartCount(name)

    End Sub

    Public Sub FrameSideSecLHS(SectionWidth As Decimal, SectionHeight As Decimal, SectionHtSet As Integer(), WallHt As Integer, BoxHt As Integer, Stamp As Integer, lastpart As Integer)

        'Variables
        Dim VerSecNo As Integer = Ceiling((WallHt - 100) / 1500)
        Dim PartNo As Integer = Convert.ToInt16(Stamp)
        Dim SideHoleCenter As Decimal

        Dim TempSecHt As Integer = 0
        If PartNo > 1 Then
            For i = 0 To PartNo - 2
                TempSecHt += (SectionHtSet(i))
            Next
        End If

        Dim TempBoxHt As Integer = 0
        Dim NoOfBox As Integer = 0
        If PartNo > 1 Then
            While TempBoxHt < TempSecHt
                NoOfBox += 1
                TempBoxHt = NoOfBox * BoxHt
            End While
        End If

        If PartNo <> lastpart Then
            SideHoleCenter = (BoxHt - 100) / 1000
        Else
            SideHoleCenter = (TempBoxHt - 100 - TempSecHt) / 1000
        End If

        'Open Doc
        Part = swApp.OpenDoc6("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Frame - Sample\_07_side lhs bot-l.SLDPRT", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("_07_side lhs bot-l", False, longstatus)
        Part = swApp.ActiveDoc

        'Re-Dim Part
        boolstatus = Part.Extension.SelectByID2("Height@BaseFlange@_07_side bot-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        myDimension = Part.Parameter("Height@BaseFlange")
        myDimension.SystemValue = SectionHeight - 0.001
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_07_side bot-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        myDimension = Part.Parameter("Width@BaseFlange")
        myDimension.SystemValue = SectionWidth
        Part.ClearSelection2(True)

        SideLHSRHSHoles(SectionHeight - 0.001)

        boolstatus = Part.EditRebuild3()
        Part.ViewZoomtofit2()

        If PartNo = lastpart Then
            If SectionHeight - 0.001 > BoxHt / 1000 Then
                'Interbolting Holes Side - 1
                If SideHoleCenter <= SectionHeight Then

                    If (SideHoleCenter + 0.025) < SectionHeight Then
                        Part = swApp.ActiveDoc
                        boolstatus = Part.Extension.SelectByID2("SideHoles1", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                        Part.EditSketch()

                        boolstatus = Part.Extension.SelectByID2("SideHoleDist@SideHoles1@_07_side lhs bot-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                        myDimension = Part.Parameter("SideHoleDist@SideHoles1")
                        myDimension.SystemValue = SideHoleCenter + 0.024
                        Part.ClearSelection2(True)

                    End If

                    Part.ViewOrientationUndo()
                End If

                'Interbolting Holes Side - 2
                If SideHoleCenter + (BoxHt / 1000) <= SectionHeight Then
                    boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    Part.SketchManager.InsertSketch(True)
                    Part.ClearSelection2(True)

                    skSegment = Part.SketchManager.CreateCircle(-0.025, SideHoleCenter + (BoxHt / 1000) - 0.025, 0, -0.025 + 0.0045, SideHoleCenter + (BoxHt / 1000) - 0.025, 0)
                    Part.SketchAddConstraints("sgFIXED")
                    Part.ClearSelection2(True)

                    If (SideHoleCenter + (BoxHt / 1000) + 0.025) < SectionHeight Then
                        skSegment = Part.SketchManager.CreateCircle(-0.025, SideHoleCenter + (BoxHt / 1000) + 0.025, 0, -0.025 + 0.0045, SideHoleCenter + (BoxHt / 1000) + 0.025, 0)
                        Part.SketchAddConstraints("sgFIXED")
                        Part.ClearSelection2(True)
                    End If

                    myFeature = Part.FeatureManager.FeatureCut4(True, False, True, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
                    Part.SelectionManager.EnableContourSelection = False
                    Part.ClearSelection2(True)

                    Part.ViewOrientationUndo()
                End If
            End If
        Else
            'Interbolting Holes Side - 1
            If SideHoleCenter <= SectionHeight Then

                If (SideHoleCenter + 0.025) < SectionHeight Then
                    Part = swApp.ActiveDoc
                    boolstatus = Part.Extension.SelectByID2("SideHoles1", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                    Part.EditSketch()

                    boolstatus = Part.Extension.SelectByID2("SideHoleDist@SideHoles1@_07_side lhs bot-l.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                    myDimension = Part.Parameter("SideHoleDist@SideHoles1")
                    myDimension.SystemValue = SideHoleCenter + 0.024
                    Part.ClearSelection2(True)

                End If

                Part.ViewOrientationUndo()
            End If

            'Interbolting Holes Side - 2
            If SideHoleCenter + (BoxHt / 1000) <= SectionHeight Then
                boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                Part.SketchManager.InsertSketch(True)
                Part.ClearSelection2(True)

                skSegment = Part.SketchManager.CreateCircle(-0.025, SideHoleCenter + (BoxHt / 1000) - 0.025, 0, -0.025 + 0.0045, SideHoleCenter + (BoxHt / 1000) - 0.025, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)

                If (SideHoleCenter + (BoxHt / 1000) + 0.025) < SectionHeight Then
                    skSegment = Part.SketchManager.CreateCircle(-0.025, SideHoleCenter + (BoxHt / 1000) + 0.025, 0, -0.025 + 0.0045, SideHoleCenter + (BoxHt / 1000) + 0.025, 0)
                    Part.SketchAddConstraints("sgFIXED")
                    Part.ClearSelection2(True)
                End If

                myFeature = Part.FeatureManager.FeatureCut4(True, False, True, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
                Part.SelectionManager.EnableContourSelection = False
                Part.ClearSelection2(True)

                Part.ViewOrientationUndo()
            End If
        End If

        'Interbolting Holes Back - 1
        If SideHoleCenter - 0.11 <= SectionHeight Then
            boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            Part.SketchManager.InsertSketch(True)
            Part.ClearSelection2(True)

            If SideHoleCenter + 0.1 < SectionHeight Then
                skSegment = Part.SketchManager.CreateCircle(0.025, SideHoleCenter + 0.1, 0, 0.025 + 0.0045, SideHoleCenter + 0.1, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)
            End If

            If SideHoleCenter - 0.1 > 0.01 Then
                skSegment = Part.SketchManager.CreateCircle(0.025, SideHoleCenter - 0.1, 0, 0.025 + 0.0045, SideHoleCenter - 0.1, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)
            End If

            myFeature = Part.FeatureManager.FeatureCut4(True, False, True, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
            Part.SelectionManager.EnableContourSelection = False
            Part.ClearSelection2(True)

            Part.ViewOrientationUndo()
        End If

        'Interbolting Holes Back - 2
        If SideHoleCenter + (BoxHt / 1000) - 0.09 <= SectionHeight Then
            boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            Part.SketchManager.InsertSketch(True)
            Part.ClearSelection2(True)

            If (SideHoleCenter + (BoxHt / 1000) + 0.1) < SectionHeight Then
                skSegment = Part.SketchManager.CreateCircle(0.025, SideHoleCenter + (BoxHt / 1000) + 0.1, 0, 0.025 + 0.0045, SideHoleCenter + (BoxHt / 1000) + 0.1, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)
            End If

            skSegment = Part.SketchManager.CreateCircle(0.025, SideHoleCenter + (BoxHt / 1000) - 0.1, 0, 0.025 + 0.0045, SideHoleCenter + (BoxHt / 1000) - 0.1, 0)
            Part.SketchAddConstraints("sgFIXED")
            Part.ClearSelection2(True)

            myFeature = Part.FeatureManager.FeatureCut4(True, False, True, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
            Part.SelectionManager.EnableContourSelection = False
            Part.ClearSelection2(True)

            Part.ViewOrientationUndo()
        End If

        'Top Blank Top Hole
        If PartNo = (SectionHtSet.Length) And (WallHt - 50 - (BoxHt * NoOfBox)) >= 200 Then
            boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            Part.SketchManager.InsertSketch(True)
            Part.ClearSelection2(True)

            skSegment = Part.SketchManager.CreateCircle(0.025, SectionHeight - 0.05, 0, 0.025 + 0.0045, SectionHeight - 0.05, 0)
            Part.SketchAddConstraints("sgFIXED")
            Part.ClearSelection2(True)

            myFeature = Part.FeatureManager.FeatureCut4(True, False, True, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
            Part.SelectionManager.EnableContourSelection = False
            Part.ClearSelection2(True)

            Part.ViewOrientationUndo()
        End If

        'Save File
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_09" & Convert.ToChar(Stamp + 64) & "_LHS-L.SLDPRT", 0, 2)
        Part = Nothing
        StdFunc.CloseActiveDoc()

        'BOM Entries
        bomData.EnterValuesInCNC(JobNo & "_09" & Convert.ToChar(Stamp + 64) & "_LHS-L", (SectionWidth * 1000) - 1 + 157.21, (SectionHeight * 1000) - 2 + 85.6, "4.0", 1, Client, AHUName, JobNo)

        'Add entry to predictive Database
        Dim name As String = "FrameSideSecLHS_" & Convert.ToString(Truncate(SectionHeight * 1000)) & "_" & Convert.ToString(Truncate(WallHt)) & "_" & Convert.ToString(Truncate(BoxHt)) & "_" & Stamp
        predictivedb.AHUPartCount(name)

    End Sub

    Public Sub FrameMidVerSec(SectionHeight As Decimal, SectionHtSet As Integer(), WallHt As Integer, BoxHt As Integer, Stamp As Integer, lastpart As Integer)

        'Variables
        Dim VerSecNo As Integer = Ceiling((WallHt - 100) / 1500)
        Dim PartNo As Integer = Convert.ToInt16(Stamp)
        Dim SideHoleCenter As Decimal

        Dim TempSecHt As Integer = 0
        If PartNo > 1 Then
            For i = 0 To PartNo - 2
                TempSecHt += (SectionHtSet(i))
            Next
        End If

        Dim TempBoxHt As Integer = 0
        Dim NoOfBox As Integer = 0
        If PartNo > 1 Then
            While TempBoxHt < TempSecHt
                NoOfBox += 1
                TempBoxHt = NoOfBox * BoxHt
            End While
        End If

        If PartNo <> lastpart Then
            SideHoleCenter = (BoxHt - 100) / 1000
        Else
            SideHoleCenter = (TempBoxHt - 100 - TempSecHt) / 1000
        End If

        'Open File
        Part = swApp.OpenDoc6("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Frame - Sample\_11_vertical mid-c.SLDPRT", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("_11_vertical mid-c", False, longstatus)
        Part = swApp.ActiveDoc

        'Re-Dim Part
        boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_11_vertical mid-c.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        myDimension = Part.Parameter("Height@BaseFlange")
        myDimension.SystemValue = SectionHeight - 0.001
        Part.ClearSelection2(True)

        boolstatus = Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Interbolting Holes Side - 1
        If SideHoleCenter <= SectionHeight Then

            If (SideHoleCenter + 0.025) < SectionHeight Then

                Part = swApp.ActiveDoc
                boolstatus = Part.Extension.SelectByID2("SideHoles1", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                Part.EditSketch()

                boolstatus = Part.Extension.SelectByID2("SideHoleDist@SideHoles1@_11_vertical mid-c.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                myDimension = Part.Parameter("SideHoleDist@SideHoles1")
                myDimension.SystemValue = SideHoleCenter + 0.024
                Part.ClearSelection2(True)

            End If

            If (SideHoleCenter - 0.025) > 0.011 Then
                skSegment = Part.SketchManager.CreateCircle(-0.025, SideHoleCenter - 0.025, 0, -0.025 + 0.0045, SideHoleCenter - 0.025, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)
            End If

            'myFeature = Part.FeatureManager.FeatureCut4(False, False, False, 2, 2, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
            'Part.SelectionManager.EnableContourSelection = False
            'Part.ClearSelection2(True)

            Part.ViewOrientationUndo()
        End If

        'Interbolting Holes Side - 2
        If SideHoleCenter + (BoxHt / 1000) - 0.025 <= SectionHeight Then
            boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            Part.SketchManager.InsertSketch(True)
            Part.ClearSelection2(True)

            If (SideHoleCenter + (BoxHt / 1000) + 0.025) < SectionHeight Then
                skSegment = Part.SketchManager.CreateCircle(-0.025, SideHoleCenter + (BoxHt / 1000) + 0.025, 0, -0.025 + 0.0045, SideHoleCenter + (BoxHt / 1000) + 0.025, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)
            End If

            skSegment = Part.SketchManager.CreateCircle(-0.025, SideHoleCenter + (BoxHt / 1000) - 0.025, 0, -0.025 + 0.0045, SideHoleCenter + (BoxHt / 1000) - 0.025, 0)
            Part.SketchAddConstraints("sgFIXED")
            Part.ClearSelection2(True)

            skSegment = Part.SketchManager.CreateCircle(-0.025, SideHoleCenter + (BoxHt / 1000) - 0.025 + 0.05, 0, -0.025 + 0.0045, SideHoleCenter + (BoxHt / 1000) - 0.025 + 0.05, 0)

            myFeature = Part.FeatureManager.FeatureCut4(False, False, False, 2, 2, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
            Part.SelectionManager.EnableContourSelection = False
            Part.ClearSelection2(True)

            Part.ViewOrientationUndo()
        End If



        'Interbolting Holes Back - 1
        If SideHoleCenter - 0.11 <= SectionHeight Then
            boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            Part.SketchManager.InsertSketch(True)
            Part.ClearSelection2(True)

            If SideHoleCenter + 0.1 < SectionHeight Then
                skSegment = Part.SketchManager.CreateCircle(-0.025, SideHoleCenter + 0.1, 0, -0.025 + 0.0045, SideHoleCenter + 0.1, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)

                skSegment = Part.SketchManager.CreateCircle(0.025, SideHoleCenter + 0.1, 0, 0.025 + 0.0045, SideHoleCenter + 0.1, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)
            End If

            If SideHoleCenter > 0.11 Then
                skSegment = Part.SketchManager.CreateCircle(-0.025, SideHoleCenter - 0.1, 0, -0.025 + 0.0045, SideHoleCenter - 0.1, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)

                skSegment = Part.SketchManager.CreateCircle(0.025, SideHoleCenter - 0.1, 0, 0.025 + 0.0045, SideHoleCenter - 0.1, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)
            End If

            myFeature = Part.FeatureManager.FeatureCut4(True, False, True, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
            Part.SelectionManager.EnableContourSelection = False
            Part.ClearSelection2(True)

            Part.ViewOrientationUndo()
        End If

        'Interbolting Holes Back - 2
        If SideHoleCenter + (BoxHt / 1000) - 0.09 <= SectionHeight Then
            boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            Part.SketchManager.InsertSketch(True)
            Part.ClearSelection2(True)

            If (SideHoleCenter + (BoxHt / 1000) + 0.1) < SectionHeight Then
                skSegment = Part.SketchManager.CreateCircle(-0.025, SideHoleCenter + (BoxHt / 1000) + 0.1, 0, -0.025 + 0.0045, SideHoleCenter + (BoxHt / 1000) + 0.1, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)

                skSegment = Part.SketchManager.CreateCircle(0.025, SideHoleCenter + (BoxHt / 1000) + 0.1, 0, 0.025 + 0.0045, SideHoleCenter + (BoxHt / 1000) + 0.1, 0)
                Part.SketchAddConstraints("sgFIXED")
                Part.ClearSelection2(True)
            End If

            skSegment = Part.SketchManager.CreateCircle(-0.025, SideHoleCenter + (BoxHt / 1000) - 0.1, 0, -0.025 + 0.0045, SideHoleCenter + (BoxHt / 1000) - 0.1, 0)
            Part.SketchAddConstraints("sgFIXED")
            Part.ClearSelection2(True)

            skSegment = Part.SketchManager.CreateCircle(0.025, SideHoleCenter + (BoxHt / 1000) - 0.1, 0, 0.025 + 0.0045, SideHoleCenter + (BoxHt / 1000) - 0.1, 0)
            Part.SketchAddConstraints("sgFIXED")
            Part.ClearSelection2(True)

            myFeature = Part.FeatureManager.FeatureCut4(True, False, True, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
            Part.SelectionManager.EnableContourSelection = False
            Part.ClearSelection2(True)

            Part.ViewOrientationUndo()
        End If

        'Top Blank Top Hole
        If PartNo = (SectionHtSet.Length) And (WallHt - 50 - (BoxHt * NoOfBox)) >= 200 Then
            boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            Part.SketchManager.InsertSketch(True)
            Part.ClearSelection2(True)

            skSegment = Part.SketchManager.CreateCircle(-0.025, SectionHeight - 0.05, 0, -0.025 + 0.0045, SectionHeight - 0.05, 0)
            Part.SketchAddConstraints("sgFIXED")
            Part.ClearSelection2(True)

            skSegment = Part.SketchManager.CreateCircle(0.025, SectionHeight - 0.05, 0, 0.025 + 0.0045, SectionHeight - 0.05, 0)
            Part.SketchAddConstraints("sgFIXED")
            Part.ClearSelection2(True)

            myFeature = Part.FeatureManager.FeatureCut4(True, False, True, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
            Part.SelectionManager.EnableContourSelection = False
            Part.ClearSelection2(True)

            Part.ViewOrientationUndo()
        End If

        'Save File
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_10" & Convert.ToChar(Stamp + 64) & "_Mid_Ver-C.SLDPRT", 0, 2)
        Part = Nothing
        StdFunc.CloseActiveDoc()

        'BOM Entries
        bomData.EnterValuesInCNC(JobNo & "_10" & Convert.ToChar(Stamp + 64) & "_Mid_Ver-C", "185.60", (SectionHeight * 1000) - 2 + 85.6, "4.0", 1, Client, AHUName, JobNo)

        'Add entry to predictive Database
        Dim name As String = "FrameMidVerSec_" & Convert.ToString(Truncate(SectionHeight * 1000)) & "_" & Convert.ToString(Truncate(WallHt)) & "_" & Convert.ToString(Truncate(BoxHt)) & "_" & Stamp
        predictivedb.AHUPartCount(name)

    End Sub

    Public Sub FrameMidHorSec(BoxWth As Decimal, SecName As String)

        Part = swApp.OpenDoc6("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Frame - Sample\_13_horizontal mid-c.SLDPRT", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("_13_horizontal mid-c", False, longstatus)
        Part = swApp.ActiveDoc

        'Re-Dim Part
        boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_13_horizontal mid-c.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        myDimension = Part.Parameter("Width@BaseFlange")
        myDimension.SystemValue = BoxWth - 0.102
        Part.ClearSelection2(True)

        boolstatus = Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Save File
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_11_Mid_Hor" & SecName & "-C.SLDPRT", 0, 2)
        Part = Nothing
        StdFunc.CloseActiveDoc()

        'BOM Entries
        bomData.EnterValuesInCNC(JobNo & "_11_Mid_Hor" & SecName & "-C", (BoxWth * 1000) - 16.4, "185.60", " 4.0", 1, Client, AHUName, JobNo)

        'Add entry to predictive Database
        Dim name As String = "FrameMidHorSec_" & Convert.ToString(Truncate(BoxWth * 1000))
        predictivedb.AHUPartCount(name)

    End Sub




    Public Sub FrameSubAssy(WallWth As Decimal, WallHt As Decimal, BoxWth As Decimal, BoxHt As Decimal, FanWthNo As Integer, FanHtNo As Integer, HorSecWth As Integer(), VerSecHt As Integer(),
                                SideBlankWth As Decimal, TopBlankHt As Decimal, Door As String, DoorWth As Decimal, DoorHt As Decimal)

        'Variables
        Dim VerSecRHSDis As Decimal
        Dim MidVerSecPattX As Integer
        Dim TempTopBlkHt As Decimal
        Dim TempTopBlkWth As Decimal

        If SideBlankWth < 0.2 Then
            VerSecRHSDis = SideBlankWth + BoxWth
            MidVerSecPattX = FanWthNo - 1
        Else
            VerSecRHSDis = SideBlankWth
            MidVerSecPattX = FanWthNo + 1
        End If

        'Open Part Files
        For i = 0 To UBound(HorSecWth)
            Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_06" & Convert.ToChar(i + 1 + 64) & "_Bot-L.SLDPRT", 1, 0, "", longstatus, longwarnings)
            Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_07" & Convert.ToChar(i + 1 + 64) & "_Top-L.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Next
        For i = 0 To UBound(VerSecHt)
            Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_08" & Convert.ToChar(i + 1 + 64) & "_RHS-L.SLDPRT", 1, 0, "", longstatus, longwarnings)
            Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_09" & Convert.ToChar(i + 1 + 64) & "_LHS-L.SLDPRT", 1, 0, "", longstatus, longwarnings)
            Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_10" & Convert.ToChar(i + 1 + 64) & "_Mid_Ver-C.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Next
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_11_Mid_Hor-C.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_11_Mid_Hor_Door-C.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_11_Mid_Hor_BLANK-C.SLDPRT", 1, 0, "", longstatus, longwarnings)

        'New Assembly Document
        Dim assyTemp As String = swApp.GetUserPreferenceStringValue(9)
        Part = swApp.NewDocument(assyTemp, 0, 0, 0)
        swApp.ActivateDoc2("Assem3", False, longstatus)
        Part = swApp.ActiveDoc
        Assy = Part

        'Add Components
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_11_Mid_Hor-C.SLDPRT", 0, BoxHt / 2, 0.025)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_11_Mid_Hor_Door-C.SLDPRT", -1 * WallWth, DoorHt / 2, 0.4)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_11_Mid_Hor_BLANK-C.SLDPRT", 2 * BoxWth, BoxHt, 0.4)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_11_Mid_Hor_BLANK-C.SLDPRT", -2 * BoxWth, BoxHt, 0.4)
        For i = 0 To UBound(HorSecWth)
            boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_06" & Convert.ToChar(i + 1 + 64) & "_Bot-L.SLDPRT", -1 * (1.8 * i), -1 * (BoxHt), 0.4)
        Next
        For i = 0 To UBound(VerSecHt)
            boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_09" & Convert.ToChar(i + 1 + 64) & "_LHS-L.SLDPRT", -1 * BoxWth * FanWthNo, (1.8 * i), 0.4)
        Next
        For i = 0 To UBound(HorSecWth)
            boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_07" & Convert.ToChar(i + 1 + 64) & "_Top-L.SLDPRT", -1 * (1.8 * i), (BoxHt * FanHtNo), 0.4)
        Next
        For i = 0 To UBound(VerSecHt)
            boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_08" & Convert.ToChar(i + 1 + 64) & "_RHS-L.SLDPRT", BoxWth, (1.8 * i), 0.4)
        Next
        For i = 0 To UBound(VerSecHt)
            boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_10" & Convert.ToChar(i + 1 + 64) & "_Mid_Ver-C.SLDPRT", -1, (1.8 * i), 0.4)
        Next

        Part.ViewZoomtofit2()

        'Save Assembly Doc
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_FrameSubAssembly.SLDASM", 0, 2)
        Part = Nothing

        swApp.CloseAllDocuments(True)

        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_FrameSubAssembly.SLDASM", 2, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2(JobNo & "_FrameSubAssembly", False, longstatus)
        Part = swApp.ActiveDoc
        Assy = Part

        Part.ViewZoomtofit2()

        'Create Axis
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

        'Mates
        'Box Horizontal Section Pattern
        boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, True, 4, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2(JobNo & "_11_Mid_Hor-C-1@" & JobNo & "_FrameSubAssembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)

        Dim Yqty As Integer = 0      'Y axis LP Quantity
        Dim Xqty As Integer = 0      'X axis LP Quantity
        Dim TempXqty As Integer

        '-x-x-x- Top Clearance Calculation -x-x-x-
        If TopBlankHt < 0.2 Then
            Yqty = FanHtNo - 1
        End If

        TempTopBlkHt = TopBlankHt / 1000

        If TopBlankHt > BoxHt Then
            Yqty = FanHtNo
            While TempTopBlkHt > BoxHt
                Yqty = Yqty + 1
                TempTopBlkHt = TempTopBlkHt - BoxHt
            End While
        End If

        If TopBlankHt <= BoxHt And TopBlankHt >= 0.2 Then
            Yqty = FanHtNo
        End If

        '-x-x-x- Side Clearance Calculation -x-x-x-
        TempTopBlkWth = SideBlankWth

        If SideBlankWth < BoxWth Then
            Xqty = FanWthNo
            TempXqty = 0
        Else
            Xqty = FanWthNo
            TempXqty = FanWthNo
            While TempTopBlkWth > BoxWth
                Xqty = Xqty + 1
                TempTopBlkWth = TempTopBlkWth - BoxWth
            End While
        End If

        myFeature = Part.FeatureManager.FeatureLinearPattern2(Yqty, BoxHt, Xqty, BoxWth, True, True, "NULL", "NULL", False)

        Part.ClearSelection2(True)
        Part.ViewZoomtofit2()

        If SideBlankWth > BoxWth Then
            Dim TotalXYqty1 As Integer = Xqty * Yqty
            Dim TotalXYqty2 As Integer = (Xqty - TempXqty + 1) * Yqty
            Dim TotalXY As Integer = TotalXYqty1 + TotalXYqty2 - 1

            Part = swApp.ActiveDoc
            boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, True, 2, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, True, 4, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(JobNo & "_11_Mid_Hor-C-1@" & JobNo & "_FrameSubAssembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)

            myFeature = Part.FeatureManager.FeatureLinearPattern2(Xqty - TempXqty + 1, BoxWth, Yqty, BoxHt, False, True, "NULL", "NULL", False)

            Part.ClearSelection2(True)
            Part.ViewZoomtofit2()

            'Delete Extra Mid Hor C
            For a = TotalXYqty1 + 1 To TotalXY - Yqty
                boolstatus = Part.Extension.SelectByID2(JobNo & "_11_Mid_Hor-C-" & a & "@" & JobNo & "_FrameSubAssembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
                Part.EditDelete()
            Next

        End If

        'Box Vertical Section - 1
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_11_Mid_Hor-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_10A_Mid_Ver-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0, 0.001, 0.001, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_11_Mid_Hor-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_10A_Mid_Ver-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        If SideBlankWth < 0.2 Then
            myMate = Assy.AddMate5(5, 0, True, BoxWth / 2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Else
            myMate = Assy.AddMate5(5, 0, False, BoxWth / 2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        End If
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_11_Mid_Hor-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_10A_Mid_Ver-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 0, True, BoxHt - 0.101, BoxHt - 0.1, BoxHt - 0.1, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Box Vertical Section - 2+
        For i = 1 To UBound(VerSecHt)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_10" & Convert.ToChar(i + 64) & "_Mid_Ver-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_10" & Convert.ToChar(i + 1 + 64) & "_Mid_Ver-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0.374985, 0.001, 0.001, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_10" & Convert.ToChar(i + 64) & "_Mid_Ver-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_10" & Convert.ToChar(i + 1 + 64) & "_Mid_Ver-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0.49, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_10" & Convert.ToChar(i + 64) & "_Mid_Ver-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_10" & Convert.ToChar(i + 1 + 64) & "_Mid_Ver-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(5, 0, False, (VerSecHt(i - 1)) / 1000, VerSecHt(i - 1) / 1000, VerSecHt(i - 1) / 1000, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            'myMate = Assy.AddMate5(5, 0, False, (VerSecHt(i - 1) - 1) / 1000, VerSecHt(i - 1) / 1000, VerSecHt(i - 1) / 1000, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)
        Next

        Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Box Vertical Section Pattern
        boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
        For i = 0 To UBound(VerSecHt)
            boolstatus = Part.Extension.SelectByID2("" & JobNo & "_10" & Convert.ToChar(i + 1 + 64) & "_Mid_Ver-C-1@" & JobNo & "_FrameSubAssembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
        Next

        Dim MidVerSecPattXOld As Integer = MidVerSecPattX
        Dim TempSideBlankWth As Decimal = SideBlankWth
        If SideBlankWth > BoxWth Then
            While TempSideBlankWth > BoxWth
                MidVerSecPattX = MidVerSecPattX + 1
                TempSideBlankWth = TempSideBlankWth - BoxWth
            End While
        End If

        myFeature = Part.FeatureManager.FeatureLinearPattern2(MidVerSecPattX, BoxWth, 1, 0.05, True, False, "NULL", "NULL", False)
        Part.ClearSelection2(True)

        Part.ViewZoomtofit2()

        If SideBlankWth > BoxWth Then

            boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
            For i = 0 To UBound(VerSecHt)
                boolstatus = Part.Extension.SelectByID2("" & JobNo & "_10" & Convert.ToChar(i + 1 + 64) & "_Mid_Ver-C-1@" & JobNo & "_FrameSubAssembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            Next

            myFeature = Part.FeatureManager.FeatureLinearPattern2(MidVerSecPattX - MidVerSecPattXOld + 1, BoxWth, 1, 0.05, False, False, "NULL", "NULL", False)
            Part.ClearSelection2(True)

        End If




        'Door Horizontal Section
        If Door = "YES" Then
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_11_Mid_Hor-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_11_Mid_Hor_Door-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0.37499, 0.001, 0.001, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_11_Mid_Hor-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_11_Mid_Hor_Door-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(5, 0, True, ((FanWthNo - 0.5) * BoxWth) + (DoorWth / 2), ((FanWthNo - 0.5) * BoxWth) + (DoorWth / 2), ((FanWthNo - 0.5) * BoxWth) + (DoorWth / 2), 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_11_Mid_Hor-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_11_Mid_Hor_Door-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(5, 0, False, DoorHt - BoxHt, DoorHt - BoxHt, DoorHt - BoxHt, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            Part.EditRebuild3()
            Part.ViewZoomtofit2()
        End If

        'Door Horizontal Section Pattern
        If Door = "YES" & TopBlankHt >= 200 Then
            boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(JobNo & "_11_Mid_Hor_Door-C-1@" & JobNo & "_FrameSubAssembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2(2, (FanHtNo * BoxHt) - DoorHt, 1, 0.05, True, False, "NULL", "NULL", False)
            Part.ClearSelection2(True)

            Part.ViewZoomtofit2()
        End If







        'Dim J As Integer = 0
        ''Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_11_Mid_Hor_BLANK-C.SLDPRT", 1, 0, "", longstatus, longwarnings)
        ''swApp.ActivateDoc2(JobNo & "_FrameSubAssembly", False, longstatus)
        'Part = swApp.ActiveDoc
        'For a = 1 To Yqty * 2
        '    boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_11_Mid_Hor_BLANK-C.SLDPRT", 2 * BoxWth, BoxHt, 0.4)
        'Next

        'For i = 1 To Yqty * 2

        'If i Mod 2 = 0 Then
        '    'BLANK Horizontal Section - Even
        '    boolstatus = Part.Extension.SelectByID2("" & JobNo & "_11_Mid_Hor_BLANK-C-" & i & "@" & JobNo & "_FrameSubAssembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)

        '    boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_11_Mid_Hor_BLANK-C-" & i - 1 & "@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        '    boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_11_Mid_Hor_BLANK-C-" & i & "@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        '    myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        '    Part.ClearSelection2(True)

        '    'boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_09A_LHS-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        '    boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_11_Mid_Hor_BLANK-C-" & i - 1 & "@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        '    boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_11_Mid_Hor_BLANK-C-" & i & "@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        '    myMate = Assy.AddMate5(5, 0, False, SideBlankWth / 2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        '    Part.ClearSelection2(True)

        '    boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_11_Mid_Hor_BLANK-C-" & i - 1 & "@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        '    boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_11_Mid_Hor_BLANK-C-" & i & "@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        '    myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        '    Part.ClearSelection2(True)

        '    Part.EditRebuild3()
        '    Part.ViewZoomtofit2()
        '    'End If
        'Else
        'BLANK Horizontal Section - Odd
        boolstatus = Part.Extension.SelectByID2(JobNo & "_11_Mid_Hor-C-1@" & JobNo & "_FrameSubAssembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)

        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_11_Mid_Hor-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, False, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_11_Mid_Hor_BLANK-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_11_Mid_Hor-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, False, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_11_Mid_Hor_BLANK-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 0, False, (BoxWth + SideBlankWth) / 2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_11_Mid_Hor-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, False, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_11_Mid_Hor_BLANK-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        Part.EditRebuild3()
        Part.ViewZoomtofit2()


        'End If

        'Next

        'BLANK Horizontal Section Pattern
        If WallHt - (BoxHt * FanHtNo) >= 0.2 Then
            boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(JobNo & "_11_Mid_Hor_BLANK-C-1@" & JobNo & "_FrameSubAssembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            If Truncate(FanHtNo / 2) > 1 Then
                myFeature = Part.FeatureManager.FeatureLinearPattern2(Truncate(FanHtNo / 2), 2 * BoxHt, 1, 0.05, True, False, "NULL", "NULL", False)
            End If
            If FanHtNo Mod 2 > 0 Then
                myFeature = Part.FeatureManager.FeatureLinearPattern2(2, (FanHtNo - 2) * BoxHt, 1, 0.05, True, False, "NULL", "NULL", False)
            End If
            Part.ClearSelection2(True)

            Part.ViewZoomtofit2()
        End If




        'Right Frame Sections - 1
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_08A_RHS-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_10A_Mid_Ver-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_08A_RHS-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_10A_Mid_Ver-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_08A_RHS-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_10A_Mid_Ver-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 0, True, VerSecRHSDis, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Right Frame Sections - 2+
        For i = 1 To UBound(VerSecHt)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_08" & Convert.ToChar(i + 64) & "_RHS-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_08" & Convert.ToChar(i + 1 + 64) & "_RHS-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0.001, 0.001, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_08" & Convert.ToChar(i + 64) & "_RHS-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_08" & Convert.ToChar(i + 1 + 64) & "_RHS-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0.001, 0.001, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_08" & Convert.ToChar(i + 64) & "_RHS-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_08" & Convert.ToChar(i + 1 + 64) & "_RHS-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(5, 0, False, VerSecHt(i - 1) / 1000, VerSecHt(i - 1) / 1000, VerSecHt(i - 1) / 1000, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)
        Next

        Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Left Frame Sections - 1
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_08A_RHS-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_09A_LHS-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_08A_RHS-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_09A_LHS-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_08A_RHS-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_09A_LHS-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 0, True, WallWth - 0.1, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Left Frame Sections - 2+
        For i = 1 To UBound(VerSecHt)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_09" & Convert.ToChar(i + 64) & "_LHS-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_09" & Convert.ToChar(i + 1 + 64) & "_LHS-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0.001, 0.001, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_09" & Convert.ToChar(i + 64) & "_LHS-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_09" & Convert.ToChar(i + 1 + 64) & "_LHS-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_09" & Convert.ToChar(i + 64) & "_LHS-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_09" & Convert.ToChar(i + 1 + 64) & "_LHS-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(5, 0, False, VerSecHt(i - 1) / 1000, VerSecHt(i - 1) / 1000, VerSecHt(i - 1) / 1000, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)
        Next

        Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Bottom Frame Sections - 1
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_11_Mid_Hor-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_06A_Bot-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 0, True, BoxHt - 0.05, BoxHt - 0.05, BoxHt - 0.05, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_06A_Bot-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_08A_RHS-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_06A_Bot-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_08A_RHS-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 0, True, 0.05, 0.05, 0.05, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Bottom Frame Sections - 2+
        For i = 1 To UBound(HorSecWth)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_06" & Convert.ToChar(i + 64) & "_Bot-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_06" & Convert.ToChar(i + 1 + 64) & "_Bot-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0.349985, 0.001, 0.001, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_06" & Convert.ToChar(i + 64) & "_Bot-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_06" & Convert.ToChar(i + 1 + 64) & "_Bot-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0.51, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_06" & Convert.ToChar(i + 64) & "_Bot-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_06" & Convert.ToChar(i + 1 + 64) & "_Bot-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(5, 0, True, HorSecWth(i - 1) / 1000, HorSecWth(i - 1) / 1000, HorSecWth(i - 1) / 1000, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)
        Next

        Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Top Frame Sections - 1
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_06A_Bot-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_07A_Top-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_06A_Bot-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_07A_Top-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 0, False, WallHt - 0.1, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_06A_Bot-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_07A_Top-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, True, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Top Frame Sections - 2+
        For i = 1 To UBound(HorSecWth)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_07" & Convert.ToChar(i + 64) & "_Top-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_07" & Convert.ToChar(i + 1 + 64) & "_Top-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0.349985, 0.001, 0.001, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_07" & Convert.ToChar(i + 64) & "_Top-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_07" & Convert.ToChar(i + 1 + 64) & "_Top-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0.51, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_07" & Convert.ToChar(i + 64) & "_Top-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_07" & Convert.ToChar(i + 1 + 64) & "_Top-L-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(5, 0, True, HorSecWth(i - 1) / 1000, HorSecWth(i - 1) / 1000, HorSecWth(i - 1) / 1000, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)
        Next

        Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Save File
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_FrameSubAssembly.SLDASM", 0, 2)
        Part = Nothing
        StdFunc.CloseActiveDoc()

    End Sub

    Public Sub FrameHorSecDrawings(HorSecNos As Integer)
        Exit Sub
        ' Variables
        Dim FilePathBotSec(HorSecNos - 1) As String
        Dim FileNameBotSec(HorSecNos - 1) As String
        Dim FilePathTopSec(HorSecNos - 1) As String
        Dim FileNameTopSec(HorSecNos - 1) As String
        For i = 0 To HorSecNos - 1
            FilePathBotSec(i) = "C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_06" & Convert.ToChar(i + 65) & "_Bot-L.SLDPRT"
            FileNameBotSec(i) = JobNo & "_06" & Convert.ToChar(i + 65) & "_Bot-L"
            FilePathTopSec(i) = "C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_07" & Convert.ToChar(i + 65) & "_Top-L.SLDPRT"
            FileNameTopSec(i) = JobNo & "_07" & Convert.ToChar(i + 65) & "_Top-L"
        Next
        Dim FilePathMidHor As String = "C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_11_Mid_Hor-C.SLDPRT"
        Dim FileNameMidHor As String = JobNo & "_11_Mid_Hor-C"

        Dim TotalSections As Integer = 2 * HorSecNos + 1
        Dim FilePath(TotalSections - 1) As String
        FilePathBotSec.CopyTo(FilePath, 0)
        FilePathTopSec.CopyTo(FilePath, FilePathBotSec.Length)
        FilePath(UBound(FilePath)) = "C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_11_Mid_Hor-C.SLDPRT"
        Dim FileName(TotalSections - 1) As String
        FileNameBotSec.CopyTo(FileName, 0)
        FileNameTopSec.CopyTo(FileName, FileNameBotSec.Length)
        FileName(UBound(FileName)) = JobNo & "_11_Mid_Hor-C"

        ' Bottom L
        Dim xDimBotSec(HorSecNos - 1), yDimBotSec(HorSecNos - 1), zDimBotSec(HorSecNos - 1) As Decimal
        Dim xDimFlatBotSec(HorSecNos - 1), yDimFlatBotSec(HorSecNos - 1), zDimFlatBotSec(HorSecNos - 1) As Decimal
        For i = 0 To UBound(FilePathBotSec)
            ' Open File
            Part = swApp.OpenDoc6(FilePathBotSec(i), 1, 0, "", longstatus, longwarnings)
            swApp.ActivateDoc2(FileNameBotSec(i), False, longstatus)
            Part = swApp.ActiveDoc

            Part.ViewZoomtofit2()

            'Bounding Box
            Dim BBoxBotSec As Object = StdFunc.BoundingBox()
            xDimBotSec(i) = Abs(BBoxBotSec(0)) + Abs(BBoxBotSec(3))
            yDimBotSec(i) = Abs(BBoxBotSec(1)) + Abs(BBoxBotSec(4))
            zDimBotSec(i) = Abs(BBoxBotSec(2)) + Abs(BBoxBotSec(5))

            'Bounding Box - Flat
            boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditUnsuppress2()

            Dim BBoxFlatBotSec As Object = StdFunc.BoundingBox()
            xDimFlatBotSec(i) = Abs(BBoxFlatBotSec(0)) + Abs(BBoxFlatBotSec(3))
            yDimFlatBotSec(i) = Abs(BBoxFlatBotSec(1)) + Abs(BBoxFlatBotSec(4))
            zDimFlatBotSec(i) = Abs(BBoxFlatBotSec(2)) + Abs(BBoxFlatBotSec(5))

            boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditSuppress2()

            swApp.CloseDoc(FileNameBotSec(i) & ".SLDPRT")
        Next

        ' Top L
        Dim xDimTopSec(HorSecNos - 1), yDimTopSec(HorSecNos - 1), zDimTopSec(HorSecNos - 1) As Decimal
        Dim xDimFlatTopSec(HorSecNos - 1), yDimFlatTopSec(HorSecNos - 1), zDimFlatTopSec(HorSecNos - 1) As Decimal
        For i = 0 To UBound(FilePathTopSec)
            ' Open File
            Part = swApp.OpenDoc6(FilePathTopSec(i), 1, 0, "", longstatus, longwarnings)
            swApp.ActivateDoc2(FileNameTopSec(i), False, longstatus)
            Part = swApp.ActiveDoc

            Part.ViewZoomtofit2()

            'Bounding Box
            Dim BBoxTopSec As Object = StdFunc.BoundingBox()
            xDimTopSec(i) = Abs(BBoxTopSec(0)) + Abs(BBoxTopSec(3))
            yDimTopSec(i) = Abs(BBoxTopSec(1)) + Abs(BBoxTopSec(4))
            zDimTopSec(i) = Abs(BBoxTopSec(2)) + Abs(BBoxTopSec(5))

            'Bounding Box - Flat
            boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditUnsuppress2()

            Dim BBoxFlatTopSec As Object = StdFunc.BoundingBox()
            xDimFlatTopSec(i) = Abs(BBoxFlatTopSec(0)) + Abs(BBoxFlatTopSec(3))
            yDimFlatTopSec(i) = Abs(BBoxFlatTopSec(1)) + Abs(BBoxFlatTopSec(4))
            zDimFlatTopSec(i) = Abs(BBoxFlatTopSec(2)) + Abs(BBoxFlatTopSec(5))

            boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditSuppress2()

            swApp.CloseDoc(FileNameTopSec(i) & ".SLDPRT")
        Next

        ' Mid Horizontal C
        ' Open File
        Part = swApp.OpenDoc6(FilePathMidHor, 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2(FileNameMidHor, False, longstatus)
        Part = swApp.ActiveDoc

        Part.ViewZoomtofit2()

        'Bounding Box
        Dim BBoxMidHor As Object = StdFunc.BoundingBox()
        Dim xDimMidHor As Decimal = Abs(BBoxMidHor(0)) + Abs(BBoxMidHor(3))
        Dim yDimMidHor As Decimal = Abs(BBoxMidHor(1)) + Abs(BBoxMidHor(4))
        Dim zDimMidHor As Decimal = Abs(BBoxMidHor(2)) + Abs(BBoxMidHor(5))

        'Bounding Box - Flat
        boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditUnsuppress2()

        Dim BBoxFlatMidHor As Object = StdFunc.BoundingBox()
        Dim xDimFlatMidHor As Decimal = Abs(BBoxFlatMidHor(0)) + Abs(BBoxFlatMidHor(3))
        Dim yDimFlatMidHor As Decimal = Abs(BBoxFlatMidHor(1)) + Abs(BBoxFlatMidHor(4))
        Dim zDimFlatMidHor As Decimal = Abs(BBoxFlatMidHor(2)) + Abs(BBoxFlatMidHor(5))

        boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditSuppress2()

        swApp.CloseDoc(FileNameMidHor & ".SLDPRT")

        ' Sheet Scale
        Dim PerSheet As Integer = 4
        If yDimFlatBotSec(0) > 0.25 Or yDimFlatTopSec(0) > 0.25 Then
            PerSheet = 3
        End If
        Dim SScaleX, SScaleY As Integer
        SScaleX = Ceiling((xDimFlatBotSec(0) + zDimBotSec(0) + xDimBotSec(0)) / (0.297 - (0.03 + 0.025 + 0.025 + 0.03)))
        SScaleY = Ceiling((yDimFlatBotSec(0) + zDimBotSec(0)) / ((0.21 / PerSheet) - (0.015 + 0.02 + 0.005)))
        Dim SScale As Integer = SScaleX
        If SScaleX < SScaleY Then
            SScale = SScaleY
        End If

        ' Adjust Values for Scale
        For i = 0 To UBound(FilePathBotSec)
            xDimBotSec(i) /= SScale
            yDimBotSec(i) /= SScale
            zDimBotSec(i) /= SScale

            xDimFlatBotSec(i) /= SScale
            yDimFlatBotSec(i) /= SScale
            zDimFlatBotSec(i) /= SScale

            xDimTopSec(i) /= SScale
            yDimTopSec(i) /= SScale
            zDimTopSec(i) /= SScale

            xDimFlatTopSec(i) /= SScale
            yDimFlatTopSec(i) /= SScale
            zDimFlatTopSec(i) /= SScale
        Next

        xDimMidHor /= SScale
        yDimMidHor /= SScale
        zDimMidHor /= SScale

        xDimFlatMidHor /= SScale
        yDimFlatMidHor /= SScale
        zDimFlatMidHor /= SScale

        Dim xDim(UBound(FileName)), yDim(UBound(FileName)), zDim(UBound(FileName)) As Decimal
        xDimBotSec.CopyTo(xDim, 0)
        xDimTopSec.CopyTo(xDim, xDimBotSec.Length)
        xDim(UBound(xDim)) = xDimMidHor
        yDimBotSec.CopyTo(yDim, 0)
        yDimTopSec.CopyTo(yDim, yDimBotSec.Length)
        yDim(UBound(yDim)) = yDimMidHor
        zDimBotSec.CopyTo(zDim, 0)
        zDimTopSec.CopyTo(zDim, zDimBotSec.Length)
        zDim(UBound(zDim)) = zDimMidHor

        Dim xDimFlat(UBound(FileName)), yDimFlat(UBound(FileName)), zDimFlat(UBound(FileName)) As Decimal
        xDimFlatBotSec.CopyTo(xDimFlat, 0)
        xDimFlatTopSec.CopyTo(xDimFlat, xDimFlatBotSec.Length)
        xDimFlat(UBound(xDimFlat)) = xDimFlatMidHor
        yDimFlatBotSec.CopyTo(yDimFlat, 0)
        yDimFlatTopSec.CopyTo(yDimFlat, yDimFlatBotSec.Length)
        yDimFlat(UBound(yDimFlat)) = yDimFlatMidHor
        zDimFlatBotSec.CopyTo(zDimFlat, 0)
        zDimFlatTopSec.CopyTo(zDimFlat, zDimFlatBotSec.Length)
        zDimFlat(UBound(zDimFlat)) = zDimFlatMidHor

        ' Get Margins
        Dim marginX As Decimal = (0.297 - (xDimFlat(0) + 0.025 + zDim(0) + 0.025 + xDim(0))) / 2
        If marginX < 0.02 Then marginX = 0.02
        Dim marginY As Decimal = (0.21 / PerSheet) - (yDimFlat(0) + 0.02 + zDim(0) + 0.005)
        If marginY < 0.02 Then marginY = 0.02

        ' Calculate View Placements
        Dim xFrontFlat As Decimal = marginX + xDimFlat(0) / 2
        Dim xRightSec As Decimal = xFrontFlat + xDimFlat(0) / 2 + 0.025 + zDim(0) / 2
        Dim xIso As Decimal = xRightSec + zDim(0) / 2 + 0.025 + xDim(0) / 2
        Dim yTopSec(UBound(FileName)) As Decimal
        yTopSec(0) = marginY + zDim(0) / 2
        Dim yFrontFlat(UBound(FileName)) As Decimal
        yFrontFlat(0) = yTopSec(0) + zDim(0) / 2 + 0.02 + yDimFlat(0) / 2
        Dim yViewInc As Decimal = 0.02 + zDim(0) + 0.02 + yDimFlat(0)
        Dim yNote(UBound(FileName)) As Decimal
        yNote(0) = yTopSec(0) + zDim(0) / 2 + 0.01

        ' Open Parts
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_FrameSubAssembly.SLDASM", swDocumentTypes_e.swDocASSEMBLY, 0, "", longstatus, longwarnings)

        ' Start Drawing
        Part = swApp.NewDocument(DrawTemp, 2, 0.297, 0.21)
        Draw = Part
        boolstatus = Draw.SetupSheet5("Sheet1", 11, 12, 1, SScale, False, DrawSheet, 0.297, 0.21, "Default", False)

        boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingBalloonStyle, 0, swBalloonStyle_e.swBS_None)

        Dim X, Y, Z As Decimal
        Dim ViewCount As Integer = 0
        Dim SheetCount As Integer = 1
        For i = 0 To UBound(FilePath)

            If ViewCount + 1 > PerSheet Then
                SheetCount += 1
                boolstatus = Draw.NewSheet3("Sheet" & SheetCount, 11, 12, 1, SScale, False, DrawSheet, 0.297, 0.21, "Default")
                ViewCount = 0
                boolstatus = Part.ActivateSheet("Sheet" & SheetCount)
            End If

            'View Placements
            yTopSec(i) = yTopSec(0) + ViewCount * yViewInc
            yFrontFlat(i) = yFrontFlat(0) + ViewCount * yViewInc
            yNote(i) = yNote(0) + ViewCount * yViewInc

            ' Views
            'Front Flat
            boolstatus = Draw.CreateFlatPatternViewFromModelView(FilePath(i), "Default", xFrontFlat, yFrontFlat(i), 0)
            boolstatus = Part.Extension.SelectByID2("Drawing View" & 1 + i * 5, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Draw.ActivateView("Drawing View" & 1 + i * 5)
            Part.ClearSelection2(True)

            'Front - Outside
            myView = Draw.CreateDrawViewFromModelView3(FilePath(i), "*Front", -xFrontFlat, yFrontFlat(i), 0)
            boolstatus = Draw.ActivateView("Drawing View" & 2 + i * 5)
            Part.ClearSelection2(True)

            'Right - Section
            boolstatus = Draw.ActivateView("Drawing View" & 2 + i * 5)
            skSegment = Part.SketchManager.CreateLine(0, (yDim(i) * SScale / 2) + 0.15, 0, 0, -(yDim(i) * SScale / 2) - 0.15, 0)
            boolstatus = Part.Extension.SelectByID2("Line" & 1 + ViewCount * 2, "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
            excludedComponents = vbEmpty
            myView = Draw.CreateSectionViewAt5(xRightSec, yFrontFlat(i), 0, Convert.ToChar(65 + (2 * i)), swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
            boolstatus = Draw.ActivateView("Drawing View" & 3 + i * 5)
            Part.ClearSelection2(True)

            'Top - Section
            boolstatus = Draw.ActivateView("Drawing View" & 2 + i * 5)
            skSegment = Part.SketchManager.CreateLine((xDim(i) * SScale / 2) + 0.15, 0, 0, -(xDim(i) * SScale / 2) - 0.15, 0, 0)
            boolstatus = Part.Extension.SelectByID2("Line" & ViewCount * 2 + 2, "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
            excludedComponents = vbEmpty
            myView = Draw.CreateSectionViewAt5(xFrontFlat, yTopSec(i), 0, Convert.ToChar(66 + (2 * i)), swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
            boolstatus = Draw.ActivateView("Drawing View" & 4 + i * 5)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Drawing View" & 4 + i * 5, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
            DrawView = Part.SelectionManager.GetSelectedObject5(1)
            boolstatus = Part.Extension.SelectByID2("Drawing View" & 1 + i * 5, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
            BaseView = Part.SelectionManager.GetSelectedObject5(1)
            boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewVerticalCenter, BaseView)
            Part.ClearSelection2(True)

            'Isometric
            myView = Draw.CreateDrawViewFromModelView3(FilePath(i), "*Dimetric", xIso, yFrontFlat(i), 0)
            boolstatus = Draw.ActivateView("Drawing View" & 5 + i * 5)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Drawing View" & 5 + i * 5, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
            DrawView = Part.SelectionManager.GetSelectedObject6(1, -1)
            boolstatus = DrawView.SetDisplayMode4(False, swDisplayMode_e.swSHADED, True, True, False)

            ' Dimentions
            'Height - Front Flat
            boolstatus = Part.Extension.SelectByRay(xFrontFlat, yFrontFlat(i) + (yDimFlat(i) / 2), -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
            boolstatus = Part.Extension.SelectByRay(xFrontFlat, yFrontFlat(i) - (yDimFlat(i) / 2), -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
            myDisplayDim = Part.AddDimension2(xFrontFlat - (xDimFlat(i) / 2) - 0.015, yFrontFlat(i), 0)
            Part.ClearSelection2(True)

            'Width - Front Flat
            boolstatus = Part.Extension.SelectByRay(xFrontFlat + (xDimFlat(i) / 2), yFrontFlat(i), -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
            boolstatus = Part.Extension.SelectByRay(xFrontFlat - (xDimFlat(i) / 2), yFrontFlat(i), -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
            myDisplayDim = Part.AddDimension2(xFrontFlat, yFrontFlat(i) - (yDimFlat(i) / 2) - 0.01, 0)
            Part.ClearSelection2(True)

            'Height - Top Section
            boolstatus = Part.Extension.SelectByRay(xFrontFlat, yTopSec(i) - (0.025 / SScale), -7000, 0, 0, -1, 0.00001, 1, False, 0, 0)
            boolstatus = Part.Extension.SelectByRay(xFrontFlat - (xDim(i) / 2) + (0.002 / SScale), yTopSec(i) + (0.025 / SScale), -7000, 0, 0, -1, 0.00001, 1, True, 0, 0)
            myDisplayDim = Part.AddDimension2(xFrontFlat - (xDim(i) / 2) - 0.015, yTopSec(i) - (zDim(i) / 2) - 0.01, 0)
            Part.ClearSelection2(True)

            'Width - Top Section
            boolstatus = Part.Extension.SelectByRay(xFrontFlat + (xDim(i) / 2), yTopSec(i), -7000, 0, 0, -1, 0.00001, 1, False, 0, 0)
            boolstatus = Part.Extension.SelectByRay(xFrontFlat - (xDim(i) / 2), yTopSec(i), -7000, 0, 0, -1, 0.00001, 1, True, 0, 0)
            myDisplayDim = Part.AddDimension2(xFrontFlat, yTopSec(i) - (zDim(i) / 2) - 0.01, 0)
            Part.ClearSelection2(True)

            ' Note
            X = Round(xDimFlat(i) * SScale * 1000, 2)
            Y = Round(yDimFlat(i) * SScale * 1000, 2)
            Z = Round(zDimFlat(i) * SScale * 1000, 2)
            myNote = Draw.CreateText2(FileName(i) & vbNewLine & "Qty - " & vbNewLine & X & "mm x " & Y & "mm x " & Z & "mm", xRightSec, yNote(i), 0, 0.004, 0)

            ViewCount += 1
        Next

        boolstatus = Part.ActivateSheet("Sheet1")

        ' Save As
        Part.ViewZoomtofit2()
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Drawings\Support Structure\" & JobNo & "_06_07_11_Hor Sections.SLDDRW", 0, 2)
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Drawings\Support Structure\" & JobNo & "_06_07_11_Hor Sections.PDF", 0, 2)
        swApp.CloseAllDocuments(True)

    End Sub

    Public Sub FrameVerSecDrawings(VerSecNos As Integer)
        Exit Sub
        ' Variables
        Dim FilePathRHSSec(VerSecNos - 1) As String
        Dim FileNameRHSSec(VerSecNos - 1) As String
        Dim FilePathLHSSec(VerSecNos - 1) As String
        Dim FileNameLHSSec(VerSecNos - 1) As String
        Dim FilePathMidVer(VerSecNos - 1) As String
        Dim FileNameMidVer(VerSecNos - 1) As String
        For i = 0 To VerSecNos - 1
            FilePathRHSSec(i) = "C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_08" & Convert.ToChar(i + 65) & "_RHS-L.SLDPRT"
            FileNameRHSSec(i) = JobNo & "_08" & Convert.ToChar(i + 65) & "_RHS-L"
            FilePathLHSSec(i) = "C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_09" & Convert.ToChar(i + 65) & "_LHS-L.SLDPRT"
            FileNameLHSSec(i) = JobNo & "_09" & Convert.ToChar(i + 65) & "_LHS-L"
            FilePathMidVer(i) = "C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_10" & Convert.ToChar(i + 65) & "_Mid_Ver-C.SLDPRT"
            FileNameMidVer(i) = JobNo & "_10" & Convert.ToChar(i + 65) & "_Mid_Ver-C"
        Next

        Dim TotalSections As Integer = 3 * VerSecNos
        Dim FilePath(TotalSections - 1) As String
        FilePathRHSSec.CopyTo(FilePath, 0)
        FilePathLHSSec.CopyTo(FilePath, FilePathRHSSec.Length)
        FilePathMidVer.CopyTo(FilePath, FilePathRHSSec.Length + FilePathLHSSec.Length)
        Dim FileName(TotalSections - 1) As String
        FileNameRHSSec.CopyTo(FileName, 0)
        FileNameLHSSec.CopyTo(FileName, FileNameRHSSec.Length)
        FileNameMidVer.CopyTo(FileName, FileNameRHSSec.Length + FileNameLHSSec.Length)

        ' RHS L
        Dim xDimRHSSec(VerSecNos - 1), yDimRHSSec(VerSecNos - 1), zDimRHSSec(VerSecNos - 1) As Decimal
        Dim xDimFlatRHSSec(VerSecNos - 1), yDimFlatRHSSec(VerSecNos - 1), zDimFlatRHSSec(VerSecNos - 1) As Decimal
        For i = 0 To UBound(FilePathRHSSec)
            ' Open File
            Part = swApp.OpenDoc6(FilePathRHSSec(i), 1, 0, "", longstatus, longwarnings)
            swApp.ActivateDoc2(FileNameRHSSec(i), False, longstatus)
            Part = swApp.ActiveDoc

            Part.ViewZoomtofit2()

            'Bounding Box
            Dim BBoxRHSSec As Object = StdFunc.BoundingBox()
            xDimRHSSec(i) = Abs(BBoxRHSSec(0)) + Abs(BBoxRHSSec(3))
            yDimRHSSec(i) = Abs(BBoxRHSSec(1)) + Abs(BBoxRHSSec(4))
            zDimRHSSec(i) = Abs(BBoxRHSSec(2)) + Abs(BBoxRHSSec(5))

            'Bounding Box - Flat
            boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditUnsuppress2()

            Dim BBoxFlatRHSSec As Object = StdFunc.BoundingBox()
            xDimFlatRHSSec(i) = Abs(BBoxFlatRHSSec(0)) + Abs(BBoxFlatRHSSec(3))
            yDimFlatRHSSec(i) = Abs(BBoxFlatRHSSec(1)) + Abs(BBoxFlatRHSSec(4))
            zDimFlatRHSSec(i) = Abs(BBoxFlatRHSSec(2)) + Abs(BBoxFlatRHSSec(5))

            boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditSuppress2()

            swApp.CloseDoc(FileNameRHSSec(i) & ".SLDPRT")
        Next

        ' LHS L
        Dim xDimLHSSec(VerSecNos - 1), yDimLHSSec(VerSecNos - 1), zDimLHSSec(VerSecNos - 1) As Decimal
        Dim xDimFlatLHSSec(VerSecNos - 1), yDimFlatLHSSec(VerSecNos - 1), zDimFlatLHSSec(VerSecNos - 1) As Decimal
        For i = 0 To UBound(FilePathLHSSec)
            ' Open File
            Part = swApp.OpenDoc6(FilePathLHSSec(i), 1, 0, "", longstatus, longwarnings)
            swApp.ActivateDoc2(FileNameLHSSec(i), False, longstatus)
            Part = swApp.ActiveDoc

            Part.ViewZoomtofit2()

            'Bounding Box
            Dim BBoxLHSSec As Object = StdFunc.BoundingBox()
            xDimLHSSec(i) = Abs(BBoxLHSSec(0)) + Abs(BBoxLHSSec(3))
            yDimLHSSec(i) = Abs(BBoxLHSSec(1)) + Abs(BBoxLHSSec(4))
            zDimLHSSec(i) = Abs(BBoxLHSSec(2)) + Abs(BBoxLHSSec(5))

            'Bounding Box - Flat
            boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditUnsuppress2()

            Dim BBoxFlatLHSSec As Object = StdFunc.BoundingBox()
            xDimFlatLHSSec(i) = Abs(BBoxFlatLHSSec(0)) + Abs(BBoxFlatLHSSec(3))
            yDimFlatLHSSec(i) = Abs(BBoxFlatLHSSec(1)) + Abs(BBoxFlatLHSSec(4))
            zDimFlatLHSSec(i) = Abs(BBoxFlatLHSSec(2)) + Abs(BBoxFlatLHSSec(5))

            boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditSuppress2()

            swApp.CloseDoc(FileNameLHSSec(i) & ".SLDPRT")
        Next

        ' Mid Vertical C
        Dim xDimMidVer(VerSecNos - 1), yDimMidVer(VerSecNos - 1), zDimMidVer(VerSecNos - 1) As Decimal
        Dim xDimFlatMidVer(VerSecNos - 1), yDimFlatMidVer(VerSecNos - 1), zDimFlatMidVer(VerSecNos - 1) As Decimal
        For i = 0 To UBound(FilePathMidVer)
            ' Open File
            Part = swApp.OpenDoc6(FilePathMidVer(i), 1, 0, "", longstatus, longwarnings)
            swApp.ActivateDoc2(FileNameMidVer(i), False, longstatus)
            Part = swApp.ActiveDoc

            Part.ViewZoomtofit2()

            'Bounding Box
            Dim BBoxMidVer As Object = StdFunc.BoundingBox()
            xDimMidVer(i) = Abs(BBoxMidVer(0)) + Abs(BBoxMidVer(3))
            yDimMidVer(i) = Abs(BBoxMidVer(1)) + Abs(BBoxMidVer(4))
            zDimMidVer(i) = Abs(BBoxMidVer(2)) + Abs(BBoxMidVer(5))

            'Bounding Box - Flat
            boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditUnsuppress2()

            Dim BBoxFlatMidVer As Object = StdFunc.BoundingBox()
            xDimFlatMidVer(i) = Abs(BBoxFlatMidVer(0)) + Abs(BBoxFlatMidVer(3))
            yDimFlatMidVer(i) = Abs(BBoxFlatMidVer(1)) + Abs(BBoxFlatMidVer(4))
            zDimFlatMidVer(i) = Abs(BBoxFlatMidVer(2)) + Abs(BBoxFlatMidVer(5))

            boolstatus = Part.Extension.SelectByID2("Flat-Pattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditSuppress2()

            swApp.CloseDoc(FileNameMidVer(i) & ".SLDPRT")
        Next

        ' Sheet Scale
        Dim PerSheet As Integer = 2
        Dim SScaleX, SScaleY As Integer
        SScaleX = Ceiling((xDimFlatRHSSec(0) + zDimRHSSec(0) + xDimRHSSec(0)) / ((0.297 / 2) - (0.02 + 0.025 + 0.025)))
        SScaleY = Ceiling((yDimFlatRHSSec(0) + zDimRHSSec(0)) / (0.21 - (0.03 + 0.02 + 0.03)))
        Dim SScale As Integer = SScaleX
        If SScaleX < SScaleY Then
            SScale = SScaleY
        End If

        ' Adjust Values for Scale
        For i = 0 To UBound(FilePathRHSSec)
            xDimRHSSec(i) /= SScale
            yDimRHSSec(i) /= SScale
            zDimRHSSec(i) /= SScale

            xDimFlatRHSSec(i) /= SScale
            yDimFlatRHSSec(i) /= SScale
            zDimFlatRHSSec(i) /= SScale

            xDimLHSSec(i) /= SScale
            yDimLHSSec(i) /= SScale
            zDimLHSSec(i) /= SScale

            xDimFlatLHSSec(i) /= SScale
            yDimFlatLHSSec(i) /= SScale
            zDimFlatLHSSec(i) /= SScale

            xDimMidVer(i) /= SScale
            yDimMidVer(i) /= SScale
            zDimMidVer(i) /= SScale

            xDimFlatMidVer(i) /= SScale
            yDimFlatMidVer(i) /= SScale
            zDimFlatMidVer(i) /= SScale
        Next

        Dim xDim(UBound(FileName)), yDim(UBound(FileName)), zDim(UBound(FileName)) As Decimal
        xDimRHSSec.CopyTo(xDim, 0)
        xDimLHSSec.CopyTo(xDim, xDimRHSSec.Length)
        xDimLHSSec.CopyTo(xDim, xDimRHSSec.Length + xDimLHSSec.Length)
        yDimRHSSec.CopyTo(yDim, 0)
        yDimLHSSec.CopyTo(yDim, yDimRHSSec.Length)
        yDimLHSSec.CopyTo(yDim, yDimRHSSec.Length + yDimLHSSec.Length)
        zDimRHSSec.CopyTo(zDim, 0)
        zDimLHSSec.CopyTo(zDim, zDimRHSSec.Length)
        zDimLHSSec.CopyTo(zDim, zDimRHSSec.Length + zDimLHSSec.Length)

        Dim xDimFlat(UBound(FileName)), yDimFlat(UBound(FileName)), zDimFlat(UBound(FileName)) As Decimal
        xDimFlatRHSSec.CopyTo(xDimFlat, 0)
        xDimFlatLHSSec.CopyTo(xDimFlat, xDimFlatRHSSec.Length)
        xDimFlatLHSSec.CopyTo(xDimFlat, xDimFlatRHSSec.Length + xDimFlatLHSSec.Length)
        yDimFlatRHSSec.CopyTo(yDimFlat, 0)
        yDimFlatLHSSec.CopyTo(yDimFlat, yDimFlatRHSSec.Length)
        yDimFlatLHSSec.CopyTo(yDimFlat, yDimFlatRHSSec.Length + yDimFlatLHSSec.Length)
        zDimFlatRHSSec.CopyTo(zDimFlat, 0)
        zDimFlatLHSSec.CopyTo(zDimFlat, zDimFlatRHSSec.Length)
        zDimFlatLHSSec.CopyTo(zDimFlat, zDimFlatRHSSec.Length + zDimFlatLHSSec.Length)

        ' Get Margins
        Dim marginX As Decimal = (0.297 / 2) - (xDimFlatRHSSec(0) + 0.025 + zDimRHSSec(0) + 0.025 + xDimRHSSec(0))
        If marginX < 0.02 Then marginX = 0.02
        If marginX > 0.03 Then marginX = 0.03
        Dim marginY As Decimal = (0.21 - (yDimFlatRHSSec(0) + 0.02 + zDimRHSSec(0))) / 2
        If marginY < 0.02 Then marginY = 0.02

        ' Calculate View Placements
        Dim xFrontFlat(UBound(FileName)) As Decimal
        xFrontFlat(0) = marginX + xDimFlat(0) / 2
        Dim xRightSec(UBound(FileName)) As Decimal
        xRightSec(0) = xFrontFlat(0) + xDimFlat(0) / 2 + 0.025 + zDim(0) / 2
        Dim xIso(UBound(FileName)) As Decimal
        xIso(0) = xRightSec(0) + zDim(0) / 2 + 0.025 + xDim(0) / 2
        Dim yTopSec As Decimal = marginY + zDim(0) / 2
        Dim yFrontFlat As Decimal = yTopSec + zDim(0) / 2 + 0.02 + yDimFlat(0) / 2
        Dim yViewInc As Decimal = 0.297 / 2
        Dim xNote(UBound(FileName)) As Decimal
        xNote(0) = xFrontFlat(0) + xDim(0) / 2 + 0.015
        Dim yNote As Decimal = yTopSec

        ' Open Parts
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_FrameSubAssembly.SLDASM", swDocumentTypes_e.swDocASSEMBLY, 0, "", longstatus, longwarnings)
        'For i = 0 To UBound(FilePath)
        '    Part = swApp.OpenDoc6(FilePath(i), swDocumentTypes_e.swDocPART, 0, "", longstatus, longwarnings)
        'Next

        ' Start Drawing
        Part = swApp.NewDocument(DrawTemp, 2, 0.297, 0.21)
        Draw = Part
        boolstatus = Draw.SetupSheet5("Sheet1", 11, 12, 1, SScale, False, DrawSheet, 0.297, 0.21, "Default", False)

        boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingBalloonStyle, 0, swBalloonStyle_e.swBS_None)

        Dim X, Y, Z As Decimal
        Dim ViewCount As Integer = 0
        Dim SheetCount As Integer = 1
        For i = 0 To UBound(FilePath)

            If ViewCount + 1 > PerSheet Then
                SheetCount += 1
                boolstatus = Draw.NewSheet3("Sheet" & SheetCount, 11, 12, 1, SScale, False, DrawSheet, 0.297, 0.21, "Default")
                ViewCount = 0
                boolstatus = Part.ActivateSheet("Sheet" & SheetCount)
            End If

            'View Placements
            xFrontFlat(i) = xFrontFlat(0) + ViewCount * yViewInc
            xRightSec(i) = xRightSec(0) + ViewCount * yViewInc
            xIso(i) = xIso(0) + ViewCount * yViewInc
            xNote(i) = xNote(0) + ViewCount * yViewInc

            ' Views
            'Front Flat
            boolstatus = Draw.CreateFlatPatternViewFromModelView(FilePath(i), "Default", xFrontFlat(i), yFrontFlat, 0)
            boolstatus = Part.Extension.SelectByID2("Drawing View" & 1 + i * 5, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = Draw.ActivateView("Drawing View" & 1 + i * 5)
            Part.ClearSelection2(True)

            'Front - Outside
            myView = Draw.CreateDrawViewFromModelView3(FilePath(i), "*Front", -xFrontFlat(i), yFrontFlat, 0)
            boolstatus = Draw.ActivateView("Drawing View" & 2 + i * 5)
            Part.ClearSelection2(True)

            'Right - Section
            boolstatus = Draw.ActivateView("Drawing View" & 2 + i * 5)
            skSegment = Part.SketchManager.CreateLine(0, (yDim(i) * SScale / 2) + 0.15, 0, 0, -(yDim(i) * SScale / 2) - 0.15, 0)
            boolstatus = Part.Extension.SelectByID2("Line" & 1 + ViewCount * 2, "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
            excludedComponents = vbEmpty
            myView = Draw.CreateSectionViewAt5(xRightSec(i), yFrontFlat, 0, Convert.ToChar(65 + (2 * i)), swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
            boolstatus = Draw.ActivateView("Drawing View" & 3 + i * 5)
            Part.ClearSelection2(True)

            'Top - Section
            boolstatus = Draw.ActivateView("Drawing View" & 2 + i * 5)
            skSegment = Part.SketchManager.CreateLine((xDim(i) * SScale / 2) + 0.15, 0, 0, -(xDim(i) * SScale / 2) - 0.15, 0, 0)
            boolstatus = Part.Extension.SelectByID2("Line" & ViewCount * 2 + 2, "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
            excludedComponents = vbEmpty
            myView = Draw.CreateSectionViewAt5(xFrontFlat(i), yTopSec, 0, Convert.ToChar(66 + (2 * i)), swCreateSectionViewAtOptions_e.swCreateSectionView_DisplaySurfaceCut, excludedComponents, 0)
            boolstatus = Draw.ActivateView("Drawing View" & 4 + i * 5)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Drawing View" & 4 + i * 5, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
            DrawView = Part.SelectionManager.GetSelectedObject5(1)
            boolstatus = Part.Extension.SelectByID2("Drawing View" & 1 + i * 5, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
            BaseView = Part.SelectionManager.GetSelectedObject5(1)
            boolstatus = DrawView.AlignWithView(swAlignViewTypes_e.swAlignViewVerticalCenter, BaseView)
            Part.ClearSelection2(True)

            'Isometric
            myView = Draw.CreateDrawViewFromModelView3(FilePath(i), "*Dimetric", xIso(i), yFrontFlat, 0)
            boolstatus = Draw.ActivateView("Drawing View" & 5 + i * 5)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Drawing View" & 5 + i * 5, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
            DrawView = Part.SelectionManager.GetSelectedObject6(1, -1)
            boolstatus = DrawView.SetDisplayMode4(False, swDisplayMode_e.swSHADED, True, True, False)

            ' Dimentions
            'Height - Front Flat
            boolstatus = Part.Extension.SelectByRay(xFrontFlat(i), yFrontFlat + (yDimFlat(i) / 2), -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
            boolstatus = Part.Extension.SelectByRay(xFrontFlat(i), yFrontFlat - (yDimFlat(i) / 2), -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
            myDisplayDim = Part.AddDimension2(xFrontFlat(i) - (xDimFlat(i) / 2) - 0.015, yFrontFlat, 0)
            Part.ClearSelection2(True)

            'Width - Front Flat
            boolstatus = Part.Extension.SelectByRay(xFrontFlat(i) + (xDimFlat(i) / 2), yFrontFlat, -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
            boolstatus = Part.Extension.SelectByRay(xFrontFlat(i) - (xDimFlat(i) / 2), yFrontFlat, -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
            myDisplayDim = Part.AddDimension2(xFrontFlat(i), yFrontFlat - (yDimFlat(i) / 2) - 0.01, 0)
            Part.ClearSelection2(True)

            'Height - Top Section
            boolstatus = Part.Extension.SelectByRay(xFrontFlat(i), yTopSec - (0.025 / SScale), -7000, 0, 0, -1, 0.00001, 1, False, 0, 0)
            boolstatus = Part.Extension.SelectByRay(xFrontFlat(i) - (xDim(i) / 2) + (0.002 / SScale), yTopSec + (0.025 / SScale), -7000, 0, 0, -1, 0.00001, 1, True, 0, 0)
            myDisplayDim = Part.AddDimension2(xFrontFlat(i) - (xDim(i) / 2) - 0.015, yTopSec - (zDim(i) / 2) - 0.01, 0)
            Part.ClearSelection2(True)

            'Width - Top Section
            boolstatus = Part.Extension.SelectByRay(xFrontFlat(i) + (xDim(i) / 2), yTopSec, -7000, 0, 0, -1, 0.00001, 1, False, 0, 0)
            boolstatus = Part.Extension.SelectByRay(xFrontFlat(i) - (xDim(i) / 2), yTopSec, -7000, 0, 0, -1, 0.00001, 1, True, 0, 0)
            myDisplayDim = Part.AddDimension2(xFrontFlat(i), yTopSec - (zDim(i) / 2) - 0.01, 0)
            Part.ClearSelection2(True)

            ' Note
            X = Round(xDimFlat(i) * SScale * 1000, 2)
            Y = Round(yDimFlat(i) * SScale * 1000, 2)
            Z = Round(zDimFlat(i) * SScale * 1000, 2)
            myNote = Draw.CreateText2(FileName(i) & vbNewLine & "Qty - " & vbNewLine & X & "mm x " & Y & "mm x " & Z & "mm", xNote(i), yNote, 0, 0.004, 0)

            ViewCount += 1
        Next

        boolstatus = Part.ActivateSheet("Sheet1")

        ' Save As
        Part.ViewZoomtofit2()
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Drawings\Support Structure\" & JobNo & "_08_09_10_Ver Sections.SLDDRW", 0, 2)
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Drawings\Support Structure\" & JobNo & "_08_09_10_Ver Sections.PDF", 0, 2)
        swApp.CloseAllDocuments(True)

    End Sub

    Public Sub FrameAssyDrawings()
        'Exit Sub
        ' Variables
        Dim FilePath As String = "C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_FrameSubAssembly.SLDASM"
        Dim FileName As String = JobNo & "_FrameSubAssembly"

        'Open File
        Part = swApp.OpenDoc6(FilePath, 2, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2(FileName, False, longstatus)
        Part = swApp.ActiveDoc

        Part.ViewZoomtofit2()

        'Bounding Box
        Dim BBox As Object = StdFunc.BoundingBoxOfAssembly()
        Dim xDim As Decimal = Abs(BBox(0)) + Abs(BBox(3))
        Dim yDim As Decimal = Abs(BBox(1)) + Abs(BBox(4))
        Dim zDim As Decimal = Abs(BBox(2)) + Abs(BBox(5))

        swApp.CloseAllDocuments(True)

        ' Sheet Scale
        Dim SScaleX As Integer = Ceiling(xDim / (0.2 - (0.03 + 0.015))) 'Using only 200mm width
        Dim SScaleY As Integer = Ceiling(yDim / (0.21 - (0.03 + 0.03)))
        Dim SScale As Integer = SScaleX
        If SScaleX < SScaleY Then
            SScale = SScaleY
        End If

        ' Adjust Values for Scale
        xDim /= SScale
        yDim /= SScale
        zDim /= SScale

        ' Get Margins
        Dim marginX As Decimal = 0.2 - (xDim + 0.015)
        If marginX < 0.03 Then marginX = 0.03

        Dim marginY As Decimal = (0.21 - (yDim + 0.04 + zDim)) / 2
        If marginY < 0.03 Then marginY = 0.03

        ' Calculate View Placements
        Dim xFront As Decimal = marginX + xDim / 2
        Dim yFront As Decimal = 0.21 / 2

        ' Open Parts
        Part = swApp.OpenDoc6(FilePath, 1, 0, "", longstatus, longwarnings)

        ' Start Drawing
        Part = swApp.NewDocument(DrawTemp, 2, 0.297, 0.21)
        Draw = Part
        boolstatus = Draw.SetupSheet5("Sheet1", 11, 12, 1, SScale, False, DrawSheet, 0.297, 0.21, "Default", False)

        swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swAutomaticScaling3ViewDrawings, False)

        ' Views
        'Front
        myView = Draw.CreateDrawViewFromModelView3(FilePath, "*Front", xFront, yFront, 0)
        boolstatus = Draw.ActivateView("Drawing View1")
        Part.ClearSelection2(True)

        '' Dimentions
        ''Height
        'boolstatus = Part.Extension.SelectByRay(xFront + (xDim / 2) - (0.025 / SScale), yFront + (yDim / 2), -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        'boolstatus = Part.Extension.SelectByRay(xFront, yFront - (yDim / 2), -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        'myDisplayDim = Part.AddDimension2(xFront + (xDim / 2) + 0.015, yFront, 0)
        'Part.ClearSelection2(True)

        ''Width
        'boolstatus = Part.Extension.SelectByRay(xFront + (xDim / 2), yFront, -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        'boolstatus = Part.Extension.SelectByRay(xFront - (xDim / 2), yFront + (yDim / 2) - (0.025 / SScale), -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        'myDisplayDim = Part.AddDimension2(xFront, yFront + (yDim / 2) + 0.015, 0)
        'Part.ClearSelection2(True)

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
        boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingBalloonStyle, 0, swBalloonStyle_e.swBS_None)

        myNote = Draw.CreateText2(FileName, xFront, 0.02, 0, 0.004, 0)

        ' Save As
        Part.ViewZoomtofit2()
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Drawings\Support Structure\" & JobNo & "_Support Structure.SLDDRW", 0, 2)
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Drawings\Support Structure\" & JobNo & "_Support Structure.PDF", 0, 2)
        swApp.CloseAllDocuments(True)

    End Sub


    '-x-x-x-x-x-x-x-x-x-x- LHS RHS 4.2mm Holes -x-x-x-x-x-x-x-x-x-x-x-x-x-x-x
    Public Sub SideLHSRHSHoles(HoleSpc As Decimal)

        '4.2mm holes--------------------------------------------------------

        Dim HoleY As Integer = SideHoleNumber(HoleSpc, 0.04)
        Dim HoleYDist As Decimal = SideHoleDist(HoleY, HoleSpc, 0.04)

        boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("4.2mm Hole", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
        myFeature = Part.FeatureManager.FeatureLinearPattern2(HoleY + 1, HoleYDist, 0, 0, False, True, "NULL", "NULL", False)
        Part.ClearSelection2(True)


    End Sub


    '-x-x-x-x-x-x-x-x-x-x- TOP BOT 4.2mm Holes -x-x-x-x-x-x-x-x-x-x-x-x-x-x-x
    Public Sub SideTOPBOTHoles(HoleSpc As Decimal)

        '4.2mm holes--------------------------------------------------------

        Dim HoleY As Integer = SideHoleNumber(HoleSpc, 0.04)
        Dim HoleYDist As Decimal = SideHoleDist(HoleY, HoleSpc, 0.04)

        boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, False, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("4.2mm Hole", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
        myFeature = Part.FeatureManager.FeatureLinearPattern2(HoleY + 1, HoleYDist, 0, 0, True, True, "NULL", "NULL", False)
        Part.ClearSelection2(True)


    End Sub


    Function SideHoleNumber(Length As Decimal, HoleClearence As Decimal) As Decimal

        Dim MinDis As Decimal = 0.115
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
            HoleNo = Truncate(ActualLength / MinDis)
            If ActualLength / HoleNo < MinDis Then
                HoleNo = 1
            End If
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





#End Region

    Public Sub BlankSheet(Height As Decimal, Width As Decimal, PartName As String)

        'Open File
        Part = swApp.OpenDoc6("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Box - Sample\14_blank sheet.SLDPRT", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("14_blank sheet", False, longstatus)
        Part = swApp.ActiveDoc

        'ReDim
        boolstatus = Part.Extension.SelectByID2("BlankWidth@BaseFlange@14_blank sheet.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 0)
        myDimension = Part.Parameter("BlankWidth@BaseFlange")
        myDimension.SystemValue = Width
        Part.ClearSelection2(True)
        boolstatus = Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("BlankHeight@BaseFlange@14_blank sheet.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 0)
        myDimension = Part.Parameter("BlankHeight@BaseFlange")
        myDimension.SystemValue = Height
        Part.ClearSelection2(True)
        boolstatus = Part.EditRebuild3()

        Part.ViewZoomtofit2()

        If Width >= 0.6 Then
            boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            Part.SketchManager.InsertSketch(True)
            Part.ClearSelection2(True)

            skSegment = Part.SketchManager.CreateCircle(0, (Height / 2) - 0.025, 0, 0.0045, (Height / 2) - 0.025, 0)
            Part.SketchAddConstraints("sgFIXED")
            Part.ClearSelection2(True)

            skSegment = Part.SketchManager.CreateCircle(0, -1 * ((Height / 2) - 0.025), 0, 0.0045, -1 * ((Height / 2) - 0.025), 0)
            Part.SketchAddConstraints("sgFIXED")
            Part.ClearSelection2(True)

            myFeature = Part.FeatureManager.FeatureCut4(True, False, True, 2, 0, 0.01, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, True, True, False, 0, 0, False, False)
            Part.SelectionManager.EnableContourSelection = False
            Part.ClearSelection2(True)

            Part.ViewOrientationUndo()
        End If

        'Save File
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\AHU Box\" & JobNo & "_14_blank sheet_" & PartName & ".SLDPRT", 0, 2)
        Part.ViewZoomtofit2()
        Part = Nothing
        StdFunc.CloseActiveDoc()

        'BOM Entries
        bomData.EnterValuesInCNC(JobNo & "_14_blank sheet_" & PartName, (Width * 1000) + 28.81, (Height * 1000) + 28.81, "2.0", 1, Client, AHUName, JobNo)

        'Add entry to predictive Database
        Dim name As String = "BoxBlank_" & Convert.ToString(Truncate(Width * 1000)) & "x" & Convert.ToString(Truncate(Height * 1000)) & "_" & PartName
        predictivedb.AHUPartCount(name)

    End Sub

    Public Sub BlankSheetDrawing(PartName As String)
        Exit Sub
        ' Variables
        Dim FilePath As String = "C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\AHU Box\" & JobNo & "_14_blank sheet_" & PartName & ".SLDPRT"
        Dim FileName As String = JobNo & "_14_blank sheet_" & PartName

        'Open File
        Part = swApp.OpenDoc6(FilePath, 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2(FileName, False, longstatus)
        Part = swApp.ActiveDoc

        Part.ViewZoomtofit2()

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
        Dim marginX As Decimal = (0.297 - (xDimFlat + 0.04 + zDim + 0.04 + xDim)) / 2
        If marginX < 0.03 Then marginX = 0.03
        Dim marginY As Decimal = (0.21 - (yDimFlat + 0.04 + zDim)) / 2
        If marginY < 0.03 Then marginY = 0.03

        ' Calculate View Placements
        Dim xFrontFlat As Decimal = marginX + xDimFlat / 2
        Dim yTopSec As Decimal = marginY + zDim / 2
        Dim yFrontFlat As Decimal = yTopSec + zDim / 2 + 0.04 + yDimFlat / 2
        Dim xRightSec As Decimal = xFrontFlat + xDimFlat / 2 + 0.04 + zDim / 2
        Dim xIso As Decimal = xRightSec + zDim / 2 + 0.04 + xDim / 2

        '' Open Parts
        'Part = swApp.OpenDoc6(FilePath, 1, 0, "", longstatus, longwarnings)

        ' Start Drawing
        Part = swApp.NewDocument(DrawTemp, 2, 0.297, 0.21)
        Draw = Part
        boolstatus = Draw.SetupSheet5("Sheet1", 11, 12, 1, SScale, False, DrawSheet, 0.297, 0.21, "Default", False)

        ' Views
        boolstatus = Draw.CreateFlatPatternViewFromModelView(FilePath, "Default", xFrontFlat, yFrontFlat, 0)
        boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Draw.ActivateView("Drawing View1")
        Part.ClearSelection2(True)

        'Front - Outside
        myView = Part.CreateDrawViewFromModelView3(FilePath, "*Front", -xFrontFlat, yFrontFlat, 0)
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
        myView = Part.CreateDrawViewFromModelView3(FilePath, "*Dimetric", xIso, yFrontFlat, 0)
        boolstatus = Draw.ActivateView("Drawing View2")
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Drawing View5", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        DrawView = Part.SelectionManager.GetSelectedObject6(1, -1)
        boolstatus = DrawView.SetDisplayMode4(False, swDisplayMode_e.swSHADED, True, True, False)

        ' Dimentions
        'Height - Flat
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, yFrontFlat + (yDimFlat / 2), -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, yFrontFlat - (yDimFlat / 2), -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat - (xDimFlat / 2) - 0.015, yFrontFlat, 0)
        Part.ClearSelection2(True)

        'Width - Flat
        boolstatus = Part.Extension.SelectByRay(xFrontFlat + (xDimFlat / 2), yFrontFlat, -7000, 0, 0, -1, 0.001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat - (xDimFlat / 2), yFrontFlat, -7000, 0, 0, -1, 0.001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat, yFrontFlat + (yDimFlat / 2) + 0.015, 0)
        Part.ClearSelection2(True)

        'Hight - Top Section
        boolstatus = Part.Extension.SelectByRay(xFrontFlat, yTopSec - (zDim / 2), -7000, 0, 0, -1, 0.00001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat - (xDim / 2) + (0.001 / SScale), yTopSec + (zDim / 2), -7000, 0, 0, -1, 0.00001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat - (xDim / 2) - 0.015, yTopSec + (zDim / 2) + 0.015, 0)
        Part.ClearSelection2(True)

        'Width - Top Section
        boolstatus = Part.Extension.SelectByRay(xFrontFlat - (xDim / 2), yTopSec, -7000, 0, 0, -1, 0.00001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xFrontFlat + (xDim / 2), yTopSec, -7000, 0, 0, -1, 0.00001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xFrontFlat, yTopSec + (zDim / 2) + 0.015, 0)
        Part.ClearSelection2(True)

        'Height - Right Section
        boolstatus = Part.Extension.SelectByRay(xRightSec, yFrontFlat + (yDim / 2), -7000, 0, 0, -1, 0.00001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xRightSec, yFrontFlat - (yDim / 2), -7000, 0, 0, -1, 0.00001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xRightSec - (zDim / 2) - 0.015, yFrontFlat, 0)
        Part.ClearSelection2(True)

        'Width - Right Section
        boolstatus = Part.Extension.SelectByRay(xRightSec - (zDim / 2), yFrontFlat + (yDim / 2) - (0.001 / SScale), -7000, 0, 0, -1, 0.00001, 1, False, 0, 0)
        boolstatus = Part.Extension.SelectByRay(xRightSec + (zDim / 2), yFrontFlat, -7000, 0, 0, -1, 0.00001, 1, True, 0, 0)
        myDisplayDim = Part.AddDimension2(xRightSec - (zDim / 2) - 0.015, yFrontFlat + (yDim / 2) + 0.015, 0)
        Part.ClearSelection2(True)

        ' Note
        Dim X, Y, Z As Decimal
        boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingBalloonStyle, 0, swBalloonStyle_e.swBS_None)

        X = Round(xDimFlat * SScale * 1000, 2)
        Y = Round(yDimFlat * SScale * 1000, 2)
        Z = Round(zDimFlat * SScale * 1000, 2)
        myNote = Draw.CreateText2(FileName & vbNewLine & "Qty - " & vbNewLine & Z & "mm x " & Y & "mm x " & X & "mm", xRightSec, yTopSec, 0, 0.004, 0)

        ' Save As
        Part.ViewZoomtofit2()
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Drawings\AHU Box\" & FileName & ".SLDDRW", 0, 2)
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Drawings\AHU Box\" & FileName & ".PDF", 0, 2)
        swApp.CloseAllDocuments(True)

    End Sub

#Region "Door Models"

    Public Sub DoorFrontSheet(Width As Decimal, Height As Decimal, Client As String, AHUName As String, JobNo As String)

        'Open
        Part = swApp.OpenDoc6("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Door - Sample\_06_accsess door.sldprt", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("_06_accsess door", False, longstatus)
        Part = swApp.ActiveDoc

        'Re-Dim
        boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_06_accsess door.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 0)
        myDimension = Part.Parameter("Width@BaseFlange")
        myDimension.SystemValue = Width
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Height@BaseFlange@_06_accsess door.SLDPRT", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        myDimension = Part.Parameter("Height@BaseFlange")
        myDimension.SystemValue = Height
        Part.ClearSelection2(True)

        boolstatus = Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Save
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\" & JobNo & "_06_accsess door.SLDPRT", 0, 2)
        Part = Nothing
        StdFunc.CloseActiveDoc()

        'BOM Entries
        bomData.EnterValuesInCNC(JobNo & "_06_accsess door", (Width * 1000) + 42.81, (Height * 1000) + 42.81, "2.0", 1, Client, AHUName, JobNo)

        'Add entry to predictive Database
        Dim name As String = "DoorFrontSheet_" & Convert.ToString(Math.Truncate(Width * 1000)) & "x" & Convert.ToString(Math.Truncate(Height * 1000))
        predictivedb.AHUPartCount(name)

    End Sub

    Public Sub DoorVerticalCSec(Height As Decimal, Client As String, AHUName As String, JobNo As String, DoorSide As String)

        Dim HoleDis As Decimal
        If DoorSide = "RHS" Then
            HoleDis = 0.023
        Else
            HoleDis = 0.027
        End If

        Part = swApp.OpenDoc6("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Door - Sample\_07_door ver support -c.sldprt", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("_07_door ver support -c", False, longstatus)
        Part = swApp.ActiveDoc

        'ReDim Height
        boolstatus = Part.Extension.SelectByID2("Height@BaseFlange@_07_door ver support -c.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 0)
        myDimension = Part.Parameter("Height@BaseFlange")
        myDimension.SystemValue = Height - 0.004
        Part.ClearSelection2(True)

        boolstatus = Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Hole Distance
        boolstatus = Part.Extension.SelectByID2("HingeHole@HingeHoles@_07_door ver support -c.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 0)
        myDimension = Part.Parameter("HingeHole@HingeHoles")
        myDimension.SystemValue = HoleDis
        Part.ClearSelection2(True)

        boolstatus = Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Save
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\" & JobNo & "_07_door ver support -c.SLDPRT", 0, 2)
        Part = Nothing
        StdFunc.CloseActiveDoc()

        'BOM Entries
        bomData.EnterValuesInCNC(JobNo & "_07_door ver support -c", "99.61", (Height * 1000) - 4, "2.0", 1, Client, AHUName, JobNo)

        'Add entry to predictive Database
        Dim name As String = "DoorVerticalCSec_" & Convert.ToString(Math.Truncate(Height * 1000))
        predictivedb.AHUPartCount(name)

    End Sub

    Public Sub DoorHorizontalCSec(Width As Decimal, Client As String, AHUName As String, JobNo As String)

        'Open File
        Part = swApp.OpenDoc6("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Door - Sample\_08_door hor support -c.sldprt", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("_08_door hor support -c", False, longstatus)
        Part = swApp.ActiveDoc

        'Re-Dim
        boolstatus = Part.Extension.SelectByID2("Width@BaseFlange@_08_door hor support -c.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 0)
        myDimension = Part.Parameter("Width@BaseFlange")
        myDimension.SystemValue = Width - 0.1 - 0.004
        Part.ClearSelection2(True)

        boolstatus = Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Save
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\" & JobNo & "_08_door hor support -c.SLDPRT", 0, 2)
        Part = Nothing
        StdFunc.CloseActiveDoc()

        'BOM Entries
        bomData.EnterValuesInCNC(JobNo & "_08_door hor support -c", (Width * 1000) - 104, "99.61", "2.0", 1, Client, AHUName, JobNo)

        'Add entry to predictive Database
        Dim name As String = "DoorHorizontalCSec_" & Convert.ToString(Math.Truncate(Width * 1000))
        predictivedb.AHUPartCount(name)

    End Sub

    Public Sub DoorVerSupportCSec(Height As Decimal, Client As String, AHUName As String, JobNo As String)

        Part = swApp.OpenDoc6("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Door - Sample\_09_hinge fixing-c.sldprt", 1, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2("_09_hinge fixing-c", False, longstatus)
        Part = swApp.ActiveDoc

        boolstatus = Part.Extension.SelectByID2("Height@BaseFlange@_09_hinge fixing-c.SLDPRT", "DIMENSION", 0, 0, 0, True, 0, Nothing, 0)
        myDimension = Part.Parameter("Height@BaseFlange")
        myDimension.SystemValue = Height
        Part.ClearSelection2(True)

        boolstatus = Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Save File
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\" & JobNo & "_09_hinge fixing-c.SLDPRT", 0, 2)
        Part = Nothing
        StdFunc.CloseActiveDoc()

        'BOM Entries
        bomData.EnterValuesInCNC(JobNo & "_09_hinge fixing-c", "105.61", (Height * 1000), "2.0", 1, Client, AHUName, JobNo)

        'Add entry to predictive Database
        Dim name As String = "DoorVerSupportCSec_" & Convert.ToString(Math.Truncate(Height * 1000))
        predictivedb.AHUPartCount(name)

    End Sub

    Public Sub DoorSubAssy(Width As Decimal, Height As Decimal, Client As String, AHUName As String, JobNo As String, DoorPos As String)

        'Open Part Files
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\" & JobNo & "_06_accsess door.sldprt", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\" & JobNo & "_07_door ver support -c.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\" & JobNo & "_08_door hor support -c.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\" & JobNo & "_09_hinge fixing-c.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Door - Sample\_enox hinge.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Door - Sample\_handle knobe.SLDPRT", 1, 0, "", longstatus, longwarnings)

        'New Assembly Document
        Dim assyTemp As String = swApp.GetUserPreferenceStringValue(9)
        Part = swApp.NewDocument(assyTemp, 0, 0, 0)
        swApp.ActivateDoc2("Assem1", False, longstatus)
        Part = swApp.ActiveDoc
        Assy = Part

        'Incert Parts
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\" & JobNo & "_06_accsess door.sldprt", 0, 0, 0.0125)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\" & JobNo & "_07_door ver support -c.SLDPRT", -1, 0.5, 0.5)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\" & JobNo & "_07_door ver support -c.SLDPRT", -0.5, 0.1, 1)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\" & JobNo & "_08_door hor support -c.SLDPRT", 0.4, 0.1, 1)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\" & JobNo & "_08_door hor support -c.SLDPRT", 0.5, 0.1, 1)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\" & JobNo & "_08_door hor support -c.SLDPRT", 0.6, 0.1, 1)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\" & JobNo & "_09_hinge fixing-c.SLDPRT", 1.1, -0.5, 0.8)
        boolstatus = Assy.AddComponent("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Door - Sample\_enox hinge.SLDPRT", 0.8, 0.7, 0.2)
        boolstatus = Assy.AddComponent("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Door - Sample\_enox hinge.SLDPRT", 0.8, 0.7, 0.3)
        boolstatus = Assy.AddComponent("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Door - Sample\_enox hinge.SLDPRT", 0.8, 0.7, 0.4)
        boolstatus = Assy.AddComponent("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Door - Sample\_handle knobe.SLDPRT", -0.8, 0.7, 0.6)
        boolstatus = Assy.AddComponent("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Door - Sample\_handle knobe.SLDPRT", -0.8, 0.7, 0.7)
        boolstatus = Assy.AddComponent("C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Door - Sample\_handle knobe.SLDPRT", -0.8, 0.7, 0.8)

        'Close Part Files
        swApp.CloseAllDocuments(False)

        'Save Assembly Document
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\" & JobNo & "_DoorSubAssy.SLDASM", 0, 2)
        Part = Nothing
        StdFunc.CloseActiveDoc()

        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\" & JobNo & "_DoorSubAssy.SLDASM", 2, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2(JobNo & "_Box Assembly", False, longstatus)
        Part = swApp.ActiveDoc
        Assy = Part

        Part.ViewZoomtofit2()

        'Mates
        'Vertical Section 1
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_06_accsess door-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_07_door ver support -c-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 1, False, 0.025, 0.025, 0.025, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_06_accsess door-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_07_door ver support -c-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0.39309481042652, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_06_accsess door-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_07_door ver support -c-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 1, True, (Width / 2) - 0.025 - 0.002, (Width / 2) - 0.025 - 0.002, (Width / 2) - 0.025 - 0.002, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        'Vertical Section 2
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_06_accsess door-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_07_door ver support -c-2@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 1, False, 0.025, 0.025, 0.025, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_06_accsess door-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_07_door ver support -c-2@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0.39309481042652, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_06_accsess door-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_07_door ver support -c-2@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 1, False, (Width / 2) - 0.025 - 0.002, (Width / 2) - 0.025 - 0.002, (Width / 2) - 0.025 - 0.002, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        Part.ViewZoomtofit2()

        'Horizontal Section - Top
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_06_accsess door-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_08_door hor support -c-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 1, False, 0.025, 0.025, 0.025, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_06_accsess door-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_08_door hor support -c-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 1, False, (Height / 2) - 0.025 - 0.002, (Height / 2) - 0.025 - 0.002, (Height / 2) - 0.025 - 0.002, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_06_accsess door-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_08_door hor support -c-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        'Horizontal Section - Middle
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_06_accsess door-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_08_door hor support -c-2@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 1, False, 0.025, 0.025, 0.025, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_06_accsess door-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_08_door hor support -c-2@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 1, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_06_accsess door-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_08_door hor support -c-2@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        'Horizontal Section - Bottom
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_06_accsess door-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_08_door hor support -c-3@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 1, False, 0.025, 0.025, 0.025, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_06_accsess door-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_08_door hor support -c-3@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 1, True, (Height / 2) - 0.025 - 0.002, (Height / 2) - 0.025 - 0.002, (Height / 2) - 0.025 - 0.002, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_06_accsess door-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_08_door hor support -c-3@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        Part.ViewZoomtofit2()

        'Hinge Channel
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_06_accsess door-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_09_hinge fixing-c-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 1, False, 0.025, 0.025, 0.025, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        If DoorPos = "RHS" Then
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_06_accsess door-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_09_hinge fixing-c-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(5, 1, False, (Width / 2) + 0.005 + 0.025, (Width / 2) + 0.005 + 0.025, (Width / 2) + 0.005 + 0.025, 0.001, 0.001, 0.393094810426527, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)
            Part.EditRebuild3()
        Else
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_06_accsess door-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_09_hinge fixing-c-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(5, 1, True, (Width / 2) + 0.005 + 0.025, (Width / 2) + 0.005 + 0.025, (Width / 2) + 0.005 + 0.025, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)
            Part.EditRebuild3()
        End If

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_06_accsess door-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_09_hinge fixing-c-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0.5, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        Part.ViewZoomtofit2()

        'Hinge 1
        boolstatus = Part.Extension.SelectByID2("Plane1@_enox hinge-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_07_door ver support -c-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        If DoorPos = "RHS" Then
            boolstatus = Part.Extension.SelectByID2("Axis1@_enox hinge-1@" & JobNo & "_DoorSubAssy", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Axis1@" & JobNo & "_07_door ver support -c-2@" & JobNo & "_DoorSubAssy", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)
            Part.EditRebuild3()

            boolstatus = Part.Extension.SelectByID2("Top Plane@_enox hinge-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_06_accsess door-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(3, 0, False, 0, 0, 0, 0.001, 0.001, 0.39309481042652, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)
            Part.EditRebuild3()
        Else
            boolstatus = Part.Extension.SelectByID2("Axis1@_enox hinge-1@" & JobNo & "_DoorSubAssy", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Axis1@" & JobNo & "_07_door ver support -c-1@" & JobNo & "_DoorSubAssy", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)
            Part.EditRebuild3()

            boolstatus = Part.Extension.SelectByID2("Top Plane@_enox hinge-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_06_accsess door-2@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(3, 1, False, 0, 0, 0, 0.001, 0.001, 0.39309481042652, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)
            Part.EditRebuild3()
        End If

        'Hinge 2
        boolstatus = Part.Extension.SelectByID2("Plane1@_enox hinge-2@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_07_door ver support -c-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0.1739975, 0.001, 0.001, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        If DoorPos = "RHS" Then
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_07_door ver support -c-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@_enox hinge-2@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)
            Part.EditRebuild3()
        Else
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_07_door ver support -c-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@_enox hinge-2@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 1, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)
            Part.EditRebuild3()
        End If

        boolstatus = Part.Extension.SelectByID2("Right Plane@_enox hinge-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@_enox hinge-2@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0, 0.001, 0.001, 0.001, 0.001, 0.39309481042652, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        'Hinge 3
        boolstatus = Part.Extension.SelectByID2("Plane1@_enox hinge-3@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_07_door ver support -c-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0.1739975, 0.001, 0.001, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        If DoorPos = "RHS" Then
            boolstatus = Part.Extension.SelectByID2("Top Plane@_enox hinge-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@_enox hinge-3@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(5, 0, True, Height - 0.2, Height - 0.2, Height - 0.2, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)
            Part.EditRebuild3()
        Else
            boolstatus = Part.Extension.SelectByID2("Top Plane@_enox hinge-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@_enox hinge-3@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(5, 0, False, Height - 0.2, Height - 0.2, Height - 0.2, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)
            Part.EditRebuild3()
        End If

        boolstatus = Part.Extension.SelectByID2("Right Plane@_enox hinge-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@_enox hinge-3@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0, 0.001, 0.001, 0.001, 0.001, 0.39309481042652, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        Part.ViewZoomtofit2()

        'Knob 1
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_07_door ver support -c-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Plane1@_handle knobe-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0.584941537399259, 0.001, 0.001, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        If DoorPos = "RHS" Then
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_07_door ver support -c-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@_handle knobe-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(3, 1, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)
            Part.EditRebuild3()

            boolstatus = Part.Extension.SelectByID2("Axis1@" & JobNo & "_07_door ver support -c-1@" & JobNo & "_DoorSubAssy", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Axis1@_handle knobe-1@" & JobNo & "_DoorSubAssy", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)
            Part.EditRebuild3()
        Else
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_07_door ver support -c-2@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@_handle knobe-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(3, 1, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)
            Part.EditRebuild3()

            boolstatus = Part.Extension.SelectByID2("Axis1@" & JobNo & "_07_door ver support -c-2@" & JobNo & "_DoorSubAssy", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Axis1@_handle knobe-1@" & JobNo & "_DoorSubAssy", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)
            Part.EditRebuild3()
        End If

        'Knob 2
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_07_door ver support -c-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Plane1@_handle knobe-2@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0.584941537399259, 0.001, 0.001, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Right Plane@_handle knobe-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@_handle knobe-2@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 1, False, 0.454383520248185, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_07_door ver support -c-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@_handle knobe-2@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 1, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        'Knob 3
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_07_door ver support -c-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Plane1@_handle knobe-3@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0.584941537399259, 0.001, 0.001, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Right Plane@_handle knobe-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@_handle knobe-3@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 1, False, 0.454383520248185, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        boolstatus = Part.Extension.SelectByID2("Top Plane@_handle knobe-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@_handle knobe-3@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(5, 1, True, Height - 0.2, Height - 0.2, Height - 0.2, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()

        Part.ViewZoomtofit2()

        'Save File
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\" & JobNo & "_DoorSubAssy.SLDASM", 0, 2)
        Part = Nothing
        StdFunc.CloseActiveDoc()

    End Sub

#End Region

    Public Sub BoxAHUFinalAssy(BoxWth As Decimal, BoxHt As Decimal, FanNos As Integer, FanHtNo As Integer, FanWthNo As Integer, SideClearDis As Decimal, TopClearDis As Decimal,
                                    Door As String, DoorSide As String, DoorWth As Decimal, DoorHt As Decimal)

        'Top & Side Clearance
        Dim TopClearNo As Integer = Truncate(FanWthNo / 2)
        Dim SideClearNo As Integer = Truncate(FanHtNo / 2)

        'If Door = "YES" Then
        '    SideClearDis = (BoxWth / 2) + ((WallWth - (FanWthNo * BoxWth) - DoorWth) / 2)
        'End If

        'Door
        Dim DoorHtDis As Decimal = (DoorHt - BoxHt) / 2
        Dim DoorWthDis As Decimal = (BoxWth * (FanWthNo - 0.5)) + (DoorWth - 0.025) 'Distance for Hinge Support
        Dim DoorBlankHtDis As Decimal = DoorHt - (BoxHt / 2) + (((BoxHt * FanHtNo) - DoorHt) / 2)
        Dim DoorBlankWthDis As Decimal = (BoxWth * (FanWthNo - 0.5)) + (DoorWth / 2)

        'Top Corner Door
        Dim CornerWth As Decimal
        Dim BoxRem As Integer = (FanWthNo Mod 2)
        If Door = "YES" Then
            If BoxRem > 0 Then
                CornerWth = (DoorWth + BoxWth) / 2
            Else
                CornerWth = DoorWth / 2
            End If
        Else
            If BoxRem > 0 Then
                CornerWth = BoxWth / 2
            Else
                CornerWth = 0
            End If
        End If

        'Open Part Files
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_Box Assembly.SLDASM", 2, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\AHU Box\" & JobNo & "_14_blank sheet_Box.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\AHU Box\" & JobNo & "_14_blank sheet_Side Clearance -1.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\AHU Box\" & JobNo & "_14_blank sheet_Side Clearance -2.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\AHU Box\" & JobNo & "_14_blank sheet_Top Clearance -1.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\AHU Box\" & JobNo & "_14_blank sheet_Top Clearance -2.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\AHU Box\" & JobNo & "_14_blank sheet_Top Corner.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Access Door\" & JobNo & "_DoorSubAssy.SLDASM", 2, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\AHU Box\" & JobNo & "_14_blank sheet_Door Blank.SLDPRT", 1, 0, "", longstatus, longwarnings)
        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_FrameSubAssembly.SLDASM", 2, 0, "", longstatus, longwarnings)

        'New Assembly File
        Dim assyTemp As String = swApp.GetUserPreferenceStringValue(9)
        Part = swApp.NewDocument(assyTemp, 0, 0, 0)
        swApp.ActivateDoc2("Assem4", False, longstatus)
        Part = swApp.ActiveDoc
        Assy = Part

        'Incert Components
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Motor Box\" & JobNo & "_Box Assembly.SLDASM", 0, 0, 0)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\AHU Box\" & JobNo & "_14_blank sheet_Box.SLDPRT", 0.4, 0.4, -0.8)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\AHU Box\" & JobNo & "_14_blank sheet_Side Clearance -1.SLDPRT", 1.3, 0.2, -0.6)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\AHU Box\" & JobNo & "_14_blank sheet_Side Clearance -1.SLDPRT", -1.3, 0.2, -0.6)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\AHU Box\" & JobNo & "_14_blank sheet_Side Clearance -2.SLDPRT", 1.3, 0.8, -0.6)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\AHU Box\" & JobNo & "_14_blank sheet_Side Clearance -2.SLDPRT", -1.3, 0.8, -0.6)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\AHU Box\" & JobNo & "_14_blank sheet_Top Clearance -1.SLDPRT", 1, 0.2, -0.5)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\AHU Box\" & JobNo & "_14_blank sheet_Top Clearance -2.SLDPRT", 2, 0.2, -0.5)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\AHU Box\" & JobNo & "_14_blank sheet_Top Corner.SLDPRT", 0, 1.5, -0.7)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\AHU Box\" & JobNo & "_14_blank sheet_Top Corner.SLDPRT", -2, 1.5, -0.7)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Access Door\" & JobNo & "_DoorSubAssy.SLDASM", 2.2, 0.3, -0.95)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\AHU Box\" & JobNo & "_14_blank sheet_Door Blank.SLDPRT", -0.5, 0.2, -1.1)
        boolstatus = Assy.AddComponent("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Support Structure\" & JobNo & "_FrameSubAssembly.SLDASM", 0, 0.5, -1.3)

        'Save Assembly Document
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\" & JobNo & "_AHU Final Assembly.SLDASM", 0, 2)
        Part = Nothing
        swApp.CloseAllDocuments(True)

        Part = swApp.OpenDoc6("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\" & JobNo & "_AHU Final Assembly.SLDASM", 2, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2(JobNo & "_AHU Final Assembly", False, longstatus)
        Part = swApp.ActiveDoc
        Assy = Part

        Part.ViewZoomtofit2()

        'Create Axis
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

        'Mates
        'Box Pattern
        boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, True, 4, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2(JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
        myFeature = Part.FeatureManager.FeatureLinearPattern2(FanHtNo, BoxHt, FanWthNo, BoxWth, True, False, "NULL", "NULL", False)
        Part.ClearSelection2(True)
        Part.ViewZoomtofit2()

        'Delete Extra Box
        Dim i As Integer = FanHtNo * FanWthNo
        While i > FanNos
            boolstatus = Part.Extension.SelectByID2(JobNo & "_Box Assembly-" & i & "@" & JobNo & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
            Part.EditDelete()
            i -= 1
        End While

        'Box Blanks
        boolstatus = Part.Extension.SelectByID2(JobNo & "_14_blank sheet_Box-1@" & JobNo & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        If boolstatus = True Then
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_14_blank sheet_Box-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(5, 0, False, BoxHt * (FanHtNo - 1), BoxHt * (FanHtNo - 1), BoxHt * (FanHtNo - 1), 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)
            Part.EditRebuild3()

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_14_blank sheet_Box-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(5, 0, False, BoxWth * (FanWthNo - ((FanWthNo * FanHtNo) - FanNos)), BoxWth * (FanWthNo - ((FanWthNo * FanHtNo) - FanNos)), BoxWth * (FanWthNo - ((FanWthNo * FanHtNo) - FanNos)), 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)
            Part.EditRebuild3()

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_14_blank sheet_Box-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)
            Part.EditRebuild3()

            'Box Blank Pattern
            boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(JobNo & "_14_blank sheet_Box-1@" & JobNo & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2((FanWthNo * FanHtNo) - FanNos, BoxWth, 1, 0.05, False, False, "NULL", "NULL", False)
            Part.ClearSelection2(True)
        End If
        Part.ClearSelection2(True)

        'Side Clearance -1-1
        boolstatus = Part.Extension.SelectByID2(JobNo & "_14_blank sheet_Side Clearance -1-1@" & JobNo & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        If boolstatus = True Then
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_14_blank sheet_Side Clearance -1-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(5, 0, False, BoxHt / 2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_14_blank sheet_Side Clearance -1-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(5, 0, True, (BoxWth + SideClearDis) / 2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_14_blank sheet_Side Clearance -1-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            'Side Clearance Pattern
            boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(JobNo & "_14_blank sheet_Side Clearance -1-1@" & JobNo & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2(SideClearNo, BoxHt * 2, 1, 0.05, True, False, "NULL", "NULL", False)
            Part.ClearSelection2(True)

            Part.EditRebuild3()
            Part.ViewZoomtofit2()
        End If

        Part.ClearSelection2(True)

        'Side Clearance -1-2
        boolstatus = Part.Extension.SelectByID2(JobNo & "_14_blank sheet_Side Clearance -1-2@" & JobNo & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        If boolstatus = True Then
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_14_blank sheet_Side Clearance -1-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_14_blank sheet_Side Clearance -1-2@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_Box Assembly-" & FanWthNo & "@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_14_blank sheet_Side Clearance -1-2@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(5, 0, False, (BoxWth + SideClearDis) / 2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_14_blank sheet_Side Clearance -1-2@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            'Side Clearance Pattern
            boolstatus = Part.Extension.SelectByID2("Y-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(JobNo & "_14_blank sheet_Side Clearance -1-2@" & JobNo & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2(SideClearNo, BoxHt * 2, 1, 0.05, True, False, "NULL", "NULL", False)
            Part.ClearSelection2(True)

            Part.EditRebuild3()
            Part.ViewZoomtofit2()
        End If

        Part.ClearSelection2(True)

        'Side Clearance -2-1
        boolstatus = Part.Extension.SelectByID2(JobNo & "_14_blank sheet_Side Clearance -2-1@" & JobNo & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        If boolstatus = True Then
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_14_blank sheet_Side Clearance -2-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            If FanHtNo Mod 2 = 0 Then
                myMate = Assy.AddMate5(5, 0, False, (BoxHt * (FanHtNo - 1.5)) + (TopClearDis / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Else
                myMate = Assy.AddMate5(5, 0, False, BoxHt * (FanHtNo - 1), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            End If
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_14_blank sheet_Side Clearance -1-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_14_blank sheet_Side Clearance -2-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, True, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_14_blank sheet_Side Clearance -1-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_14_blank sheet_Side Clearance -2-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            Part.EditRebuild3()
            Part.ViewZoomtofit2()
        End If

        Part.ClearSelection2(True)

        'Side Clearance -2-2
        boolstatus = Part.Extension.SelectByID2(JobNo & "_14_blank sheet_Side Clearance -2-2@" & JobNo & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        If boolstatus = True Then
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_14_blank sheet_Side Clearance -2-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_14_blank sheet_Side Clearance -2-2@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_14_blank sheet_Side Clearance -1-2@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_14_blank sheet_Side Clearance -2-2@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, True, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_14_blank sheet_Side Clearance -1-2@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_14_blank sheet_Side Clearance -2-2@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            Part.EditRebuild3()
            Part.ViewZoomtofit2()
        End If

        Part.ClearSelection2(True)

        'Top Clearance -1
        boolstatus = Part.Extension.SelectByID2(JobNo & "_14_blank sheet_Top Clearance -1-1@" & JobNo & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        If boolstatus = True Then
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_14_blank sheet_Top Clearance -1-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(5, 0, False, (BoxHt * (FanHtNo - 0.5)) + TopClearDis / 2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_14_blank sheet_Top Clearance -1-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(5, 0, False, BoxWth / 2, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_14_blank sheet_Top Clearance -1-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            'Top Clearance Pattern
            boolstatus = Part.Extension.SelectByID2("X-Axis", "AXIS", 0, 0, 0, False, 2, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2(JobNo & "_14_blank sheet_Top Clearance -1-1@" & JobNo & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
            myFeature = Part.FeatureManager.FeatureLinearPattern2(TopClearNo, BoxWth * 2, 1, 0.05, False, False, "NULL", "NULL", False)
            Part.ClearSelection2(True)

            Part.EditRebuild3()
            Part.ViewZoomtofit2()
        End If

        Part.ClearSelection2(True)

        'Top Clearance -2
        boolstatus = Part.Extension.SelectByID2(JobNo & "_14_blank sheet_Top Clearance -2-1@" & JobNo & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        If boolstatus = True Then
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_14_blank sheet_Top Clearance -1-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_14_blank sheet_Top Clearance -2-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0.001, 0.001, 0.001, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            If Door = "YES" Then
                If FanWthNo Mod 2 = 0 Then
                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_14_blank sheet_Door Blank-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_14_blank sheet_Top Clearance -2-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
                    myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                Else
                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
                    boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_14_blank sheet_Top Clearance -2-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
                    myMate = Assy.AddMate5(5, 0, False, (BoxWth * (FanWthNo - 1)) + ((BoxWth + DoorWth) / 2), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                End If
            Else
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
                boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_14_blank sheet_Top Clearance -2-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
                myMate = Assy.AddMate5(5, 0, False, BoxWth * (FanWthNo - 1), 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            End If
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_14_blank sheet_Top Clearance -1-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_14_blank sheet_Top Clearance -2-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            Part.EditRebuild3()
            Part.ViewZoomtofit2()
        End If

        Part.ClearSelection2(True)

        'Top Corner Blank -1-1
        boolstatus = Part.Extension.SelectByID2(JobNo & "_14_blank sheet_Top Corner-1@" & JobNo & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        If boolstatus = True Then
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_14_blank sheet_Top Clearance -1-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_14_blank sheet_Top Corner-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_14_blank sheet_Side Clearance -1-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_14_blank sheet_Top Corner-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, True, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_14_blank sheet_Side Clearance -1-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_14_blank sheet_Top Corner-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)
        End If

        Part.ClearSelection2(True)

        'Top Corner Blank -1-2
        boolstatus = Part.Extension.SelectByID2(JobNo & "_14_blank sheet_Top Corner-2@" & JobNo & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        If boolstatus = True Then
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_14_blank sheet_Top Corner-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_14_blank sheet_Top Corner-2@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_14_blank sheet_Side Clearance -1-2@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_14_blank sheet_Top Corner-2@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, True, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_14_blank sheet_Top Corner-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_14_blank sheet_Top Corner-2@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
            Part.ClearSelection2(True)
        End If

        Part.ClearSelection2(True)

        Part.EditRebuild3()

        'Door
        If Door = "YES" Then
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_DoorSubAssy-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(5, 0, False, DoorHtDis, DoorHtDis, DoorHtDis, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_DoorSubAssy-1@" & JobNo & "_AHU Final Assembly/" & JobNo & "_09_hinge fixing-c-1@" & JobNo & "_DoorSubAssy", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(5, 1, False, DoorWthDis, DoorWthDis, DoorWthDis, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_DoorSubAssy-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)
        End If
        Part.EditRebuild3()

        'Door Top Blank
        If Door = "YES" Then
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_14_blank sheet_Door Blank-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(5, 0, False, DoorBlankHtDis, DoorBlankHtDis, DoorBlankHtDis, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_14_blank sheet_Door Blank-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(5, 0, False, DoorBlankWthDis, DoorBlankWthDis, DoorBlankWthDis, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_14_blank sheet_Door Blank-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)
        End If
        Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Top Corner Blank - Door
        boolstatus = Part.Extension.SelectByID2(JobNo & "_14_blank sheet_Top Corner Door-1@" & JobNo & "_AHU Final Assembly", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        If boolstatus = True Then
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_14_blank sheet_Top Corner Door-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(5, 0, False, (BoxHt * (FanHtNo - 1)) + TopClearDis, (BoxHt * (FanHtNo - 1)) + TopClearDis, (BoxHt * (FanHtNo - 1)) + TopClearDis, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_14_blank sheet_Door Blank-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_14_blank sheet_Top Corner Door-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_14_blank sheet_Top Corner Door-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
            Part.ClearSelection2(True)

            Part.EditRebuild3()
            Part.ViewZoomtofit2()
        End If

        'Frame Assembly
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Front Plane@" & JobNo & "_FrameSubAssembly-1@" & JobNo & "_AHU Final Assembly/" & JobNo & "_11_Mid_Hor-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 1, False, 0, 0.001, 0.001, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Top Plane@" & JobNo & "_FrameSubAssembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        myMate = Assy.AddMate5(0, 0, False, 0, 0, 0, 0.001, 0.001, 0.39309481042652, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        'boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        'boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_FrameSubAssembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_Box Assembly-1@" & JobNo & "_AHU Final Assembly", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane@" & JobNo & "_FrameSubAssembly-1@" & JobNo & "_AHU Final Assembly/" & JobNo & "_11_Mid_Hor-C-1@" & JobNo & "_FrameSubAssembly", "PLANE", 0, 0, 0, True, 0, Nothing, 0)

        myMate = Assy.AddMate5(0, 1, False, 0, 0, 0, 0.001, 0.001, 0, 0.5235987755983, 0.5235987755983, False, False, 0, longstatus)
        Part.ClearSelection2(True)

        Part.EditRebuild3()
        Part.ViewZoomtofit2()

        'Save Assembly Document
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\" & JobNo & "_AHU Final Assembly.SLDASM", 0, 2)
        Part = Nothing
        StdFunc.CloseActiveDoc()

    End Sub

    Public Sub BoxAHUFinalAssyDrawings()
        'Exit Sub
        ' Variables
        Dim FilePath As String = "C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\" & JobNo & "_AHU Final Assembly.SLDASM"
        Dim FileName As String = JobNo & "_AHU Final Assembly"

        'Open File
        Part = swApp.OpenDoc6(FilePath, 2, 0, "", longstatus, longwarnings)
        swApp.ActivateDoc2(FileName, False, longstatus)
        Part = swApp.ActiveDoc

        Part.ViewZoomtofit2()

        'Bounding Box
        Dim BBox As Object = StdFunc.BoundingBoxOfAssembly()
        Dim xDim As Decimal = Abs(BBox(0)) + Abs(BBox(3))
        Dim yDim As Decimal = Abs(BBox(1)) + Abs(BBox(4))
        Dim zDim As Decimal = Abs(BBox(2)) + Abs(BBox(5))

        swApp.CloseAllDocuments(True)

        ' Sheet Scale
        Dim SScaleX As Integer = Ceiling(xDim / (0.2 - (0.03 + 0.015))) 'Using only 200mm width
        Dim SScaleY As Integer = Ceiling(yDim / (0.21 - (0.03 + 0.03)))
        Dim SScale As Integer = SScaleX
        If SScaleX < SScaleY Then
            SScale = SScaleY
        End If

        ' Adjust Values for Scale
        xDim /= SScale
        yDim /= SScale
        zDim /= SScale

        ' Get Margins
        Dim marginX As Decimal = 0.2 - (xDim + 0.015)
        If marginX < 0.03 Then marginX = 0.03

        Dim marginY As Decimal = (0.21 - (yDim + 0.04 + zDim)) / 2
        If marginY < 0.03 Then marginY = 0.03

        ' Calculate View Placements
        Dim xFront As Decimal = marginX + xDim / 2
        Dim yFront As Decimal = 0.21 / 2

        ' Open Parts
        Part = swApp.OpenDoc6(FilePath, 1, 0, "", longstatus, longwarnings)

        ' Start Drawing
        Part = swApp.NewDocument(DrawTemp, 2, 0.297, 0.21)
        Draw = Part
        boolstatus = Draw.SetupSheet5("Sheet1", 11, 12, 1, SScale, False, DrawSheet, 0.297, 0.21, "Default", False)

        swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swAutomaticScaling3ViewDrawings, False)

        ' Views
        'Front
        myView = Draw.CreateDrawViewFromModelView3(FilePath, "*Front", xFront, yFront, 0)
        boolstatus = Draw.ActivateView("Drawing View1")
        Part.ClearSelection2(True)

        '' Dimentions
        ''Height
        'boolstatus = Part.Extension.SelectByRay(xFront + (xDim / 2) - (0.025 / SScale), yFront + (yDim / 2), -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        'boolstatus = Part.Extension.SelectByRay(xFront, yFront - (yDim / 2), -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        'myDisplayDim = Part.AddDimension2(xFront + (xDim / 2) + 0.015, yFront, 0)
        'Part.ClearSelection2(True)

        ''Width
        'boolstatus = Part.Extension.SelectByRay(xFront + (xDim / 2), yFront, -7000, 0, 0, -1, 0.0001, 1, False, 0, 0)
        'boolstatus = Part.Extension.SelectByRay(xFront - (xDim / 2), yFront + (yDim / 2) - (0.025 / SScale), -7000, 0, 0, -1, 0.0001, 1, True, 0, 0)
        'myDisplayDim = Part.AddDimension2(xFront, yFront + (yDim / 2) + 0.015, 0)
        'Part.ClearSelection2(True)

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
        boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingBalloonStyle, 0, swBalloonStyle_e.swBS_None)

        myNote = Draw.CreateText2(FileName, xFront, 0.02, 0, 0.004, 0)

        ' Save As
        Part.ViewZoomtofit2()
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Drawings\Support Structure\" & JobNo & "_Support Structure.SLDDRW", 0, 2)
        longstatus = Part.SaveAs3("C:\AHU Automation - Output\" & Client & "\" & AHUName & "\" & JobNo & "\Drawings\Support Structure\" & JobNo & "_Support Structure.PDF", 0, 2)
        swApp.CloseAllDocuments(True)

    End Sub

End Class