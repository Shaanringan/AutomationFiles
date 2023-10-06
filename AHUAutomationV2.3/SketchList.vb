Imports SolidWorks.Interop.sldworks

Public Class SketchList

    Dim swApp As New SldWorks

    Dim Part As ModelDoc2
    Dim boolstatus As Boolean
    Dim myFeature As Object
    Dim myBlockDefinition As SketchBlockDefinition

    Sub FanCutout(FanDia As Integer, FanNos As Integer, CutoutXDis As Decimal(), CutoutYDis As Decimal)

        Part = swApp.ActiveDoc

        Dim BlockUtil As MathUtility
        Dim Blockpoint As MathPoint
        Dim BlockPath As String = "C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Motor Cutout\" & FanDia & "_motor plate cutout.SLDBLK"
        Dim ArData(2) As Double

        For i = 0 To FanNos - 1
            ArData(0) = CutoutXDis(i)
            ArData(1) = CutoutYDis
            ArData(2) = 0

            BlockUtil = swApp.GetMathUtility
            Blockpoint = BlockUtil.CreatePoint(ArData)
            myBlockDefinition = Part.SketchManager.MakeSketchBlockFromFile(Blockpoint, BlockPath, False, 1, 0)
        Next

        Part.ViewOrientationUndo()

    End Sub

    Function IdentifierJobNo(ByVal jNo As String) As Char

        Dim IdentFirst As Char

        Dim temp As Char() = jNo.ToCharArray

        IdentFirst = temp(UBound(temp))

        Return IdentFirst

    End Function

    Sub Identifier(ByVal Symbol As Char, ByVal CutoutXDis As Decimal, ByVal CutoutYDis As Decimal, ByVal RotAngle As Double)

        Part = swApp.ActiveDoc

        Dim BlockUtil As MathUtility
        Dim Blockpoint As MathPoint
        Dim BlockPath As String = "C:\Program Files (x86)\Crescent Engineering\Automation\AHU - Library\Identifing Symbols\" & Symbol & ".SLDBLK"

        Dim ArData(2) As Double
        ArData(0) = CutoutXDis
        ArData(1) = CutoutYDis
        ArData(2) = 0

        BlockUtil = swApp.GetMathUtility
        Blockpoint = BlockUtil.CreatePoint(ArData)
        myBlockDefinition = Part.SketchManager.MakeSketchBlockFromFile(Blockpoint, BlockPath, False, 1, RotAngle)

        Part.ViewOrientationUndo()

    End Sub

End Class