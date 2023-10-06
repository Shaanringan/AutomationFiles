Imports System.Collections.Generic
Imports MySql.Data.MySqlClient

Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Imports SolidWorksTools.File

Public Class Standard_Functions

    'SolidWorks Variables
    Dim swApp As New SldWorks

    Dim Part As ModelDoc2
    Dim Assy As AssemblyDoc

    'Connect to MySQL Server
    Dim MysqlConn As MySqlConnection
    Dim query As String
    Dim command As MySqlCommand
    Dim reader As MySqlDataReader

#Region "Database"

    Public Function GetFromTable(ByVal ToGetColumnName As String, ByVal TableName As String, ByVal ReffColName As String, ByVal ReffColEntry As String) As String

        'Connect to MySQL
        MysqlConn = New MySqlConnection With {.ConnectionString = "server = 127.0.0.1; userid = root; password = ; database = aadtech_erp_db"}
        MysqlConn.Open()

        'Get Value
        query = "SELECT " & ToGetColumnName & " FROM " & TableName & " WHERE " & ReffColName & " = '" & ReffColEntry & "'"
        command = New MySqlCommand(query, MysqlConn)
        Dim strVal As String = command.ExecuteScalar.ToString

        'Close MySQL Connection
        MysqlConn.Close()
        MysqlConn.Dispose()

        Return strVal

    End Function

    Function GetRowCount(ByVal TableName As String) As Integer

        'Connect to MySQL
        MysqlConn = New MySqlConnection With {.ConnectionString = "server = 127.0.0.1; userid = root; password = ; database = aadtech_erp_db"}
        MysqlConn.Open()

        'System Test
        query = "SELECT COUNT(*) FROM " & TableName
        command = New MySqlCommand(query, MysqlConn)
        Dim Count As Integer = command.ExecuteScalar

        'Close MySQL Connection
        MysqlConn.Close()
        MysqlConn.Dispose()

        Return Count

    End Function

    Function GetColumnArray(ByVal ColumnName As String, ByVal TableName As String) As String()

        'Connect to MySQL
        MysqlConn = New MySqlConnection With {.ConnectionString = "server = 127.0.0.1; userid = root; password = ; database = aadtech_erp_db"}
        MysqlConn.Open()

        'System Test
        query = "SELECT " & ColumnName & " FROM " & TableName
        command = New MySqlCommand(query, MysqlConn)
        reader = command.ExecuteReader

        Dim valList As New List(Of String)
        While reader.Read
            valList.Add(reader(ColumnName).ToString)
        End While
        Dim valReturn() As String = valList.ToArray

        'Close MySQL Connection
        MysqlConn.Close()
        MysqlConn.Dispose()

        Return valReturn

    End Function

    Function GetColumnArray(ByVal ColumnName As String, ByVal TableName As String, ByVal ReffColName As String, ByVal ReffColEntry As String) As String()

        'Connect to MySQL
        MysqlConn = New MySqlConnection With {.ConnectionString = "server = 127.0.0.1; userid = root; password = ; database = aadtech_erp_db"}
        MysqlConn.Open()

        'System Test
        query = "SELECT " & ColumnName & " FROM " & TableName & " WHERE " & ReffColName & " = '" & ReffColEntry & "'"
        command = New MySqlCommand(query, MysqlConn)
        reader = command.ExecuteReader

        Dim valList As New List(Of String)
        While reader.Read
            valList.Add(reader(ColumnName).ToString)
        End While
        Dim valReturn() As String = valList.ToArray

        'Close MySQL Connection
        MysqlConn.Close()
        MysqlConn.Dispose()

        Return valReturn

    End Function

    Function JobNo() As String

        'Connect to MySQL
        MysqlConn = New MySqlConnection With {.ConnectionString = "server = 127.0.0.1; userid = root; password = ; database = aadtech_erp_db"}
        MysqlConn.Open()

        'Get Value
        query = "SELECT job_no FROM `counter_master_data_new` ORDER BY job_no DESC LIMIT 1"
        command = New MySqlCommand(query, MysqlConn)
        Dim CurrJobNo As String = command.ExecuteScalar.ToString

        Return CurrJobNo

    End Function

#End Region

#Region "SolidWorks"

    Public Sub CloseActiveDoc()

        Part = swApp.ActiveDoc

        Dim fileName As String
        fileName = Part.GetTitle()
        swApp.CloseDoc(fileName)

    End Sub

    Public Sub PackAndGo(ByVal OutputFolder As String, ByVal PreFix As String)

        Dim swPackAndGo As PackAndGo
        Dim pgFileNames As Object
        Dim pgFileStatus As Object
        Dim pgGetFileNames As Object
        Dim pgDocumentStatus As Object
        Dim status As Boolean
        Dim namesCount As Long
        Dim myPath As String
        Dim statuses As Object

        ' Open assembly
        Part = swApp.ActiveDoc 'ActivateDoc(openFile)

        ' Get Pack and Go object
        swPackAndGo = Part.Extension.GetPackAndGo

        ' Get number of documents in assembly
        namesCount = swPackAndGo.GetDocumentNamesCount

        ' Include any drawings, SOLIDWORKS Simulation results, and SOLIDWORKS Toolbox components
        swPackAndGo.IncludeDrawings = True
        swPackAndGo.IncludeSimulationResults = True
        swPackAndGo.IncludeToolboxComponents = True

        ' Get current paths and filenames of the assembly's documents
        status = swPackAndGo.GetDocumentNames(pgFileNames)

        ' Get current save-to paths and filenames of the assembly's documents
        status = swPackAndGo.GetDocumentSaveToNames(pgFileNames, pgFileStatus)

        ' Set folder where to save the files
        myPath = OutputFolder
        status = swPackAndGo.SetSaveToName(True, myPath)

        ' Flatten the Pack and Go folder structure; save all files to the root directory
        swPackAndGo.FlattenToSingleFolder = True

        ' Add a prefix and suffix to the new Pack and Go filenames
        swPackAndGo.AddPrefix = PreFix & "_"
        'swPackAndGo.AddSuffix = "_PackAndGo"

        ' Verify document paths and filenames after adding prefix and suffix
        ReDim pgGetFileNames(namesCount - 1)
        ReDim pgDocumentStatus(namesCount - 1)
        status = swPackAndGo.GetDocumentSaveToNames(pgGetFileNames, pgDocumentStatus)

        ' Pack and Go
        statuses = Part.Extension.SavePackAndGo(swPackAndGo)

    End Sub

    Function SheetScale(ByVal X As Decimal, ByVal Y As Decimal, ByVal Z As Decimal, ByVal FlatX As Decimal, ByVal SheetX As Decimal, ByVal SheetY As Decimal, ByVal Parts As Integer) As Integer

        'Actual Width of Views
        Dim SizeActualX As Decimal = Y + X + X
        Dim SizeActualY As Decimal = Y + Z

        Dim SizeFlatX As Decimal = FlatX * Parts

        If SizeFlatX > SizeActualX Then
            SizeActualX = SizeFlatX
        End If

        'Adjusted Sheet Size Required
        If SizeFlatX > SizeActualX Then
            SheetX -= 0.06 + 0.05 * Parts
        Else
            SheetX -= 0.06 + 0.05 + 0.05 + 0.05
        End If
        SheetY -= 0.06 + 0.05 + 0.05

        'Calculate Scale
        Dim ScaleX As Decimal = SizeActualX / SheetX
        Dim ScaleY As Decimal = SizeActualY / SheetY

        'Set Sheet Scale
        Dim Scale As Decimal = ScaleX
        If ScaleY > ScaleX Then
            Scale = ScaleY
        End If

        Return Math.Ceiling(Scale)

    End Function

    Function ShtScale(X As Decimal, Z As Decimal, SheetX As Decimal, SheetY As Decimal) As Integer

        'Actual Width of Views
        Dim SizeActualX As Decimal = X
        Dim SizeActualY As Decimal = Z

        'Adjusted Sheet Size Required
        SheetX -= 0.02 + 0.06 + 0.06 + 0.02
        SheetY -= 0.02 + 0.06 + 0.06 + 0.02

        'Calculate Scale
        Dim ScaleX As Decimal = SizeActualX / SheetX
        Dim ScaleY As Decimal = SizeActualY / SheetY

        'Set Sheet Scale
        Dim Scale As Decimal = ScaleX
        If ScaleY > ScaleX Then
            Scale = ScaleY
        End If

        Return Math.Ceiling(Scale)

    End Function


    Function SheetScale(ByVal X As Decimal, ByVal Y As Decimal, ByVal Z As Decimal, ByVal SheetX As Decimal, ByVal SheetY As Decimal) As Integer

        'Actual Width of Views
        Dim SizeActualX As Decimal = 0.1 + Y + 0.05 + X + 0.05 + Y + 0.05 + X + 0.1
        Dim SizeActualY As Decimal = 0.1 + Y + 0.05 + Z + 0.05 + Y + 0.1

        'Calculate Scale
        Dim ScaleX As Decimal = SizeActualX / SheetX
        Dim ScaleY As Decimal = SizeActualY / SheetY

        'Set Sheet Scale
        Dim Scale As Decimal = ScaleX
        If ScaleY > ScaleX Then
            Scale = ScaleY
        End If

        Return Math.Ceiling(Scale)

    End Function

    Function SheetScaleMicaWorksheet(ByVal X As Decimal, ByVal Y As Decimal, ByVal Z As Decimal, ByVal FlatX As Decimal, ByVal SheetX As Decimal, ByVal SheetY As Decimal, ByVal Parts As Integer) As Integer

        'Actual Width of Views
        Dim SizeActualX As Decimal = Y + X + X
        Dim SizeActualY As Decimal = Z

        Dim SizeFlatX As Decimal = 0.06 + (FlatX + 0.05) * Parts

        If SizeFlatX > SizeActualX Then
            SizeActualX = SizeFlatX
        End If

        'Adjusted Sheet Size Required
        SheetX -= 0.06 + 0.05 + 0.05 + 0.05
        SheetY -= 0.06 + 0.05

        'Calculate Scale
        Dim ScaleX As Decimal = SizeActualX / SheetX
        Dim ScaleY As Decimal = SizeActualY / SheetY

        'Set Sheet Scale
        Dim Scale As Decimal = ScaleX
        If ScaleY > ScaleX Then
            Scale = ScaleY
        End If

        Return Math.Ceiling(Scale)

    End Function

    Function BoundingBox() As Object

        Dim swModel As ModelDoc2
        Dim swPart As PartDoc
        Dim swBody As Body2
        Dim vBodies As Object
        Dim vPts As Object
        Dim j As Integer

        swModel = swApp.ActiveDoc
        swPart = swModel

        'Get all bodies in the part
        vBodies = swPart.GetBodies2(swBodyType_e.swAllBodies, False)

        'Traverse all bodies
        For i = 0 To UBound(vBodies)
            swBody = vBodies(i)
            If swBody.Name.Chars(0) <> "<" Then
                Exit For
            End If
        Next
        vPts = swBody.GetBodyBox

        'Round all values
        For j = 0 To UBound(vPts)
            vPts(j) = Decimal.Round(vPts(j), 5)
        Next j

        Return vPts

    End Function

    Function BoundingBoxOfPart() As Decimal()

        Dim swModel As ModelDoc2
        Dim swPart As PartDoc
        Dim swBody As Body2
        Dim vBodies As Object
        Dim vPts As Object

        swModel = swApp.ActiveDoc
        swPart = swModel

        'Get all bodies in the part
        vBodies = swPart.GetBodies2(swBodyType_e.swAllBodies, False)

        'Traverse all bodies
        Dim xMax, yMax, zMax, xMin, yMin, zMin As Decimal
        For i = 0 To UBound(vBodies)
            swBody = vBodies(i)
            vPts = swBody.GetBodyBox
            If xMax < vPts(0) Then
                xMax = vPts(0)
            End If
            If yMax < vPts(1) Then
                yMax = vPts(1)
            End If
            If zMax < vPts(2) Then
                zMax = vPts(2)
            End If
            If xMin > vPts(3) Then
                xMin = vPts(3)
            End If
            If yMin > vPts(4) Then
                yMin = vPts(4)
            End If
            If zMin > vPts(5) Then
                zMin = vPts(5)
            End If
        Next

        Dim maxBoxSize As Decimal() = {xMax, yMax, zMax, xMin, yMin, zMin}

        Return maxBoxSize

    End Function

    Function BoundingBoxOfAssembly() As Object

        Dim vBox As Object

        Part = swApp.ActiveDoc
        Assy = Part

        vBox = Assy.GetBox(swBoundingBoxOptions_e.swBoundingBoxIncludeSketches)

        Return vBox

    End Function

#End Region

#Region "General"

    Public Sub FolderCreationPanel(ClientName As String, JobNo As String)

        If Not IO.Directory.Exists("C:\AHU Automation - Output") Then
            IO.Directory.CreateDirectory("C:\AHU Automation - Output")
        End If

        If Not IO.Directory.Exists("C:\AHU Automation - Output\" & ClientName) Then
            IO.Directory.CreateDirectory("C:\AHU Automation - Output\" & ClientName)
        End If

        If Not IO.Directory.Exists("C:\AHU Automation - Output\" & ClientName & "\" & JobNo) Then
            IO.Directory.CreateDirectory("C:\AHU Automation - Output\" & ClientName & "\" & JobNo)
        End If

        If Not IO.Directory.Exists("C:\AHU Automation - Output\" & ClientName & "\" & JobNo & "\AHU") Then
            IO.Directory.CreateDirectory("C:\AHU Automation - Output\" & ClientName & "\" & JobNo & "\AHU")
        End If

        If Not IO.Directory.Exists("C:\AHU Automation - Output\" & ClientName & "\" & JobNo & "\PDF Drawings") Then
            IO.Directory.CreateDirectory("C:\AHU Automation - Output\" & ClientName & "\" & JobNo & "\PDF Drawings")
        End If

        If Not IO.Directory.Exists("C:\AHU Automation - Output\" & ClientName & "\" & JobNo & "\eDrawings") Then
            IO.Directory.CreateDirectory("C:\AHU Automation - Output\" & ClientName & "\" & JobNo & "\eDrawings")
        End If

    End Sub

    Public Sub FolderCreationBox(ClientName As String, AHUName As String, JobNo As String)

        If Not IO.Directory.Exists("C:\AHU Automation - Output") Then
            IO.Directory.CreateDirectory("C:\AHU Automation - Output")
        End If

        If Not IO.Directory.Exists("C:\AHU Automation - Output\" & ClientName) Then
            IO.Directory.CreateDirectory("C:\AHU Automation - Output\" & ClientName)
        End If

        If Not IO.Directory.Exists("C:\AHU Automation - Output\" & ClientName & "\" & AHUName) Then
            IO.Directory.CreateDirectory("C:\AHU Automation - Output\" & ClientName & "\" & AHUName)
        End If

        If Not IO.Directory.Exists("C:\AHU Automation - Output\" & ClientName & "\" & AHUName & "\" & JobNo) Then
            IO.Directory.CreateDirectory("C:\AHU Automation - Output\" & ClientName & "\" & AHUName & "\" & JobNo)
        End If

        If Not IO.Directory.Exists("C:\AHU Automation - Output\" & ClientName & "\" & AHUName & "\" & JobNo & "\Access Door") Then
            IO.Directory.CreateDirectory("C:\AHU Automation - Output\" & ClientName & "\" & AHUName & "\" & JobNo & "\Access Door")
        End If

        If Not IO.Directory.Exists("C:\AHU Automation - Output\" & ClientName & "\" & AHUName & "\" & JobNo & "\AHU Box") Then
            IO.Directory.CreateDirectory("C:\AHU Automation - Output\" & ClientName & "\" & AHUName & "\" & JobNo & "\AHU Box")
        End If

        If Not IO.Directory.Exists("C:\AHU Automation - Output\" & ClientName & "\" & AHUName & "\" & JobNo & "\Drawings") Then
            IO.Directory.CreateDirectory("C:\AHU Automation - Output\" & ClientName & "\" & AHUName & "\" & JobNo & "\Drawings")
            IO.Directory.CreateDirectory("C:\AHU Automation - Output\" & ClientName & "\" & AHUName & "\" & JobNo & "\Drawings\Access Door")
            IO.Directory.CreateDirectory("C:\AHU Automation - Output\" & ClientName & "\" & AHUName & "\" & JobNo & "\Drawings\AHU Box")
            IO.Directory.CreateDirectory("C:\AHU Automation - Output\" & ClientName & "\" & AHUName & "\" & JobNo & "\Drawings\Motor Box")
            IO.Directory.CreateDirectory("C:\AHU Automation - Output\" & ClientName & "\" & AHUName & "\" & JobNo & "\Drawings\Support Structure")
        End If

        If Not IO.Directory.Exists("C:\AHU Automation - Output\" & ClientName & "\" & AHUName & "\" & JobNo & "\Motor") Then
            IO.Directory.CreateDirectory("C:\AHU Automation - Output\" & ClientName & "\" & AHUName & "\" & JobNo & "\Motor")
        End If

        If Not IO.Directory.Exists("C:\AHU Automation - Output\" & ClientName & "\" & AHUName & "\" & JobNo & "\Motor Box") Then
            IO.Directory.CreateDirectory("C:\AHU Automation - Output\" & ClientName & "\" & AHUName & "\" & JobNo & "\Motor Box")
        End If

        If Not IO.Directory.Exists("C:\AHU Automation - Output\" & ClientName & "\" & AHUName & "\" & JobNo & "\Support Structure") Then
            IO.Directory.CreateDirectory("C:\AHU Automation - Output\" & ClientName & "\" & AHUName & "\" & JobNo & "\Support Structure")
        End If

    End Sub

#End Region

End Class