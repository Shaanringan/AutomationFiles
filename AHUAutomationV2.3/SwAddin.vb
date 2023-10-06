Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Imports SolidWorks.Interop.swpublished
Imports SolidWorksTools
Imports SolidWorksTools.File

Imports System.Net.NetworkInformation

<Guid("127edebc-9f80-4ef1-a3dc-2a95683cb178")> _
<ComVisible(True)>
<SwAddin(
        Description:="AHUAutomationV2.3 description",
        Title:="AHUAutomationV2.3",
        LoadAtStartup:=True
        )>
Public Class SwAddin
    Implements SolidWorks.Interop.swpublished.SwAddin

#Region "Local Variables"
    Dim WithEvents iSwApp As SldWorks
    Dim iCmdMgr As ICommandManager
    Dim addinID As Integer
    Dim openDocs As Hashtable
    Dim SwEventPtr As SldWorks
    Dim ppage As UserPMPage
    Dim iBmp As BitmapHandler
    Dim frame As IFrame
    Dim bRet As Boolean
    Dim registerID As Integer

    Public Const mainCmdGroupID As Integer = 0
    Public Const mainItemID1 As Integer = 0
    Public Const mainItemID2 As Integer = 1
    Public Const flyoutGroupID As Integer = 91

    ' Public Properties
    ReadOnly Property SwApp() As SldWorks
        Get
            Return iSwApp
        End Get
    End Property

    ReadOnly Property CmdMgr() As ICommandManager
        Get
            Return iCmdMgr
        End Get
    End Property

    ReadOnly Property OpenDocumentsTable() As Hashtable
        Get
            Return openDocs
        End Get
    End Property
#End Region

#Region "SolidWorks Registration"

    <ComRegisterFunction()> Public Shared Sub RegisterFunction(ByVal t As Type)

        ' Get Custom Attribute: SwAddinAttribute
        Dim attributes() As Object
        Dim SWattr As SwAddinAttribute = Nothing

        attributes = System.Attribute.GetCustomAttributes(GetType(SwAddin), GetType(SwAddinAttribute))

        If attributes.Length > 0 Then
            SWattr = DirectCast(attributes(0), SwAddinAttribute)
        End If
        Try
            Dim hklm As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine
            Dim hkcu As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser

            Dim keyname As String = "SOFTWARE\SolidWorks\Addins\{" + t.GUID.ToString() + "}"
            Dim addinkey As Microsoft.Win32.RegistryKey = hklm.CreateSubKey(keyname)
            addinkey.SetValue(Nothing, 0)
            addinkey.SetValue("Description", SWattr.Description)
            addinkey.SetValue("Title", SWattr.Title)

            keyname = "Software\SolidWorks\AddInsStartup\{" + t.GUID.ToString() + "}"
            addinkey = hkcu.CreateSubKey(keyname)
            addinkey.SetValue(Nothing, SWattr.LoadAtStartup, Microsoft.Win32.RegistryValueKind.DWord)
        Catch nl As System.NullReferenceException
            Console.WriteLine("There was a problem registering this dll: SWattr is null.\n " & nl.Message)
            System.Windows.Forms.MessageBox.Show("There was a problem registering this dll: SWattr is null.\n" & nl.Message)
        Catch e As System.Exception
            Console.WriteLine("There was a problem registering this dll: " & e.Message)
            System.Windows.Forms.MessageBox.Show("There was a problem registering this dll: " & e.Message)
        End Try
    End Sub

    <ComUnregisterFunction()> Public Shared Sub UnregisterFunction(ByVal t As Type)
        Try
            Dim hklm As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine
            Dim hkcu As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser

            Dim keyname As String = "SOFTWARE\SolidWorks\Addins\{" + t.GUID.ToString() + "}"
            hklm.DeleteSubKey(keyname)

            keyname = "Software\SolidWorks\AddInsStartup\{" + t.GUID.ToString() + "}"
            hkcu.DeleteSubKey(keyname)
        Catch nl As System.NullReferenceException
            Console.WriteLine("There was a problem unregistering this dll: SWattr is null.\n " & nl.Message)
            System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: SWattr is null.\n" & nl.Message)
        Catch e As System.Exception
            Console.WriteLine("There was a problem unregistering this dll: " & e.Message)
            System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: " & e.Message)
        End Try

    End Sub

#End Region

#Region "ISwAddin Implementation"

    Function ConnectToSW(ByVal ThisSW As Object, ByVal Cookie As Integer) As Boolean Implements SolidWorks.Interop.swpublished.SwAddin.ConnectToSW
        iSwApp = ThisSW
        addinID = Cookie

        ' Setup callbacks
        iSwApp.SetAddinCallbackInfo(0, Me, addinID)

        ' Setup the Command Manager
        iCmdMgr = iSwApp.GetCommandManager(Cookie)
        AddCommandMgr()

        'Setup the Event Handlers
        SwEventPtr = iSwApp
        openDocs = New Hashtable
        AttachEventHandlers()

        'Setup Sample Property Manager
        AddPMP()

        ConnectToSW = True
    End Function

    Function DisconnectFromSW() As Boolean Implements SolidWorks.Interop.swpublished.SwAddin.DisconnectFromSW

        RemoveCommandMgr()
        RemovePMP()
        DetachEventHandlers()

        System.Runtime.InteropServices.Marshal.ReleaseComObject(iCmdMgr)
        iCmdMgr = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(iSwApp)
        iSwApp = Nothing
        'The addin _must_ call GC.Collect() here in order to retrieve all managed code pointers 
        GC.Collect()
        GC.WaitForPendingFinalizers()

        GC.Collect()
        GC.WaitForPendingFinalizers()

        DisconnectFromSW = True
    End Function
#End Region

#Region "UI Methods"

    Public Sub AddCommandMgr()

        Dim cmdGroup As ICommandGroup

        If iBmp Is Nothing Then
            iBmp = New BitmapHandler()
        End If

        Dim thisAssembly As Assembly

        Dim cmdIndex0 As Integer
        Dim Title As String = "AHU Automation V2.3"
        Dim ToolTip As String = "AHU Automation V2.3"

        Dim docTypes() As Integer = {swDocumentTypes_e.swDocASSEMBLY, swDocumentTypes_e.swDocDRAWING, swDocumentTypes_e.swDocPART}

        thisAssembly = System.Reflection.Assembly.GetAssembly(Me.GetType())

        Dim cmdGroupErr As Integer = 0
        Dim ignorePrevious As Boolean = False

        Dim registryIDs As Object = Nothing
        Dim getDataResult As Boolean = iCmdMgr.GetGroupDataFromRegistry(mainCmdGroupID, registryIDs)

        Dim knownIDs As Integer() = New Integer(1) {mainItemID1, mainItemID2}

        If getDataResult Then
            If Not CompareIDs(registryIDs, knownIDs) Then 'if the IDs don't match, reset the commandGroup
                ignorePrevious = True
            End If
        End If

        cmdGroup = iCmdMgr.CreateCommandGroup2(mainCmdGroupID, Title, ToolTip, "", -1, ignorePrevious, cmdGroupErr)
        If cmdGroup Is Nothing Or thisAssembly Is Nothing Then
            Throw New NullReferenceException()
        End If

        ' Add bitmaps to your project and set them as embedded resources or provide a direct path to the bitmaps
        Dim mainIcons(6) As String
        Dim icons(6) As String
        icons(0) = iBmp.CreateFileFromResourceBitmap("AHUAutomationV2._3.toolbar20x.png", thisAssembly)
        icons(1) = iBmp.CreateFileFromResourceBitmap("AHUAutomationV2._3.toolbar32x.png", thisAssembly)
        icons(2) = iBmp.CreateFileFromResourceBitmap("AHUAutomationV2._3.toolbar40x.png", thisAssembly)
        icons(3) = iBmp.CreateFileFromResourceBitmap("AHUAutomationV2._3.toolbar64x.png", thisAssembly)
        icons(4) = iBmp.CreateFileFromResourceBitmap("AHUAutomationV2._3.toolbar96x.png", thisAssembly)
        icons(5) = iBmp.CreateFileFromResourceBitmap("AHUAutomationV2._3.toolbar128x.png", thisAssembly)

        mainIcons(0) = iBmp.CreateFileFromResourceBitmap("AHUAutomationV2._3.mainicon_20.png", thisAssembly)
        mainIcons(1) = iBmp.CreateFileFromResourceBitmap("AHUAutomationV2._3.mainicon_32.png", thisAssembly)
        mainIcons(2) = iBmp.CreateFileFromResourceBitmap("AHUAutomationV2._3.mainicon_40.png", thisAssembly)
        mainIcons(3) = iBmp.CreateFileFromResourceBitmap("AHUAutomationV2._3.mainicon_64.png", thisAssembly)
        mainIcons(4) = iBmp.CreateFileFromResourceBitmap("AHUAutomationV2._3.mainicon_96.png", thisAssembly)
        mainIcons(5) = iBmp.CreateFileFromResourceBitmap("AHUAutomationV2._3.mainicon_128.png", thisAssembly)

        cmdGroup.IconList = icons
        cmdGroup.MainIconList = mainIcons

        Dim menuToolbarOption As Integer = swCommandItemType_e.swMenuItem Or swCommandItemType_e.swToolbarItem

        cmdIndex0 = cmdGroup.AddCommandItem2("AHUDesign", -1, "AHU Design Automation", "AHU Design", 0, "AHUDesign", "", mainItemID1, menuToolbarOption)

        cmdGroup.HasToolbar = True
        cmdGroup.HasMenu = True
        cmdGroup.Activate()

        thisAssembly = Nothing

    End Sub

    Public Sub RemoveCommandMgr()
        Try
            iBmp.Dispose()
            iCmdMgr.RemoveCommandGroup(mainCmdGroupID)
            iCmdMgr.RemoveFlyoutGroup(flyoutGroupID)
        Catch e As Exception
        End Try
    End Sub

    Function AddPMP() As Boolean
        ppage = New UserPMPage
        ppage.Init(iSwApp, Me)
    End Function

    Function RemovePMP() As Boolean
        ppage = Nothing
    End Function

    Function CompareIDs(ByVal storedIDs() As Integer, ByVal addinIDs() As Integer) As Boolean

        Dim storeList As New List(Of Integer)(storedIDs)
        Dim addinList As New List(Of Integer)(addinIDs)

        addinList.Sort()
        storeList.Sort()

        If Not addinList.Count = storeList.Count Then

            Return False
        Else

            For i As Integer = 0 To addinList.Count - 1
                If Not addinList(i) = storeList(i) Then

                    Return False
                End If
            Next
        End If

        Return True
    End Function
#End Region

#Region "Event Methods"
    Sub AttachEventHandlers()
        AttachSWEvents()

        'Listen for events on all currently open docs
        AttachEventsToAllDocuments()
    End Sub

    Sub DetachEventHandlers()
        DetachSWEvents()

        'Close events on all currently open docs
        Dim docHandler As DocumentEventHandler
        Dim key As ModelDoc2
        Dim numKeys As Integer
        numKeys = openDocs.Count
        If numKeys > 0 Then
            Dim keys() As Object = New Object(numKeys - 1) {}

            'Remove all document event handlers
            openDocs.Keys.CopyTo(keys, 0)
            For Each key In keys
                docHandler = openDocs.Item(key)
                docHandler.DetachEventHandlers() 'This also removes the pair from the hash
                docHandler = Nothing
                key = Nothing
            Next
        End If
    End Sub

    Sub AttachSWEvents()
        Try
            AddHandler iSwApp.ActiveDocChangeNotify, AddressOf Me.SldWorks_ActiveDocChangeNotify
            AddHandler iSwApp.DocumentLoadNotify2, AddressOf Me.SldWorks_DocumentLoadNotify2
            AddHandler iSwApp.FileNewNotify2, AddressOf Me.SldWorks_FileNewNotify2
            AddHandler iSwApp.ActiveModelDocChangeNotify, AddressOf Me.SldWorks_ActiveModelDocChangeNotify
            AddHandler iSwApp.FileOpenPostNotify, AddressOf Me.SldWorks_FileOpenPostNotify
        Catch e As Exception
            Console.WriteLine(e.Message)
        End Try
    End Sub

    Sub DetachSWEvents()
        Try
            RemoveHandler iSwApp.ActiveDocChangeNotify, AddressOf Me.SldWorks_ActiveDocChangeNotify
            RemoveHandler iSwApp.DocumentLoadNotify2, AddressOf Me.SldWorks_DocumentLoadNotify2
            RemoveHandler iSwApp.FileNewNotify2, AddressOf Me.SldWorks_FileNewNotify2
            RemoveHandler iSwApp.ActiveModelDocChangeNotify, AddressOf Me.SldWorks_ActiveModelDocChangeNotify
            RemoveHandler iSwApp.FileOpenPostNotify, AddressOf Me.SldWorks_FileOpenPostNotify
        Catch e As Exception
            Console.WriteLine(e.Message)
        End Try
    End Sub

    Sub AttachEventsToAllDocuments()
        Dim modDoc As ModelDoc2
        modDoc = iSwApp.GetFirstDocument()
        While Not modDoc Is Nothing
            If Not openDocs.Contains(modDoc) Then
                AttachModelDocEventHandler(modDoc)
            Else
                Dim docHandler As DocumentEventHandler = openDocs(modDoc)
                If Not docHandler Is Nothing Then
                    docHandler.ConnectModelViews()
                End If
            End If
            modDoc = modDoc.GetNext()
        End While
    End Sub

    Function AttachModelDocEventHandler(ByVal modDoc As ModelDoc2) As Boolean
        If modDoc Is Nothing Then
            Return False
        End If
        Dim docHandler As DocumentEventHandler = Nothing

        If Not openDocs.Contains(modDoc) Then
            Select Case modDoc.GetType
                Case swDocumentTypes_e.swDocPART
                    docHandler = New PartEventHandler()
                Case swDocumentTypes_e.swDocASSEMBLY
                    docHandler = New AssemblyEventHandler()
                Case swDocumentTypes_e.swDocDRAWING
                    docHandler = New DrawingEventHandler()
            End Select

            docHandler.Init(iSwApp, Me, modDoc)
            docHandler.AttachEventHandlers()
            openDocs.Add(modDoc, docHandler)
        End If
    End Function

    Sub DetachModelEventHandler(ByVal modDoc As ModelDoc2)
        Dim docHandler As DocumentEventHandler
        docHandler = openDocs.Item(modDoc)
        openDocs.Remove(modDoc)
        modDoc = Nothing
        docHandler = Nothing
    End Sub
#End Region

#Region "Event Handlers"
    Function SldWorks_ActiveDocChangeNotify() As Integer
        'TODO: Add your implementation here
    End Function

    Function SldWorks_DocumentLoadNotify2(ByVal docTitle As String, ByVal docPath As String) As Integer

    End Function

    Function SldWorks_FileNewNotify2(ByVal newDoc As Object, ByVal doctype As Integer, ByVal templateName As String) As Integer
        AttachEventsToAllDocuments()
    End Function

    Function SldWorks_ActiveModelDocChangeNotify() As Integer
        'TODO: Add your implementation here
    End Function

    Function SldWorks_FileOpenPostNotify(ByVal FileName As String) As Integer
        AttachEventsToAllDocuments()
    End Function
#End Region

#Region "UI Callbacks"

    Function DevSystem() As Boolean

        Dim GetID() As NetworkInterface = NetworkInterface.GetAllNetworkInterfaces
        Dim macID As String = GetID(0).GetPhysicalAddress.ToString

        Select Case macID
            Case "D8BBC17369E4" ' PranavMSI
                Return True
            Case "FC3497DC87F3" ' CELaptop02
                Return True
            Case "18C04D0893E8" ' CEWorkstation
                Return True
            Case "D8BBC178F619" ' CELaptop03
                Return True
            Case "8291334D73D9" ' CE ASUS S14
                Return True
            Case "D493901B40F7" ' CELaptop06
                Return True
            Case "C01850A4D647" ' Pranali Laptop
                Return True
            Case "C01850A4CFBB" ' Gaurang Laptop
                Return True
            Case "2CEA7F056329" ' Aadesh Laptop
                Return True
            Case "38F3ABB7D203" ' Shantanu Laptop
                Return True
            Case Else
                Return False
        End Select

    End Function

    Function CheckLicense() As Boolean

        Dim checkSWSerial As Integer
        Dim SWSerial As Object
        SWSerial = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\Software\SolidWorks\Licenses\Serial Numbers", "SolidWorks", Nothing)

        Dim readvalue As String = SWSerial.ToString
        Select Case readvalue
            Case "9000 0143 7954 5465 RYFN JYHH" 'Pravin Patil
                checkSWSerial = 1
            Case "9000 0099 9999 7937 NB8W P4KC" 'Ravindra & Sanjay
                checkSWSerial = 2
            Case "9000 0112 5184 8155 ZDDH RKDC" 'Reshma
                checkSWSerial = 3
            Case "0018 0000 0010 9647 NKHW WBH3"
                checkSWSerial = 4
            Case Else
                checkSWSerial = 100
        End Select

        Dim checkMACid As Integer
        Dim GetID() As NetworkInterface = NetworkInterface.GetAllNetworkInterfaces
        Dim macID As String = GetID(0).GetPhysicalAddress.ToString

        Select Case macID
            Case "B88198B7E4E6" 'Pravin Patil
                checkMACid = 1
            Case "D89EF336EE00" 'Ravindra
                checkMACid = 2
            Case "64006A27663C" 'Sanjay
                checkMACid = 2
            Case "D89EF30A6E6F" 'Reshma
                checkMACid = 3
            Case "E0D55EAA48F3"
                checkMACid = 2
            Case "8C8CAA3DD787" 'Siddhant
                checkMACid = 4
            Case "D8BBC178F619" 'Mayur
                checkMACid = 4
            Case "18C04D0893E8" 'CEWorkstation
                checkMACid = 4
            Case "D493901B40F7" ' Shantanu
                checkMACid = 4
            Case "C01850A4D647" ' Pranali
                checkMACid = 4
            Case "C01850A4CFBB" ' Gaurang Laptop
                checkMACid = 4
            Case "2CEA7F056329" ' Aadesh Laptop
                checkMACid = 4
                Return True
            Case Else
                checkMACid = 0
        End Select

        If checkMACid = checkSWSerial Then
            Return True
        Else
            Return False
        End If

    End Function

    Sub AHUDesign()

        Dim panel As New AHUForm

        Dim DevLicense As Boolean = DevSystem()
        Dim SystemLicense As Boolean = CheckLicense()

        Dim start As Boolean
        If DevLicense Then
            start = True
        ElseIf SystemLicense Then
            start = True
        Else
            start = False
        End If

        If start Then
            panel.Show()
        Else
            MsgBox("Authorised license not found for this Addin." & vbNewLine & "Please contact your system administrator.")
        End If

    End Sub

    Sub PopupCallbackFunction()
        bRet = iSwApp.ShowThirdPartyPopupMenu(registerID, 500, 500)
    End Sub

    Function PopupEnable() As Integer
        If iSwApp.ActiveDoc Is Nothing Then
            PopupEnable = 0
        Else
            PopupEnable = 1
        End If
    End Function

    Sub TestCallback()
        Debug.Print("Test callback")
    End Sub
    Function EnableTest() As Integer
        If iSwApp.ActiveDoc Is Nothing Then
            EnableTest = 0
        Else
            EnableTest = 1
        End If
    End Function

    Sub ShowPMP()
        If Not ppage Is Nothing Then
            ppage.Show()
        End If
    End Sub

    Function PMPEnable() As Integer
        If iSwApp.ActiveDoc Is Nothing Then
            PMPEnable = 0
        Else
            PMPEnable = 1
        End If
    End Function

    Sub FlyoutCallback()

        Dim flyGroup As FlyoutGroup = iCmdMgr.GetFlyoutGroup(flyoutGroupID)
        flyGroup.RemoveAllCommandItems()

        flyGroup.AddCommandItem(System.DateTime.Now.ToLongTimeString(), "test", 0, "FlyoutCommandItem1", "FlyoutEnableCommandItem1")

    End Sub

    Function FlyoutEnable() As Integer
        Return 1
    End Function

    Sub FlyoutCommandItem1()
        iSwApp.SendMsgToUser("Flyout command 1")
    End Sub

    Function FlyoutEnableCommandItem1() As Integer
        Return 1
    End Function


#End Region

End Class

