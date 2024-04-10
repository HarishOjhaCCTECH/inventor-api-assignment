Imports System.Runtime.InteropServices
Imports Inventor
Imports Microsoft.Win32

Namespace InventorAddIn2
    <ProgIdAttribute("InventorAddIn2.StandardAddInServer"),
    GuidAttribute("f2bd0cc4-8ae0-46af-a151-c7fdb5828cce")>
    Public Class StandardAddInServer
        Implements Inventor.ApplicationAddInServer

        Private WithEvents m_uiEvents As UserInterfaceEvents
        Private WithEvents m_sampleButton As ButtonDefinition

        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate
            g_inventorApplication = addInSiteObject.Application
            m_uiEvents = g_inventorApplication.UserInterfaceManager.UserInterfaceEvents

            If firstTime Then
                AddToUserInterface()
            End If
        End Sub

        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate
            m_uiEvents = Nothing
            g_inventorApplication = Nothing

            System.GC.Collect()
            System.GC.WaitForPendingFinalizers()
        End Sub

        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation
            Get
                Return Nothing
            End Get
        End Property

        Public Sub ExecuteCommand(ByVal commandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand
        End Sub

        Private Sub AddToUserInterface()
            ' Get the part ribbon.
            Dim partRibbon As Ribbon = g_inventorApplication.UserInterfaceManager.Ribbons.Item("Part")

            ' Get the "Tools" tab.
            Dim toolsTab As RibbonTab = partRibbon.RibbonTabs.Item("id_TabTools")

            ' Create a new panel.
            Dim customPanel As RibbonPanel = toolsTab.RibbonPanels.Add("Sample", "MysSample", AddInClientID)

            ' Add a button.
            Dim controlDefs As ControlDefinitions = g_inventorApplication.CommandManager.ControlDefinitions
            m_sampleButton = controlDefs.AddButtonDefinition("Create Sketch", "Create Sketch", CommandTypesEnum.kShapeEditCmdType, AddInClientID)
            customPanel.CommandControls.AddButton(m_sampleButton)
        End Sub

        Private Sub m_uiEvents_OnResetRibbonInterface(Context As NameValueMap) Handles m_uiEvents.OnResetRibbonInterface
            AddToUserInterface()
        End Sub

        Private Sub m_sampleButton_OnExecute(Context As NameValueMap) Handles m_sampleButton.OnExecute
            SketchCreation()
        End Sub

        Public Sub SketchCreation()
            Dim partDoc As PartDocument = TryCast(g_inventorApplication.ActiveDocument, PartDocument)
            If partDoc IsNot Nothing Then
                Dim partDef As PartComponentDefinition = partDoc.ComponentDefinition
                Dim sketch As PlanarSketch = partDef.Sketches.Add(partDef.WorkPlanes.Item(3))
                Dim tg As TransientGeometry = g_inventorApplication.TransientGeometry

                ' Draw rectangles by center point.
                sketch.SketchLines.AddAsTwoPointCenteredRectangle(tg.CreatePoint2d(0, 0), tg.CreatePoint2d(8, 3))
                sketch.SketchLines.AddAsThreePointCenteredRectangle(tg.CreatePoint2d(20, 0), tg.CreatePoint2d(28, 3), tg.CreatePoint2d(24, 9))

                g_inventorApplication.ActiveView.Fit()
            Else
                MsgBox("No active part document found.")
            End If
        End Sub

        Public Function AddInClientID() As String
            Dim guid As String = ""
            Try
                Dim t As Type = GetType(InventorAddIn2.StandardAddInServer)
                Dim customAttributes() As Object = t.GetCustomAttributes(GetType(GuidAttribute), False)
                Dim guidAttribute As GuidAttribute = CType(customAttributes(0), GuidAttribute)
                guid = "{" + guidAttribute.Value.ToString() + "}"
            Catch
            End Try
            Return guid
        End Function
    End Class
End Namespace


Public Module Globals
    ' Inventor application object.
    Public g_inventorApplication As Inventor.Application

#Region "Function to get the add-in client ID."
    ' This function uses reflection to get the GuidAttribute associated with the add-in.
    Public Function AddInClientID() As String
        Dim guid As String = ""
        Try
            Dim t As Type = GetType(InventorAddIn2.StandardAddInServer)
            Dim customAttributes() As Object = t.GetCustomAttributes(GetType(GuidAttribute), False)
            Dim guidAttribute As GuidAttribute = CType(customAttributes(0), GuidAttribute)
            guid = "{" + guidAttribute.Value.ToString() + "}"
        Catch
        End Try

        Return guid
    End Function
#End Region

#Region "hWnd Wrapper Class"
    ' This class is used to wrap a Win32 hWnd as a .Net IWind32Window class.
    ' This is primarily used for parenting a dialog to the Inventor window.
    '
    ' For example:
    ' myForm.Show(New WindowWrapper(g_inventorApplication.MainFrameHWND))
    '
    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window
        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle() As IntPtr _
          Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property

        Private _hwnd As IntPtr
    End Class
#End Region

#Region "Image Converter"
    ' Class used to convert bitmaps and icons from their .Net native types into
    ' an IPictureDisp object which is what the Inventor API requires. A typical
    ' usage is shown below where MyIcon is a bitmap or icon that's available
    ' as a resource of the project.
    '
    ' Dim smallIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.MyIcon)

    Public NotInheritable Class PictureDispConverter
        <DllImport("OleAut32.dll", EntryPoint:="OleCreatePictureIndirect", ExactSpelling:=True, PreserveSig:=False)> _
        Private Shared Function OleCreatePictureIndirect( _
            <MarshalAs(UnmanagedType.AsAny)> ByVal picdesc As Object, _
            ByRef iid As Guid, _
            <MarshalAs(UnmanagedType.Bool)> ByVal fOwn As Boolean) As stdole.IPictureDisp
        End Function

        Shared iPictureDispGuid As Guid = GetType(stdole.IPictureDisp).GUID

        Private NotInheritable Class PICTDESC
            Private Sub New()
            End Sub

            'Picture Types
            Public Const PICTYPE_BITMAP As Short = 1
            Public Const PICTYPE_ICON As Short = 3

            <StructLayout(LayoutKind.Sequential)> _
            Public Class Icon
                Friend cbSizeOfStruct As Integer = Marshal.SizeOf(GetType(PICTDESC.Icon))
                Friend picType As Integer = PICTDESC.PICTYPE_ICON
                Friend hicon As IntPtr = IntPtr.Zero
                Friend unused1 As Integer
                Friend unused2 As Integer

                Friend Sub New(ByVal icon As System.Drawing.Icon)
                    Me.hicon = icon.ToBitmap().GetHicon()
                End Sub
            End Class

            <StructLayout(LayoutKind.Sequential)> _
            Public Class Bitmap
                Friend cbSizeOfStruct As Integer = Marshal.SizeOf(GetType(PICTDESC.Bitmap))
                Friend picType As Integer = PICTDESC.PICTYPE_BITMAP
                Friend hbitmap As IntPtr = IntPtr.Zero
                Friend hpal As IntPtr = IntPtr.Zero
                Friend unused As Integer

                Friend Sub New(ByVal bitmap As System.Drawing.Bitmap)
                    Me.hbitmap = bitmap.GetHbitmap()
                End Sub
            End Class
        End Class

        Public Shared Function ToIPictureDisp(ByVal icon As System.Drawing.Icon) As stdole.IPictureDisp
            Dim pictIcon As New PICTDESC.Icon(icon)
            Return OleCreatePictureIndirect(pictIcon, iPictureDispGuid, True)
        End Function

        Public Shared Function ToIPictureDisp(ByVal bmp As System.Drawing.Bitmap) As stdole.IPictureDisp
            Dim pictBmp As New PICTDESC.Bitmap(bmp)
            Return OleCreatePictureIndirect(pictBmp, iPictureDispGuid, True)
        End Function
    End Class
#End Region

End Module
