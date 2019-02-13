Imports Inventor
Imports System.Runtime.InteropServices
Imports Microsoft.Win32

Namespace BatchProgram
    <ProgIdAttribute("BatchProgram.StandardAddInServer"), _
    GuidAttribute(g_simpleAddInClientID)> _
    Public Class StandardAddInServer
        Implements Inventor.ApplicationAddInServer

        '*********************************************************************************
        '* The two declarations below are related to adding buttons to Inventor's UI.
        '* They can be deleted if this add-in doesn't have a UI and only runs in the 
        '* background handling events.
        '*********************************************************************************

        ' Declaration of the object for the UserInterfaceEvents to be able to handle
        ' if the user resets the ribbon so the button can be added back in.
        Private WithEvents m_uiEvents As UserInterfaceEvents

        ' Declaration of the button definition with events to handle the click event.
        ' For additional commands this declaration along with other sections of code
        ' that apply to the button can be duplicated from this example.
        Private WithEvents m_sampleButton As ButtonDefinition


#Region "ApplicationAddInServer Members"

        ' This method is called by Inventor when it loads the AddIn. The AddInSiteObject provides access  
        ' to the Inventor Application object. The FirstTime flag indicates if the AddIn is loaded for
        ' the first time. However, with the introduction of the ribbon this argument is always true.
        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate
            Try
                ' Initialize AddIn members.
                g_inventorApplication = addInSiteObject.Application

                ' Connect to the user-interface events to handle a ribbon reset.
                m_uiEvents = g_inventorApplication.UserInterfaceManager.UserInterfaceEvents

                '*********************************************************************************
                '* The remaining code in this Sub is all for adding the add-in into Inventor's UI.
                '* It can be deleted if this add-in doesn't have a UI and only runs in the 
                '* background handling events.
                '*********************************************************************************

                ' Create the button definition using the CreateButtonDefinition function to simplify this step.
                m_sampleButton = Utilities.CreateButtonDefinition("Command" & vbCr & "Name", "niftyCommandID", "", "ButtonResources\SampleButton")

                ' Add to the user interface, if it's the first time.
                ' If this add-in doesn't have a UI but runs in the background listening
                ' to events, you can delete this.
                If firstTime Then
                    AddToUserInterface()
                End If
            Catch ex As Exception
                MsgBox("Unexpected failure in the activation of the add-in ""BatchProgram""" & vbCrLf & vbCrLf & ex.Message)
            End Try
        End Sub

        ' This method is called by Inventor when the AddIn is unloaded. The AddIn will be
        ' unloaded either manually by the user or when the Inventor session is terminated.
        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate
            ' Release objects.
            m_sampleButton = Nothing
            m_uiEvents = Nothing
            g_inventorApplication = Nothing

            System.GC.Collect()
            System.GC.WaitForPendingFinalizers()
        End Sub

        ' This property is provided to allow the AddIn to expose an API of its own to other 
        ' programs. Typically, this  would be done by implementing the AddIn's API
        ' interface in a class and returning that class object through this property.
        ' Typically it's not used, like in this case, and returns Nothing.
        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation
            Get
                Return Nothing
            End Get
        End Property

        ' Note:this method is now obsolete, you should use the 
        ' ControlDefinition functionality for implementing commands.
        Public Sub ExecuteCommand(ByVal commandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand
            ' Not used.
        End Sub

#End Region

#Region "User interface definition"
        ' Adds whatever is needed by this add-in to the user-interface.  This is 
        ' called when the add-in loaded and also if the user interface is reset.
        Private Sub AddToUserInterface()
            ' This sample code illustrates creating a button on a new panel of the Tools tab of 
            ' the Part ribbon. You'll need to change this to create the UI that your add-in needs.

            ' Get the part ribbon.
            Dim partRibbon As Ribbon = g_inventorApplication.UserInterfaceManager.Ribbons.Item("Part")

            ' Get the "Tools" tab.
            Dim toolsTab As RibbonTab = partRibbon.RibbonTabs.Item("id_TabTools")

            ' Check to see if the "MySample" panel already exists and create it if it doesn't.
            Dim customPanel As RibbonPanel = Nothing
            Try
                customPanel = toolsTab.RibbonPanels.Item("MySample")
            Catch ex As Exception
            End Try

            If customPanel Is Nothing Then
                ' Create a new panel.
                customPanel = toolsTab.RibbonPanels.Add("Sample", "MySample", g_addInClientID)
            End If

            ' Add a button.
            If Not m_sampleButton Is Nothing Then
                customPanel.CommandControls.AddButton(m_sampleButton, True)
            End If
        End Sub

        Private Sub m_uiEvents_OnResetRibbonInterface(Context As NameValueMap) Handles m_uiEvents.OnResetRibbonInterface
            ' The ribbon was reset, so add back the add-ins user-interface.
            AddToUserInterface()
        End Sub

        ' Sample handler for the button.
        Private Sub m_sampleButton_OnExecute(Context As NameValueMap) Handles m_sampleButton.OnExecute
            CommandFunctions.SampleCommandFunction()
        End Sub
#End Region

    End Class
End Namespace


Public Module Globals
    ' Inventor application object.
    Public g_inventorApplication As Inventor.Application

    ' The unique ID for this add-in.  If this add-in is copied to create a new add-in
    ' you need to update this ID along with the ID in the .manifest file, the .addin file
    ' and create a new ID for the typelib GUID in AssemblyInfo.vb
    Public Const g_simpleAddInClientID As String = "9326db0c-2b3b-4169-a489-a66e31a9bdd5"
    Public Const g_addInClientID As String = "{" & g_simpleAddInClientID & "}"
End Module