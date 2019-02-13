Imports System.Runtime.InteropServices


Public Module Utilities

#Region "Simplified Button Creation"
    ''' <summary>
    ''' Function to simplify the creation of a button definition.  The big advantage
    ''' to using this function is that you don't have to deal with converting images
    ''' but instead just reference a folder on disk where this routine reads the images.
    ''' </summary>
    ''' <param name="DisplayName">
    ''' The name of the command as it will be displayed on the button. 
    ''' </param>
    ''' <param name="InternalName">
    ''' The internal name of the command. This needs to be unique with respect to ALL other
    ''' commands. It's best to incorporate a company name to help with uniqueness.
    ''' </param>
    ''' <param name="ToolTip">
    ''' The tooltip that will be used for the command.
    ''' 
    ''' This is optional and the display name will be used as the
    ''' tooltip if no tooltip is specified. Like in the DisplayName argument, you can use
    ''' returns to force line breaks.
    ''' </param>
    ''' <param name="IconFolder">
    ''' The folder that contains the icon files. This can be a full path or a path that is
    ''' relative to the location of the add-in dll. The folder should contain the files 
    ''' 16x16.png and 32x32.png. Each command will have its own folder so they can have 
    ''' their own icons.
    ''' 
    ''' This is optional and if no icon is specified then no icon will be displayed on the
    ''' button and it will be only text.
    ''' </param>
    ''' <returns>
    ''' Returns the newly created button definition or Nothing in case of failure.
    ''' </returns>
    Public Function CreateButtonDefinition(DisplayName As String,
                                           InternalName As String,
                                           Optional ToolTip As String = "",
                                           Optional IconFolder As String = "") As Inventor.ButtonDefinition

        ' Check to see if a command already exists is the specified internal name.
        Dim testDef As Inventor.ButtonDefinition = Nothing
        Try
            testDef = g_inventorApplication.CommandManager.ControlDefinitions.Item(InternalName)
        Catch ex As Exception
        End Try

        If Not testDef Is Nothing Then
            MsgBox("Error when loading the add-in ""BatchProgram"". A command already exists with the same internal name. Each add-in must have a unique internal name. Change the internal name in the call to CreateButtonDefinition.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Inventor Add-In Template")
            Return Nothing
        End If

        ' Check to see if the provided folder is a full or relative path.
        If iconFolder <> "" Then
            If Not System.IO.Directory.Exists(iconFolder) Then
                ' The folder provided doesn't exist, so assume it is a relative path and
                ' build up the full path.
                Dim dllPath As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)

                IconFolder = System.IO.Path.Combine(dllPath, IconFolder)
            End If
        End If

        ' Get the images from the specified icon folder.
        Dim iPicDisp16x16 As stdole.IPictureDisp = Nothing
        Dim iPicDisp32x32 As stdole.IPictureDisp = Nothing
        If IconFolder <> "" Then
            If System.IO.Directory.Exists(IconFolder) Then
                Dim filename16x16 As String = System.IO.Path.Combine(IconFolder, "16x16.png")
                Dim filename32x32 As String = System.IO.Path.Combine(IconFolder, "32x32.png")

                If System.IO.File.Exists(filename16x16) Then
                    Try
                        Dim image16x16 As New System.Drawing.Bitmap(filename16x16)
                        iPicDisp16x16 = Utilities.ConvertImage.ConvertImageToIPictureDisp(image16x16)
                    Catch ex As Exception
                        MsgBox("Unable to load the 16x16.png image from """ & IconFolder & """." & vbCrLf & "No small icon will be used.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Error Loading Icon")
                    End Try
                Else
                    MsgBox("The icon for the small button does not exist: """ & filename16x16 & """." & vbCrLf & "No small icon will be used.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Error Loading Icon")
                End If

                If System.IO.File.Exists(filename32x32) Then
                    Try
                        Dim image32x32 As New System.Drawing.Bitmap(filename32x32)
                        iPicDisp32x32 = Utilities.ConvertImage.ConvertImageToIPictureDisp(image32x32)
                    Catch ex As Exception
                        MsgBox("Unable to load the 32x32.png image from """ & IconFolder & """." & vbCrLf & "No large icon will be used.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Error Loading Icon")
                    End Try
                Else
                    MsgBox("The icon for the large button does not exist: """ & filename32x32 & """." & vbCrLf & "No large icon will be used.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Error Loading Icon")
                End If
            End If
        End If

        Try
            ' Get the ControlDefinitions collection.
            Dim controlDefs As Inventor.ControlDefinitions = g_inventorApplication.CommandManager.ControlDefinitions

            ' Create the command defintion.
            Dim btnDef As Inventor.ButtonDefinition = controlDefs.AddButtonDefinition(DisplayName,
                                                                                  InternalName,
                                                                                  Inventor.CommandTypesEnum.kShapeEditCmdType,
                                                                                  g_addInClientID,
                                                                                  "",
                                                                                  ToolTip,
                                                                                  iPicDisp16x16,
                                                                                  iPicDisp32x32)
            Return btnDef
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
#End Region


#Region "hWnd Wrapper Class"
    ' This class is used to wrap a Win32 hWnd as a .Net IWind32Window class.
    ' This is primarily used for parenting a dialog to the Inventor window.
    ' This provides the expected behavior when the Inventor window is collapsed
    ' and activated.
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
    ' Class used to convert bitmaps and icons between their .Net native types
    ' and an IPictureDisp object which is what the Inventor API requires.

    <Global.System.Security.Permissions.PermissionSetAttribute _
    (Global.System.Security.Permissions.SecurityAction.Demand, Name:="FullTrust")>
    Public Class ConvertImage
        Inherits System.Windows.Forms.AxHost
        Public Sub New()
            MyBase.New("59EE46BA-677D-4d20-BF10-8D8067CB8B32")
        End Sub

        Public Shared Function ConvertImageToIPictureDisp(ByVal Image As System.Drawing.Image) As stdole.IPictureDisp
            Try
                Return (GetIPictureFromPicture(Image))
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        Public Shared Function ConvertIPictureDispToImage(ByVal IPict As stdole.IPictureDisp) As System.Drawing.Image
            Try
                Return (GetPictureFromIPictureDisp(IPict))
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
    End Class
#End Region

End Module
