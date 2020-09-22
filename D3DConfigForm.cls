VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "D3DConfigForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'D3DConfigForm.cls -- Driver class for D3DConfig. This is the
'main interface to the DLL. Use this class to create
'the configuration form.
'AUTHOR: Devin Watson


'Private instance of our form
Private WithEvents m_ConfigForm As frmConfigScreen
Attribute m_ConfigForm.VB_VarHelpID = -1

'These events are fired as a relay from the Form,
'which stays private. This way, we can have as much
'or as little of the form's functionality to pass
'through to whatever links into the DLL.
Event OnCloseClick()
Event OnOKClick(AdapterSelected As Long, ModeSelected As Long, Fullscreen As Boolean)


Public Property Let Caption(strVal As String)
    'Sets the title bar caption on the form.
    strVal = Trim(strVal)
    m_ConfigForm.Caption = strVal
End Property


Public Property Get Caption() As String
    'Returns the title bar caption of the form.
    Caption = m_ConfigForm.Caption
End Property


Public Sub CloseWindow()
    'Closes out the window. This will also destroy it, instead
    'of just hiding it.
    If Not m_ConfigForm Is Nothing Then
        Unload m_ConfigForm
        Set m_ConfigForm = Nothing
    End If
End Sub

Public Sub Refresh()
    'This refreshes the form's
    'list of adapters and modes.
    If Not m_ConfigForm Is Nothing Then
        m_ConfigForm.BuildLists
    End If
End Sub

Public Sub Show()
    'Code guards are good. It is a good
    'idea to always check to see if your
    'objects have been instantiated
    'correctly.
    If Not m_ConfigForm Is Nothing Then
        m_ConfigForm.Visible = True
        m_ConfigForm.ZOrder 0
    End If
End Sub

Public Property Get Visible() As Boolean
    Visible = m_ConfigForm.Visible
End Property

Public Property Let Visible(blnVal As Boolean)
    m_ConfigForm.Visible = blnVal
End Property

Private Sub Class_Initialize()
    'We really only want initialize
    'the form.
    Set m_ConfigForm = New frmConfigScreen
End Sub


Private Sub Class_Terminate()

    If Not m_ConfigForm Is Nothing Then
        Unload m_ConfigForm
        Set m_ConfigForm = Nothing
    End If
End Sub


Private Sub m_ConfigForm_cmdCloseClicked()
    'This is one of the custom events
    'that is raised by the form which
    'we catch here. Not much to do with
    'this one, except make it invisible,
    'and fire an even off to the calling
    'program to let it know they clicked
    'Close.
    m_ConfigForm.Visible = False
    RaiseEvent OnCloseClick
End Sub


Private Sub m_ConfigForm_cmdOKClicked(SelectedAdapter As Long, SelectedMode As Long, Fullscreen As Boolean)
    'This event is fired off when
    'the user clicks on the Ok button,
    'which should signify they have
    'made their decision, and want to
    'continue. We fire off another
    'event to the calling program, which
    'then takes the same data this event has.
    m_ConfigForm.Visible = False
    RaiseEvent OnOKClick(SelectedAdapter, SelectedMode, Fullscreen)
End Sub


