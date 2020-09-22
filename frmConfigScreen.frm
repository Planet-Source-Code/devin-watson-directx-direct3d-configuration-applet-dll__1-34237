VERSION 5.00
Begin VB.Form frmConfigScreen 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Direct3D Configuration"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmConfigScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   438
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   468
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkFullscreen 
      Caption         =   "Run in Fullscreen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   2775
   End
   Begin VB.PictureBox picLogo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1035
      Left            =   1800
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   7
      Top             =   60
      Width           =   3000
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   435
      Left            =   60
      TabIndex        =   3
      Top             =   6000
      Width           =   1215
   End
   Begin VB.ListBox lstModes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Lists all of the video modes possible for the selected card"
      Top             =   2700
      Width           =   2775
   End
   Begin VB.ComboBox cmbAdapters 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Lists all video cards installed on your machine"
      Top             =   1800
      Width           =   2835
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   435
      Left            =   5760
      TabIndex        =   4
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label lblInstructions 
      BackStyle       =   0  'Transparent
      Caption         =   "lblInstructions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3435
      Left            =   3180
      TabIndex        =   8
      Top             =   1740
      Width           =   3495
   End
   Begin VB.Label lblModes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Available Video Modes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   2340
      Width           =   2160
   End
   Begin VB.Label lblAdapters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Video Cards:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   1500
      Width           =   1170
   End
End
Attribute VB_Name = "frmConfigScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'frmConfigScreen.frm -- Private form which represents
'the applet window. This form is not called directly due
'to the ActiveX DLL model, but instead is controlled
'via the D3DConfigForm class.
'AUTHOR: Devin Watson
'NOTE: To easily change the strings and the logo
'picture, use the Resource editor. I put everything there
'in order to better facilitate modifications.
'Also, a little terminology: "adapter" in DirectX
'refers to a video card. You can use them
'interchangably.

'Private references to DirectX/Direct3D needed for
'enumerating devices.
Private mDX As DirectX8
Private mD3D As Direct3D8

'Private variables for various settings
Private mblnFullscreen As Boolean       'Whether or not to run in fullscreen
Private CurMode As Long                 'The currently selected mode
Private CurAdapter As Long              'The currently selected video adapter

'The array works like this:
'The index of the array is the adapter number,
'and the value stored at that index is the
'number of modes available. That way,
'we just need to query on change.
Private Adapters() As Long

'Some events we will be raising to the
'class, which in turn will raise other events
'from the class to the user.
Event cmdCloseClicked()
Event cmdOKClicked(SelectedAdapter As Long, SelectedMode As Long, Fullscreen As Boolean)

Public Sub BuildLists()
    'This constructs the lists of
    'adapters and modes.
    
    Dim numAdapters As Long
    Dim Adapter As Long
    Dim Info As D3DADAPTER_IDENTIFIER8
    Dim I As Integer
    Dim tmpStr As String
    Dim NumModes As Long
    
    'As a safeguard, we clear out
    'all of our visual displays,
    'in case this sub is called
    'out of cycle.
    cmbAdapters.Clear
    lstModes.Clear
    
    'First, we need to determine how many adapters
    '(video cards) there are on the computer
    numAdapters = mD3D.GetAdapterCount
    
    'And from that, we redimension
    'our array to that number.
    ReDim Adapters(numAdapters)

    'Then, for each adapter, we need
    'to build up a list of information
    'on each adapter.
    For Adapter = 0 To numAdapters - 1
    
        'This retrieves descriptive information about the
        'adapter (video card)
        tmpStr = ""
        mD3D.GetAdapterIdentifier Adapter, 0, Info
        
        'The description field in the structure
        'is one that is good to use for presentation.
        'However, since it is a byte array (ASCII codes)
        'we need to convert it from the codes to the
        'actual display text for presentation.
        'I've optimized this loop a little with some
        'unrolling, because I noticed that it ran
        'somewhat slowly otherwise.
        For I = 0 To 511
            tmpStr = tmpStr & (Chr(Info.Description(I)))
            I = I + 1
            tmpStr = tmpStr & (Chr(Info.Description(I)))
        Next I
        
        'This gets rid of some extra
        'space that might be in the string.
        tmpStr = Trim(tmpStr)
        
        'If we are looking at the default adapter,
        'we should let the user know that.
        If Adapter = D3DADAPTER_DEFAULT Then
            tmpStr = tmpStr & " (Default)"
        End If
        
        'And we add that item to the list. It's
        'important that the lists be left unsorted,
        'because everything is 0-based, including the
        'array, and so we have a 1:1 relationship
        'between the adapters, modes, etc.
        cmbAdapters.AddItem tmpStr
        
        'And finally, we capture the number of
        'modes possible for this adapter.
        NumModes = mD3D.GetAdapterModeCount(Adapter)
        Adapters(Adapter) = NumModes
    Next Adapter
    
    'Just to prime the pump a little,
    'we set drop-down list to display
    'the first adapter, then
    'call its Click() method
    'to get the modes update.
    cmbAdapters.ListIndex = 0
    cmbAdapters_Click
End Sub


Private Sub LoadResources()
    'This subroutine will load all
    'of the resources from the resource file
    'into the form. To make a change, simply
    'edit the file D3DConfig.res file
    'with the Resource Editor.
    
    'Load the logo picture into the form
    picLogo.Picture = LoadResPicture(101, vbResBitmap)
    'Center the picture on the form
    picLogo.Left = (Me.ScaleWidth / 2) - (picLogo.Width / 2)
    
    'Load the string resources for the buttons.
    cmdClose.Caption = LoadResString(102)
    cmdOK.Caption = LoadResString(103)
    chkFullscreen.Caption = LoadResString(104)
    lblAdapters = LoadResString(105)
    lblModes = LoadResString(106)
    lblInstructions = LoadResString(107)
End Sub

Private Sub chkFullscreen_Click()
    'Since check boxes can have a third
    'value (vbGrayed), it is a good idea
    'to explicitly check the values of it
    'against known constants.
    If chkFullscreen.Value = vbChecked Then
        mblnFullscreen = True
    ElseIf chkFullscreen.Value = vbUnchecked Then
        mblnFullscreen = False
    End If
End Sub

Private Sub cmbAdapters_Click()
    'Here, we get the individual modes for
    'the selected card.
    Dim tmpMode As D3DDISPLAYMODE
    Dim I As Long
    Dim Idx As Long
    Dim tmpStr As String
    
    Idx = cmbAdapters.ListIndex
    CurAdapter = Idx
    lstModes.Clear
    
    For I = 0 To Adapters(Idx) - 1
        tmpStr = ""
        
        mD3D.EnumAdapterModes Idx, I, tmpMode
        tmpStr = tmpMode.Width & "x" & tmpMode.Height
        
        'For target rendering devices (meaning screens),
        'there can only be 4 possible choices, and out
        'of those, there's just 16-bit and 32-bit.
        
        If (tmpMode.Format = D3DFMT_X1R5G5B5) _
            Or (tmpMode.Format = D3DFMT_R5G6B5) Then
                tmpStr = tmpStr & " 16-bit"
        End If
            
        If (tmpMode.Format = D3DFMT_X8R8G8B8) _
            Or (tmpMode.Format = D3DFMT_A8R8G8B8) Then
                tmpStr = tmpStr & " 32-bit"
        End If
        
        If tmpMode.RefreshRate <> 0 Then
            tmpStr = tmpStr & " @ " & tmpMode.RefreshRate
        End If
        lstModes.AddItem tmpStr
    Next I
    lstModes.ListIndex = 0
    CurMode = 0
End Sub


Private Sub cmdClose_Click()
    RaiseEvent cmdCloseClicked
End Sub

Private Sub cmdOK_Click()
    RaiseEvent cmdOKClicked(CurAdapter, CurMode, mblnFullscreen)
End Sub

Private Sub Form_Load()
    'Initialize any local data here for the form.
    Set mDX = New DirectX8
    Set mD3D = mDX.Direct3DCreate
    
    LoadResources
    'Build up the arrays and create our
    'information lists
    BuildLists
    mblnFullscreen = True
    chkFullscreen.Value = vbChecked
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'Make sure we get rid of anything here we need to.
    If Not mD3D Is Nothing Then
        Set mD3D = Nothing
    End If
    
    If Not mDX Is Nothing Then
        Set mDX = Nothing
    End If
End Sub


Private Sub lstModes_Click()
    CurMode = lstModes.ListIndex
End Sub


