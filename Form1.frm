VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Network Chat"
   ClientHeight    =   2745
   ClientLeft      =   4440
   ClientTop       =   3180
   ClientWidth     =   5250
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   5250
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   2205
      Top             =   3780
   End
   Begin VB.ComboBox cboComputers 
      Height          =   315
      Left            =   105
      TabIndex        =   2
      Top             =   210
      Width           =   5055
   End
   Begin VB.TextBox Text2 
      Height          =   960
      Left            =   105
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   945
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      Height          =   645
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1995
      Width           =   5055
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3045
      Top             =   3780
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   1365
      Picture         =   "Form1.frx":000C
      Top             =   3780
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   420
      Picture         =   "Form1.frx":0316
      Top             =   3780
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Pick or Type a machine to talk to:"
      Height          =   225
      Left            =   105
      TabIndex        =   4
      Top             =   0
      Width           =   5055
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   -210
      X2              =   16905
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Label Label1 
      Caption         =   "From:"
      Height          =   225
      Left            =   105
      TabIndex        =   3
      Top             =   735
      Width           =   5055
   End
   Begin VB.Menu mnuAppPopup 
      Caption         =   "AppPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopWhenMin 
         Caption         =   "Popup when minimized"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private lMinHeight As Long
Private lMinWidth As Long
Private bResizeOff As Boolean
'Private colMessages As String

Private Declare Function SetForegroundWindow Lib "user32" _
      (ByVal hWnd As Long) As Long
      
Private Declare Function Shell_NotifyIcon Lib "shell32" _
      Alias "Shell_NotifyIconA" _
      (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'constants required by Shell_NotifyIcon API call:
Const NIM_ADD = &H0
Const NIM_MODIFY = &H1
Const NIM_DELETE = &H2
Const NIF_MESSAGE = &H1
Const NIF_ICON = &H2
Const NIF_TIP = &H4
Const WM_MOUSEMOVE = &H200
Const WM_RBUTTONDBLCLK = &H206
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_LBUTTONDBLCLK = &H203
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_MBUTTONDBLCLK = &H209
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private nid As NOTIFYICONDATA

Private Sub UpdateIcon(Value As Long)
   ' Used to add, modify and delete icon.
   With nid
      .cbSize = Len(nid)
      .hWnd = Me.hWnd
      .uID = vbNull
      .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
      .uCallbackMessage = WM_MOUSEMOVE
      .hIcon = Me.Icon
      .szTip = App.Title & vbNullChar
   End With
   Shell_NotifyIcon Value, nid
End Sub

Private Sub Form_Load()
  Me.Icon = Me.Image1
  
  lMinHeight = Me.Height
  lMinWidth = Me.Width
  
  'load saved off computers into combobox
  LoadComputers
  
  'set winsock properties
  Winsock1.Protocol = sckUDPProtocol
  Winsock1.LocalPort = 6421
  Winsock1.RemotePort = 6421
End Sub

Private Sub Form_Resize()
  Dim lWidth As Long
  Dim lHeight As Long
  Const Unit = 105
  
  'this is here so when the mnuShow_Click event is fired, the form wont minimize and hide again
  If bResizeOff = False Then
    If Me.WindowState = vbMinimized Then
      Me.Hide
      UpdateIcon NIM_ADD
    Else
      UpdateIcon NIM_DELETE
    End If
  End If
    
  'generic resize logic
  With Me
    If .WindowState = vbMinimized Then Exit Sub
    If .Height < lMinHeight Then .Height = lMinHeight
    If .Width < lMinWidth Then .Width = lMinWidth
  
    lWidth = .ScaleWidth
    lHeight = .ScaleHeight
    
    .cboComputers.Width = lWidth - 2 * Unit
    .Text1.Width = lWidth - 2 * Unit
    .Text2.Width = lWidth - 2 * Unit
    
    .Text2.Height = lHeight - 17 * Unit
    .Text1.Top = .Text2.Top + .Text2.Height + Unit
    
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'remove icon from system tray
  UpdateIcon NIM_DELETE
  
  'save off computers added to combobox to an XML file
  PersistComputers
  
  Winsock1.Close
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim Result As Long
   Dim msg As Long
       
   'really interesting stuff here...i got it from MSDN
   If Me.ScaleMode = vbPixels Then
      msg = X
   Else
      msg = X / Screen.TwipsPerPixelX
   End If

   'handles mouse events when form is minimized, hidden and icon is in the system tray
   Select Case msg
      Case WM_RBUTTONDBLCLK
      Case WM_RBUTTONDOWN
      Case WM_RBUTTONUP
         PopupMenu mnuAppPopup
      Case WM_LBUTTONDBLCLK
          mnuShow_Click
      Case WM_LBUTTONDOWN
      Case WM_LBUTTONUP
      Case WM_MBUTTONDBLCLK
      Case WM_MBUTTONDOWN
      Case WM_MBUTTONUP
      Case WM_MOUSEMOVE
      Case Else
   End Select
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuPopWhenMin_Click()
  'this menu item is used so that if it is checked and the app is in the system tray
  'and a new message is recieved the app will unhide and show in normal state.
  'if this menu item is unchecked and the app is in the system tray and the app recieves
  'a new message, the icon will blink until the user brings it up from the tray to
  'see the new message
  If Me.mnuPopWhenMin.Checked = True Then
    Me.mnuPopWhenMin.Checked = False
  Else
    Me.mnuPopWhenMin.Checked = True
  End If
End Sub

Private Sub mnuShow_Click()
  Dim Result As Long
  'this menu event will unhide the app from the system tray and show it in a normal state
  Me.Timer1.Enabled = False
  Me.Icon = Me.Image1
  UpdateIcon NIM_DELETE
  bResizeOff = True
  Me.WindowState = vbNormal
  Result = SetForegroundWindow(Me.hWnd)
  Me.Show
  bResizeOff = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  'when the user hits the enter key when typing in a message, the
  'app will try to send the message to the computer selected in the combobox
  If KeyAscii = Asc(vbCrLf) Then
    If Len(Me.cboComputers.Text) = 0 Then
      MsgBox "Please pick a computer to send message to"
      Exit Sub
    End If
    
    Me.MousePointer = 11
    On Error Resume Next
    Winsock1.SendData Winsock1.LocalHostName & "|" & Text1.Text
    If Err.Number <> 0 Then
      MsgBox "There was an error sending your message" & vbCrLf & "Check to make sure the Machine Name is correct", vbCritical + vbOKOnly, App.Title
    Else
      Text1.Text = ""
    End If
    On Error GoTo 0
    Me.MousePointer = 0
  End If
End Sub

Private Sub Timer1_Timer()
  Static bool As Boolean
  'used to flash the icon when the app is in the system tray and a message is waiting for the user
  If bool = True Then
    Me.Icon = Me.Image1
    bool = False
  Else
    Me.Icon = Me.Image2
    bool = True
  End If
  UpdateIcon NIM_MODIFY
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
  Dim sText As String
  Dim sFrom As String
  Dim sMsg As String
  Dim iPlace As Integer
  'this takes the data that arived from another computer displays it
  
  Winsock1.GetData sText
  iPlace = InStr(1, sText, "|", vbBinaryCompare)
  sFrom = Mid(sText, 1, iPlace - 1)
  sMsg = Mid(sText, iPlace + 1)
  Label1.Caption = "From: " & sFrom
  Text2.Text = sMsg
  
  If Me.mnuPopWhenMin.Checked = True Then
    mnuShow_Click
  Else
    Me.Timer1.Enabled = True
  End If
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  MsgBox Description & vbCrLf & Number
End Sub

Private Sub Label2_DblClick()
  'this will remove the selected computer in the combobox
  If Me.cboComputers.ListCount = 0 Then Exit Sub
  Me.cboComputers.RemoveItem Me.cboComputers.ListIndex
  If Me.cboComputers.ListCount = 0 Then Exit Sub
  Me.cboComputers.ListIndex = 0
  cboComputers_Click
End Sub

Private Sub cboComputers_Click()
  'sets the remote host of the winsock control to the computer selected in the combobox
  If Me.cboComputers.ListCount = 0 Then Exit Sub
  Winsock1.RemoteHost = Me.cboComputers.Text
End Sub

Private Sub cboComputers_Validate(Cancel As Boolean)
  Dim sComputer As String
  Dim X As Integer
  Dim bFound As Boolean
  'makes sure no computer is listed twice
  
  bFound = False
  sComputer = Me.cboComputers.Text
  For X = 1 To Me.cboComputers.ListCount
    If sComputer = Me.cboComputers.List(X - 1) Then
      bFound = True
      Exit For
    End If
  Next
  
  If bFound = False Then
    Me.cboComputers.AddItem sComputer
  End If
  cboComputers_Click
End Sub

'this loads the list of computers saved off in the xml file into the combobox.
'its written so that it can be used by either the msxml.dll version 2 or 3
Private Sub LoadComputers()
  Dim X As Long
  'Dim oXML2 As MSXML2.DOMDocument
  'Dim oXML As MSXML.DOMDocument
  Dim oXML As Object
  
  If Len(Dir(App.Path & "\netchat.xml")) = 0 Then Exit Sub
  
  'Set oXML = New MSXML2.DOMDocument
  'Set oXML = New MSXML.DOMDocument
  On Error Resume Next
  Set oXML = CreateObject("MSXML2.DOMDocument")
  If oXML Is Nothing Then Set oXML = CreateObject("MSXML.DOMDocument")
  If oXML Is Nothing Then
    MsgBox "Error loading chat application"
    End
  End If
  On Error GoTo 0
  
  oXML.async = False
  If oXML.Load(App.Path & "\netchat.xml") = False Then
    MsgBox "There was an error loading saved computers"
  Else
    For X = 0 To oXML.documentElement.childNodes.length - 1
      Me.cboComputers.AddItem oXML.documentElement.childNodes.Item(X).Text
    Next
  End If
  
  Set oXML = Nothing
End Sub

'this saves off the list of computers in the combobox into
'an xml file for the next time the app is started.
'its written so that it can be used by either the msxml.dll version 2 or 3
Private Sub PersistComputers()
  Dim X As Integer
  'Dim oXML As MSXML2.DOMDocument
  'Dim oMain As MSXML2.IXMLDOMNode
  'Dim oComputer As MSXML2.IXMLDOMNode
  Dim oXML As Object
  Dim oMain As Object
  Dim oComputer As Object
    
  'Set oXML = New MSXML2.DOMDocument
  On Error Resume Next
  Set oXML = CreateObject("MSXML2.DOMDocument")
  If oXML Is Nothing Then Set oXML = CreateObject("MSXML.DOMDocument")
  If oXML Is Nothing Then
    MsgBox "Error closing chat application"
    End
  End If
  On Error GoTo 0
  
  Set oMain = oXML.createNode(1, "netchat", "")
  oXML.appendChild oMain
  
  For X = 1 To Me.cboComputers.ListCount
    Set oComputer = oXML.createNode(1, "computer", "")
    oComputer.Text = Me.cboComputers.List(X - 1)
    oMain.appendChild oComputer
  Next
  oXML.save App.Path & "\netchat.xml"
  Set oXML = Nothing
  Set oMain = Nothing
  Set oComputer = Nothing
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    Me.PopupMenu mnuAppPopup
  End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    Me.PopupMenu mnuAppPopup
  End If
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    Me.PopupMenu mnuAppPopup
  End If
End Sub
