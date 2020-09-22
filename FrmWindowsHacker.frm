VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmWindowsHacker 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "#"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   570
   ClientWidth     =   5805
   Icon            =   "FrmWindowsHacker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   5160
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Bye"
      Height          =   495
      Left            =   2280
      TabIndex        =   16
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Change Internet Explorer Settings"
      Height          =   1815
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   5535
      Begin VB.CommandButton Command3 
         Caption         =   "Apply"
         Height          =   375
         Left            =   4440
         TabIndex        =   15
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Internet Start Page:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Internet Explorer Title:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Change Windows Registered Owner && Organization"
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   5535
      Begin VB.CommandButton Command2 
         Caption         =   "Apply"
         Height          =   375
         Left            =   4440
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Organization:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Owner:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Change The Speed That The Windows Start Menu Appears"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.Frame Frame4 
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   4560
         TabIndex        =   18
         Top             =   240
         Width           =   855
         Begin VB.Shape Shape1 
            BackColor       =   &H000000FF&
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   855
            Left            =   0
            Top             =   720
            Width           =   375
         End
         Begin VB.Shape Shape2 
            FillColor       =   &H80000004&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   0
            Top             =   720
            Width           =   855
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Apply"
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   393216
         Min             =   1
         Max             =   1000
         SelStart        =   1
         TickStyle       =   3
         Value           =   1
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Note that when you apply a change to this feature, you will need to restart your pc in order for the changes to take effect."
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   3015
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "labcomputers@lineone.net"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3720
      TabIndex        =   19
      Top             =   5520
      Width           =   1890
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Created By Lee Bailey"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3840
      TabIndex        =   17
      Top             =   5280
      Width           =   1560
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu MnuAbout 
         Caption         =   "&About"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "FrmWindowsHacker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastPosition As Integer

Private Sub Command1_Click()
    Me.MousePointer = 11
    ChangeStartMenuScrollSpeed Slider1.Value
    Me.MousePointer = 0
End Sub

Private Sub Command2_Click()
    
    If Trim(Text1.Text) = "" Then
        MsgBox "Registered Owner Cannot Be Blank!", vbExclamation, "Whacker"
        Exit Sub
    End If
    
    If Trim(Text2.Text) = "" Then
        MsgBox "Registered Organization Cannot Be Blank!", vbExclamation, "Whacker"
        Exit Sub
    End If
    
    Me.MousePointer = 11
    ChangeWindowsRegisteredOwner Trim(Text2.Text), Trim(Text1.Text)
    Me.MousePointer = 0
End Sub

Private Sub Command3_Click()
    Me.MousePointer = 11
    If Trim(Text4.Text) <> "" Then SetIEStartPage Trim(Text4.Text)
    SetIEWindowTitle Trim(Text3.Text)
    Me.MousePointer = 0
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "Whacker - V" + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + " - FREEWARE!"
    Slider1.Value = Int(GetSettingString(HKEY_CURRENT_USER, "Control Panel\Desktop", "MenuShowDelay", "1"))
    Text1.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
    Text2.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOrganization")
    Text3.Text = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Window Title")
    Text4.Text = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Start Page")
    LastPosition = 0
    Timer1.Interval = Slider1.Value
    Timer1.Enabled = True
End Sub

Private Sub MnuAbout_Click()
Dim StrAbout As String
    StrAbout = Me.Caption + Chr$(13) + Chr$(13)
    StrAbout = StrAbout + "Change your Windows and Internet Explorer settings that are not available to the user with this cool tool." + Chr$(13)
    StrAbout = StrAbout + "Finally you can get rid of that Provided By message in the Internet Explorer window." + Chr$(13) + Chr$(13)
    StrAbout = StrAbout + "Use this program at your own risk!  I will not be held responsible for any damage caused by using this tool!"
    MsgBox StrAbout, vbInformation, "About Whacker"
End Sub

Private Sub Slider1_Change()
    Timer1.Enabled = False
    Timer1.Interval = Slider1.Value
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    LastPosition = LastPosition + 1
    If LastPosition = 7 Then
        LastPosition = 0
        Shape1.Top = 720
    Else
        Shape1.Top = Shape1.Top - 120
    End If
End Sub
