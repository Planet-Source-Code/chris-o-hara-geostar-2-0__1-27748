VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "GeoStar"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   9195
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog comdiag 
      Left            =   4920
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame con 
      Height          =   5355
      Left            =   5520
      TabIndex        =   1
      Top             =   360
      Width           =   3615
      Begin VB.OptionButton Option3 
         Caption         =   "Options"
         Height          =   255
         Left            =   2400
         TabIndex        =   32
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Exit"
         Height          =   495
         Left            =   2400
         TabIndex        =   31
         Top             =   4680
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Grid Star"
         Height          =   255
         Left            =   1260
         TabIndex        =   29
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "GeoStar"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         Caption         =   "Clear"
         Height          =   495
         Left            =   1320
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   27
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton Command11 
         Appearance      =   0  'Flat
         Caption         =   "Create"
         Height          =   495
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   26
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Height          =   2775
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3375
         Begin VB.TextBox aspins 
            Height          =   285
            Left            =   1440
            TabIndex        =   6
            Text            =   "5"
            Top             =   2280
            Width           =   975
         End
         Begin VB.TextBox aslength 
            Height          =   285
            Left            =   1440
            TabIndex        =   5
            Text            =   "540"
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox aangle 
            Height          =   285
            Left            =   1440
            TabIndex        =   4
            Text            =   "45"
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox alength 
            Height          =   285
            Left            =   1440
            TabIndex        =   3
            Text            =   "1200"
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Density:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Width between:"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "# of Sides:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   8
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Side length:"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   1215
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   2640
            Picture         =   "frmMain.frx":08CA
            Top             =   360
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   2640
            Picture         =   "frmMain.frx":1594
            Top             =   960
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   2
            Left            =   2640
            Picture         =   "frmMain.frx":225E
            Top             =   1560
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   3
            Left            =   2640
            Picture         =   "frmMain.frx":2F28
            Top             =   2160
            Width           =   480
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "BackColour"
         Height          =   1155
         Left            =   120
         TabIndex        =   11
         Top             =   3360
         Width           =   1575
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   375
            Left            =   1200
            TabIndex        =   43
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox Textb 
            Height          =   285
            Left            =   720
            TabIndex        =   37
            Text            =   "255"
            Top             =   780
            Width           =   375
         End
         Begin VB.TextBox Textg 
            Height          =   285
            Left            =   720
            TabIndex        =   38
            Text            =   "255"
            Top             =   540
            Width           =   375
         End
         Begin VB.TextBox Textr 
            Height          =   285
            Left            =   720
            TabIndex        =   39
            Text            =   "255"
            Top             =   300
            Width           =   375
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Blue:"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   780
            Width           =   495
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Green:"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   540
            Width           =   495
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Red:"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   300
            Width           =   495
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "ForeColour"
         Height          =   1155
         Left            =   1920
         TabIndex        =   12
         Top             =   3360
         Width           =   1575
         Begin VB.CommandButton Command2 
            Caption         =   "..."
            Height          =   375
            Left            =   1200
            TabIndex        =   44
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   720
            TabIndex        =   13
            Text            =   "r"
            Top             =   780
            Width           =   375
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   720
            TabIndex        =   14
            Text            =   "r"
            Top             =   540
            Width           =   375
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   720
            TabIndex        =   15
            Text            =   "r"
            Top             =   300
            Width           =   375
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Red:"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   300
            Width           =   495
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Green:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   540
            Width           =   495
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Blue:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   780
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2775
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   3375
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   1440
            TabIndex        =   22
            Text            =   "8"
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1440
            TabIndex        =   21
            Text            =   "30"
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1440
            TabIndex        =   20
            Text            =   "3000"
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Point size:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Line segments:"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "  Points:"
            Height          =   255
            Left            =   480
            TabIndex        =   23
            Top             =   480
            Width           =   855
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   4
            Left            =   2640
            Picture         =   "frmMain.frx":3BF2
            Top             =   360
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   5
            Left            =   2640
            Picture         =   "frmMain.frx":48BC
            Top             =   960
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   6
            Left            =   2640
            Picture         =   "frmMain.frx":5586
            Top             =   1560
            Width           =   480
         End
         Begin VB.Line Line1 
            X1              =   3120
            X2              =   240
            Y1              =   2400
            Y2              =   2400
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2775
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   3375
         Begin VB.CheckBox sh 
            Alignment       =   1  'Right Justify
            Caption         =   "Shadow"
            Height          =   255
            Left            =   480
            TabIndex        =   35
            Top             =   1080
            Value           =   1  'Checked
            Width           =   1155
         End
         Begin VB.TextBox ps 
            Height          =   285
            Left            =   1440
            MaxLength       =   1
            TabIndex        =   34
            Text            =   "1"
            Top             =   480
            Width           =   975
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   8
            Left            =   2640
            Picture         =   "frmMain.frx":6250
            Top             =   960
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   7
            Left            =   2640
            Picture         =   "frmMain.frx":6F1A
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label11 
            Caption         =   "Line Width:"
            Height          =   255
            Left            =   480
            TabIndex        =   36
            Top             =   480
            Width           =   855
         End
      End
   End
   Begin VB.PictureBox P1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      DrawWidth       =   2
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   0
      ScaleHeight     =   5385
      ScaleWidth      =   5385
      TabIndex        =   0
      Top             =   360
      Width           =   5415
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5640
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7F80
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":831C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":86B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8A54
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8DF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":918C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9528
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":98C4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    'Declare PI
    Const PI = 3.141592654
    
    'Declare public variables
    Dim gstar As Boolean
    Dim override As Boolean
    

Private Sub Command1_Click()

    'Declare Variables
    Dim col, r, g, b, x

    'Set values
    r = 0
    g = 0
    b = 0
        
    On Error Resume Next
    
    'Show common dialog
    comdiag.CancelError = True
    comdiag.ShowColor
     
    'Get RGB values
    If Not Err Then
    
        col = comdiag.Color
   
        For x = 1 To 513 Step 1
        
            If col >= 65536 Then
                col = col - 65536
                b = b + 1
            ElseIf col >= 256 Then
                col = col - 256
                g = g + 1
            Else
                r = col
            End If
            
        Next x
        
    End If
    
    'Write values
    Textr.Text = r
    Textg.Text = g
    Textb.Text = b

End Sub

Private Sub Command2_Click()

    'Declare Variables
    Dim col, r, g, b, x

    'Set values
    r = 0
    g = 0
    b = 0
        
    On Error Resume Next
    
    'Show common dialog
    comdiag.CancelError = True
    comdiag.ShowColor
     
    'Get RGB Values
    If Not Err Then
    
        col = comdiag.Color
   
        For x = 1 To 513 Step 1
        
            If col >= 65536 Then
                col = col - 65536
                b = b + 1
            ElseIf col >= 256 Then
                col = col - 256
                g = g + 1
            Else
                r = col
            End If
            
        Next x
        
    End If
    
    'Write values
    Text3.Text = r
    Text4.Text = g
    Text5.Text = b

End Sub

Private Sub Form_Load()
    
    'Setup Toolbar
    Toolbar1.ImageList = ImageList2
    
    Toolbar1.Buttons.Item(1).Image = 1
    Toolbar1.Buttons.Item(2).Image = 3
    Toolbar1.Buttons.Item(4).Image = 4
    Toolbar1.Buttons.Item(5).Image = 6
    Toolbar1.Buttons.Item(7).Image = 7
    Toolbar1.Buttons.Item(8).Image = 8
    Toolbar1.Buttons.Item(10).Image = 9
    
    Toolbar1.Buttons.Item(7).Value = 1
    Toolbar1.Buttons.Item(8).Value = 0
    
    Toolbar1.Buttons.Item(1).Key = "NEW"
    Toolbar1.Buttons.Item(2).Key = "EXPORT"
    Toolbar1.Buttons.Item(4).Key = "CREATE"
    Toolbar1.Buttons.Item(5).Key = "OVERRIDE"
    Toolbar1.Buttons.Item(7).Key = "GEO"
    Toolbar1.Buttons.Item(8).Key = "GRID"
    Toolbar1.Buttons.Item(10).Key = "EXIT"
    
    'Geostar is selected type
    gstar = True
    
    'Reset controls
    sreset
    
End Sub





Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    'Declare Variables
    Dim sfname As String

    'See what button was pressed
    Select Case Button.Key

        Case "NEW"
        
            'Create a new star
            P1.Cls
            sreset
            
        Case "CREATE"
        
            'Create star
            Command11_Click
            
        Case "EXPORT"
                        
            'Show Common Dialog
            comdiag.Filter = "Bitmap|*.bmp"
            comdiag.DialogTitle = "Save as..."
            comdiag.InitDir = App.Path
            comdiag.ShowSave
            sfname = comdiag.filename
            
            'See if user cancelled
            If sfname = "" Then Exit Sub
                    
            'See if file exists
            If FileExist(sfname) = True Then
                MsgBox "The File already exist!", vbCritical, "Error"
                Exit Sub
            End If
                
            'Save file
            If (LCase$(Right$(sfname, 4)) = ".bmp") Then
                SavePicture P1.Image, sfname
            Else
                sfname = sfname & ".bmp"
                SavePicture P1.Image, sfname
            End If
            
        Case "OVERRIDE"
        
            'Activate override
            overide
        
        Case "GEO"
        
            Option1_Click
            Option1.Value = True
            Option2.Value = False
            Toolbar1.Buttons.Item(7).Value = 1
            Toolbar1.Buttons.Item(8).Value = 0
            
        Case "GRID"
            Option2_Click
            Option2.Value = True
            Option1.Value = False
            Toolbar1.Buttons.Item(7).Value = 0
            Toolbar1.Buttons.Item(8).Value = 1
        
        Case "EXIT"
            
            'Exit
            Unload Me
            End

    End Select

End Sub

Private Sub Form_Resize()

    'If form is smaller than original height then stop it!
    If Me.Height < 6180 Then Me.Height = 6180
    If Me.Width < 9315 Then Me.Width = 9315
    
    'Resize all controls to follow form
    P1.Height = Height - 400 - 360
    con.Left = Width - con.Width - 195
    P1.Width = con.Left - 120
    con.Height = Height - 465 - 360
    Command10.Top = con.Height - Command10.Height - 180
    Command11.Top = con.Height - Command11.Height - 180
    Command9.Top = con.Height - Command9.Height - 180
    
    
    'Clear screen
    P1.Cls
    

End Sub



Private Sub Command10_Click()

    'Clear screen
    P1.Cls

End Sub
Private Sub Command9_Click()

    'Exit
    Unload Me
    End

End Sub

Private Sub Command11_Click()
    
    'Do Backcolor
    P1.BackColor = RGB(CInt(Textr.Text), CInt(Textg.Text), CInt(Textb.Text))
    
    'Draw star
     P1.DrawWidth = CInt(ps.Text)
    If gstar = False Then Gridstar
    If gstar = True Then GeoStar

End Sub

Public Function sreset()

    'Reset all controls and clear screen
    Text3.Text = "r"
    Text4.Text = "r"
    Text5.Text = "r"
    alength.Text = 1200
    aangle.Text = 4
    aslength.Text = 540
    aspins.Text = 50
    Text7.Text = 8
    Text1.Text = 30
    Text2.Text = 3000
     P1.Cls
    ps.Text = 1
    sh.Value = 1
    
End Function

Private Sub Option2_Click()

    'Change to Gridstar
    If Option1.Value = True Then gstar = True
    If Option1.Value = False Then gstar = False
    Frame1.Visible = True
    Frame2.Visible = False
    Frame5.Visible = False
    
    Toolbar1.Buttons.Item(8).Style = tbrCheck
    Toolbar1.Buttons.Item(7).Style = tbrCheck
    
    Toolbar1.Buttons.Item(7).Value = tbrUnpressed
    Toolbar1.Buttons.Item(8).Value = tbrPressed

End Sub

Private Sub Option1_Click()

    'Change to Geostar
    If Option1.Value = True Then gstar = True
    If Option1.Value = False Then gstar = False
    Frame2.Visible = True
    Frame1.Visible = False
    Frame5.Visible = False
    
    Toolbar1.Buttons.Item(8).Style = tbrCheck
    Toolbar1.Buttons.Item(7).Style = tbrCheck
    
    Toolbar1.Buttons.Item(7).Value = tbrPressed
    Toolbar1.Buttons.Item(8).Value = tbrUnpressed

End Sub

Private Sub Option3_Click()

    'Hide frames, show Option frame
    Frame5.Visible = True
    Frame1.Visible = False
    Frame2.Visible = False

End Sub

Private Sub Text1_Gotfocus()

    'Select all of the text
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)

End Sub

Private Sub Text1_LostFocus()

    'If override is active, exit
    If override = True Then Exit Sub
    On Error GoTo ErrorHandler
    
    'Check to see if value is greator than limit
    If CDbl(Text1.Text) > 100 Then Text1.Text = "100"
    If CDbl(Text1.Text) < 1 Then Text1.Text = "1"
    
    Exit Sub
    
ErrorHandler:

    
End Sub

Private Sub Text2_Gotfocus()

    'Select all of the text
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)

End Sub

Private Sub Text2_LostFocus()

    'If override is active, exit
    If override = True Then Exit Sub
    On Error GoTo ErrorHandler
    
    'Check to see if value is greator than limit
    If CDbl(Text2.Text) < 500 Then Text2.Text = "500"
    
    Exit Sub
    
ErrorHandler:

    
End Sub

Private Sub Text3_Gotfocus()

    'Select all of the text
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)

End Sub

Private Sub Text3_LostFocus()
    
    'If override is active, exit
    If override = True Then Exit Sub
    If Text3.Text = "r" Then Exit Sub
        
    On Error GoTo ErrorHandler
    
    'Check to see if value is greator than limit
    If CDbl(Text3.Text) > 255 Then Text3.Text = "255"
    If CDbl(Text3.Text) < 0 Then Text3.Text = "0"
    
    Exit Sub
    
ErrorHandler:

    
End Sub

Private Sub Text4_Gotfocus()

    'Select all of the text
    Text4.SelStart = 0
    Text4.SelLength = Len(Text4.Text)

End Sub



Private Sub Text4_LostFocus()

    'If override is active, exit
    If override = True Then Exit Sub
    If Text4.Text = "r" Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    'Check to see if value is greator than limit
    If CDbl(Text4.Text) > 255 Then Text4.Text = "255"
    If CDbl(Text4.Text) < 0 Then Text4.Text = "0"
    
    Exit Sub
    
ErrorHandler:
    

End Sub

Private Sub Text5_GotFocus()

    'Select all of the text
    Text5.SelStart = 0
    Text5.SelLength = Len(Text5.Text)

End Sub


Private Sub Text5_LostFocus()

    'If override is active, exit
    If override = True Then Exit Sub
    If Text5.Text = "r" Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    'Check to see if value is greator than limit
    If CDbl(Text5.Text) > 255 Then Text5.Text = "255"
    If CDbl(Text5.Text) < 0 Then Text5.Text = "0"
    
    Exit Sub
    
ErrorHandler:


End Sub

Private Sub Text7_Gotfocus()

    'If override is active, exit
    If override = True Then Exit Sub
    
    'Select all of the text
    Text7.SelStart = 0
    Text7.SelLength = Len(Text7.Text)
    
End Sub


Private Sub Text7_LostFocus()

    'If override is active, exit
    If override = True Then Exit Sub
    On Error GoTo ErrorHandler
    
    'Check to see if value is greator than limit
    If CDbl(Text7.Text) > 1080 Then Text7.Text = "1080"
    If CDbl(Text7.Text) < 3 Then Text7.Text = "3"
    Exit Sub
    
ErrorHandler:
    

End Sub

Private Function overide()

    'Declare Variable
    Dim msg As String

    If override = True Then
    
        'Turn off override
        override = False
        
    Else
                
        'Provide warning msg
        msg = MsgBox("Warning: Override has been activated!" & Chr(13) & Chr(13) & "Override gives you complete freedom over star properties. " & Chr(13) & "Using high numbers could cause your computer to freeze, " & Chr(13) & "So please use at your own risk!" & Chr(13) & Chr(13) & "Do you wish to Continue?", vbExclamation + vbYesNo, "Override")
        
        'Check results
        If msg = vbYes Then
        
            'Turn On Override
            override = True
            
        Else
            
            'Turn Off Override
            override = False
            
        End If
        
    End If
    

End Function

Private Sub Form_Unload(Cancel As Integer)

    'Exit
    Unload Me
    End

End Sub

Public Function FileExist(sfname As String) As Boolean

    'See if file exists
    Dim temp1 As Long
    
    On Error Resume Next
    temp1 = GetAttr(sfname)
    
    If Err Then
    
        FileExist = False
        
    Else
    
        FileExist = True
        
    End If
    
End Function







Private Sub Textr_Gotfocus()

    'Select all of the text
    Textr.SelStart = 0
    Textr.SelLength = Len(Textr.Text)

End Sub

Private Sub Textr_LostFocus()
    
    'If override is active, exit
    If override = True Then Exit Sub
        
    On Error GoTo ErrorHandler
    
    'Check to see if value is greator than limit
    If CDbl(Textr.Text) > 255 Then Textr.Text = "255"
    If CDbl(Textr.Text) < 0 Then Textr.Text = "0"
    
    Exit Sub
    
ErrorHandler:

    
End Sub

Private Sub Textg_Gotfocus()

    'Select all of the text
    Textg.SelStart = 0
    Textg.SelLength = Len(Textg.Text)

End Sub



Private Sub Textg_LostFocus()

    'If override is active, exit
    If override = True Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    'Check to see if value is greator than limit
    If CDbl(Textg.Text) > 255 Then Textg.Text = "255"
    If CDbl(Textg.Text) < 0 Then Textg.Text = "0"
    
    Exit Sub
    
ErrorHandler:
    

End Sub

Private Sub Textb_GotFocus()

    'Select all of the text
    Textb.SelStart = 0
    Textb.SelLength = Len(Textb.Text)

End Sub


Private Sub Textb_LostFocus()

    'If override is active, exit
    If override = True Then Exit Sub
    If Textb.Text = "r" Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    'Check to see if value is greator than limit
    If CDbl(Textb.Text) > 255 Then Textb.Text = "255"
    If CDbl(Textb.Text) < 0 Then Textb.Text = "0"
    
    Exit Sub
    
ErrorHandler:


End Sub

Public Function Gridstar()

    'Declare Variables
    Dim x1
    Dim x2
    Dim y1
    Dim y2
    Dim Midx
    Dim Midy
    Dim inte
    Dim length
    Dim points
    Dim r
    Dim g
    Dim b
    Dim temp
    Dim angle
    Dim x
    Dim Y
    Dim msg
            
    'Clear screen
     P1.Cls
     
    'Declare points
    points = CDbl(Text7.Text)
    angle = 360 / points
        
    'Declare density and size of star
    length = CDbl(Text2.Text)
    inte = CDbl(Text1.Text)
                        
                        
             
 'Do shadow
 If sh.Value = 1 Then
                        
    'Get middle of form
    Midx = P1.Width / 2
    Midy = P1.Height / 2 + 10


    'Create Basic outline shape of star
    For x = 1 To points Step 1
    
        'Determine line co-ordinates
        x1 = Midx
        y1 = Midy
        x2 = length * Cos(PI / 180 * (angle * x - 90)) + Midx
        y2 = length * Sin(PI / 180 * (angle * x - 90)) + Midy

            'Detect if colour red should be random or not
            If Text3.Text = "r" Then
                Randomize
                r = 255 * Rnd
            Else
                r = CDbl(Text3.Text)
            End If
            
            'Detect if colour green should be random or not
            If Text4.Text = "r" Then
                Randomize
                g = 255 * Rnd
            Else
                g = CDbl(Text4.Text)
            End If
            
            'Detect if colour blue should be random or not
            If Text5.Text = "r" Then
                Randomize
                b = 255 * Rnd
            Else
                b = CDbl(Text5.Text)
            End If
        
        'Draw line
         P1.Line (x1, y1)-(x2, y2), RGB(150, 150, 150)

    Next x


    'Determine which pie section of the star to draw
    For x = 1 To points Step 1
    
        'Determine which lines to draw
        For Y = 0 To length Step (length / inte)
        
            'Determine line co-ordinates
            x1 = (length - Y) * Cos(PI / 180 * (angle * x - 90)) + Midx
            y1 = (length - Y) * Sin(PI / 180 * (angle * x - 90)) + Midy
            x2 = (length - (length - Y)) * Cos(PI / 180 * (angle * (x + 1) - 90)) + Midx
            y2 = (length - (length - Y)) * Sin(PI / 180 * (angle * (x + 1) - 90)) + Midy

                'Detect if colour red should be random or not
                If Text3.Text = "r" Then
                    Randomize
                    r = 255 * Rnd
                Else
                    r = CDbl(Text3.Text)
                End If
                
                'Detect if colour green should be random or not
                If Text4.Text = "r" Then
                    Randomize
                    g = 255 * Rnd
                Else
                    g = CDbl(Text4.Text)
                End If
                
                'Detect if colour blue should be random or not
                If Text5.Text = "r" Then
                    Randomize
                    b = 255 * Rnd
                Else
                    b = CDbl(Text5.Text)
                End If
                
            'Draw line
             P1.Line (x1, y1)-(x2, y2), RGB(150, 150, 150)

        Next Y
    Next x
    
End If
                        
'Do normal Star

    'Get middle of form
    Midx = P1.Width / 2
    Midy = P1.Height / 2


    'Create Basic outline shape of star
    For x = 1 To points Step 1
    
        'Determine line co-ordinates
        x1 = Midx
        y1 = Midy
        x2 = length * Cos(PI / 180 * (angle * x - 90)) + Midx
        y2 = length * Sin(PI / 180 * (angle * x - 90)) + Midy

            'Detect if colour red should be random or not
            If Text3.Text = "r" Then
                Randomize
                r = 255 * Rnd
            Else
                r = CDbl(Text3.Text)
            End If
            
            'Detect if colour green should be random or not
            If Text4.Text = "r" Then
                Randomize
                g = 255 * Rnd
            Else
                g = CDbl(Text4.Text)
            End If
            
            'Detect if colour blue should be random or not
            If Text5.Text = "r" Then
                Randomize
                b = 255 * Rnd
            Else
                b = CDbl(Text5.Text)
            End If
        
        'Draw line
         P1.Line (x1, y1)-(x2, y2), RGB(r, g, b)

    Next x


    'Determine which pie section of the star to draw
    For x = 1 To points Step 1
    
        'Determine which lines to draw
        For Y = 0 To length Step (length / inte)
        
            'Determine line co-ordinates
            x1 = (length - Y) * Cos(PI / 180 * (angle * x - 90)) + Midx
            y1 = (length - Y) * Sin(PI / 180 * (angle * x - 90)) + Midy
            x2 = (length - (length - Y)) * Cos(PI / 180 * (angle * (x + 1) - 90)) + Midx
            y2 = (length - (length - Y)) * Sin(PI / 180 * (angle * (x + 1) - 90)) + Midy

                'Detect if colour red should be random or not
                If Text3.Text = "r" Then
                    Randomize
                    r = 255 * Rnd
                Else
                    r = CDbl(Text3.Text)
                End If
                
                'Detect if colour green should be random or not
                If Text4.Text = "r" Then
                    Randomize
                    g = 255 * Rnd
                Else
                    g = CDbl(Text4.Text)
                End If
                
                'Detect if colour blue should be random or not
                If Text5.Text = "r" Then
                    Randomize
                    b = 255 * Rnd
                Else
                    b = CDbl(Text5.Text)
                End If
                
            'Draw line
             P1.Line (x1, y1)-(x2, y2), RGB(r, g, b)

        Next Y
    Next x

End Function


Public Function GeoStar()

    'Clear Screen
     P1.Cls

    'Declare Variables
    Dim angle, x, r, g, b, x1, x2, y1, y2, length, b1, b2, Y, spins, slength
    
    'Get values
    length = alength.Text
    angle = aangle.Text
    spins = aspins.Text
    slength = aslength.Text
    
    'Change values to integers
    angle = 360 / CInt(angle)
    length = CInt(length)
    spins = CInt(spins)
    slength = CInt(slength)
    
'Do Shadow
If sh.Value = 1 Then

    'See which shape to draw
    For Y = 0 To 360 Step (360 / spins)
    
            'Get coordinates
            b1 = slength * Cos(PI / 180 * (Y)) + P1.Height / 2 + 15
            b2 = slength * Sin(PI / 180 * (Y)) + P1.Width / 2
            x2 = b2
            y2 = b1
        
        'Draw shape
        For x = 0 To 360 / angle Step 1
                    
            'Get coordinates
            x1 = x2
            y1 = y2
            
            x2 = length * Cos(PI / 180 * ((angle * x) - Y)) + b2
            y2 = length * Sin(PI / 180 * ((angle * x) - Y)) + b1
            
            'Draw Line
            If x <> 0 Then P1.Line (x1, y1)-(x2, y2), RGB(150, 150, 150)
            
        Next x
            
    Next Y

End If

'Draw Normal star
For Y = 0 To 360 Step (360 / spins)

        'Get coordinates
        b1 = slength * Cos(PI / 180 * (Y)) + P1.Height / 2
        b2 = slength * Sin(PI / 180 * (Y)) + P1.Width / 2
        x2 = b2
        y2 = b1
    
    'Draw Shape
    For x = 0 To 360 / angle Step 1
                
        'Get coordinates
        x1 = x2
        y1 = y2
        
        x2 = length * Cos(PI / 180 * ((angle * x) - Y)) + b2
        y2 = length * Sin(PI / 180 * ((angle * x) - Y)) + b1
        
            'Detect if colour red should be random or not
            If Text3.Text = "r" Then
                Randomize
                r = 255 * Rnd
            Else
                r = CDbl(Text3.Text)
            End If
            
            'Detect if colour green should be random or not
            If Text4.Text = "r" Then
                Randomize
                g = 255 * Rnd
            Else
                g = CDbl(Text4.Text)
            End If
            
            'Detect if colour blue should be random or not
            If Text5.Text = "r" Then
                Randomize
                b = 255 * Rnd
            Else
                b = CDbl(Text5.Text)
            End If
        
        'Draw line
        If x <> 0 Then P1.Line (x1, y1)-(x2, y2), RGB(r, g, b)
        
    Next x
        
Next Y

End Function
