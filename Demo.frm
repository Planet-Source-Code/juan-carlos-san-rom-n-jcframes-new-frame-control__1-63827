VERSION 5.00
Object = "*\AjcFrames.vbp"
Begin VB.Form Demo 
   AutoRedraw      =   -1  'True
   Caption         =   "jcFrames"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6900
   Icon            =   "Demo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   6900
   StartUpPosition =   2  'CenterScreen
   Begin jcFramesOCX.jcFrames jcFrames 
      Height          =   1905
      Left            =   2820
      Top             =   150
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   3360
      Caption         =   ""
      IconAlignment   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "Exit Demo"
         Height          =   375
         Left            =   1410
         TabIndex        =   27
         Top             =   1350
         Width           =   1185
      End
   End
   Begin jcFramesOCX.jcFrames jcFrames7 
      Height          =   4725
      Left            =   180
      Top             =   2190
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   8334
      FillColor       =   14745599
      TextBoxColor    =   11595760
      Style           =   3
      Caption         =   "jcFrames v1.1 - main features"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Demo.frx":08CA
      IconSize        =   32
      Begin jcFramesOCX.jcFrames jcFrames2 
         Height          =   1875
         Index           =   1
         Left            =   2550
         Top             =   630
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   3307
         FrameColor      =   6974058
         BackColor       =   14745599
         FillColor       =   14745599
         Style           =   0
         Caption         =   "Properties for GradientFrame style"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox CboColorTo 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "Demo.frx":11A4
            Left            =   2070
            List            =   "Demo.frx":11DA
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   1410
            Width           =   1575
         End
         Begin VB.ComboBox CboColorFrom 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "Demo.frx":121C
            Left            =   2070
            List            =   "Demo.frx":1252
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1050
            Width           =   1575
         End
         Begin VB.ComboBox CboThemeColor 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":128F
            Left            =   2070
            List            =   "Demo.frx":12A5
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   690
            Width           =   1575
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ColorTo:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   300
            TabIndex        =   32
            Top             =   1440
            Width           =   600
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ColorFrom:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   300
            TabIndex        =   30
            Top             =   1080
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ThemeColor:"
            Height          =   195
            Left            =   300
            TabIndex        =   4
            Top             =   750
            Width           =   900
         End
      End
      Begin jcFramesOCX.jcFrames jcFrames2 
         Height          =   1875
         Index           =   0
         Left            =   2550
         Top             =   630
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   3307
         FrameColor      =   6974058
         BackColor       =   14745599
         FillColor       =   14745599
         Style           =   0
         Caption         =   "Properties for XpDefault style"
         Alignment       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox CboFrameColor 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":12DE
            Left            =   2070
            List            =   "Demo.frx":1310
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   330
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FrameColor:"
            Height          =   195
            Left            =   300
            TabIndex        =   5
            Top             =   390
            Width           =   840
         End
      End
      Begin jcFramesOCX.jcFrames jcFrames2 
         Height          =   1875
         Index           =   2
         Left            =   2550
         Top             =   630
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   3307
         FrameColor      =   6974058
         BackColor       =   14745599
         FillColor       =   14745599
         Style           =   0
         RoundedCornerTxtBox=   -1  'True
         Caption         =   "Properties for TextBox style"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox CboFillColor 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":133C
            Left            =   2070
            List            =   "Demo.frx":1375
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   1410
            Width           =   1575
         End
         Begin VB.ComboBox CboTextBoxColor 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":13AD
            Left            =   2070
            List            =   "Demo.frx":13F7
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   690
            Width           =   1575
         End
         Begin VB.ComboBox CboRoundTxtBox 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":1435
            Left            =   2070
            List            =   "Demo.frx":143F
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1050
            Width           =   1575
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FillColor:"
            Height          =   195
            Left            =   300
            TabIndex        =   9
            Top             =   1440
            Width           =   585
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RoundedCornerTxtBox:"
            Height          =   195
            Left            =   300
            TabIndex        =   8
            Top             =   1080
            Width           =   1665
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TextBoxColor:"
            Height          =   195
            Left            =   300
            TabIndex        =   6
            Top             =   750
            Width           =   990
         End
      End
      Begin jcFramesOCX.jcFrames jcFrames3 
         Height          =   1905
         Left            =   150
         Top             =   2640
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   3360
         FrameColor      =   6974058
         BackColor       =   14745599
         FillColor       =   14745599
         TextBoxColor    =   11595760
         Style           =   2
         RoundedCornerTxtBox=   -1  'True
         Caption         =   "General features"
         TextBoxHeight   =   18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox CboIconAlign 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "Demo.frx":1450
            Left            =   4680
            List            =   "Demo.frx":145A
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1110
            Width           =   1425
         End
         Begin VB.ComboBox CboTextBoxHeight 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":147D
            Left            =   1680
            List            =   "Demo.frx":149C
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1470
            Width           =   855
         End
         Begin VB.ComboBox CboIconSize 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "Demo.frx":14C4
            Left            =   4680
            List            =   "Demo.frx":14D1
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   750
            Width           =   1035
         End
         Begin VB.ComboBox CboPicture 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":14E1
            Left            =   4680
            List            =   "Demo.frx":14EB
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   420
            Width           =   1035
         End
         Begin VB.ComboBox CboRoundCorner 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":14F8
            Left            =   1680
            List            =   "Demo.frx":1502
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   750
            Width           =   1035
         End
         Begin VB.ComboBox CboCaptionAlig 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":1513
            Left            =   1680
            List            =   "Demo.frx":1520
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   420
            Width           =   1425
         End
         Begin VB.ComboBox CboTextColor 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":154D
            Left            =   1680
            List            =   "Demo.frx":1576
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1110
            Width           =   1305
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Icon Alignment:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   3540
            TabIndex        =   35
            Top             =   1170
            Width           =   1095
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TextBoxHeight:"
            Height          =   195
            Left            =   330
            TabIndex        =   26
            Top             =   1530
            Width           =   1095
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IconSize:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   3540
            TabIndex        =   23
            Top             =   810
            Width           =   660
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Picture:"
            Height          =   195
            Left            =   3540
            TabIndex        =   21
            Top             =   480
            Width           =   540
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RoundedCorner:"
            Height          =   195
            Left            =   330
            TabIndex        =   15
            Top             =   810
            Width           =   1170
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caption Alignment:"
            Height          =   195
            Left            =   330
            TabIndex        =   14
            Top             =   480
            Width           =   1320
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TextColor:"
            Height          =   195
            Left            =   330
            TabIndex        =   13
            Top             =   1170
            Width           =   720
         End
      End
      Begin jcFramesOCX.jcFrames jcFrames1 
         Height          =   1905
         Left            =   150
         Top             =   600
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   3360
         FrameColor      =   6974058
         BackColor       =   14745599
         FillColor       =   14745599
         TextBoxColor    =   11595760
         Style           =   2
         RoundedCornerTxtBox=   -1  'True
         Caption         =   "Frame styles"
         TextBoxHeight   =   18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.OptionButton OptCaptionStyle 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0FFFF&
            Caption         =   "Messenger"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   4
            Left            =   120
            TabIndex        =   33
            Tag             =   "Messenger style"
            Top             =   1530
            Width           =   2085
         End
         Begin VB.OptionButton OptCaptionStyle 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0FFFF&
            Caption         =   "Windows"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   3
            Left            =   120
            TabIndex        =   28
            Tag             =   "Windows style"
            Top             =   1245
            Width           =   2085
         End
         Begin VB.OptionButton OptCaptionStyle 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0FFFF&
            Caption         =   "TextBox (from EZFrame)"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   2
            Left            =   120
            TabIndex        =   2
            Tag             =   "Text box style"
            ToolTipText     =   "Thanks to ElectroZ for his frame style"
            Top             =   960
            Width           =   2085
         End
         Begin VB.OptionButton OptCaptionStyle 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0FFFF&
            Caption         =   "GradientFrame"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   1
            Tag             =   "Gradient frame style"
            Top             =   675
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton OptCaptionStyle 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0FFFF&
            Caption         =   "XpDefault"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Tag             =   "Xp default style"
            Top             =   390
            Width           =   1815
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   2430
      Picture         =   "Demo.frx":159A
      Top             =   960
      Width           =   240
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "jcFrames v1.1 with 5 styles"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   915
      Left            =   360
      TabIndex        =   18
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "jcFrames v1.1 with 5 styles"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   915
      Left            =   390
      TabIndex        =   19
      Top             =   750
      Width           =   1815
   End
End
Attribute VB_Name = "Demo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blnFormLoaded As Boolean

Private Sub CboCaptionAlig_Click()
    If blnFormLoaded = True Then jcFrames.Alignment = CboCaptionAlig.ListIndex
End Sub

Private Sub CboColorFrom_Click()
    If CboColorFrom.ListIndex = -1 Then Exit Sub
    jcFrames.ColorFrom = CboColorFrom.ItemData(CboColorFrom.ListIndex)
End Sub

Private Sub CboColorTo_Click()
    If CboColorTo.ListIndex = -1 Then Exit Sub
    jcFrames.ColorTo = CboColorTo.ItemData(CboColorTo.ListIndex)
End Sub

Private Sub CboFillColor_Click()
    If blnFormLoaded = True Then
        If CboFillColor.ListIndex = 4 Then
            jcFrames.FillColor = jcFrames.BackColor
        Else
            jcFrames.FillColor = CboFillColor.ItemData(CboFillColor.ListIndex)
        End If
    End If
End Sub

Private Sub CboFrameColor_Click()
    If blnFormLoaded = True Then jcFrames.FrameColor = CboFrameColor.ItemData(CboFrameColor.ListIndex)
End Sub

Private Sub CboIconAlign_Click()
    If CboIconAlign.ListIndex = -1 Then Exit Sub
    If blnFormLoaded = True Then jcFrames.IconAlignment = CboIconAlign.ListIndex
End Sub

Private Sub CboIconSize_Click()
    If CboIconSize.ListIndex = -1 Then Exit Sub
    jcFrames.IconSize = Val(CboIconSize.Text)
End Sub

Private Sub CboPicture_Click()
    Select Case CboPicture.Text
        Case "Yes"
            Set jcFrames.Picture = LoadPicture(App.Path & "\103_56.ico")
            CboIconSize.Enabled = True
            CboIconAlign.Enabled = True
            Label16.Enabled = True
            Label13.Enabled = True
            CboIconSize.ListIndex = 0
            CboIconAlign.ListIndex = 0
        Case "No"
            Set jcFrames.Picture = Nothing
            CboIconSize.Enabled = False
            CboIconAlign.Enabled = False
            Label16.Enabled = False
            Label13.Enabled = False
            CboIconSize.ListIndex = -1
            CboIconAlign.ListIndex = -1
    End Select
End Sub

Private Sub CboRoundCorner_Click()
    If blnFormLoaded = True Then jcFrames.RoundedCorner = CboRoundCorner.ListIndex
End Sub

Private Sub CboRoundTxtBox_Click()
    If blnFormLoaded = True Then jcFrames.RoundedCornerTxtBox = CboRoundTxtBox.ListIndex
End Sub

Private Sub CboTextBoxColor_Click()
    If blnFormLoaded = True Then jcFrames.TextboxColor = CboTextBoxColor.ItemData(CboTextBoxColor.ListIndex)
End Sub

Private Sub CboTextBoxHeight_Click()
    If blnFormLoaded = True Then jcFrames.TextBoxHeight = Val(CboTextBoxHeight.Text)
End Sub

Private Sub CboTextColor_Click()
    If blnFormLoaded = True Then jcFrames.TextColor = CboTextColor.ItemData(CboTextColor.ListIndex)
End Sub

Private Sub CboThemeColor_Click()
    If blnFormLoaded = True Then
        If CboThemeColor.ListIndex = 5 Then
            CboColorFrom.Enabled = True
            CboColorTo.Enabled = True
            CboColorFrom.ListIndex = 4
            CboColorTo.ListIndex = 4
            Label14.Enabled = True
            Label15.Enabled = True
            Label2.Enabled = True
            CboFrameColor.Enabled = True
        Else
            CboColorFrom.Enabled = False
            CboColorTo.Enabled = False
            CboColorFrom.ListIndex = -1
            CboColorTo.ListIndex = -1
            Label14.Enabled = False
            Label15.Enabled = False
            Label2.Enabled = False
            CboFrameColor.Enabled = False
        End If
        jcFrames.ThemeColor = CboThemeColor.ListIndex
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    blnFormLoaded = False
    CboThemeColor.ListIndex = 0
    CboTextBoxHeight.ListIndex = 2
    CboCaptionAlig.ListIndex = 2
    CboRoundCorner.ListIndex = 1
    CboTextColor.ListIndex = 0
    CboPicture.ListIndex = 1
    CboFrameColor.ListIndex = 0
    CboRoundTxtBox.ListIndex = 1
    CboTextBoxColor.ListIndex = 3
    CboFillColor.ListIndex = 4
    blnFormLoaded = True
    jcFrames.Caption = OptCaptionStyle(1).Tag
    Set Label2.Container = jcFrames2(1)
    Set CboFrameColor.Container = jcFrames2(1)
    Label2.Enabled = False
    CboFrameColor.Enabled = False
End Sub

Private Sub OptCaptionStyle_Click(Index As Integer)
    Dim i As Integer
    
    If blnFormLoaded = False Then Exit Sub
    jcFrames.Style = Index
    jcFrames.Caption = OptCaptionStyle(Index).Tag
    
    'blnFormLoaded = False
    
    For i = 0 To 2
        If i = Index Then
            jcFrames2(i).Visible = True
            Set Label2.Container = jcFrames2(i)
            Set CboFrameColor.Container = jcFrames2(i)
        Else
            jcFrames2(i).Visible = False
        End If
    Next i
    
    If Index = 3 Then
        jcFrames2(2).Visible = True
        Set Label2.Container = jcFrames2(2)
        Set CboFrameColor.Container = jcFrames2(2)
        jcFrames2(2).Caption = "Properties for Windows style"
    Else
        jcFrames2(2).Caption = "Properties for TextBox style"
    End If
    
    If Index = 4 Then
        jcFrames2(1).Visible = True
        Set Label2.Container = jcFrames2(1)
        Set CboFrameColor.Container = jcFrames2(1)
        jcFrames2(1).Caption = "Properties for Messenger style"
    Else
        jcFrames2(1).Caption = "Properties for GradientFrame style"
    End If
    
    If Index = 1 Or Index = 4 Then
        If CboThemeColor.ListIndex = 5 Then
            Label2.Enabled = True
            CboFrameColor.Enabled = True
        Else
            Label2.Enabled = False
            CboFrameColor.Enabled = False
        End If
    Else
        Label2.Enabled = True
        CboFrameColor.Enabled = True
    End If
    
    Select Case Index
        Case 0
            CboPicture.Visible = False
            CboIconSize.Visible = False
            CboIconAlign.Visible = False
            Label16.Visible = False
            Label12.Visible = False
            Label13.Visible = False
            CboTextBoxHeight.Visible = False
            Label4.Visible = False
            
            'Default conditions
            CboCaptionAlig.ListIndex = 0
            CboTextColor.ListIndex = 1
            CboFrameColor.ListIndex = 3
        Case 1, 4
            CboPicture.Visible = True
            CboIconSize.Visible = True
            CboIconAlign.Visible = True
            Label16.Visible = True
            Label12.Visible = True
            Label13.Visible = True
            CboTextBoxHeight.Visible = True
            Label4.Visible = True
            
            If Index = 4 Then
                CboRoundCorner.ListIndex = 0
            Else
                CboRoundCorner.ListIndex = 1
            End If
            
            'Default conditions
            CboCaptionAlig.ListIndex = 2
            CboTextColor.ListIndex = 0
            CboTextBoxHeight.ListIndex = 2

        Case 2, 3
            CboPicture.Visible = True
            CboIconSize.Visible = True
            CboIconAlign.Visible = True
            Label16.Visible = True
            Label12.Visible = True
            Label13.Visible = True
            CboTextBoxHeight.Visible = True
            Label4.Visible = True
    
            'Default conditions
            CboCaptionAlig.ListIndex = 2
            CboTextColor.ListIndex = 0
            If Index = 2 Then
                CboRoundTxtBox.ListIndex = 1
                CboFillColor.ListIndex = 4
                CboFrameColor.ListIndex = 5
            Else
                CboRoundTxtBox.ListIndex = 0
                CboFillColor.ListIndex = 5
                CboFrameColor.ListIndex = 0
            End If
            CboRoundCorner.ListIndex = 1
            CboTextBoxColor.ListIndex = 4
            CboTextBoxHeight.ListIndex = 2
    End Select
    'blnFormLoaded = True
    jcFrames.BackColor = BackColor
End Sub
