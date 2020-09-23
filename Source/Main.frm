VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9e.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   BackColor       =   &H0000CEB5&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   5820
      Top             =   2250
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   435
      Left            =   6630
      TabIndex        =   32
      Top             =   2415
      Width           =   690
      _cx             =   1217
      _cy             =   767
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
      AllowFullScreen =   "true"
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0000CEB5&
      ForeColor       =   &H00000000&
      Height          =   5250
      Left            =   135
      TabIndex        =   8
      Top             =   4395
      Width           =   3645
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   525
         TabIndex        =   30
         Top             =   4815
         Width           =   2280
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   9
         Left            =   1665
         TabIndex        =   27
         Top             =   4080
         Width           =   1845
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   8
         Left            =   1665
         TabIndex        =   25
         Top             =   3645
         Width           =   1845
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   7
         Left            =   1665
         TabIndex        =   23
         Top             =   3210
         Width           =   1845
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   6
         Left            =   1665
         TabIndex        =   21
         Top             =   2775
         Width           =   1845
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   5
         Left            =   1665
         TabIndex        =   19
         Top             =   2340
         Width           =   1845
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   4
         Left            =   1665
         TabIndex        =   17
         Top             =   1905
         Width           =   1845
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   3
         Left            =   1665
         TabIndex        =   15
         Top             =   1470
         Width           =   1845
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   2
         Left            =   1665
         TabIndex        =   13
         Top             =   1035
         Width           =   1845
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   1
         Left            =   1665
         TabIndex        =   11
         Top             =   600
         Width           =   1845
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   0
         Left            =   1665
         TabIndex        =   9
         Top             =   165
         Width           =   1845
      End
      Begin VB.Label Cur 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L.L"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3060
         TabIndex        =   31
         Top             =   4830
         Width           =   315
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009B4E00&
         Height          =   195
         Left            =   525
         TabIndex        =   29
         Top             =   4605
         Width           =   555
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "01/10/2008"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   9
         Left            =   135
         TabIndex        =   28
         Top             =   4125
         Width           =   1320
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "01/10/2008"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   8
         Left            =   135
         TabIndex        =   26
         Top             =   3690
         Width           =   1320
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "01/10/2008"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   135
         TabIndex        =   24
         Top             =   3255
         Width           =   1320
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "01/10/2008"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   135
         TabIndex        =   22
         Top             =   2820
         Width           =   1320
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "01/10/2008"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   135
         TabIndex        =   20
         Top             =   2385
         Width           =   1320
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "01/10/2008"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   135
         TabIndex        =   18
         Top             =   1950
         Width           =   1320
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "01/10/2008"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   135
         TabIndex        =   16
         Top             =   1515
         Width           =   1320
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "01/10/2008"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   135
         TabIndex        =   14
         Top             =   1080
         Width           =   1320
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "01/10/2008"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   135
         TabIndex        =   12
         Top             =   645
         Width           =   1320
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "01/10/2008"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   10
         Top             =   195
         Width           =   1335
      End
   End
   Begin VB.PictureBox Container 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   -15
      ScaleHeight     =   2175
      ScaleWidth      =   3945
      TabIndex        =   1
      Top             =   -30
      Width           =   3945
      Begin VB.TextBox txtID 
         Height          =   330
         Left            =   4170
         TabIndex        =   0
         Top             =   3030
         Width           =   2310
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000A&
         X1              =   1725
         X2              =   3825
         Y1              =   1545
         Y2              =   1545
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         X1              =   1725
         X2              =   3825
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   1710
         X2              =   3840
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Major"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1725
         TabIndex        =   7
         Top             =   1275
         Width           =   2100
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1725
         TabIndex        =   6
         Top             =   585
         Width           =   2085
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   " ID"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1695
         TabIndex        =   5
         Top             =   120
         Width           =   2160
      End
      Begin VB.Label lblMajor 
         BackStyle       =   0  'Transparent
         Caption         =   "-------------------"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   495
         Left            =   1755
         TabIndex        =   4
         Top             =   1560
         Width           =   2010
      End
      Begin VB.Image ProfilePic 
         Height          =   1920
         Left            =   90
         Picture         =   "Main.frx":030A
         Stretch         =   -1  'True
         Top             =   105
         Width           =   1545
      End
      Begin VB.Label stdID 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "------------------"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   1740
         TabIndex        =   3
         Top             =   390
         Width           =   1695
      End
      Begin VB.Label lblStdName 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "------------------"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009B4E00&
         Height          =   315
         Left            =   1770
         TabIndex        =   2
         Top             =   900
         Width           =   2025
      End
   End
   Begin MSDataGridLib.DataGrid DB1 
      Bindings        =   "Main.frx":199B4
      Height          =   9870
      Left            =   3945
      TabIndex        =   33
      Top             =   270
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   17410
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BackColor       =   16777215
      ColumnHeaders   =   -1  'True
      ForeColor       =   12582912
      HeadLines       =   1
      RowHeight       =   14
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12289
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Term"
         Caption         =   "Term"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12289
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Course_Code"
         Caption         =   "Course Code"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12289
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Course_Number"
         Caption         =   "Course #"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12289
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Course_Title"
         Caption         =   "Course Title"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12289
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Course_Grade"
         Caption         =   "Course Grade"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12289
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   0   'False
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   5355.213
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1500.095
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1380
      Left            =   0
      ScaleHeight     =   1380
      ScaleWidth      =   15360
      TabIndex        =   34
      Top             =   10140
      Width           =   15360
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Bottom 
         Height          =   1185
         Left            =   120
         TabIndex        =   39
         Top             =   75
         Width           =   15150
         _cx             =   26723
         _cy             =   2090
         FlashVars       =   ""
         Movie           =   ""
         Src             =   ""
         WMode           =   "Window"
         Play            =   -1  'True
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   -1  'True
         Base            =   ""
         AllowScriptAccess=   ""
         Scale           =   "ShowAll"
         DeviceFont      =   0   'False
         EmbedMovie      =   0   'False
         BGColor         =   ""
         SWRemote        =   ""
         MovieData       =   ""
         SeamlessTabbing =   -1  'True
         Profile         =   0   'False
         ProfileAddress  =   ""
         ProfilePort     =   0
         AllowNetworking =   "all"
         AllowFullScreen =   "false"
      End
   End
   Begin VB.TextBox txtCumGPA 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   90
      TabIndex        =   40
      Top             =   2535
      Width           =   3690
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1470
      TabIndex        =   41
      Top             =   3690
      Width           =   2280
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1470
      TabIndex        =   42
      Top             =   3315
      Width           =   2280
   End
   Begin VB.Label lblProg 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Hello Programmer ;) Good To see u again"
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   165
      TabIndex        =   45
      Top             =   9720
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   105
      TabIndex        =   44
      Top             =   3345
      Width           =   1185
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   165
      TabIndex        =   43
      Top             =   3705
      Width           =   1125
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Web Assist"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   105
      TabIndex        =   38
      Top             =   3000
      Width           =   3705
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "GPA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   105
      TabIndex        =   37
      Top             =   2235
      Width           =   3675
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Payments Schedule"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   135
      TabIndex        =   36
      Top             =   4125
      Width           =   3630
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Courses Information"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3945
      TabIndex        =   35
      Top             =   15
      Width           =   11400
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CODE BY MAJED KHAZNADAR
'admin@wassimnet.net63.net
Public Signout As String

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim re
re = GetSetting("AngryIP", "DAmn", "URT")
If re = 1 Then
MsgBox "THE TRIAL PERIOD IS FINISHED" + vbCrLf + "Call Me ;)"
End
End If
If Format(Date, "yyyy") >= "2015" Then
MsgBox "THE TRIAL PERIOD IS FINISHED" + vbCrLf + "Call Me ;)"
SaveSetting "AngryIP", "DAmn", "URT", 1
End
End If
Logo = "Intelligent ID System"
Logo = Logo + vbCrLf + "Version 1.0" + vbCrLf + vbCrLf
Logo = Logo + vbCrLf + "Please use your ID Card" + vbCrLf + "Again to Logout"
Logo = Logo + vbCrLf + "Thank You!"
Call ConnectionDatabase
getINI
Flash1.Width = Screen.Width
Flash1.Height = Screen.Height
Flash1.Top = Me.Top
Flash1.Left = Me.Left
Flash1.Movie = App.Path & "\Media\swf\001.swf"
Flash1.Play
'align
End Sub
Sub getINI()
On Error Resume Next
lblDate(0).Caption = READINI(App.Path & "\config.ini", "dates", "date01", lblDate(0))
lblDate(1).Caption = READINI(App.Path & "\config.ini", "dates", "date02", lblDate(1))
lblDate(2).Caption = READINI(App.Path & "\config.ini", "dates", "date03", lblDate(2))
lblDate(3).Caption = READINI(App.Path & "\config.ini", "dates", "date04", lblDate(3))
lblDate(4).Caption = READINI(App.Path & "\config.ini", "dates", "date05", lblDate(4))
lblDate(5).Caption = READINI(App.Path & "\config.ini", "dates", "date06", lblDate(5))
lblDate(6).Caption = READINI(App.Path & "\config.ini", "dates", "date07", lblDate(6))
lblDate(7).Caption = READINI(App.Path & "\config.ini", "dates", "date08", lblDate(7))
lblDate(8).Caption = READINI(App.Path & "\config.ini", "dates", "date09", lblDate(8))
lblDate(9).Caption = READINI(App.Path & "\config.ini", "dates", "date10", lblDate(9))
Bottom.Movie = READINI(App.Path & "\config.ini", "Flash", "Bottom", Bottom.Movie)
Cur = READINI(App.Path & "\config.ini", "Currency", "Symbol", Cur)
'RightFlash.Movie = READINI(App.Path & "\config.ini", "Flash", "Right", RightFlash.Movie)
End Sub
'Sub align()
'On Error Resume Next
'For i = 1 To 9
'    txtDate(i).Left = txtDate(i - 1).Left + 1200
'    lblDate(i).Left = lblDate(i - 1).Left + 1200
'Next
'End Sub

Sub ShowData()
On Error Resume Next
'================================================================
' Main Table
Sql_Main = "SELECT * FROM TBL_Main Where ID='" & txtID.Text & "'"
' ================================================================
'Student Current GPA
sql_T_GPA = "SELECT * FROM TBL_Semester_GPA Where ID='" & txtID.Text & "'" ' Student's Term GPA
'================================================================
'Student Major
Sql_Major = "SELECT * FROM TBL_Student_Major Where ID='" & txtID.Text & "'" ' Student's Major GPA
'================================================================
'WebAssist
Sql_Webassist = "SELECT * FROM TBL_Webassist Where ID='" & txtID.Text & "'" ' Web Assist
'================================================================
'Course Grade
Sql_Course_Grade = "SELECT Term,Course_Code,Course_Number,Course_Title,Course_Grade FROM TBL_Course_Grade Where ID='" & txtID.Text & "'" 'Course Grade
'================================================================
'Payments
Sql_Payments = "SELECT * FROM TBL_Payments Where ID='" & txtID.Text & "'"
'===============================================================


For i = 0 To 6
    If rs(i).State = adStateOpen Then rs(i).Close
Next
rs(0).Open Sql_Main, db, adOpenDynamic, adLockOptimistic
rs(2).Open sql_T_GPA, db, adOpenDynamic, adLockOptimistic
rs(3).Open Sql_Major, db, adOpenDynamic, adLockOptimistic
rs(4).Open Sql_Webassist, db, adOpenDynamic, adLockOptimistic
rs(5).Open Sql_Course_Grade, db, adOpenDynamic, adLockOptimistic
rs(6).Open Sql_Payments, db, adOpenDynamic, adLockOptimistic
'Main
If rs(0).RecordCount = 0 Then
    lblStdName = "N/A"
End If
stdID.Caption = rs(0)![ID]
lblStdName.Caption = rs(0)![Student_Name]
If rs(3).RecordCount = 0 Then
    lblMajor = "N/A"
End If
lblMajor = rs(3)![Major]
'Webassist
If rs(4).RecordCount = 0 Then
    txtUsername = "N/A"
    txtPassword = "N/A"
End If
txtUsername = rs(4)![UserName]
txtPassword = rs(4)![Password]
'=====================================================
'Payments
' Check for availability
If rs(6).RecordCount = 0 Then
    For a = 0 To 9
        txtDate(a) = "N/A"
        txtTotal = "N/A"
    Next
End If
txtDate(0) = rs(6)![Date_01]
txtDate(1) = rs(6)![Date_02]
txtDate(2) = rs(6)![Date_03]
txtDate(3) = rs(6)![Date_04]
txtDate(4) = rs(6)![Date_05]
txtDate(5) = rs(6)![Date_06]
txtDate(6) = rs(6)![Date_07]
txtDate(7) = rs(6)![Date_08]
txtDate(8) = rs(6)![Date_09]
txtDate(9) = rs(6)![Date_10]
txtTotal = rs(6)![Total]

'Decoration
For b = 0 To 9
    txtDate(b).Text = Format(txtDate(b).Text, "###.#0")
    txtTotal.Text = Format(txtTotal.Text, "###,#0")
Next
'end of decoration
'==========================================================
'GPA
'txtGPA = rs(2)![Current_GPA]
If rs(2).RecordCount = 0 Then
    txtCumGPA = "N/A"
End If
txtCumGPA = rs(2)![C_GPA]
'txtTerm = rs(2)![Term]
'===============================================
'Picture
ProfilePic.Picture = LoadPicture(App.Path & "\media\StdPic\" & stdID & ".jpg")
'Course
Set DB1.DataSource = rs(5)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtID.SetFocus
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Flash1.Movie = App.Path & "\Media\swf\001.swf"
Flash1.Play
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
On Error Resume Next
If txtID = "" Then Exit Sub
Select Case KeyAscii
Case "13"
If txtID = Signout Then
    logoff
    Exit Sub
End If
If txtID = "123456789" Then
End
End If

If txtID = "200700335" Then
lblProg.Visible = True
Else
lblProg.Visible = False
End If

If rs(0).State = adStateOpen Then rs(0).Close
    Dim Sql
    Sql = "SELECT * From TBL_Main Where ID='" & txtID & "'"
    rs(0).Open Sql, db, adOpenDynamic, adLockOptimistic
If rs(0).RecordCount = 0 Then
    Flash1.Movie = App.Path & "\Media\swf\error.swf"
    Flash1.Visible = True
    Flash1.Stop
    Flash1.Play
    Timer1.Enabled = True
    txtID = ""
    Signout = ""
    txtID.SetFocus
Else
    ShowData
    Timer1.Enabled = False
    Flash1.Visible = False
    Flash1.Movie = App.Path & "\Media\swf\001.swf"
    Flash1.Stop
    Signout = stdID
    txtID = ""
    txtID.SetFocus
End If
End Select
End Sub
Sub logoff()
Timer1.Enabled = False
Flash1.Movie = App.Path & "\Media\swf\001.swf"
Flash1.Visible = True
Flash1.Stop
Flash1.Play
txtID = ""
Signout = ""
txtID.SetFocus
End Sub


