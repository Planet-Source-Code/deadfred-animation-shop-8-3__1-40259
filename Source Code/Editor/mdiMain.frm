VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Animation Shop 8.3"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14250
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  'Manual
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer setResize 
      Interval        =   1
      Left            =   3480
      Top             =   3600
   End
   Begin MSComDlg.CommonDialog GetFile 
      Left            =   1080
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox SideFrame 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   8115
      Left            =   10350
      ScaleHeight     =   8115
      ScaleWidth      =   3900
      TabIndex        =   13
      Top             =   1230
      Visible         =   0   'False
      Width           =   3900
      Begin VB.ListBox List1 
         Height          =   255
         Left            =   360
         Sorted          =   -1  'True
         TabIndex        =   186
         Top             =   7680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame Frame 
         Caption         =   "Position03"
         Height          =   5775
         Index           =   15
         Left            =   240
         TabIndex        =   137
         Tag             =   "Animation2"
         Top             =   360
         Width           =   3615
         Begin Project1.EnterGrid grGrid 
            Height          =   3615
            Left            =   120
            TabIndex        =   195
            Top             =   240
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   6376
         End
         Begin VB.CommandButton cmdBF 
            Caption         =   "Next frame"
            Height          =   375
            Index           =   1
            Left            =   2160
            TabIndex        =   139
            Top             =   5160
            Width           =   1215
         End
         Begin VB.CommandButton cmdBF 
            Caption         =   "Previous frame"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   138
            Top             =   5160
            Width           =   1215
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Scenes08"
         Height          =   5775
         Index           =   14
         Left            =   240
         TabIndex        =   135
         Tag             =   "Animation1"
         Top             =   360
         Width           =   3615
         Begin VB.Frame FrameButtons 
            BorderStyle     =   0  'None
            Caption         =   "Frame10"
            Height          =   855
            Left            =   120
            TabIndex        =   187
            Top             =   3600
            Width           =   3495
            Begin VB.CommandButton cmdScene 
               Caption         =   "Rename"
               Height          =   375
               Index           =   0
               Left            =   2400
               TabIndex        =   188
               Top             =   360
               Width           =   1095
            End
            Begin VB.CommandButton cmdScene 
               Caption         =   "Remove"
               Height          =   375
               Index           =   1
               Left            =   2400
               TabIndex        =   189
               Top             =   0
               Width           =   1095
            End
            Begin VB.CommandButton cmdScene 
               Caption         =   "Move down"
               Height          =   375
               Index           =   4
               Left            =   1200
               TabIndex        =   190
               Top             =   360
               Width           =   1215
            End
            Begin VB.CommandButton cmdScene 
               Caption         =   "Move up"
               Height          =   375
               Index           =   5
               Left            =   1200
               TabIndex        =   192
               Top             =   0
               Width           =   1215
            End
            Begin VB.CommandButton cmdScene 
               Caption         =   "Add frame"
               Height          =   375
               Index           =   2
               Left            =   0
               TabIndex        =   191
               Top             =   360
               Width           =   1215
            End
            Begin VB.CommandButton cmdScene 
               Caption         =   "Add scene"
               Height          =   375
               Index           =   3
               Left            =   0
               TabIndex        =   193
               Top             =   0
               Width           =   1215
            End
         End
         Begin MSComctlLib.TreeView trFrames 
            Height          =   2895
            Left            =   120
            TabIndex        =   136
            Top             =   240
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   5106
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   441
            LabelEdit       =   1
            Style           =   7
            ImageList       =   "EditIcons"
            Appearance      =   1
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Shading02"
         Height          =   5775
         Index           =   9
         Left            =   240
         TabIndex        =   127
         Tag             =   "3DView2"
         Top             =   360
         Visible         =   0   'False
         Width           =   3615
         Begin VB.CommandButton cmdRender 
            Caption         =   "&Render"
            Height          =   375
            Left            =   720
            TabIndex        =   134
            ToolTipText     =   "Start the drawing process. This may take a few minutes"
            Top             =   5160
            Width           =   2415
         End
         Begin VB.OptionButton opShadeMode 
            Caption         =   "2D checker with depth shading"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   133
            Top             =   840
            Width           =   2775
         End
         Begin VB.OptionButton opShadeMode 
            Caption         =   "3D Texture mapping"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   132
            Top             =   1200
            Width           =   2775
         End
         Begin VB.OptionButton opShadeMode 
            Caption         =   "3D checker with depth shading"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   131
            Top             =   1560
            Width           =   2775
         End
         Begin VB.OptionButton opShadeMode 
            Caption         =   "3D checker with depth shading II"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   130
            Top             =   1920
            Width           =   2775
         End
         Begin VB.OptionButton opShadeMode 
            Caption         =   "Depth Shading"
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   129
            Top             =   2280
            Width           =   2775
         End
         Begin VB.OptionButton optWireFrame 
            Caption         =   "Wireframe"
            Height          =   255
            Left            =   360
            TabIndex        =   128
            Top             =   480
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Preview02"
         Height          =   5775
         Index           =   10
         Left            =   240
         TabIndex        =   120
         Tag             =   "PreView2"
         Top             =   360
         Visible         =   0   'False
         Width           =   3615
         Begin VB.ComboBox lstLightStyle 
            Height          =   315
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   181
            ToolTipText     =   "Selects the light pattern used on the model"
            Top             =   2500
            Width           =   1815
         End
         Begin VB.ComboBox sldShade 
            Height          =   315
            ItemData        =   "mdiMain.frx":0442
            Left            =   1500
            List            =   "mdiMain.frx":0452
            Style           =   2  'Dropdown List
            TabIndex        =   180
            ToolTipText     =   "The method used to draw the model"
            Top             =   1560
            Width           =   1815
         End
         Begin VB.CheckBox ckShowJoint 
            Alignment       =   1  'Right Justify
            Caption         =   "Show Joints"
            Height          =   255
            Left            =   360
            TabIndex        =   179
            ToolTipText     =   "Includes the skeliton as part of the model"
            Top             =   2040
            Width           =   1335
         End
         Begin MSComctlLib.Slider sldLight 
            Height          =   255
            Index           =   0
            Left            =   1440
            TabIndex        =   121
            ToolTipText     =   "The level of background light"
            Top             =   600
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            _Version        =   393216
            Max             =   100
            SelStart        =   50
            TickFrequency   =   10
            Value           =   50
         End
         Begin MSComctlLib.Slider sldLight 
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   122
            ToolTipText     =   "The level of directional light to highlight details"
            Top             =   1080
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            _Version        =   393216
            Max             =   100
            SelStart        =   50
            TickFrequency   =   10
            Value           =   50
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Ambiant light"
            Height          =   375
            Left            =   240
            TabIndex        =   126
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Directional light"
            Height          =   375
            Left            =   240
            TabIndex        =   125
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Image quality"
            Height          =   255
            Left            =   360
            TabIndex        =   124
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Light style"
            Height          =   255
            Left            =   360
            TabIndex        =   123
            Top             =   2520
            Width           =   855
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Edit13"
         Height          =   5775
         Index           =   11
         Left            =   240
         TabIndex        =   109
         Tag             =   "Edit1"
         Top             =   360
         Width           =   3615
         Begin VB.OptionButton optEdit 
            Caption         =   "Randomize"
            Height          =   255
            Index           =   18
            Left            =   360
            TabIndex        =   178
            Top             =   3240
            Width           =   2415
         End
         Begin VB.OptionButton optEdit 
            Caption         =   "Reverse faces"
            Height          =   255
            Index           =   17
            Left            =   360
            TabIndex        =   110
            Top             =   2520
            Width           =   2175
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit"
            Height          =   375
            Index           =   0
            Left            =   720
            TabIndex        =   119
            Top             =   5160
            Width           =   2295
         End
         Begin VB.Frame Frame7 
            Height          =   615
            Left            =   240
            TabIndex        =   116
            Top             =   3000
            Width           =   3135
         End
         Begin VB.OptionButton optEdit 
            Caption         =   "Select"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   115
            Top             =   480
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optEdit 
            Caption         =   "Move face"
            Height          =   255
            Index           =   14
            Left            =   360
            TabIndex        =   113
            Top             =   1320
            Width           =   2175
         End
         Begin VB.OptionButton optEdit 
            Caption         =   "Flip Vertically"
            Height          =   255
            Index           =   16
            Left            =   360
            TabIndex        =   111
            Top             =   2280
            Width           =   2175
         End
         Begin VB.OptionButton optEdit 
            Caption         =   "Move vertex"
            Height          =   255
            Index           =   13
            Left            =   360
            TabIndex        =   114
            Top             =   1080
            Width           =   2175
         End
         Begin VB.OptionButton optEdit 
            Caption         =   "Flip horizontally"
            Height          =   255
            Index           =   15
            Left            =   360
            TabIndex        =   112
            Top             =   2040
            Width           =   2175
         End
         Begin VB.Frame Frame2 
            Height          =   1095
            Left            =   240
            TabIndex        =   117
            Top             =   1800
            Width           =   3135
         End
         Begin VB.Frame Frame4 
            Height          =   855
            Left            =   240
            TabIndex        =   118
            Top             =   840
            Width           =   3135
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Galleries07"
         Height          =   7215
         Index           =   3
         Left            =   240
         TabIndex        =   169
         Tag             =   "Create1"
         Top             =   360
         Visible         =   0   'False
         Width           =   3615
         Begin Project1.DXEngine ShowGallary 
            Height          =   1815
            Left            =   120
            TabIndex        =   177
            Top             =   5160
            Width           =   3495
            _ExtentX        =   4048
            _ExtentY        =   1720
         End
         Begin Project1.Gallary Gallary 
            Height          =   615
            Left            =   240
            TabIndex        =   176
            Top             =   840
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   7646
         End
         Begin VB.ComboBox cmbGallary 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   170
            Top             =   240
            Width           =   3495
         End
         Begin VB.DirListBox Dir1 
            Height          =   315
            Left            =   240
            TabIndex        =   171
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblLoadGallary 
            Alignment       =   2  'Center
            Caption         =   "Loading Gallary. Please wait"
            Height          =   255
            Left            =   240
            TabIndex        =   185
            Top             =   3240
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Caption         =   "To create new gallaries, choose 'Gallaries' from the Tools menu"
            Height          =   495
            Left            =   360
            TabIndex        =   173
            Top             =   2520
            Width           =   3015
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Caption         =   "There are currently no Gallaries avaliable."
            Height          =   255
            Left            =   360
            TabIndex        =   172
            Top             =   2040
            Width           =   3015
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Selection06"
         Height          =   5775
         Index           =   12
         Left            =   240
         TabIndex        =   96
         Tag             =   "Edit3"
         Top             =   360
         Width           =   3615
         Begin VB.OptionButton optEdit 
            Caption         =   "Add vertex"
            Height          =   255
            Index           =   8
            Left            =   360
            TabIndex        =   98
            Top             =   2760
            Width           =   1215
         End
         Begin VB.OptionButton optEdit 
            Caption         =   "Seperate vertecies"
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   101
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit"
            Height          =   375
            Index           =   2
            Left            =   720
            TabIndex        =   108
            Top             =   5040
            Width           =   2295
         End
         Begin VB.OptionButton optEdit 
            Caption         =   "Select"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   104
            Top             =   480
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optEdit 
            Caption         =   "Combine objects"
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   102
            Top             =   1320
            Width           =   1575
         End
         Begin VB.OptionButton optEdit 
            Caption         =   "Delete vertex"
            Height          =   255
            Index           =   7
            Left            =   360
            TabIndex        =   99
            Top             =   2520
            Width           =   1335
         End
         Begin VB.OptionButton optEdit 
            Caption         =   "Compress object"
            Height          =   255
            Index           =   9
            Left            =   360
            TabIndex        =   97
            Top             =   3480
            Width           =   1575
         End
         Begin VB.Frame Frame5 
            Height          =   615
            Left            =   240
            TabIndex        =   105
            Top             =   3240
            Width           =   3135
         End
         Begin VB.OptionButton optEdit 
            Caption         =   "Select sub object"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   103
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Frame Frame1 
            Height          =   1095
            Left            =   240
            TabIndex        =   107
            Top             =   840
            Width           =   3135
         End
         Begin VB.OptionButton optEdit 
            Caption         =   "Delete face"
            Height          =   255
            Index           =   6
            Left            =   360
            TabIndex        =   100
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Frame Frame3 
            Height          =   1095
            Left            =   240
            TabIndex        =   106
            Top             =   2040
            Width           =   3135
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Face05"
         Height          =   6495
         Index           =   13
         Left            =   240
         TabIndex        =   80
         Tag             =   "Edit2"
         Top             =   360
         Width           =   3615
         Begin VB.OptionButton optEdit 
            Caption         =   "Fragment new face"
            Height          =   255
            Index           =   20
            Left            =   360
            TabIndex        =   183
            Top             =   2520
            Width           =   2055
         End
         Begin VB.OptionButton optEdit 
            Caption         =   "Fragment new vertex"
            Height          =   255
            Index           =   19
            Left            =   360
            TabIndex        =   182
            Top             =   2280
            Width           =   2055
         End
         Begin VB.OptionButton optEdit 
            Caption         =   "Fragment to Triangles"
            Height          =   255
            Index           =   10
            Left            =   360
            TabIndex        =   81
            Top             =   2040
            Width           =   2055
         End
         Begin VB.OptionButton optEdit 
            Caption         =   "Extend face"
            Height          =   255
            Index           =   12
            Left            =   360
            TabIndex        =   82
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit"
            Height          =   375
            Index           =   1
            Left            =   720
            TabIndex        =   86
            Top             =   5640
            Width           =   2295
         End
         Begin VB.OptionButton optEdit 
            Caption         =   "Select"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   84
            Top             =   480
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optEdit 
            Caption         =   "Bend face"
            Height          =   255
            Index           =   11
            Left            =   360
            TabIndex        =   83
            Top             =   1080
            Width           =   2055
         End
         Begin VB.Frame Frame6 
            Height          =   855
            Left            =   240
            TabIndex        =   85
            Top             =   840
            Width           =   3135
         End
         Begin VB.Frame Frame9 
            Height          =   1695
            Left            =   240
            TabIndex        =   184
            Top             =   1800
            Width           =   3135
            Begin VB.ComboBox cmbMethod 
               Height          =   315
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   194
               Top             =   1200
               Width           =   2655
            End
         End
         Begin VB.Frame frmExtend 
            Caption         =   "Extend Face"
            Height          =   1455
            Left            =   240
            TabIndex        =   89
            Top             =   3600
            Visible         =   0   'False
            Width           =   3135
            Begin MSComctlLib.Slider sldExtend 
               Height          =   255
               Index           =   2
               Left            =   1080
               TabIndex        =   90
               Top             =   1080
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   450
               _Version        =   393216
               LargeChange     =   1
               Min             =   -10
            End
            Begin MSComctlLib.Slider sldExtend 
               Height          =   255
               Index           =   1
               Left            =   1080
               TabIndex        =   91
               Top             =   720
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   450
               _Version        =   393216
               Max             =   20
            End
            Begin MSComctlLib.Slider sldExtend 
               Height          =   255
               Index           =   0
               Left            =   1080
               TabIndex        =   92
               Top             =   360
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   450
               _Version        =   393216
               Min             =   1
               SelStart        =   1
               Value           =   1
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "Segements"
               Height          =   255
               Left            =   120
               TabIndex        =   95
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "Taper"
               Height          =   255
               Index           =   0
               Left            =   480
               TabIndex        =   94
               Top             =   720
               Width           =   495
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "Curve"
               Height          =   255
               Index           =   1
               Left            =   480
               TabIndex        =   93
               Top             =   1080
               Width           =   495
            End
         End
         Begin VB.Frame frmFragment 
            Caption         =   "New face Scale"
            Height          =   735
            Left            =   240
            TabIndex        =   87
            Top             =   3600
            Visible         =   0   'False
            Width           =   3135
            Begin MSComctlLib.Slider Slider1 
               Height          =   255
               Left            =   120
               TabIndex        =   88
               Top             =   240
               Width           =   2895
               _ExtentX        =   5106
               _ExtentY        =   450
               _Version        =   393216
            End
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Texture17"
         Height          =   5775
         Index           =   17
         Left            =   240
         TabIndex        =   65
         Tag             =   "Texture1"
         Top             =   360
         Width           =   3615
         Begin VB.PictureBox f1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   0
            Left            =   1080
            ScaleHeight     =   465
            ScaleWidth      =   705
            TabIndex        =   77
            Top             =   600
            Width           =   735
         End
         Begin VB.PictureBox f1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   1
            Left            =   1320
            ScaleHeight     =   465
            ScaleWidth      =   705
            TabIndex        =   78
            Top             =   720
            Width           =   735
         End
         Begin VB.OptionButton Mode 
            Caption         =   "Spray Can"
            Height          =   255
            Index           =   7
            Left            =   480
            TabIndex        =   76
            ToolTipText     =   "Paints using a spray can effect"
            Top             =   2280
            Width           =   1095
         End
         Begin VB.OptionButton Mode 
            Caption         =   "Line"
            Height          =   255
            Index           =   6
            Left            =   2160
            TabIndex        =   75
            ToolTipText     =   "Draws a straight line by dragging the mouse"
            Top             =   1560
            Width           =   735
         End
         Begin VB.OptionButton Mode 
            Caption         =   "Paint Picture"
            Height          =   255
            Index           =   5
            Left            =   480
            TabIndex        =   74
            Top             =   2640
            Width           =   1215
         End
         Begin VB.OptionButton Mode 
            Caption         =   "Paint"
            Height          =   255
            Index           =   4
            Left            =   480
            TabIndex        =   73
            ToolTipText     =   "Fills in an area of the picture using the forecolor as a border"
            Top             =   1920
            Width           =   735
         End
         Begin VB.OptionButton Mode 
            Caption         =   "Pen"
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   72
            ToolTipText     =   "Draws a line free hand with the mouse"
            Top             =   1560
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton Mode 
            Caption         =   "Eclipse"
            Height          =   255
            Index           =   2
            Left            =   2160
            TabIndex        =   71
            ToolTipText     =   "Draws an ellipse by dragging the mouse"
            Top             =   2280
            Width           =   855
         End
         Begin VB.OptionButton Mode 
            Caption         =   "Circle"
            Height          =   255
            Index           =   1
            Left            =   2160
            TabIndex        =   70
            ToolTipText     =   "Draws a perfect circle by dragging the mouse"
            Top             =   2640
            Width           =   735
         End
         Begin VB.OptionButton Mode 
            Caption         =   "Box"
            Height          =   255
            Index           =   0
            Left            =   2160
            TabIndex        =   69
            ToolTipText     =   "Draws a box by dragging the mouse"
            Top             =   1920
            Width           =   615
         End
         Begin VB.ComboBox lstPattern 
            Height          =   315
            ItemData        =   "mdiMain.frx":0489
            Left            =   360
            List            =   "mdiMain.frx":04A5
            Style           =   2  'Dropdown List
            TabIndex        =   68
            ToolTipText     =   "Sets the fill pattern for shapes and the paint tool"
            Top             =   3360
            Width           =   3015
         End
         Begin VB.CheckBox ckShowEdges 
            Caption         =   "Show Faces"
            Height          =   195
            Left            =   480
            TabIndex        =   66
            ToolTipText     =   "Sets whether object faces are displayed"
            Top             =   4920
            Width           =   1215
         End
         Begin MSComctlLib.Slider sldWidth 
            Height          =   255
            Left            =   1200
            TabIndex        =   67
            ToolTipText     =   "Sets the width of lines as they are drawn"
            Top             =   4200
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Min             =   1
            SelStart        =   1
            Value           =   1
         End
         Begin VB.Label Label13 
            Caption         =   "Line Width"
            Height          =   255
            Left            =   240
            TabIndex        =   79
            Top             =   4200
            Width           =   855
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Edit Skeliton12"
         Height          =   5775
         Index           =   4
         Left            =   240
         TabIndex        =   60
         Tag             =   "Skeliton1"
         Top             =   360
         Visible         =   0   'False
         Width           =   3615
         Begin VB.OptionButton opChangeJ 
            Caption         =   "Change Target of joint"
            Height          =   255
            Left            =   600
            TabIndex        =   63
            ToolTipText     =   "Drag from one joint to another to set the target"
            Top             =   5160
            Width           =   1935
         End
         Begin VB.OptionButton opAddJ 
            Caption         =   "Add joint to model"
            Height          =   255
            Left            =   600
            TabIndex        =   62
            ToolTipText     =   "Creates a new joint whee you click"
            Top             =   4800
            Width           =   1575
         End
         Begin VB.OptionButton opSelectJ 
            Caption         =   "Select tool"
            Height          =   255
            Left            =   600
            TabIndex        =   61
            ToolTipText     =   "As thought you have the 'Select Tool' on the toolbar selected"
            Top             =   4440
            Value           =   -1  'True
            Width           =   1095
         End
         Begin MSComctlLib.TreeView Joints 
            Height          =   4095
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   7223
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   529
            Style           =   7
            FullRowSelect   =   -1  'True
            ImageList       =   "EditIcons"
            Appearance      =   1
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Paint16"
         Height          =   5775
         Index           =   16
         Left            =   240
         TabIndex        =   55
         Tag             =   "3DView3"
         Top             =   360
         Width           =   3615
         Begin VB.CheckBox ckEnablePaint 
            Caption         =   "Enable Paint Brush"
            Height          =   255
            Left            =   240
            TabIndex        =   59
            Top             =   480
            Width           =   1695
         End
         Begin VB.PictureBox pcPaint 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   0
            Left            =   480
            ScaleHeight     =   705
            ScaleWidth      =   1065
            TabIndex        =   58
            ToolTipText     =   "Back-Colour"
            Top             =   1920
            Width           =   1095
         End
         Begin VB.PictureBox pcPaint 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   1
            Left            =   720
            ScaleHeight     =   705
            ScaleWidth      =   1065
            TabIndex        =   57
            ToolTipText     =   "Fore-Colour"
            Top             =   2160
            Width           =   1095
         End
         Begin VB.CheckBox ckRotateFace 
            Caption         =   "Rotate Face"
            Height          =   255
            Left            =   480
            TabIndex        =   56
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000014&
            X1              =   120
            X2              =   3600
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000010&
            X1              =   120
            X2              =   3600
            Y1              =   1183
            Y2              =   1183
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Select09"
         Height          =   5775
         Index           =   2
         Left            =   240
         TabIndex        =   45
         Tag             =   "Select1"
         Top             =   360
         Width           =   3615
         Begin VB.CheckBox chkSelect 
            Caption         =   "Select &vertecies"
            Height          =   255
            Index           =   6
            Left            =   480
            TabIndex        =   54
            ToolTipText     =   "Allows individual vertecies to be selected or deselected"
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CommandButton CmdAline 
            Height          =   495
            Index           =   3
            Left            =   2040
            Picture         =   "mdiMain.frx":0523
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Moves the selection right"
            Top             =   4320
            Width           =   495
         End
         Begin VB.CommandButton CmdAline 
            Height          =   495
            Index           =   2
            Left            =   1560
            Picture         =   "mdiMain.frx":082D
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Moves the selection down"
            Top             =   4800
            Width           =   495
         End
         Begin VB.CommandButton CmdAline 
            Height          =   495
            Index           =   1
            Left            =   1080
            Picture         =   "mdiMain.frx":0B37
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Moves the selection left"
            Top             =   4320
            Width           =   495
         End
         Begin VB.CommandButton CmdAline 
            Height          =   495
            Index           =   0
            Left            =   1560
            Picture         =   "mdiMain.frx":0E41
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Moves the selection up"
            Top             =   3840
            Width           =   495
         End
         Begin VB.CheckBox chkSelect 
            Caption         =   "Select joint &groups"
            Height          =   255
            Index           =   4
            Left            =   480
            TabIndex        =   49
            ToolTipText     =   "Selects the objects attached to the selected joint"
            Top             =   2520
            Width           =   2175
         End
         Begin VB.CheckBox chkSelect 
            Caption         =   "&Boxband Select must be complete"
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   48
            ToolTipText     =   "When on, a box band must fully suround an object"
            Top             =   2040
            Width           =   2895
         End
         Begin VB.CheckBox chkSelect 
            Caption         =   "Select &Joints"
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   47
            ToolTipText     =   "Allows joints to be selected"
            Top             =   1080
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CheckBox chkSelect 
            Caption         =   "Select &Objects"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   46
            ToolTipText     =   "Allows objects to be selected"
            Top             =   600
            Value           =   1  'Checked
            Width           =   1455
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Scale Brush10"
         Height          =   5775
         Index           =   6
         Left            =   240
         TabIndex        =   33
         Tag             =   "Scale1"
         Top             =   360
         Visible         =   0   'False
         Width           =   3615
         Begin VB.CommandButton cmdScale 
            Caption         =   "Scale brush"
            Height          =   375
            Left            =   720
            TabIndex        =   41
            ToolTipText     =   "Click to set the changes to your model"
            Top             =   4560
            Width           =   2295
         End
         Begin VB.TextBox PresetScale 
            Height          =   375
            Left            =   600
            TabIndex        =   37
            Text            =   "1"
            ToolTipText     =   "The size that the object will become compared to before the opperation"
            Top             =   4080
            Width           =   2415
         End
         Begin VB.OptionButton SklMode 
            Caption         =   "Custom scale"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   36
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton SklMode 
            Caption         =   "Preset scale"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   35
            Top             =   2400
            Width           =   1575
         End
         Begin VB.ListBox Scales 
            Height          =   1230
            ItemData        =   "mdiMain.frx":114B
            Left            =   600
            List            =   "mdiMain.frx":1173
            TabIndex        =   34
            Top             =   2760
            Width           =   2415
         End
         Begin MSComctlLib.Slider sclXDim 
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   38
            Top             =   720
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   450
            _Version        =   393216
            Min             =   50
            Max             =   150
            SelStart        =   100
            TickFrequency   =   10
            Value           =   100
         End
         Begin MSComctlLib.Slider sclXDim 
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   39
            Top             =   1200
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   450
            _Version        =   393216
            Min             =   50
            Max             =   150
            SelStart        =   100
            TickFrequency   =   10
            Value           =   100
         End
         Begin MSComctlLib.Slider sclXDim 
            Height          =   255
            Index           =   2
            Left            =   720
            TabIndex        =   40
            Top             =   1680
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   450
            _Version        =   393216
            Min             =   50
            Max             =   150
            SelStart        =   100
            TickFrequency   =   10
            Value           =   100
         End
         Begin VB.Label Label1 
            Caption         =   "Z"
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   44
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Y"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   43
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "X"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   42
            Top             =   720
            Width           =   255
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "AI Links01"
         Height          =   5775
         Index           =   7
         Left            =   240
         TabIndex        =   29
         Tag             =   "Skeliton2"
         Top             =   360
         Width           =   3615
         Begin VB.OptionButton opAddG 
            Caption         =   "Create AI link"
            Height          =   255
            Left            =   600
            TabIndex        =   31
            ToolTipText     =   "Drag the mouse between two joints to create an AI link"
            Top             =   4560
            Width           =   1335
         End
         Begin VB.OptionButton opSelectG 
            Caption         =   "Select"
            Height          =   255
            Left            =   600
            TabIndex        =   30
            ToolTipText     =   "Select objects and joints by clicking on them"
            Top             =   4200
            Value           =   -1  'True
            Width           =   855
         End
         Begin MSFlexGridLib.MSFlexGrid AIGrid 
            Height          =   3375
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   5953
            _Version        =   393216
            Rows            =   0
            Cols            =   5
            FixedRows       =   0
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Rotate Brush11"
         Height          =   5775
         Index           =   5
         Left            =   240
         TabIndex        =   14
         Tag             =   "Rotate1"
         Top             =   360
         Visible         =   0   'False
         Width           =   3615
         Begin VB.CommandButton cmdRotate 
            Caption         =   "Rotate Brush"
            Height          =   375
            Left            =   720
            TabIndex        =   27
            ToolTipText     =   "Click to set the changes to your model"
            Top             =   4920
            Width           =   2295
         End
         Begin VB.TextBox txtCustomAngle 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   2040
            TabIndex        =   25
            Text            =   "0"
            ToolTipText     =   "Enter a value to rotate the selected objects through"
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton QuickSpin 
            Height          =   495
            Index           =   0
            Left            =   840
            Picture         =   "mdiMain.frx":11AA
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Quick 5* rotate left"
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton QuickSpin 
            Height          =   495
            Index           =   1
            Left            =   1920
            Picture         =   "mdiMain.frx":14B4
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Quick 5* rotate right"
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton QuickSpin 
            Height          =   495
            Index           =   2
            Left            =   840
            Picture         =   "mdiMain.frx":17BE
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Quick 90* rotate left"
            Top             =   1800
            Width           =   855
         End
         Begin VB.CommandButton QuickSpin 
            Height          =   495
            Index           =   3
            Left            =   1920
            Picture         =   "mdiMain.frx":1AC8
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Quick 90* rotate left"
            Top             =   1800
            Width           =   855
         End
         Begin VB.OptionButton optGetCenter 
            Caption         =   "Around object centers"
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   20
            ToolTipText     =   "Rotate each object around its own center"
            Top             =   3240
            Width           =   1935
         End
         Begin VB.OptionButton optGetCenter 
            Caption         =   "Around joint"
            Height          =   255
            Index           =   4
            Left            =   720
            TabIndex        =   19
            ToolTipText     =   "Rotate the objects around the joint they are attached to"
            Top             =   3960
            Width           =   1215
         End
         Begin VB.OptionButton optGetCenter 
            Caption         =   "Around world center"
            Height          =   255
            Index           =   2
            Left            =   720
            TabIndex        =   18
            ToolTipText     =   "Rotate around the 0,0,0 location at the center of the world"
            Top             =   3480
            Width           =   1815
         End
         Begin VB.OptionButton optGetCenter 
            Caption         =   "Around selection center"
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   17
            ToolTipText     =   "Rotate around the collective center of the selected objects"
            Top             =   3000
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton optGetCenter 
            Caption         =   "Custom"
            Height          =   255
            Index           =   3
            Left            =   720
            TabIndex        =   16
            ToolTipText     =   "Rotate around whereever you click the mouse"
            Top             =   3720
            Width           =   975
         End
         Begin VB.OptionButton optGetCenter 
            Caption         =   "Around vertex center"
            Enabled         =   0   'False
            Height          =   255
            Index           =   5
            Left            =   720
            TabIndex        =   15
            ToolTipText     =   "Rotate around the collective center of the selected vertecies"
            Top             =   4200
            Width           =   1815
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   375
            Left            =   2640
            TabIndex        =   26
            Top             =   360
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   661
            _Version        =   393216
            OrigLeft        =   2400
            OrigTop         =   360
            OrigRight       =   2640
            OrigBottom      =   735
            Max             =   359
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
         Begin VB.Label Label3 
            Caption         =   "Rotate throught"
            Height          =   255
            Left            =   720
            TabIndex        =   28
            Top             =   405
            Width           =   1215
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Create Brush02"
         Height          =   5775
         Index           =   1
         Left            =   240
         TabIndex        =   152
         Tag             =   "Create2"
         Top             =   360
         Visible         =   0   'False
         Width           =   3615
         Begin VB.OptionButton EditLine 
            Caption         =   "Extend line"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   155
            ToolTipText     =   "Alters the profile of the new object"
            Top             =   4680
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.CommandButton cmdCreate 
            Caption         =   "Cancel Brush"
            Height          =   375
            Index           =   1
            Left            =   2040
            TabIndex        =   158
            ToolTipText     =   "Click to cancel the selected options"
            Top             =   5160
            Width           =   1215
         End
         Begin VB.CommandButton cmdCreate 
            Caption         =   "Create Brush"
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   157
            ToolTipText     =   "Click to create the brush"
            Top             =   5160
            Width           =   1215
         End
         Begin VB.OptionButton EditLine 
            Caption         =   "Move Profile"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   154
            ToolTipText     =   "Moves the whole line that defines some objects"
            Top             =   4440
            Width           =   1335
         End
         Begin MSComctlLib.Slider ShpProp 
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   153
            ToolTipText     =   "Sets the angle of the whole object"
            Top             =   3120
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Max             =   71
            TickFrequency   =   5
         End
         Begin MSComctlLib.ListView ShapeList 
            Height          =   1695
            Left            =   120
            TabIndex        =   159
            Top             =   240
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   2990
            View            =   2
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComctlLib.Slider ShpProp 
            Height          =   255
            Index           =   4
            Left            =   1920
            TabIndex        =   160
            ToolTipText     =   "Sets the size of the bottom of the object"
            Top             =   3720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            Min             =   1
            Max             =   20
            SelStart        =   20
            TickFrequency   =   2
            Value           =   20
         End
         Begin MSComctlLib.Slider ShpProp 
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   161
            ToolTipText     =   "Sets the number of edges"
            Top             =   2520
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Min             =   3
            Max             =   25
            SelStart        =   3
            Value           =   3
         End
         Begin MSComctlLib.Slider ShpProp 
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   162
            ToolTipText     =   "Sets the size of the top of the object"
            Top             =   3720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            Min             =   1
            Max             =   20
            SelStart        =   20
            TickFrequency   =   2
            Value           =   20
         End
         Begin MSComctlLib.Slider ShpProp 
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   163
            ToolTipText     =   "Sets the number of faces around the axis"
            Top             =   3720
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Min             =   3
            Max             =   25
            SelStart        =   3
            Value           =   3
         End
         Begin VB.OptionButton EditLine 
            Caption         =   "Move Axis"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   156
            ToolTipText     =   "Sets the position of the axis"
            Top             =   4200
            Width           =   1095
         End
         Begin VB.Label ShpName 
            Caption         =   "Rotation Angle"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   168
            Top             =   2880
            Width           =   1455
         End
         Begin VB.Label ShpName 
            Caption         =   "Top face"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   167
            Top             =   3480
            Width           =   855
         End
         Begin VB.Label ShpName 
            Caption         =   "Bottom Face"
            Height          =   255
            Index           =   4
            Left            =   1680
            TabIndex        =   166
            Top             =   3480
            Width           =   1215
         End
         Begin VB.Label ShpName 
            Caption         =   "Horizontal faces"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   165
            Top             =   3480
            Width           =   1455
         End
         Begin VB.Label ShpName 
            Caption         =   "Edges"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   164
            Top             =   2280
            Width           =   975
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Wireframe04"
         Height          =   5775
         Index           =   8
         Left            =   240
         TabIndex        =   140
         Tag             =   "3DView1"
         Top             =   360
         Visible         =   0   'False
         Width           =   3615
         Begin VB.CheckBox ck3Dopt 
            Caption         =   "Name joints"
            Height          =   255
            Index           =   8
            Left            =   960
            TabIndex        =   141
            ToolTipText     =   "Dislpay the names of the joints"
            Top             =   2400
            Width           =   1215
         End
         Begin VB.CheckBox ck3Dopt 
            Caption         =   "Highlight verteies"
            Height          =   255
            Index           =   10
            Left            =   960
            TabIndex        =   142
            ToolTipText     =   "Draws a circle over each vertex"
            Top             =   1680
            Width           =   1575
         End
         Begin VB.CheckBox ck3Dopt 
            Caption         =   "Remove hidden faces"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   143
            ToolTipText     =   "Remove faces that point away from the camara"
            Top             =   1440
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox ck3Dopt 
            Caption         =   "Highlight faces"
            Height          =   255
            Index           =   9
            Left            =   960
            TabIndex        =   144
            ToolTipText     =   "Displays a circle at the center of each face"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CheckBox ck3Dopt 
            Caption         =   "Origin"
            Height          =   255
            Index           =   11
            Left            =   600
            TabIndex        =   151
            ToolTipText     =   "Display the axis names"
            Top             =   3840
            Width           =   855
         End
         Begin VB.CheckBox ck3Dopt 
            Caption         =   "Perspective"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   150
            ToolTipText     =   "Objects get smaller as they get further away"
            Top             =   480
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox ck3Dopt 
            Caption         =   "Ground plane"
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   149
            ToolTipText     =   "Draw a square representing the zero Y plain"
            Top             =   4320
            Width           =   1335
         End
         Begin VB.CheckBox ck3Dopt 
            Caption         =   "Y - Clipping"
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   148
            ToolTipText     =   "Remove faces that go below the zero Y plain"
            Top             =   3360
            Width           =   1095
         End
         Begin VB.CheckBox ck3Dopt 
            Caption         =   "Draw skeliton"
            Height          =   255
            Index           =   6
            Left            =   600
            TabIndex        =   147
            ToolTipText     =   "Draw the skeliton in the model"
            Top             =   2160
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox ck3Dopt 
            Caption         =   "Draw AI links"
            Height          =   255
            Index           =   7
            Left            =   600
            TabIndex        =   146
            ToolTipText     =   "Display the AI links"
            Top             =   2880
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox ck3Dopt 
            Caption         =   "Draw objects"
            Height          =   255
            Index           =   5
            Left            =   600
            TabIndex        =   145
            ToolTipText     =   "Draw the objects in the model"
            Top             =   960
            Value           =   1  'Checked
            Width           =   1335
         End
      End
      Begin MSComctlLib.TabStrip cmdSidebar 
         Height          =   4815
         Left            =   60
         TabIndex        =   174
         Top             =   120
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   8493
         MultiRow        =   -1  'True
         HotTracking     =   -1  'True
         TabStyle        =   1
         ImageList       =   "GrayIcons"
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList GrayIcons 
      Left            =   960
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483644
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1DD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":20EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":253E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2990
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2DE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3236
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3688
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3ADA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3DF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4246
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4698
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4AEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4F3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":538E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":57E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":5C36
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":608A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":64DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":6932
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList PlayIcons 
      Left            =   2760
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483644
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483644
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":6D86
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":71DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":762E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":7A80
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":7ED2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":8326
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":877A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":8BCE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList EditIcons 
      Left            =   2160
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":9022
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":9476
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":98CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":9D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":A170
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":A5C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":AA16
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":AD30
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":B184
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":B5D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":BA2C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList MainIcons 
      Left            =   1560
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":BE80
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":C2D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":C724
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":CA3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":CE90
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":D2E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":D5FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":D916
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":DD68
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":E1BA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cBar 
      Align           =   1  'Align Top
      Height          =   1230
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   14250
      _ExtentX        =   25135
      _ExtentY        =   2170
      BandCount       =   4
      BandBorders     =   0   'False
      VariantHeight   =   0   'False
      _CBWidth        =   14250
      _CBHeight       =   1230
      _Version        =   "6.0.8169"
      Child1          =   "tbar (0)"
      MinHeight1      =   390
      Width1          =   5745
      NewRow1         =   0   'False
      Visible1        =   0   'False
      Child2          =   "tbar (1)"
      MinHeight2      =   390
      Width2          =   2175
      NewRow2         =   0   'False
      Visible2        =   0   'False
      Child3          =   "tbar (2)"
      MinHeight3      =   390
      Width3          =   2175
      NewRow3         =   -1  'True
      Visible3        =   0   'False
      Child4          =   "tbar (3)"
      MinHeight4      =   390
      Width4          =   615
      NewRow4         =   -1  'True
      Visible4        =   0   'False
      Begin MSComctlLib.Toolbar tbar 
         Height          =   390
         Index           =   0
         Left            =   165
         TabIndex        =   11
         Top             =   30
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         ImageList       =   "MainIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Create a new model"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Open an existing model"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Save the current model"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Display the sidebar"
               ImageIndex      =   4
               Style           =   1
               Value           =   1
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   5
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Open the object editor window"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Open the entity editor window"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Open the joint editor window"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Open the Surface Editor window"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Zoom in"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
               Object.Width           =   1600
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Zoom out"
               ImageIndex      =   9
            EndProperty
         EndProperty
         Begin VB.ComboBox cmdZoomLevels 
            Height          =   315
            ItemData        =   "mdiMain.frx":E60E
            Left            =   3480
            List            =   "mdiMain.frx":E627
            Style           =   2  'Dropdown List
            TabIndex        =   12
            ToolTipText     =   "Select a level of magnification"
            Top             =   40
            Width           =   1535
         End
      End
      Begin MSComctlLib.Toolbar tbar 
         Height          =   390
         Index           =   1
         Left            =   5910
         TabIndex        =   10
         Top             =   30
         Width           =   8250
         _ExtentX        =   14552
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         ImageList       =   "EditIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Select"
               Object.ToolTipText     =   "Select"
               Object.Tag             =   "1"
               ImageIndex      =   4
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Create"
               Object.ToolTipText     =   "Insert object"
               Object.Tag             =   "1"
               ImageIndex      =   2
               Style           =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Edit"
               Object.ToolTipText     =   "Edit object"
               Object.Tag             =   "1"
               ImageIndex      =   3
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Scale"
               Object.ToolTipText     =   "Scale objects"
               Object.Tag             =   "1"
               ImageIndex      =   6
               Style           =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Rotate"
               Object.ToolTipText     =   "Rotate objects"
               Object.Tag             =   "1"
               ImageIndex      =   5
               Style           =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Skeliton"
               Object.ToolTipText     =   "Edit Skeliton"
               Object.Tag             =   "1"
               ImageIndex      =   1
               Style           =   2
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "3DView"
               Object.ToolTipText     =   "3D View"
               Object.Tag             =   "1"
               ImageIndex      =   10
               Style           =   2
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Animation"
               Object.ToolTipText     =   "Animation"
               Object.Tag             =   "1"
               ImageIndex      =   7
               Style           =   2
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "PreView"
               Object.ToolTipText     =   "Model Preview"
               Object.Tag             =   "1"
               ImageIndex      =   2
               Style           =   2
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Texture"
               Object.ToolTipText     =   "Texture Map"
               Object.Tag             =   "1"
               ImageIndex      =   9
               Style           =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbar 
         Height          =   390
         Index           =   2
         Left            =   165
         TabIndex        =   7
         Top             =   420
         Width           =   13995
         _ExtentX        =   24686
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         ImageList       =   "PlayIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Jump to start"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Rewind"
               ImageIndex      =   2
               Style           =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Stop"
               ImageIndex      =   3
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Forward"
               ImageIndex      =   4
               Style           =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Jump to end"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Play to End"
               ImageIndex      =   8
               Style           =   2
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Play on Loop"
               ImageIndex      =   6
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Reverse Play"
               ImageIndex      =   7
               Style           =   2
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
               Object.Width           =   1800
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         Begin VB.ComboBox cmbScenes 
            Height          =   315
            Left            =   3000
            Style           =   2  'Dropdown List
            TabIndex        =   9
            ToolTipText     =   "Select a scene to play"
            Top             =   45
            Width           =   1815
         End
         Begin MSComctlLib.Slider sldFrames 
            Height          =   255
            Left            =   5040
            TabIndex        =   8
            ToolTipText     =   "The currently displayed frame "
            Top             =   30
            Visible         =   0   'False
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Min             =   1
            Max             =   2
            SelStart        =   1
            Value           =   1
         End
      End
      Begin MSComctlLib.Toolbar tbar 
         Height          =   390
         Index           =   3
         Left            =   165
         TabIndex        =   1
         Top             =   810
         Width           =   13995
         _ExtentX        =   24686
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "GrayIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Select All"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Deselect All"
               ImageIndex      =   18
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
               Object.Width           =   6600
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Filter Selection"
               ImageIndex      =   19
            EndProperty
         EndProperty
         Begin VB.ComboBox cmdAttribute 
            Height          =   315
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   50
            Width           =   1455
         End
         Begin VB.Frame Frame8 
            BorderStyle     =   0  'None
            Caption         =   "Frame8"
            Height          =   375
            Left            =   840
            TabIndex        =   5
            Top             =   0
            Width           =   1695
            Begin VB.Label Label12 
               Caption         =   "Select Objects where"
               Height          =   255
               Left            =   0
               TabIndex        =   6
               Top             =   90
               Width           =   1695
            End
         End
         Begin VB.ComboBox cmbLogic 
            Height          =   315
            Left            =   3960
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   45
            Width           =   1455
         End
         Begin VB.ComboBox cmbValue 
            Height          =   315
            Left            =   5520
            TabIndex        =   2
            Top             =   50
            Width           =   1815
         End
      End
   End
   Begin MSComctlLib.StatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   175
      Top             =   9345
      Visible         =   0   'False
      Width           =   14250
      _ExtentX        =   25135
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19500
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&New"
         Index           =   1
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Open"
         Index           =   2
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Inport"
         Index           =   3
      End
      Begin VB.Menu OldFile 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Index           =   1
      End
   End
   Begin VB.Menu MenuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit 
         Caption         =   "&Settings"
         Index           =   1
         Shortcut        =   +{F1}
      End
   End
   Begin VB.Menu menuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuTools 
         Caption         =   "&Compile Folder..."
         Index           =   1
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Animation &Viewer..."
         Index           =   2
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&About"
         Index           =   1
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
         Index           =   2
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Tip of the Day"
         Index           =   3
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "On the Web"
         Index           =   4
         Begin VB.Menu mnuOnTheWeb 
            Caption         =   "Animation Shop 8.3"
            Index           =   1
         End
         Begin VB.Menu mnuOnTheWeb 
            Caption         =   "Get Registration Code"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuOnTheWeb 
            Caption         =   "Download Samples"
            Index           =   3
         End
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Edit Help"
         Index           =   6
      End
   End
   Begin VB.Menu menuPopupGallary 
      Caption         =   "menuPopupGallary"
      Visible         =   0   'False
      Begin VB.Menu mnuGallary 
         Caption         =   "View from Top"
         Index           =   1
      End
      Begin VB.Menu mnuGallary 
         Caption         =   "View from Front"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mnuGallary 
         Caption         =   "View from Side"
         Index           =   3
      End
      Begin VB.Menu mnuGallary 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuGallary 
         Caption         =   "Rename Item"
         Index           =   5
      End
      Begin VB.Menu mnuGallary 
         Caption         =   "Remove Item"
         Index           =   6
      End
   End
   Begin VB.Menu menuFrame 
      Caption         =   "Frames"
      Visible         =   0   'False
      Begin VB.Menu mnuFrame 
         Caption         =   "Insert at Top"
         Index           =   1
      End
      Begin VB.Menu mnuFrame 
         Caption         =   "Insert Above"
         Index           =   2
      End
      Begin VB.Menu mnuFrame 
         Caption         =   "Insert Below"
         Index           =   3
      End
      Begin VB.Menu mnuFrame 
         Caption         =   "Insert at Bottom"
         Index           =   4
      End
      Begin VB.Menu mnuFrame 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuFrame 
         Caption         =   "Split Scene Below"
         Index           =   6
      End
      Begin VB.Menu mnuFrame 
         Caption         =   "Combine with previous scene"
         Index           =   7
      End
      Begin VB.Menu mnuFrame 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuFrame 
         Caption         =   "Duplicate Frame"
         Index           =   9
      End
      Begin VB.Menu mnuFrame 
         Caption         =   "Ignore Frame"
         Index           =   10
      End
      Begin VB.Menu mnuFrame 
         Caption         =   "Delete Frame"
         Index           =   11
      End
      Begin VB.Menu mnuFrame 
         Caption         =   "Delete Scene"
         Index           =   12
      End
      Begin VB.Menu mnuFrame 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuFrame 
         Caption         =   "Frame Iterations"
         Index           =   14
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkSelect_Click(Index As Integer)
    Select Case Index
        Case 6
            If chkSelect(Index) = 1 Then ActiveForm.Tablet.pShowVertecies = True Else ActiveForm.Tablet.pShowVertecies = False
            ActiveForm.Tablet.Refresh
    End Select
End Sub

Private Sub Command1_Click()
  Am8(ActiveFile).Scene.MoveAnimation
  ActiveForm.Engine.RefreshView
End Sub


Private Sub ckShowEdges_Click()
    If ckShowEdges = 1 Then ActiveForm.TexMap.DrawShapes = True Else ActiveForm.TexMap.DrawShapes = False
End Sub

Private Sub cmdBF_Click(Index As Integer)
    Dim FrameON As String, SceneON As String
 '   If sldFrames.Visible = False Then Exit Sub
    If Index = 0 Then
        If sldFrames = 1 Then Exit Sub
        trFrames.SelectedItem.Previous.Selected = True
        sldFrames = sldFrames - 1
    Else
        If sldFrames = sldFrames.Max Then Exit Sub
        trFrames.SelectedItem.Next.Selected = True
        sldFrames = sldFrames + 1
    End If
    trFrames_NodeClick trFrames.SelectedItem
    'FrameON = Am8(ActiveFile).Scene.GetFrame(trFrames.SelectedItem.Key)
    'SceneON = Am8(ActiveFile).Scene.GetScene(trFrames.SelectedItem.Key)
    'grGrid.UpdateDisplay SceneON, FrameON
End Sub

'#######################################################################
'#                                                                     #
'#  This is the MDI form that makes up the main window of the program  #
'#   It contains a large amount of code as it deals with all the       #
'#  controls on the sidebar, and for creating new files and windows    #
'#                                                                     #
'#######################################################################


Private Sub MDIForm_Load()
    'This is the main code section for the form, and is run at the start of the program
    Dim Key As String, n As Integer
    'ShowGallary.pShowSkeliton = True
    ShowSidebar "Select": Visible = True
    With frmMain.ShapeList.ListItems
        .Add , , "Cube":        .Add , , "Face"
        .Add , , "Prism":       .Add , , "Cone"
        .Add , , "Dimond":      .Add , , "Torous"
        .Add , , "Wrap":        .Add , , "Grid"
        .Add , , "Sphere":      '.Add , , "Rubix Cube"
        .Add , , "Star"
    End With
    SetNewShapeMenu "Cube"
    AIGrid.ColWidth(0) = 600: AIGrid.ColWidth(1) = 0: AIGrid.ColWidth(4) = 600
    AIGrid.AddItem vbTab & vbTab & "Start" & vbTab & "End" & vbTab & "Dir"
    AIGrid.Rows = 6: AIGrid.FixedRows = 1
    frmMain.ShapeList.ListItems(1).Selected = True
    cmdAttribute.AddItem "Vertex Count": cmdAttribute.AddItem "Face Count"
    cmdAttribute.AddItem "Edge Count": cmdAttribute.AddItem "Colour"
    cmdAttribute.AddItem "Entity Name": cmdAttribute.ListIndex = 0
    cmbLogic.AddItem "is equal to": cmbLogic.AddItem "is not"
    cmbLogic.AddItem "is greater than": cmbLogic.AddItem "is less than"
    cmbLogic.ListIndex = 0: sldShade.ListIndex = 3: lstPattern.ListIndex = 0
    cmbMethod.AddItem "General Object": cmbMethod.AddItem "Cylinder"
    cmbMethod.AddItem "Cylinder Ends": cmbMethod.ListIndex = 0
    ShowGallary.pShowSkeliton = True
    GetGallaryFolders
End Sub

Public Sub CauseFormResize()
    'This allows other forms to makes this form resize itself
    MDIForm_Resize
End Sub

Private Sub MDIForm_Resize()
    'Resize all the objects on the screen. With a MDI form, you MUST use form.height
    'and not form.scaleheight to align objects, becuase otherwise it won't work right
    Dim n As Integer, ReduceHeight As Integer
    On Error Resume Next
    If sBar.Visible = True Then
        If cBar.Visible = True Then cmdSidebar.Height = Height - 1217 - cBar.Height Else cmdSidebar.Height = Height - 1117
    Else
        If cBar.Visible = True Then cmdSidebar.Height = Height - 900 - cBar.Height Else cmdSidebar.Height = Height - 800
    End If
    For n = 1 To Frame.Count
        Frame(n).Top = cmdSidebar.ClientTop: Frame(n).Left = cmdSidebar.ClientLeft
        Frame(n).Height = cmdSidebar.ClientHeight: Frame(n).Width = cmdSidebar.ClientWidth
    Next n
    Gallary.Height = Frame(1).Height - Gallary.Top - 300 - ShowGallary.Height
    ShowGallary.Top = Gallary.Height + Gallary.Top + 200
    For n = 0 To 2: cmdEdit(n).Top = Frame(1).Height - 700: Next n
    cmdCreate(0).Top = Frame(1).Height - 700:   cmdCreate(1).Top = Frame(1).Height - 700
    cmdScale.Top = Frame(1).Height - 700:       cmdRotate.Top = Frame(1).Height - 700
    cmdRender.Top = Frame(1).Height - 700:      trFrames.Width = Frame(1).Width - 200
    grGrid.Width = Frame(1).Width - 200:        grGrid.Height = Frame(1).Height - 1100
    sBar.Width = Width - 110:                   AIGrid.Height = Frame(1).Height - 1800
    cmdBF(0).Top = Frame(1).Height - 550:       cmdBF(1).Top = Frame(1).Height - 550
    Joints.Height = Frame(1).Height - 1800:     opSelectG.Top = Frame(1).Height - 1400
    opAddG.Top = Frame(1).Height - 1100:        opSelectJ.Top = Frame(1).Height - 1400
    opAddJ.Top = Frame(1).Height - 1100:        opChangeJ.Top = Frame(1).Height - 800
    trFrames.Height = Frame(1).Height - 1500
    FrameButtons.Top = trFrames.Height + 300
End Sub

Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
    'This command allows you to drag a file from Windows onto the MDI form and automaticly load it
    LoadExistingFileWithWindow Data.Files(1)
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    'When the main form is unloaded, this removes the AM8 class, and clears up any tempory files
    Set Am8 = Nothing
    Destroy App.Path & "\data\clipboard"
    Destroy App.Path & "\data\duplicate"
    End
End Sub

Public Sub UpdateHistoryMenu()
    'This command updates the file history menu on the MDI form, not the edit window forms
    Dim n As Integer, Length As Integer
    Length = Am8.FileHistory.Lenght
    For n = frmMain.OldFile.Count - 1 To 1 Step -1: Unload frmMain.OldFile(n): Next n
    If Length > Am8.FileHistory.CountHistory Then Length = Am8.FileHistory.CountHistory
    For n = 1 To Length
        Load OldFile(n): OldFile(n).Visible = True
        If Am8.FullPath = True Then OldFile(n).Caption = "&" & n & ". " & MaxLength(Am8.FileHistory(n).FilePath, 20, 3) Else OldFile(n).Caption = "&" & n & ". " & Am8.FileHistory(n).FileName
    Next n
    If Length = 0 Then OldFile(0).Visible = False Else OldFile(0).Visible = True
End Sub

Private Sub cmbGallary_Click()
    'When you change the gallary dropdown list, this changes the file list
    'box below the combo box to look at the right 'gallary'
    If cmbGallary.List(cmbGallary.ListIndex) <> "[None]" Then
        Gallary.Visible = False
        lblLoadGallary.Visible = True
        'frmMain.ShowGallary.ClearWindow
        DoEvents
        Gallary.FolderLocation = App.Path & "\data\gallarys\" & cmbGallary.List(cmbGallary.ListIndex)
        Gallary.Visible = True
        lblLoadGallary.Visible = False
    End If
    Am8.OpenGallary = cmbGallary.ListIndex
End Sub




Private Sub mnuOnTheWeb_Click(Index As Integer)
    'This menu item enables you to connect to websites on the internet so that you can
    'get updates, help and sample files
    Dim Ie As InternetExplorer
    Set Ie = New InternetExplorer
    Select Case Index
        Case 1: Ie.Navigate "www.geocities.com/animationshop8/index.htm"
        Case 2: Ie.Navigate "www.geocities.com/animationshop8/register.htm"
        Case 3: Ie.Height = 0: Ie.Width = 0: Ie.Navigate "www.geocities.com/animationshop8/samples.zip"
    End Select
    Ie.Visible = True
End Sub

Private Sub mnuHelp_Click(Index As Integer)
    'This menu item controls the Help menu, allowing you to start the help program, or other help commands
    Select Case Index
        Case 1: Am8.ShowAbout
        Case 2: Am8.ShowHelp
        Case 3: Am8.ShowTipofDay
        Case 6: Am8.ShowHelp "EditME!!"
    End Select
End Sub

Public Sub LoadExistingFileWithWindow(FileName As String)
    'This routine makes it easy to load a file from disk. Given the filename, this will
    'create the nessessary window and file, load the file and setup the window correctly
    Dim Key As String
    ActiveFile = "Model_" & Timer & Rnd
    Key = NewWindow
    Am8.File.Add ActiveFile
    If Am8(ActiveFile).LoadFromFile(FileName) = True Then
        SetUpWindow Key, ActiveFile
        Am8(ActiveFile).Saved = True
    Else
        Am8.Forms.Remove Am8.Forms.Count
        Unload ActiveForm
        Am8.File.Remove ActiveFile
        If Am8.Forms.Count = 0 Then LastWindowClose
    End If
End Sub

Private Sub SetUpWindow(Key As String, FileKey As String)
    'This sets up a form to show the default values, and assossiates the relevent parts of the
    'for with the given file object. The view menus on the form are also set up with the current values
    Dim n As Integer
    With Am8.Forms(Key)
        If Am8.ShowSidebar = True Then .mnuView(8).Checked = True
        If Am8.ShowLayers = True Then .LayerTab.Visible = True: .mnuView(9).Checked = True: .CauseFormResize
        If Am8.ShowStatusBar = True Then .mnuView(10).Checked = True
        For n = 1 To 4
            If cBar.Bands(n).Visible = True Then .mnuToolbars(n).Checked = True Else .mnuToolbars(n).Checked = False
        Next n
        .Caption = Am8(FileKey).ModelName
        .Visible = True
        .AssignWindowToFile Am8(FileKey)
        .Tablet.AssignTabletTo Am8(FileKey)
        .TexMap.AssignTexmapTo Am8(FileKey)
        .Tablet.ViewMode = 1
        .Tablet.pShowGrid = True
        .Tablet.pEnableEdit = True
        .Tablet.Refresh
        .LayerTab.AssignLayerDisplayTo Am8(FileKey)
        .Tablet.SetDragDropStyle 1
        .Engine.AssignEngineTo Am8(FileKey)
        .Engine.pAutoRotate = True
        .Engine.pDrawObjects = True
        .Engine.pPerspecitve = True
        .Engine.pAllFace = True
        .Engine.pDrawSkeliton = True
        If DirectXNotAvaliable = True Then .MainTab.Tabs.Remove 6
        Am8(FileKey).Layers(1).Default = True
        .LayerTab.Update
    End With
    UpdateLayerMenu Am8.Forms(Key)
End Sub


Private Sub ck3Dopt_Click(Index As Integer)
    Dim NewValue As Boolean
    If ck3Dopt(Index) = 1 Then NewValue = True Else NewValue = False
    Select Case Index
        Case 0: ActiveForm.Engine.pPerspecitve = NewValue
        Case 1: ActiveForm.Engine.pAllFace = NewValue
        Case 2
            If NewValue = True Then
                Am8(ActiveFile).Geometery.CreateObject "GroundPlain"
                Am8(ActiveFile).Geometery("GroundPlain").CreateObject "Grid", 1, -400, 400, 50, 50, -400, 400, 3, 3
                Am8(ActiveFile).Geometery("GroundPlain").Layer = "Main"
            Else
                Am8(ActiveFile).Geometery.RemoveObject "GroundPlain"
            End If
        Case 3: ActiveForm.Engine.pClipFaces = NewValue
        Case 6: ActiveForm.Engine.pDrawSkeliton = NewValue
        Case 5: ActiveForm.Engine.pDrawObjects = NewValue
        Case 8: ActiveForm.Engine.pLabelJoints = NewValue
        Case 9: ActiveForm.Engine.pHighlightFace = NewValue
        Case 10: ActiveForm.Engine.pHightlightVertex = NewValue
        Case 11: ActiveForm.Engine.pDrawOrigin = NewValue
    End Select
    ActiveForm.Engine.RefreshView
End Sub

Public Sub CreateNewFileWithWindow(Optional Title As String = "")
    'This creates a new blank file and a window, and sets up the window to point to that file
    Dim Key As String
    Static NewIndex As Integer
    NewIndex = NewIndex + 1
    ActiveFile = "Model_" & Timer & Rnd
    If Title = "" Then Title = "Untitled " & NewIndex
    Key = NewWindow
    Am8.File.Add ActiveFile
    With Am8(ActiveFile)
        .ModelName = Title
        .Saved = True
        .Layers.AddLayer "Main", "Main"
        .Layers("Main").Selected = True
        .Scene.CreateScene "BaseFrame", "BaseFrame"
        .Scene("BaseFrame").CreateFrame "BaseFrame"
        .Scene("BaseFrame").CreateFrame "Animate"
        .Scene("BaseFrame").CreateFrame "Inc"
        .Scene.UpdateAllScenes
    End With
    SetUpWindow Key, ActiveFile
End Sub

Public Function AddWindowToCurrentFile(ActiveFileKey As String) As String
    'This creats a new window for the given form. The new window looks at the existing window
    Dim NewKey As String
    NewKey = "j" & Timer & Rnd
    CreateWindow NewKey
    SetUpWindow NewKey, ActiveFile
End Function

Private Sub CreateWindow(WindowKey As String)
    'This routine loads a new Edit Form into the main form. It does not assign a file to the form at this
    'point, so you cannot edit anything on the form till a file is assigned
    Dim NewForm As frmEdit
    Set NewForm = New frmEdit
    NewForm.WindowKey = WindowKey
    Am8.Forms.Add NewForm, WindowKey
    Set NewForm = Nothing
End Sub

Public Function NewWindow() As String
    'This public function creates a new window with a unique Window key.
    Dim NewKey As String
    If Am8.Forms.Count = 0 Then FirstWindowOpen
    NewKey = "j" & Timer & Rnd
    CreateWindow NewKey
    Am8.Forms(NewKey).Visible = True
    NewWindow = NewKey
End Function

Public Function RemoveWindow(WindowCaption As String, Optional FileKey As String = "", Optional WindowKey As String = "") As Boolean
    'This function allows you to close a single window, or close an entire file. If you close a single window, it
    'looks if other windows are pointing at the same file, and if not, then asks you if you want to close the file
    Dim FormCount As Integer, X As frmEdit, FileName As String
    If WindowKey = "" Then 'You are trying to unload an entire file, and all the windows pointing to the file
        If Am8.File(FileKey).Saved = False Then
            Select Case MsgBox(WindowCaption & vbNewLine & vbNewLine & amConfirmSaveAndCloseFile, vbYesNoCancel + vbQuestion + vbDefaultButton3)
                Case vbYes
                    If Am8(ActiveFile).CurrentFilePath = "" Then
                        FileName = SetFileName("Am8", amSaveModelTo)
                        If FileName = "" Then
                            Exit Function
                        Else
                            Am8(ActiveFile).SaveToFile FileName
                            Am8(ActiveFile).CurrentFilePath = FileName
                            Am8(ActiveFile).ModelName = RightClip(Mid(FileName, InStrRev(FileName, "\") + 1), 4)
                        End If
                    Else
                        Am8(ActiveFile).SaveToFile Am8(ActiveFile).CurrentFilePath
                    End If
                    ActiveForm.Caption = Am8(ActiveFile).ModelName
                Case vbCancel: Exit Function
            End Select
            Am8.File.Remove FileKey
        Else
            If Am8(FileKey).NonEditableFile = False And Am8.ConfirmCloseNoSave = True Then If MsgBox(WindowCaption & vbNewLine & vbNewLine & amConfirmCloseFile, vbYesNo + vbQuestion) = vbNo Then Exit Function
            Am8.File.Remove FileKey
        End If
    Else 'You are closing a single window. If its the last window pointing to a file, unload the file as well.
        For Each X In Am8.Forms
            If X.File.Key = Am8.Forms(WindowKey).File.Key Then FormCount = FormCount + 1
        Next X
        If FormCount = 1 Then
            If Am8(Am8.Forms(WindowKey).File.Key).Saved = False Then
                Select Case MsgBox(WindowCaption & vbNewLine & vbNewLine & amConfirmSaveAndCloseFile, vbYesNoCancel + vbQuestion)
                    Case vbYes
                        If Am8(ActiveFile).CurrentFilePath = "" Then
                            FileName = SetFileName("Am8", amSaveModelTo)
                            If FileName = "" Then
                                Exit Function
                            Else
                                Am8(ActiveFile).SaveToFile FileName
                                Am8(ActiveFile).CurrentFilePath = FileName
                                Am8(ActiveFile).ModelName = RightClip(Mid(FileName, InStrRev(FileName, "\") + 1), 4)
                            End If
                        Else
                            Am8(ActiveFile).SaveToFile Am8(ActiveFile).CurrentFilePath
                        End If
                        ActiveForm.Caption = Am8(ActiveFile).ModelName
                    Case vbCancel: Exit Function
                End Select
                Am8.File.Remove Am8.Forms(WindowKey).File.Key
            Else
                If Am8(Am8.Forms(WindowKey).File.Key).NonEditableFile = False Then If Am8.ConfirmCloseNoSave = True Then If MsgBox(WindowCaption & vbNewLine & vbNewLine & amConfirmCloseFile, vbYesNo + vbQuestion) = vbNo Then Exit Function
                Am8.File.Remove Am8.Forms(WindowKey).File.Key
            End If
        End If
    End If
    If WindowKey = "" Then
        For Each X In Am8.Forms: If X.File.Key = FileKey Then Am8.Forms.Remove X.WindowKey: Unload X
        Next X: If Am8.Forms.Count = 0 Then LastWindowClose
    Else
        Am8.Forms.Remove WindowKey: If Am8.Forms.Count = 0 Then LastWindowClose
    End If
    RemoveWindow = True
End Function

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'When you try to close the MDI form, ask if each file should be closed or saved if nessesary
    Dim File As clsFile
    For Each File In Am8.File: If RemoveWindow(File.ModelName, File.Key) = False Then Cancel = 1: Exit Sub
    Next File
    SaveProgramSettings
End Sub

Private Sub FirstWindowOpen()
    'When there are no edit windows open, and a new one is created, this code runs and displays the sidebar and toolbars
    If cmbGallary.ListIndex = -1 And Am8.OpenGallary < cmbGallary.ListCount Then cmbGallary.ListIndex = Am8.OpenGallary Else If cmbGallary.ListCount > 0 Then cmbGallary.ListIndex = 0
    cBar.Visible = True: If tbar(0).buttons(5).Value = tbrPressed Then SideFrame.Visible = True
End Sub

Private Sub LastWindowClose()
    'When the last remaining edit window closes, this code hides the sidebars and toolbars
    cBar.Visible = False
    SideFrame.Visible = False
End Sub

Public Function ShowSidebar(SideBarName As String)
    'This function displays the given sidebars, ands sets the tabs along the top of the sidebar to show pages
    Dim n As Byte, m As Byte
    cmdSidebar.Tabs.Clear
    For m = 1 To Frame.Count
        For n = 1 To Frame.Count
            If LCase(RightClip(Frame(n).Tag, 1)) = LCase(SideBarName) And Val(Right(Frame(n).Tag, 1)) = m Then
                Frame(n).BorderStyle = 0
                cmdSidebar.Tabs.Add , Frame(n).Tag, RightClip(Frame(n).Caption, 2), Val(Right(Frame(n).Caption, 2))
            End If
        Next n
        If Val(tbar(1).buttons(EditButton).Tag) <= cmdSidebar.Tabs.Count Then
            If Frame(m).Tag = cmdSidebar.Tabs(Val(tbar(1).buttons(EditButton).Tag)).Key Then Frame(m).Visible = True Else Frame(m).Visible = False
        Else
            Frame(m).Visible = False
        End If
    Next m
    If cBar.Bands(2).Visible = True Then cmdSidebar.Tabs(Val(tbar(1).buttons(EditButton).Tag)).Selected = True
    cmdSidebar_Click
End Function

Private Sub cmdSidebar_Click()
    'When you click on the edit toolbar, this alters the sidebar to show the required frames
    Dim n As Integer
    If cmdSidebar.Tabs.Count > 0 Then
        For n = 1 To Frame.Count
            If Frame(n).Tag = cmdSidebar.SelectedItem.Key Then Frame(n).Visible = True Else Frame(n).Visible = False
        Next n
        If trFrames.Nodes.Count > 0 Then
            If cmdSidebar.SelectedItem.Caption = "Position" Then
                If Am8(ActiveFile).Joint.CountChildren = 0 Then
                    cmdSidebar.Tabs(1).Selected = True
                    MsgBox amMustCreateJoints, vbInformation
                    Exit Sub
                End If
                If Am8(ActiveFile).Scene.GetFrame(trFrames.SelectedItem.Key) = "" Then
                    cmdSidebar.Tabs(1).Selected = True
                    MsgBox amCantEditScene, vbInformation
                Else
                    With Am8(ActiveFile)
                        frmMain.grGrid.UpdateDisplay .Scene.GetScene(trFrames.SelectedItem.Key), .Scene.GetFrame(trFrames.SelectedItem.Key)
                    End With
                End If
            End If
        End If
        frmMain.tbar(1).buttons(EditButton).Tag = cmdSidebar.SelectedItem.Index
    End If
End Sub

Private Sub cmdCreate_Click(Index As Integer)
    'To create objects, you must press the Create Object button on the sidebar. When you do, this code creates a new
    'object, and passes that object to the CreateObject module so that a new object can be placed in the object
    Dim NewKey As String
    With ActiveForm.Tablet
        Select Case Index
            Case 0
                .ShapeX1 = .AbsoluteX(.ShapeX1): .ShapeX2 = .AbsoluteX(.ShapeX2)
                .ShapeY1 = .AbsoluteY(.ShapeY1): .ShapeY2 = .AbsoluteY(.ShapeY2)
                If (.ShapeX1 - .ShapeX2 <> 0) And (.ShapeY1 - .ShapeY2 <> 0) Or frmMain.ShapeList.SelectedItem.Text = "Wrap" Then
                    Select Case Index
                        Case 0
                            NewKey = "Custom" & Timer & Rnd
                            Am8(ActiveFile).Geometery.CreateObject NewKey
                            Select Case frmMain.ShapeList.SelectedItem.Text
                                Case "Cube": Am8(ActiveFile).Geometery(NewKey).CreateObject "Cube", .ViewMode, .ShapeX1, .ShapeX2, -50, 50, .ShapeY1, .ShapeY2
                                Case "Grid": Am8(ActiveFile).Geometery(NewKey).CreateObject "Grid", .ViewMode, .ShapeX1, .ShapeX2, -50, 50, .ShapeY1, .ShapeY2, ShpProp(3), ShpProp(4)
                                Case "Prism": Am8(ActiveFile).Geometery(NewKey).CreateObject "Prism", .ViewMode, .ShapeX1, .ShapeX2, -50, 50, .ShapeY1, .ShapeY2, ShpProp(1), ShpProp(3), ShpProp(4), ShpProp(2)
                                Case "Face": Am8(ActiveFile).Geometery(NewKey).CreateObject "Face", .ViewMode, .ShapeX1, .ShapeX2, -50, 50, .ShapeY1, .ShapeY2, ShpProp(1), ShpProp(3), ShpProp(4), ShpProp(2)
                                Case "Cone": Am8(ActiveFile).Geometery(NewKey).CreateObject "Cone", .ViewMode, .ShapeX1, .ShapeX2, -50, 50, .ShapeY1, .ShapeY2, ShpProp(1), ShpProp(3), ShpProp(4), ShpProp(2)
                                Case "Dimond": Am8(ActiveFile).Geometery(NewKey).CreateObject "Dimond", .ViewMode, .ShapeX1, .ShapeX2, -50, 50, .ShapeY1, .ShapeY2, ShpProp(1), ShpProp(3), ShpProp(4), ShpProp(2)
                                Case "Torous": Am8(ActiveFile).Geometery(NewKey).CreateObject "Tourus", .ViewMode, .ShapeX1, .ShapeX2, -50, 50, .ShapeY1, .ShapeY2, ShpProp(1), ShpProp(3), ShpProp(5), ShpProp(2)
                                Case "Sphere": Am8(ActiveFile).Geometery(NewKey).CreateObject "Sphere", .ViewMode, .ShapeX1, .ShapeX2, -50, 50, .ShapeY1, .ShapeY2, ShpProp(1), ShpProp(5).Value
                                Case "Star": Am8(ActiveFile).Geometery(NewKey).CreateObject "Star", .ViewMode, .ShapeX1, .ShapeX2, -50, 50, .ShapeY1, .ShapeY2, ShpProp(1), ShpProp(3), ShpProp(4), ShpProp(2)
                                Case "Wrap": Am8(ActiveFile).Geometery(NewKey).CreateObject "Wrap", .ViewMode
                            End Select
                    End Select
                    Am8(ActiveFile).Saved = False: Am8(ActiveFile).FindModelOutline
                    Am8(ActiveFile).Geometery(NewKey).Layer = Am8(ActiveFile).Layers.Default
                Else
                    frmMain.sBar.Panels(2) = "Object too thin to create"
                End If
                .CancelCreateObject
                .Refresh
                
        Case 1
            .ShapeX1 = 0: .ShapeX2 = 0
            .ShapeY1 = 0: .ShapeY2 = 0
            
        End Select
    End With
End Sub

Private Sub cmdAline_Click(Index As Integer)
    'This is the code for the four buttons on the Select sidebar, which moves the selected objects up, down, left and right
    With ActiveForm.Tablet
        If ActiveForm.mnuView(1).Checked = True Then
            If Index = 1 Then Am8(ActiveFile).Geometery.MoveSelected -frmSettings.txtGrid, 0, .ViewMode: Am8(ActiveFile).Joint.MoveSelected -frmSettings.txtGrid, 0, .ViewMode
            If Index = 2 Then Am8(ActiveFile).Geometery.MoveSelected 0, frmSettings.txtGrid, .ViewMode: Am8(ActiveFile).Joint.MoveSelected 0, frmSettings.txtGrid, .ViewMode
            If Index = 3 Then Am8(ActiveFile).Geometery.MoveSelected frmSettings.txtGrid, 0, .ViewMode: Am8(ActiveFile).Joint.MoveSelected frmSettings.txtGrid, 0, .ViewMode
            If Index = 0 Then Am8(ActiveFile).Geometery.MoveSelected 0, -frmSettings.txtGrid, .ViewMode: Am8(ActiveFile).Joint.MoveSelected 0, -frmSettings.txtGrid, .ViewMode
        Else
            If Index = 1 Then Am8(ActiveFile).Geometery.MoveSelected -1, 0, .ViewMode: Am8(ActiveFile).Joint.MoveSelected -1, 0, .ViewMode
            If Index = 2 Then Am8(ActiveFile).Geometery.MoveSelected 0, 1, .ViewMode: Am8(ActiveFile).Joint.MoveSelected 0, 1, .ViewMode
            If Index = 3 Then Am8(ActiveFile).Geometery.MoveSelected 1, 0, .ViewMode: Am8(ActiveFile).Joint.MoveSelected 1, 0, .ViewMode
            If Index = 0 Then Am8(ActiveFile).Geometery.MoveSelected 0, -1, .ViewMode: Am8(ActiveFile).Joint.MoveSelected 0, -1, .ViewMode
       End If
        .Refresh
    End With
End Sub

Private Sub cmdEdit_Click(Index As Integer)
    'This is the command button on the Edit sidebars. When you click on it, depending on the function selected
    'from any of the three sidebar pages, the right peice of code is selected from the Select case statement
    Dim Am As clsObject, n As Integer, StartPoint As Integer, EndPoint As Integer, FaceHolder As Integer
    With Am8(ActiveFile)
        Select Case SelectedEditButton
            Case 4: .Geometery.CombineObject: .FindModelOutline
            Case 5: .Geometery(Am8(ActiveFile).Geometery.FirstSelectedObject).SeperateVertecies
            Case 7: .Geometery(Am8(ActiveFile).Geometery.FirstSelectedObject).DeleteVertecies
            Case 9: For Each Am In .Geometery: If Am.Selected = True Then Am.CompressObject
                    Next Am
            Case 15: .Geometery.FlipSelected 1
            Case 16: .Geometery.FlipSelected 2
            Case 10, 19, 20
            
                    For Each Am In .Geometery
                        Select Case cmbMethod.ListIndex
                            Case 0
                                For n = 1 To Am.Face.Count + 1
                                    If Am.Selected = True And SelectedEditButton = 10 Then Am.FragmentFace n, 0, 0.5
                                    If Am.Selected = True And SelectedEditButton = 19 Then Am.FragmentFace n, 1, 0.5
                                    If Am.Selected = True And SelectedEditButton = 20 Then Am.FragmentFace n, 2, 0.5
                                Next n
                            
                            
                            
                            Case 1, 2
                                If cmbMethod.ListIndex = 1 Then StartPoint = 2: EndPoint = Am.Face.Count - 1: FaceHolder = 1
                                If cmbMethod.ListIndex = 2 Then StartPoint = Am.Face.Count - 1: EndPoint = Am.Face.Count: FaceHolder = Am.Face.Count - 1
                                
                                For n = StartPoint To EndPoint
                                    If Am.Selected = True And SelectedEditButton = 10 Then Am.FragmentFace FaceHolder, 0, 0.5
                                    If Am.Selected = True And SelectedEditButton = 19 Then Am.FragmentFace FaceHolder, 1, 0.5
                                    If Am.Selected = True And SelectedEditButton = 20 Then Am.FragmentFace FaceHolder, 2, 0.5
                                Next n
                            
                            
                            
                        End Select
                    Next Am
            Case 17: For Each Am In .Geometery: If Am.Selected = True Then Am.ReverseFace
                     Next Am
            Case 18: For Each Am In .Geometery: If Am.Selected = True Then Am.Randomize frmMain.ActiveForm.Tablet.ViewMode
                     Next Am
                     .FindModelOutline
        End Select
    End With
    Am8(ActiveFile).Saved = False
    frmMain.ActiveForm.Tablet.Refresh
End Sub

Private Sub cmdScale_Click()
    'This is the command button on the scale sidebar. It is used to scale by exact amounts, and scales both
    'objects and joints by the amounts given
    Dim Cx As Single, Cy As Single, Cz As Single, Am As clsObject, Jm As clsJoint
    Cx = (Am8(ActiveFile).MinX + Am8(ActiveFile).MaxX) / 2
    Cy = (Am8(ActiveFile).MinY + Am8(ActiveFile).MaxY) / 2
    Cz = (Am8(ActiveFile).MinZ + Am8(ActiveFile).MaxZ) / 2
    For Each Am In Am8(ActiveFile).Geometery
        If Am.Selected = True Then
            If SklMode(1) = False Then
                Am.Grow PresetScale, PresetScale, PresetScale, Cx, Cy, Cz
            Else
                Am.Grow sclXDim(0) * 0.01, sclXDim(1) * 0.01, sclXDim(2) * 0.01, Cx, Cy, Cz
            End If
            Am.FindObjectOutline
        End If
    Next Am
    For Each Jm In Am8(ActiveFile).Joint
        If Jm.Selected = True Then
            If SklMode(2) = True Then
                Jm.Grow PresetScale, PresetScale, PresetScale, Cx, Cy, Cz
            Else
                Jm.Grow sclXDim(0) * 0.01, sclXDim(1) * 0.01, sclXDim(2) * 0.01, Cx, Cy, Cz
            End If
        End If
    Next Jm
    Am8(ActiveFile).FindModelOutline
    Am8(ActiveFile).Saved = False
    ActiveForm.Tablet.Refresh
End Sub




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Private Sub mnuGallary_Click(Index As Integer)
    Dim NewName As String, n As Integer
    On Error GoTo FailedToRenameFile
    Select Case Index
        Case 1 To 3
            For n = 1 To 3: mnuGallary(n).Checked = False: Next n
            mnuGallary(Index).Checked = True
            Gallary.GallaryViewMode = Index
        
        Case 5
            NewName = InputBox(amNewItemName, "Rename", Gallary.GetItemName(Gallary.SelectedItem))
            If NewName <> "" Then
                Name Gallary.FolderLocation & "\" & Gallary.GetItemName(Gallary.SelectedItem) & ".cpy" As Gallary.FolderLocation & "\" & NewName & ".cpy"
                Gallary.RefreshItemGrid , 1
            End If
        
        Case 6
            If MsgBox(Gallary.GetItemName(Gallary.SelectedItem) & vbNewLine & vbNewLine & amRemoveItem, vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
                Destroy Gallary.FolderLocation & "\" & Gallary.GetItemName(Gallary.SelectedItem) & ".cpy"
                Gallary.RefreshItemGrid , 1
            End If
    End Select
Exit Sub
FailedToRenameFile:
    MsgBox amFailedToRename, vbExclamation
End Sub

Private Sub mnuFile_Click(Index As Integer)
    Dim FileName As String
    Select Case Index
        Case 1: Am8.ShowNew
        Case 2
            FileName = SelectFileName("Am8", amOpenFileName)
            If FileName <> "" Then frmMain.LoadExistingFileWithWindow FileName
        Case 3
            FileName = SelectFileName("Import", "Import file...")
            If FileName <> "" Then
                CreateNewFileWithWindow
                modImport.ImportModel FileName, Am8(ActiveFile)
                ActiveForm.Tablet.Refresh
            End If
    End Select
End Sub

Private Sub mnuEdit_Click(Index As Integer)
    Am8.ShowSettings
End Sub

Private Sub mnuExit_Click(Index As Integer)
    Unload Me
End Sub


Private Sub mnuTools_Click(Index As Integer)
    Select Case Index
        Case 1: frmCompile.RunAtStart
        Case 2: Am8.OpenAnimator
    End Select
End Sub

Private Sub opShadeMode_Click(Index As Integer)
    ActiveForm.Engine.pRenderSolid = True
    ActiveForm.Engine.ShapeFX = Index + 1
End Sub




Private Sub sldFrames_Click()
    Dim KeyOn As String, n As Integer
    trFrames.SelectedItem.FirstSibling.Selected = True
    KeyOn = trFrames.SelectedItem.Key
    For n = 1 To sldFrames - 1
        KeyOn = trFrames.Nodes(KeyOn).Next.Key
    Next n
    trFrames.Nodes(KeyOn).Selected = True
    trFrames_NodeClick trFrames.SelectedItem
End Sub


Private Sub sldFrames_Scroll()
    sldFrames_Click
End Sub

Private Sub tbar_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    'This is the area of code that controls the toolbars.
    Dim n As Integer, txtvalue, Am As clsObject, FileName As String
    'ActiveForm.Tablet.CancelCreateObject
    Select Case Index
        Case 0
            Select Case Button.Index
            
                Case 1
                    frmMain.CreateNewFileWithWindow
            
                Case 2
                    FileName = SelectFileName("Am8", amOpenFileName)
                    If FileName <> "" Then LoadExistingFileWithWindow FileName
                    
                Case 3
                    If Am8(ActiveFile).CurrentFilePath = "" Then
                        FileName = SetFileName("Am8", amSaveModelTo)
                        If FileName <> "" Then
                            Am8(ActiveFile).SaveToFile FileName
                            Am8(ActiveFile).CurrentFilePath = FileName
                            Am8(ActiveFile).ModelName = RightClip(Mid(FileName, InStrRev(FileName, "\") + 1), 4)
                        End If
                    Else
                        Am8(ActiveFile).SaveToFile Am8(ActiveFile).CurrentFilePath
                    End If
                    ActiveForm.Caption = Am8(ActiveFile).ModelName
            
                Case 5
                    'The 'Hide sidebar' button was pressed
                    If Button.Value = tbrPressed Then
                        ActiveForm.mnuView(8).Checked = True
                        SideFrame.Visible = True
                        Am8.ShowSidebar = True
                    Else
                        ActiveForm.mnuView(8).Checked = False
                        SideFrame.Visible = False
                        Am8.ShowSidebar = False
                    End If
                    MDIForm_Resize
            
                Case 7
                    frmObject.RunAtStart Am8(ActiveFile)
                
                
                Case 8
                    If Am8(ActiveFile).Geometery.CountSelected + Am8(ActiveFile).Joint.CountSelected = 0 Then MsgBox amMustSelectJointOrObject, vbInformation: Exit Sub
                    frmEntity.RunAtStart Am8(ActiveFile)
                
                Case 9
                    frmJoint.RunAtStart Am8(ActiveFile)
                
                Case 10
                    frmSurface.RunAtStart Am8(ActiveFile)
                
                
                
                Case 12
                    If ActiveForm.Tablet.ZoomLevel > 0.25 Then
                        ActiveForm.Tablet.ZoomLevel = ActiveForm.Tablet.ZoomLevel - 0.25
                        ActiveForm.Tablet.Refresh
                        ActiveForm.TexMap.Refresh
                    End If
                    cmdZoomLevels.List(5) = (ActiveForm.Tablet.ZoomLevel * 100) & "%"
                    cmdZoomLevels.ListIndex = 5
                
                Case 14
                    If ActiveForm.Tablet.ZoomLevel < 8 Then
                        ActiveForm.Tablet.ZoomLevel = ActiveForm.Tablet.ZoomLevel + 0.25
                        ActiveForm.Tablet.Refresh
                        ActiveForm.TexMap.Refresh
                    End If
                    cmdZoomLevels.List(5) = (ActiveForm.Tablet.ZoomLevel * 100) & "%"
                    cmdZoomLevels.ListIndex = 5
            End Select
        Case 1
            Select Case Button.Index
                Case 1: frmMain.sBar.Panels(2) = amSelectTool
                Case 2: frmMain.sBar.Panels(2) = amCreateTool
                Case 3: frmMain.sBar.Panels(2) = amEditTool
                Case 4: frmMain.sBar.Panels(2) = amScaleTool
                Case 5: frmMain.sBar.Panels(2) = amRotateTool
                Case 6: frmMain.sBar.Panels(2) = amSkelitonTool
            End Select
            With ActiveForm
                If .MainTab.SelectedItem.Index < 4 Then
                    .MainTab.Tabs(1).Tag = EditButton
                    .MainTab.Tabs(2).Tag = EditButton
                    .MainTab.Tabs(3).Tag = EditButton
                Else
                    .MainTab.SelectedItem.Tag = EditButton
                End If
            End With
            ShowSidebar Button.Key
            
        Case 2
            Select Case Button.Index
                Case 1: sldFrames = 1
                Case 5: sldFrames = sldFrames.Max
            End Select
            'sldFrames_Click
'            ActiveForm.Engine.RefreshView
            
        Case 3
            Select Case Button.Index
                Case 1: Am8(ActiveFile).SelectAll
                Case 2: Am8(ActiveFile).DeselectAll
                Case 5
                    For Each Am In Am8(ActiveFile).Geometery
                        Select Case cmdAttribute
                            Case "Colour"
                                If cmbLogic = "is equal to" And Am.Colour = cmbValue Then Am.Selected = True
                                If cmbLogic = "is not" And Am.Colour <> cmbValue Then Am.Selected = True
                                If cmbLogic = "is greater than" And Am.Colour > cmbValue Then Am.Selected = True
                                If cmbLogic = "is less than" And Am.Colour < cmbValue Then Am.Selected = True
            
                            Case "Face Count"
                                If cmbLogic = "is equal to" And Am.Face.Count = cmbValue Then Am.Selected = True
                                If cmbLogic = "is not" And Am.Face.Count <> cmbValue Then Am.Selected = True
                                If cmbLogic = "is greater than" And Am.Face.Count > cmbValue Then Am.Selected = True
                                If cmbLogic = "is less than" And Am.Face.Count < cmbValue Then Am.Selected = True
            
                            Case "Edge Count"
                                If cmbLogic = "is equal to" And Am.EdgeCount = cmbValue Then Am.Selected = True
                                If cmbLogic = "is not" And Am.EdgeCount <> cmbValue Then Am.Selected = True
                                If cmbLogic = "is greater than" And Am.EdgeCount > cmbValue Then Am.Selected = True
                                If cmbLogic = "is less than" And Am.EdgeCount < cmbValue Then Am.Selected = True
                    
                            Case "Vertex Count"
                                If cmbLogic = "is equal to" And Am.Vertex.Count = Val(cmbValue) Then Am.Selected = True
                                If cmbLogic = "is not" And Am.Vertex.Count <> Val(cmbValue) Then Am.Selected = True
                                If cmbLogic = "is greater than" And Am.Vertex.Count > Val(cmbValue) Then Am.Selected = True
                                If cmbLogic = "is less than" And Am.Vertex.Count < Val(cmbValue) Then Am.Selected = True
            
                        End Select
                    Next Am
                    Am8(ActiveFile).FindModelOutline
            End Select
            ActiveForm.Tablet.Refresh
    End Select
End Sub

Private Function SelectedEditButton() As Integer
    Dim n As Integer
    For n = 1 To frmMain.optEdit.Count - 1
        If optEdit(n) = True Then SelectedEditButton = n
    Next n
End Function

Private Sub optEdit_Click(Index As Integer)
    Static AvoidStartSpaceLock As Boolean
    If AvoidStartSpaceLock = True Then Exit Sub
    Dim n As Integer
    For n = 0 To optEdit.Count - 1
        If Index <> n Then optEdit(n).Value = False
    Next n
    frmExtend.Visible = False
    frmFragment.Visible = False
    If ActiveForm Is Nothing Then
    Else
        Select Case Index
            Case 7, 8, 13: ActiveForm.Tablet.pShowVertecies = True: ActiveForm.Tablet.pShowFaces = False
            Case 6, 14, 12, 11: ActiveForm.Tablet.pShowFaces = True: ActiveForm.Tablet.pShowVertecies = False
            Case Else: ActiveForm.Tablet.pShowFaces = False: ActiveForm.Tablet.pShowVertecies = False
        End Select
        ActiveForm.Tablet.Refresh
        If chkSelect(6) = 1 Then ActiveForm.Tablet.pShowVertecies = True
    End If
    Select Case Index
        Case 0, 1, 2
            sBar.Panels(2) = amEditTool01
            AvoidStartSpaceLock = True
            optEdit(0) = True:  optEdit(1) = True:    optEdit(2) = True
            AvoidStartSpaceLock = False
        Case 3: sBar.Panels(2) = amEditTool03
        Case 4: sBar.Panels(2) = amEditTool04
        Case 5: sBar.Panels(2) = amEditTool05
        Case 6: sBar.Panels(2) = amEditTool06
        Case 7: sBar.Panels(2) = amEditTool07
        Case 8: sBar.Panels(2) = amEditTool08
        Case 9: sBar.Panels(2) = amEditTool09
        Case 10, 19: sBar.Panels(2) = amEditTool10
        Case 11: sBar.Panels(2) = amEditTool11
        Case 12: sBar.Panels(2) = amEditTool12: frmExtend.Visible = True
        Case 13: sBar.Panels(2) = amEditTool13
        Case 14: sBar.Panels(2) = amEditTool14
        Case 15: sBar.Panels(2) = amEditTool15
        Case 16: sBar.Panels(2) = amEditTool16
        Case 17: sBar.Panels(2) = amEditTool17
        Case 18: sBar.Panels(2) = amEditTool18
        Case 20: frmFragment.Visible = True: sBar.Panels(2) = amEditTool10
    End Select
End Sub

Private Sub f1_DblClick(Index As Integer)
    GetFile.ShowColor
    f1(Index).BackColor = GetFile.Color
    If Index = 0 Then ActiveForm.TexMap.ForeColour = GetFile.Color
    If Index = 1 Then ActiveForm.TexMap.BackColour = GetFile.Color
End Sub

Private Sub Scales_Click()
    SklMode(2) = True
    PresetScale = Scales.Text
End Sub

Private Sub sclXDim_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    SklMode(1) = True
    If Button = 2 Then sclXDim(Index) = 100
End Sub

Private Sub Mode_Click(Index As Integer)
    ActiveForm.TexMap.DrawMode = Index + 1
End Sub

Private Sub lstPattern_Click()
    If ActiveForm Is Nothing Then Else ActiveForm.TexMap.FillPattern = lstPattern.ListIndex
End Sub

Private Sub sldWidth_Click()
    ActiveForm.TexMap.LineWidth = sldWidth
End Sub

Private Sub sldWidth_Scroll()
    ActiveForm.TexMap.LineWidth = sldWidth
End Sub

Private Sub OldFile_Click(Index As Integer)
    frmMain.CreateNewFileWithWindow Am8.FileHistory(Index).FileName
    Am8(ActiveFile).LoadFromFile Am8.FileHistory(Index).FilePath
    frmMain.ActiveForm.Tablet.Refresh
End Sub

Private Sub setResize_Timer()
    setResize.Interval = 0
    CauseFormResize
End Sub

Private Sub ShpProp_Click(Index As Integer)
    With ActiveForm.Tablet
        .DrawShapeGuide frmMain.ShapeList.SelectedItem.Text, .ShapeX1, .ShapeY1, .ShapeX2, .ShapeY2
    End With
End Sub

Private Sub ShapeList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'When you click on the list of the different shape types, the displayed options and scroll bars is updated
    SetNewShapeMenu ShapeList.SelectedItem.Text
    With ActiveForm.Tablet
        .DrawShapeGuide frmMain.ShapeList.SelectedItem.Text, .ShapeX1, .ShapeY1, .ShapeX2, .ShapeY2
    End With
End Sub

Private Sub QuickSpin_Click(Index As Integer)
    Dim n As Integer, CentreMode As Integer
    For n = 1 To optGetCenter.Count - 1
        If optGetCenter(n).Value = True Then CentreMode = n
    Next n
    Select Case Index
        Case 0: Am8(ActiveFile).RotateSelection -10, ActiveForm.Tablet.ViewMode, CentreMode
        Case 1: Am8(ActiveFile).RotateSelection 10, ActiveForm.Tablet.ViewMode, CentreMode
        Case 2: Am8(ActiveFile).RotateSelection -90, ActiveForm.Tablet.ViewMode, CentreMode
        Case 3: Am8(ActiveFile).RotateSelection 90, ActiveForm.Tablet.ViewMode, CentreMode
    End Select
    Am8(ActiveFile).FindModelOutline
    ActiveForm.Tablet.Refresh
End Sub

Private Sub ShpProp_Scroll(Index As Integer)
    With ActiveForm.Tablet
        .DrawShapeGuide frmMain.ShapeList.SelectedItem.Text, .ShapeX1, .ShapeY1, .ShapeX2, .ShapeY2
    End With
End Sub

Private Sub sldShade_Click()
    If ActiveForm Is Nothing Then Else ActiveForm.DXEngine.SetMode sldShade.ListIndex + 1
End Sub

Private Sub tbar_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 2 Then PopupMenu ActiveForm.mnuView(5)
End Sub

Private Sub sldLight_Click(Index As Integer)
    ActiveForm.DXEngine.SetLights sldLight(0), sldLight(1)
End Sub

Private Sub sldLight_Scroll(Index As Integer)
    ActiveForm.DXEngine.SetLights sldLight(0), sldLight(1)
End Sub

Private Sub ckShowJoint_Click()
    ActiveForm.DXEngine.pShowSkeliton = IntBo(ckShowJoint)
    ActiveForm.DXEngine.PlaceModelInWindow
End Sub

Private Sub lstLightStyle_Click()
    If ActiveForm Is Nothing Then Else If lstLightStyle.Text = "<None>" Then ActiveForm.DXEngine.pLightPattern = "" Else ActiveForm.DXEngine.pLightPattern = Am8.LightStyle(lstLightStyle.ListIndex).Pattern
End Sub




Private Sub cmdRender_Click()
    MousePointer = 11
    If optWireFrame = False Then ActiveForm.Engine.pRenderSolid = True
    ActiveForm.Engine.RefreshView
    MousePointer = 0
End Sub





Private Sub cmdScene_Click(Index As Integer)
    Dim NewName As String, NewKey As String, SceneON As String, FrameON As String
    Select Case Index
        Case 0
            SceneON = Am8(ActiveFile).Scene.GetScene(trFrames.SelectedItem.Key)
            If InStr(1, trFrames.SelectedItem.Key, "@") = 0 Then
                NewName = InputBox("Enter a new name for this scene", "Rename Scene", trFrames.SelectedItem.Text)
                If NewName = "" Then Exit Sub
                Am8(ActiveFile).Scene(SceneON).Name = NewName
            Else
                FrameON = Am8(ActiveFile).Scene.GetFrame(trFrames.SelectedItem.Key)
                NewName = InputBox("Enter a new name for this frame", "Rename frame", trFrames.SelectedItem.Text)
                If NewName = "" Then Exit Sub
                Am8(ActiveFile).Scene(SceneON).Frame(FrameON).Name = NewName
            End If
            trFrames.SelectedItem.Text = NewName
            Am8(ActiveFile).Scene.ListScenesInWindow cmbScenes
    
        Case 1
            SceneON = Am8(ActiveFile).Scene.GetScene(trFrames.SelectedItem.Key)
            FrameON = Am8(ActiveFile).Scene.GetFrame(trFrames.SelectedItem.Key)
            If MsgBox(Am8(ActiveFile).Scene(SceneON).Name & vbNewLine & vbNewLine & amRemoveScene, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                Am8(ActiveFile).Scene.RemoveScene SceneON
                Am8(ActiveFile).Scene.AddSceneToWindow frmMain.trFrames
            End If
        
        Case 2
            SceneON = Am8(ActiveFile).Scene.GetScene(trFrames.SelectedItem.Key)
            FrameON = Am8(ActiveFile).Scene.GetFrame(trFrames.SelectedItem.Key)
            Am8(ActiveFile).Scene(SceneON).CreateFrame "F" & (Am8(ActiveFile).Scene(SceneON).FrameCount + 1)
            Am8(ActiveFile).Scene.AddSceneToWindow frmMain.trFrames
            Am8(ActiveFile).Scene.UpdateAllScenes
            DisplayFrameScrollBar SceneON, FrameON
        
        Case 3
            NewName = InputBox("Enter a name for the new scene", "New scene")
            If NewName = "" Then Exit Sub
            NewKey = "Scene" & Timer & Rnd * 10
            Am8(ActiveFile).Scene.CreateScene NewKey, NewName
            Am8(ActiveFile).Scene.AddSceneToWindow trFrames
            Am8(ActiveFile).Scene.ListScenesInWindow cmbScenes
            trFrames.Nodes(NewKey).Selected = True
    
        Case 4
        
        Case 5
        
        Case 6
    
    End Select
    Am8(ActiveFile).Scene.ListScenesInWindow cmbScenes
    Am8(ActiveFile).Saved = False
End Sub



Private Sub trFrames_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 2 Then PopupMenu menuFrame
End Sub




Private Sub mnuFrame_Click(Index As Integer)
    Dim SceneON As String, FrameON As String, NodeOn As Node, n As Integer, NewKey As String
    SceneON = Am8(ActiveFile).Scene.GetScene(trFrames.SelectedItem.Key)
    FrameON = Am8(ActiveFile).Scene.GetFrame(trFrames.SelectedItem.Key)
    If trFrames.SelectedItem.Key = "BaseFrame@BaseFrame" Then MsgBox amNotBaseFrame, vbInformation: Exit Sub
    If Index = 2 And trFrames.SelectedItem.Key = trFrames.SelectedItem.FirstSibling.Key Then Index = 1
    If Index = 3 And trFrames.SelectedItem.Key = trFrames.SelectedItem.LastSibling.Key Then Index = 4
    Select Case Index
        Case 1
            Am8(ActiveFile).Scene(SceneON).CreateFrame "F" & (Am8(ActiveFile).Scene(SceneON).FrameCount + 1), 1
            Am8(ActiveFile).Scene.AddSceneToWindow frmMain.trFrames
        
        Case 2
            Set NodeOn = trFrames.SelectedItem.FirstSibling
            For n = 1 To trFrames.SelectedItem.Parent.Children - 1
                If NodeOn.Next.Key = trFrames.SelectedItem.Key Then
                    Am8(ActiveFile).Scene(SceneON).CreateFrame "F" & (Am8(ActiveFile).Scene(SceneON).FrameCount + 1), n + 1
                    Am8(ActiveFile).Scene.AddSceneToWindow frmMain.trFrames
                    Exit Sub
                End If
                Set NodeOn = NodeOn.Next
            Next n
        
        Case 3
            Set NodeOn = trFrames.SelectedItem.FirstSibling.Next
            For n = 2 To trFrames.SelectedItem.Parent.Children
                If NodeOn.Previous.Key = trFrames.SelectedItem.Key Then
                    Am8(ActiveFile).Scene(SceneON).CreateFrame "F" & (Am8(ActiveFile).Scene(SceneON).FrameCount + 1), n
                    Am8(ActiveFile).Scene.AddSceneToWindow frmMain.trFrames
                    Exit Sub
                End If
                Set NodeOn = NodeOn.Next
            Next n
        
        Case 4
            Am8(ActiveFile).Scene(SceneON).CreateFrame "F" & (Am8(ActiveFile).Scene(SceneON).FrameCount + 1)
            Am8(ActiveFile).Scene.AddSceneToWindow frmMain.trFrames

        Case 6
            NewKey = "Scene" & Timer & Rnd
            Am8(ActiveFile).Scene.CreateScene NewKey, Am8(ActiveFile).Scene(SceneON).Name & "2"
            Set NodeOn = trFrames.SelectedItem.Next
            Do
                If NodeOn Is Nothing Then
                Else
                    Am8(ActiveFile).Scene(SceneON).RemoveFrame Am8(ActiveFile).Scene.GetFrame(NodeOn.Key)
                    Set NodeOn = NodeOn.Next
                End If
            Loop Until NodeOn Is Nothing
            Am8(ActiveFile).Scene.AddSceneToWindow frmMain.trFrames
            trFrames.Nodes(NewKey).Selected = True
        
        Case 7
            Am8(ActiveFile).Scene.RemoveScene SceneON
            Am8(ActiveFile).Scene.AddSceneToWindow frmMain.trFrames
            
        Case 11
            If FrameON = "" Then MsgBox amSelectScene, vbInformation: Exit Sub
            Am8(ActiveFile).Scene(SceneON).RemoveFrame FrameON
            Am8(ActiveFile).Scene.AddSceneToWindow frmMain.trFrames

        Case 12
            If MsgBox(Am8(ActiveFile).Scene(SceneON).Name & vbNewLine & vbNewLine & amRemoveScene, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                Am8(ActiveFile).Scene.RemoveScene SceneON
                Am8(ActiveFile).Scene.AddSceneToWindow frmMain.trFrames
            End If
            
        Case 14
            If FrameON = "" Then MsgBox amSelectScene, vbInformation: Exit Sub
            n = InputBox("Enter the required number of frame itterations", , Am8(ActiveFile).Scene(SceneON)(FrameON).Smooth)
            Am8(ActiveFile).Scene(SceneON)(FrameON).Smooth = n
            

    End Select
    Am8(ActiveFile).Scene.ListScenesInWindow cmbScenes
End Sub



Private Sub Joints_AfterLabelEdit(Cancel As Integer, NewString As String)
    If Joints.SelectedItem.Key <> "BaseJoint" Then
        Am8(ActiveFile).Joint(Joints.SelectedItem.Key).Name = NewString
    End If
End Sub



Private Function DisplayFrameScrollBar(SceneON As String, FrameON As String)
    sldFrames.Visible = False
    If Am8(ActiveFile).Scene(SceneON).FrameCount > 1 Then
        sldFrames.Visible = True
        sldFrames.Max = Am8(ActiveFile).Scene(SceneON).FrameCount
        sldFrames.Value = Am8(ActiveFile).Scene(SceneON).FrameIndex(FrameON)
        sldFrames.TickFrequency = 1
        If sldFrames.Max > 25 Then sldFrames.TickFrequency = 2
        If sldFrames.Max > 40 Then sldFrames.TickFrequency = 5
        If sldFrames.Max > 60 Then sldFrames.TickFrequency = 10
        sldFrames.Visible = True
    End If
End Function


Private Sub cmdZoomLevels_Click()
    Dim NewLevel As String
    If cmdZoomLevels.ListIndex < 6 Then
        ActiveForm.Tablet.ZoomLevel = Val(cmdZoomLevels.Text) / 100
        ActiveForm.Tablet.Refresh
    Else
        ActiveForm.Tablet.ZoomLevel = (Int(ActiveForm.Tablet.ZoomToSelected * 100)) * 0.01
        ActiveForm.Tablet.CenterView
        ActiveForm.Tablet.Refresh
        cmdZoomLevels.List(5) = (ActiveForm.Tablet.ZoomLevel * 100) & "%"
        cmdZoomLevels.ListIndex = 5
    End If
End Sub

Private Sub trFrames_NodeClick(ByVal Node As MSComctlLib.Node)
    If Am8(ActiveFile).Scene.GetFrame(trFrames.SelectedItem.Key) <> "" Then
        With Am8(ActiveFile).Scene(Am8(ActiveFile).Scene.GetScene(trFrames.SelectedItem.Key)).Frame(Am8(ActiveFile).Scene.GetFrame(trFrames.SelectedItem.Key))
            If .Smooth = 0 Then .Smooth = 1
            DisplayFrameScrollBar Am8(ActiveFile).Scene.GetScene(trFrames.SelectedItem.Key), Am8(ActiveFile).Scene.GetFrame(trFrames.SelectedItem.Key)
            Am8(ActiveFile).Scene.CopyToAnimate Am8(ActiveFile).Scene.GetScene(trFrames.SelectedItem.Key), Am8(ActiveFile).Scene.GetFrame(trFrames.SelectedItem.Key), .Smooth
            ActiveForm.Engine.SetTimer .Smooth, 10
            ActiveForm.Engine.RefreshView
        End With
    End If
End Sub
