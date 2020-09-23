VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPresetEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Visualization Preset Editor"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frm 
      Caption         =   "Spectrum Settings"
      Height          =   4815
      Index           =   1
      Left            =   120
      TabIndex        =   39
      Top             =   600
      Width           =   6015
      Begin VB.Frame frmSpec 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3975
         Index           =   0
         Left            =   240
         TabIndex        =   41
         Top             =   600
         Width           =   5535
         Begin VB.OptionButton optSpecDrawStyle 
            Caption         =   "&Solid draw style"
            Height          =   255
            Index           =   2
            Left            =   3480
            TabIndex        =   124
            ToolTipText     =   "draw solid color bars"
            Top             =   1800
            Width           =   1455
         End
         Begin VB.CheckBox chkSpecFall 
            Caption         =   "&Bars falls faster and faster."
            Height          =   255
            Left            =   1560
            TabIndex        =   67
            Top             =   3720
            Width           =   2175
         End
         Begin VB.OptionButton optSpecDrawStyle 
            Caption         =   "&Fire draw style"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   54
            ToolTipText     =   "draw scaled gradient"
            Top             =   1800
            Width           =   1335
         End
         Begin VB.OptionButton optSpecDrawStyle 
            Caption         =   "&Normal draw style"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   53
            ToolTipText     =   "Draw normal gradient"
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Frame Frame2 
            Caption         =   "Detail"
            Height          =   1215
            Left            =   120
            TabIndex        =   43
            Top             =   0
            Width           =   5295
            Begin MSComctlLib.Slider sldSpec 
               Height          =   255
               Index           =   0
               Left            =   1320
               TabIndex        =   44
               ToolTipText     =   "How much to scale vertically."
               Top             =   360
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   450
               _Version        =   393216
               LargeChange     =   2
               Min             =   10
               Max             =   30
               SelStart        =   10
               TickFrequency   =   2
               Value           =   10
            End
            Begin MSComctlLib.Slider sldSpec 
               Height          =   255
               Index           =   1
               Left            =   1320
               TabIndex        =   47
               ToolTipText     =   "Number of values to process. "
               Top             =   720
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   450
               _Version        =   393216
               LargeChange     =   50
               Min             =   64
               Max             =   511
               SelStart        =   511
               TickFrequency   =   25
               Value           =   511
            End
            Begin VB.Label lblSpec 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "1"
               Height          =   195
               Index           =   3
               Left            =   5040
               TabIndex        =   49
               Top             =   720
               Width           =   90
            End
            Begin VB.Label lblSpec 
               AutoSize        =   -1  'True
               Caption         =   "View Values:"
               Height          =   195
               Index           =   2
               Left            =   240
               TabIndex        =   48
               Top             =   720
               Width           =   915
            End
            Begin VB.Label lblSpec 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "1"
               Height          =   195
               Index           =   1
               Left            =   5040
               TabIndex        =   46
               Top             =   360
               Width           =   90
            End
            Begin VB.Label lblSpec 
               AutoSize        =   -1  'True
               Caption         =   "Vertical Zoom:"
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   45
               Top             =   360
               Width           =   1020
            End
         End
         Begin MSComctlLib.Slider sldSpec 
            Height          =   255
            Index           =   2
            Left            =   960
            TabIndex        =   55
            ToolTipText     =   "Size of bars in pixels, when subtype is Bars."
            Top             =   2160
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   2
            Min             =   2
            Max             =   15
            SelStart        =   4
            Value           =   4
         End
         Begin MSComctlLib.Slider sldSpec 
            Height          =   255
            Index           =   3
            Left            =   960
            TabIndex        =   58
            ToolTipText     =   "Size of line in pixels, when subtype is Line."
            Top             =   2520
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   2
            Min             =   1
            SelStart        =   1
            TickFrequency   =   2
            Value           =   1
         End
         Begin MSComctlLib.Slider sldSpec 
            Height          =   255
            Index           =   4
            Left            =   1440
            TabIndex        =   61
            Top             =   3000
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   2
            Max             =   2000
            SelStart        =   100
            TickFrequency   =   100
            Value           =   100
         End
         Begin MSComctlLib.Slider sldSpec 
            Height          =   255
            Index           =   5
            Left            =   1440
            TabIndex        =   62
            Top             =   3360
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   2
            Min             =   1
            Max             =   100
            SelStart        =   1
            TickFrequency   =   10
            Value           =   1
         End
         Begin VB.Label lblSpec 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "no pause"
            Height          =   195
            Index           =   15
            Left            =   4800
            TabIndex        =   66
            Top             =   3000
            Width           =   660
         End
         Begin VB.Label lblSpec 
            AutoSize        =   -1  'True
            Caption         =   "Top Pause Time:"
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   65
            Top             =   3000
            Width           =   1215
         End
         Begin VB.Label lblSpec 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1"
            Height          =   195
            Index           =   13
            Left            =   5400
            TabIndex        =   64
            Top             =   3360
            Width           =   90
         End
         Begin VB.Label lblSpec 
            AutoSize        =   -1  'True
            Caption         =   "Drop speed:"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   63
            Top             =   3360
            Width           =   870
         End
         Begin VB.Label lblSpec 
            AutoSize        =   -1  'True
            Caption         =   "Line size:"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   60
            Top             =   2520
            Width           =   660
         End
         Begin VB.Label lblSpec 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1 pixels"
            Height          =   195
            Index           =   6
            Left            =   4920
            TabIndex        =   59
            Top             =   2520
            Width           =   525
         End
         Begin VB.Label lblSpec 
            AutoSize        =   -1  'True
            Caption         =   "Bar size:"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   57
            Top             =   2160
            Width           =   600
         End
         Begin VB.Label lblSpec 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1 pixels"
            Height          =   195
            Index           =   4
            Left            =   4920
            TabIndex        =   56
            Top             =   2160
            Width           =   525
         End
         Begin VB.Label lblSpecColor 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Line Color"
            Height          =   375
            Index           =   2
            Left            =   3720
            TabIndex        =   52
            ToolTipText     =   "Color of line when subtype is Line."
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label lblSpecColor 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Lower Gradient Color"
            Height          =   375
            Index           =   1
            Left            =   1920
            TabIndex        =   51
            ToolTipText     =   "Color of lower gradient color."
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label lblSpecColor 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Upper Gradient Color"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   50
            ToolTipText     =   "Color of upper gradient color."
            Top             =   1320
            Width           =   1695
         End
      End
      Begin VB.Frame frmSpec 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3975
         Index           =   1
         Left            =   240
         TabIndex        =   42
         Top             =   600
         Width           =   5535
         Begin VB.CheckBox chkSpecPeakFall 
            Caption         =   "&Peaks falls faster and faster."
            Height          =   255
            Left            =   1680
            TabIndex        =   70
            ToolTipText     =   "If on, peaks fall faster and faster"
            Top             =   2040
            Width           =   2415
         End
         Begin VB.CheckBox chkSpecPeaks 
            Caption         =   "&Show Peaks"
            Height          =   255
            Left            =   240
            TabIndex        =   69
            Top             =   120
            Width           =   1215
         End
         Begin MSComctlLib.Slider sldSpec 
            Height          =   255
            Index           =   6
            Left            =   1560
            TabIndex        =   71
            ToolTipText     =   "Peak pause time at it's new value"
            Top             =   1080
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   2
            Max             =   2000
            SelStart        =   100
            TickFrequency   =   100
            Value           =   100
         End
         Begin MSComctlLib.Slider sldSpec 
            Height          =   255
            Index           =   7
            Left            =   1560
            TabIndex        =   72
            ToolTipText     =   "Peak drop speed in units/frame"
            Top             =   1560
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   2
            Min             =   1
            Max             =   100
            SelStart        =   1
            TickFrequency   =   10
            Value           =   1
         End
         Begin MSComctlLib.Slider sldSpec 
            Height          =   255
            Index           =   8
            Left            =   1440
            TabIndex        =   78
            ToolTipText     =   "Fade speed"
            Top             =   3360
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   2
            Max             =   50
            TickFrequency   =   2
         End
         Begin MSComctlLib.Slider sldSpec 
            Height          =   255
            Index           =   9
            Left            =   1440
            TabIndex        =   76
            ToolTipText     =   "How much to correct the higher frequency values, making their bars larger. 0 is off."
            Top             =   2760
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   2
         End
         Begin VB.Label lblSpec 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "None"
            Height          =   195
            Index           =   9
            Left            =   5025
            TabIndex        =   123
            Top             =   2760
            Width           =   390
         End
         Begin VB.Label lblSpec 
            AutoSize        =   -1  'True
            Caption         =   "Value Correction:"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   122
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label lblSpec 
            AutoSize        =   -1  'True
            Caption         =   "Fade Speed:"
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   80
            Top             =   3360
            Width           =   915
         End
         Begin VB.Label lblSpec 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Disabled"
            Height          =   195
            Index           =   21
            Left            =   4800
            TabIndex        =   79
            Top             =   3360
            Width           =   615
         End
         Begin VB.Label lblSpec 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "no pause"
            Height          =   195
            Index           =   17
            Left            =   4830
            TabIndex        =   77
            Top             =   1080
            Width           =   660
         End
         Begin VB.Label lblSpec 
            AutoSize        =   -1  'True
            Caption         =   "Top Pause Time:"
            Height          =   195
            Index           =   16
            Left            =   240
            TabIndex        =   75
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label lblSpec 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1"
            Height          =   195
            Index           =   19
            Left            =   5400
            TabIndex        =   74
            Top             =   1560
            Width           =   90
         End
         Begin VB.Label lblSpec 
            AutoSize        =   -1  'True
            Caption         =   "Drop speed:"
            Height          =   195
            Index           =   18
            Left            =   240
            TabIndex        =   73
            Top             =   1560
            Width           =   870
         End
         Begin VB.Label lblSpecColor 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Peak Color"
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   68
            ToolTipText     =   "Color of line when subtype is Line."
            Top             =   480
            Width           =   1695
         End
      End
      Begin MSComctlLib.TabStrip tabSpec 
         Height          =   4455
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   7858
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Main"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Peaks/Fade/Correction"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame frm 
      Caption         =   "Beat Detector Settings"
      Height          =   4815
      Index           =   3
      Left            =   120
      TabIndex        =   96
      Top             =   600
      Width           =   6015
      Begin VB.CommandButton cmdBeat 
         Caption         =   "&Reset Image"
         Height          =   375
         Index           =   1
         Left            =   4680
         TabIndex        =   121
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdBeat 
         Caption         =   "&Set Image"
         Height          =   375
         Index           =   0
         Left            =   3360
         TabIndex        =   120
         Top             =   240
         Width           =   1215
      End
      Begin MSComctlLib.Slider sldBeat 
         Height          =   255
         Index           =   6
         Left            =   960
         TabIndex        =   97
         ToolTipText     =   "Lower value of range to examine for beats, def. 0"
         Top             =   720
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   10
         Max             =   510
         TickFrequency   =   10
      End
      Begin MSComctlLib.Slider sldBeat 
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   99
         ToolTipText     =   "Upper value of range to examine for beats, def. 25"
         Top             =   1080
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   10
         Min             =   1
         Max             =   511
         SelStart        =   511
         TickFrequency   =   10
         Value           =   511
      End
      Begin MSComctlLib.Slider sldBeat 
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   102
         ToolTipText     =   "Multiply Value, def. 2.5"
         Top             =   1680
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   10
         Max             =   50
         SelStart        =   50
         Value           =   50
      End
      Begin MSComctlLib.Slider sldBeat 
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   105
         ToolTipText     =   "Minimum/startup zoom value, def. 0.75"
         Top             =   2160
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         TickFrequency   =   2
      End
      Begin MSComctlLib.Slider sldBeat 
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   108
         ToolTipText     =   "Minimum zoom value increase to begin rotation at, def 0.25"
         Top             =   2640
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         TickFrequency   =   2
      End
      Begin MSComctlLib.Slider sldBeat 
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   111
         ToolTipText     =   "Move speed of rotation, in radians per frame update, def 0.05"
         Top             =   3120
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         TickFrequency   =   2
      End
      Begin MSComctlLib.Slider sldBeat 
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   114
         ToolTipText     =   "Rotation speed when subtype is Rotator, def. 0.05"
         Top             =   3600
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         TickFrequency   =   2
      End
      Begin MSComctlLib.Slider sldBeat 
         Height          =   255
         Index           =   7
         Left            =   1200
         TabIndex        =   117
         ToolTipText     =   "Fade speed"
         Top             =   4320
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         Max             =   50
         TickFrequency   =   2
      End
      Begin VB.Label lblBeat 
         AutoSize        =   -1  'True
         Caption         =   "Fade Speed:"
         Height          =   195
         Index           =   13
         Left            =   240
         TabIndex        =   119
         Top             =   4320
         Width           =   915
      End
      Begin VB.Label lblBeat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Disabled"
         Height          =   195
         Index           =   14
         Left            =   5040
         TabIndex        =   118
         Top             =   4320
         Width           =   615
      End
      Begin VB.Label lblBeat 
         AutoSize        =   -1  'True
         Caption         =   "Rot Speed:"
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   116
         Top             =   3600
         Width           =   810
      End
      Begin VB.Label lblBeat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "2,5"
         Height          =   195
         Index           =   11
         Left            =   5400
         TabIndex        =   115
         Top             =   3600
         Width           =   225
      End
      Begin VB.Label lblBeat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "2,5"
         Height          =   195
         Index           =   10
         Left            =   5400
         TabIndex        =   113
         Top             =   3120
         Width           =   225
      End
      Begin VB.Label lblBeat 
         AutoSize        =   -1  'True
         Caption         =   "Rot Move:"
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   112
         Top             =   3120
         Width           =   750
      End
      Begin VB.Label lblBeat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "2,5"
         Height          =   195
         Index           =   8
         Left            =   5400
         TabIndex        =   110
         Top             =   2640
         Width           =   225
      End
      Begin VB.Label lblBeat 
         AutoSize        =   -1  'True
         Caption         =   "Min to Rot:"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   109
         Top             =   2640
         Width           =   780
      End
      Begin VB.Label lblBeat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "2,5"
         Height          =   195
         Index           =   6
         Left            =   5400
         TabIndex        =   107
         Top             =   2160
         Width           =   225
      End
      Begin VB.Label lblBeat 
         AutoSize        =   -1  'True
         Caption         =   "Min Zoom:"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   106
         Top             =   2160
         Width           =   750
      End
      Begin VB.Label lblBeat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "2,5"
         Height          =   195
         Index           =   4
         Left            =   5400
         TabIndex        =   104
         Top             =   1680
         Width           =   225
      End
      Begin VB.Label lblBeat 
         AutoSize        =   -1  'True
         Caption         =   "Multiply:"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   103
         Top             =   1680
         Width           =   570
      End
      Begin VB.Label lblBeat 
         AutoSize        =   -1  'True
         Caption         =   "Detect values: 0-511"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   101
         Top             =   360
         Width           =   1485
      End
      Begin VB.Label lblBeat 
         AutoSize        =   -1  'True
         Caption         =   "Upper:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   100
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lblBeat 
         AutoSize        =   -1  'True
         Caption         =   "Lower:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   98
         Top             =   720
         Width           =   480
      End
   End
   Begin VB.Frame frm 
      Caption         =   "Oscilliscope Settings"
      Height          =   4815
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   6015
      Begin VB.Frame frmOsc 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3975
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   5535
         Begin VB.Frame Frame1 
            Caption         =   "Detail"
            Height          =   1335
            Left            =   120
            TabIndex        =   16
            Top             =   120
            Width           =   5295
            Begin VB.OptionButton optOscDet 
               Caption         =   "Stereo"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   19
               Top             =   360
               Width           =   855
            End
            Begin VB.OptionButton optOscDet 
               Caption         =   "Mono"
               Height          =   255
               Index           =   1
               Left            =   1440
               TabIndex        =   18
               Top             =   360
               Width           =   855
            End
            Begin MSComctlLib.Slider sldOsc 
               Height          =   255
               Index           =   0
               Left            =   720
               TabIndex        =   17
               ToolTipText     =   "Number of values to skip between each drawn value"
               Top             =   720
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   450
               _Version        =   393216
               LargeChange     =   2
               Min             =   1
               SelStart        =   1
               Value           =   1
            End
            Begin VB.Label lblOsc 
               AutoSize        =   -1  'True
               Caption         =   "Skip:"
               Height          =   195
               Index           =   1
               Left            =   240
               TabIndex        =   21
               Top             =   720
               Width           =   360
            End
            Begin VB.Label lblOsc 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "2 values"
               Height          =   195
               Index           =   0
               Left            =   4440
               TabIndex        =   20
               Top             =   720
               Width           =   600
            End
         End
         Begin MSComctlLib.Slider sldOsc 
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   8
            ToolTipText     =   "Right Channel Line Width"
            Top             =   2400
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   2
            Min             =   1
            SelStart        =   1
            Value           =   1
         End
         Begin MSComctlLib.Slider sldOsc 
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   9
            ToolTipText     =   "Left Channel Line Width"
            Top             =   2880
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   2
            Min             =   1
            SelStart        =   1
            Value           =   1
         End
         Begin MSComctlLib.Slider sldOsc 
            Height          =   255
            Index           =   3
            Left            =   1440
            TabIndex        =   22
            ToolTipText     =   "Fade speed"
            Top             =   3480
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   2
            Max             =   50
            TickFrequency   =   2
         End
         Begin VB.Label lblOsc 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Disabled"
            Height          =   195
            Index           =   3
            Left            =   4800
            TabIndex        =   24
            Top             =   3480
            Width           =   615
         End
         Begin VB.Label lblOsc 
            AutoSize        =   -1  'True
            Caption         =   "Fade Speed:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   3480
            Width           =   915
         End
         Begin VB.Label lblOsc 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1 pixel"
            Height          =   195
            Index           =   8
            Left            =   4920
            TabIndex        =   15
            Top             =   2880
            Width           =   450
         End
         Begin VB.Label lblOsc 
            AutoSize        =   -1  'True
            Caption         =   "Left Line Width:"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   14
            Top             =   2880
            Width           =   1125
         End
         Begin VB.Label lblOsc 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1 pixel"
            Height          =   195
            Index           =   6
            Left            =   4920
            TabIndex        =   13
            Top             =   2400
            Width           =   450
         End
         Begin VB.Label lblOsc 
            AutoSize        =   -1  'True
            Caption         =   "Right Line Width:"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   12
            Top             =   2400
            Width           =   1230
         End
         Begin VB.Label lblOscColor 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Left Channel Color"
            Height          =   375
            Index           =   1
            Left            =   2400
            TabIndex        =   11
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label lblOscColor 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Right Channel Color"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   1680
            Width           =   2055
         End
      End
      Begin VB.Frame frmOsc 
         BorderStyle     =   0  'None
         Caption         =   "Peaks"
         Height          =   3975
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   600
         Width           =   5535
         Begin VB.CheckBox chkOscPeakFall 
            Caption         =   "&Peaks falls faster and faster."
            Height          =   255
            Left            =   1560
            TabIndex        =   38
            ToolTipText     =   "if on, peaks fall faster and faster"
            Top             =   3000
            Width           =   2415
         End
         Begin VB.ComboBox cmbOscPeaks 
            Height          =   315
            ItemData        =   "frmPresetEdit.frx":0000
            Left            =   1920
            List            =   "frmPresetEdit.frx":000D
            Style           =   2  'Dropdown List
            TabIndex        =   29
            ToolTipText     =   "Peak style"
            Top             =   240
            Width           =   1935
         End
         Begin VB.OptionButton optPeakDetail 
            Caption         =   "Stereo"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   28
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton optPeakDetail 
            Caption         =   "Mono"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   27
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox cmbOscPeakCount 
            Height          =   315
            ItemData        =   "frmPresetEdit.frx":0030
            Left            =   120
            List            =   "frmPresetEdit.frx":003D
            Style           =   2  'Dropdown List
            TabIndex        =   26
            ToolTipText     =   "Number of peaks"
            Top             =   240
            Width           =   1455
         End
         Begin MSComctlLib.Slider sldOsc 
            Height          =   255
            Index           =   4
            Left            =   1440
            TabIndex        =   32
            ToolTipText     =   "Time to pause at the peaks new value"
            Top             =   2040
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   2
            Max             =   2000
            SelStart        =   100
            TickFrequency   =   100
            Value           =   100
         End
         Begin MSComctlLib.Slider sldOsc 
            Height          =   255
            Index           =   5
            Left            =   1440
            TabIndex        =   35
            ToolTipText     =   "Speed of drop"
            Top             =   2520
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   2
            Min             =   1
            Max             =   300
            SelStart        =   1
            TickFrequency   =   10
            Value           =   1
         End
         Begin VB.Label lblOsc 
            AutoSize        =   -1  'True
            Caption         =   "Drop speed:"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   37
            Top             =   2520
            Width           =   870
         End
         Begin VB.Label lblOsc 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1"
            Height          =   195
            Index           =   10
            Left            =   5280
            TabIndex        =   36
            Top             =   2520
            Width           =   90
         End
         Begin VB.Label lblOsc 
            AutoSize        =   -1  'True
            Caption         =   "Top Pause Time:"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   34
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblOsc 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "no pause"
            Height          =   195
            Index           =   4
            Left            =   4710
            TabIndex        =   33
            Top             =   2040
            Width           =   660
         End
         Begin VB.Label lblOscColor 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Right Channel Color"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   31
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Label lblOscColor 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Left Channel Color"
            Height          =   375
            Index           =   2
            Left            =   2400
            TabIndex        =   30
            Top             =   1200
            Width           =   2055
         End
      End
      Begin MSComctlLib.TabStrip tabsOsc 
         Height          =   4455
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   7858
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Main"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Peaks"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.ComboBox cmbSubtype 
      Height          =   315
      ItemData        =   "frmPresetEdit.frx":0061
      Left            =   4560
      List            =   "frmPresetEdit.frx":0063
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin MSComctlLib.Toolbar tbr 
      Height          =   330
      Left            =   120
      TabIndex        =   1
      Top             =   160
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   582
      ButtonWidth     =   1349
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Load"
            Key             =   "load"
            Object.ToolTipText     =   "Load"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmPresetEdit.frx":0065
      Left            =   2280
      List            =   "frmPresetEdit.frx":0075
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -240
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresetEdit.frx":00B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresetEdit.frx":040B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frm 
      Caption         =   "Volume Meter Settings"
      Height          =   4815
      Index           =   2
      Left            =   120
      TabIndex        =   81
      Top             =   600
      Width           =   6015
      Begin VB.OptionButton optVolDrawStyle 
         Caption         =   "&Solid draw style"
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   125
         ToolTipText     =   "Draw solid color bars"
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox chkVolFall 
         Caption         =   "&Bars falls faster and faster."
         Height          =   255
         Left            =   1680
         TabIndex        =   86
         Top             =   2040
         Width           =   2175
      End
      Begin VB.OptionButton optVolDrawStyle 
         Caption         =   "&Normal draw style"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   83
         ToolTipText     =   "Draw normal gradient"
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton optVolDrawStyle 
         Caption         =   "&Fire draw style"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   82
         ToolTipText     =   "Draw scaled gradient"
         Top             =   840
         Width           =   1335
      End
      Begin MSComctlLib.Slider sldVol 
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   87
         ToolTipText     =   "Top pause time"
         Top             =   1320
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         Max             =   2000
         SelStart        =   100
         TickFrequency   =   100
         Value           =   100
      End
      Begin MSComctlLib.Slider sldVol 
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   88
         ToolTipText     =   "Drop speed"
         Top             =   1680
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         Min             =   1
         Max             =   100
         SelStart        =   1
         TickFrequency   =   10
         Value           =   1
      End
      Begin MSComctlLib.Slider sldVol 
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   93
         ToolTipText     =   "Fade speed"
         Top             =   2640
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         Max             =   50
         TickFrequency   =   2
      End
      Begin VB.Label lblVol 
         AutoSize        =   -1  'True
         Caption         =   "Fade Speed:"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   95
         Top             =   2640
         Width           =   915
      End
      Begin VB.Label lblVol 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Disabled"
         Height          =   195
         Index           =   5
         Left            =   5040
         TabIndex        =   94
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblVol 
         AutoSize        =   -1  'True
         Caption         =   "Drop speed:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   92
         Top             =   1680
         Width           =   870
      End
      Begin VB.Label lblVol 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   3
         Left            =   5640
         TabIndex        =   91
         Top             =   1680
         Width           =   90
      End
      Begin VB.Label lblVol 
         AutoSize        =   -1  'True
         Caption         =   "Top Pause Time:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   90
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblVol 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "no pause"
         Height          =   195
         Index           =   1
         Left            =   5040
         TabIndex        =   89
         Top             =   1320
         Width           =   660
      End
      Begin VB.Label lblVolColor 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Upper Gradient Color"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   85
         ToolTipText     =   "Color of upper gradient color."
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblVolColor 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lower Gradient Color"
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   84
         ToolTipText     =   "Color of lower gradient color."
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Subtype:"
      Height          =   195
      Left            =   3840
      TabIndex        =   4
      Top             =   240
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Type:"
      Height          =   195
      Left            =   1800
      TabIndex        =   2
      Top             =   240
      Width           =   405
   End
End
Attribute VB_Name = "frmPresetEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkOscPeakFall_Click()
  ScopeSettings.bFall = chkOscPeakFall.Value
End Sub

Private Sub chkSpecFall_Click()
  SpectrumSettings.bFall = chkSpecFall.Value
End Sub

Private Sub chkSpecPeakFall_Click()
  SpectrumSettings.bPeakFall = chkSpecPeakFall.Value
End Sub

Private Sub chkSpecPeaks_Click()
  SpectrumSettings.bPeaks = chkSpecPeaks.Value
End Sub

Private Sub chkVolFall_Click()
  VolumeSettings.bFall = chkVolFall.Value
End Sub

Private Sub cmbOscPeakCount_Click()
  ScopeSettings.bPeakCount = cmbOscPeakCount.ListIndex
End Sub

Private Sub cmbOscPeaks_Click()
  ScopeSettings.bPeaks = cmbOscPeaks.ListIndex
End Sub

Private Sub cmbSubtype_Click()
  On Error Resume Next
  Select Case cmbType.ListIndex
    Case 0
      ScopeSettings.bType = cmbSubtype.ListIndex
    Case 1
      SpectrumSettings.bType = cmbSubtype.ListIndex
    Case 2
      VolumeSettings.bType = cmbSubtype.ListIndex
    Case 3
      BeatSettings.bType = cmbSubtype.ListIndex
  End Select
  frmMain.UpdateSpectrum
  EnableDisable
End Sub

Public Sub cmbType_Click()
  On Error Resume Next
  Dim X As Long
  For X = 0 To 3
    frm(X).Visible = (cmbType.ListIndex = X)
  Next
  Spectrum = cmbType.ListIndex + 1
  frmMain.UpdateSpectrum
  UpdateControls
  cmbSubtype.Clear
  Select Case cmbType.ListIndex
    Case 0
      cmbSubtype.AddItem "Dots"
      cmbSubtype.AddItem "Lines"
      cmbSubtype.AddItem "Solid"
      cmbSubtype.AddItem "History"
      cmbSubtype.ListIndex = ScopeSettings.bType
    Case 1
      cmbSubtype.AddItem "Normal"
      cmbSubtype.AddItem "Bars"
      cmbSubtype.AddItem "Lines"
      cmbSubtype.ListIndex = SpectrumSettings.bType
    Case 2
      cmbSubtype.AddItem "Bars"
      cmbSubtype.AddItem "History"
      cmbSubtype.ListIndex = VolumeSettings.bType
    Case 3
      cmbSubtype.AddItem "Zoomer"
      cmbSubtype.AddItem "Rotator"
      cmbSubtype.ListIndex = BeatSettings.bType
  End Select
End Sub

Private Sub cmdBeat_Click(index As Integer)
  On Error GoTo errh
  Select Case index
    Case 0
      Dim cDlg As New clsCommonDialog
      With cDlg
        Set .Parent = Me
        .CancelError = True
        .Filter = "Bitmaps (*.bmp)|*.bmp"
        .DialogTitle = "Set Image"
        .InitDir = App.Path & "\vis\data"
        .Flags = cdlOFNExplorer Or cdlOFNNoChangeDir Or cdlOFNFileMustExist
        .ShowOpen
        If LCase(sFilename(.FileName, efpFilePath)) = LCase(App.Path) & "\vis\data\" Then
          BeatSettings.sFile = sFilename(.FileName, efpFileNameAndExt)
        Else
          MsgBox "Image must be located in """ & App.Path & "\vis\data\"".", vbExclamation
        End If
        frmMain.UpdateSpectrum
        
      End With
    Case 1
      BeatSettings.sFile = ""
      frmMain.UpdateSpectrum
  End Select
errh:
End Sub

Private Sub Form_Activate()
  AlwaysOnTop Me, Settings.OnTop
  If cmbType.ListIndex <> Spectrum - 1 Then cmbType.ListIndex = Spectrum - 1
End Sub

Public Sub UpdateControls()
  On Error Resume Next
  'oscilliscope
  optOscDet(ScopeSettings.bDetail).Value = True
  lblOscColor(0).BackColor = ScopeSettings.lColorR
  lblOscColor(0).ForeColor = InvertColor(ScopeSettings.lColorR)
  lblOscColor(1).BackColor = ScopeSettings.lColorL
  lblOscColor(1).ForeColor = InvertColor(ScopeSettings.lColorL)
  sldOsc(0).Value = ScopeSettings.bSkip / 2
  sldOsc(1).Value = ScopeSettings.bBrushSizeR
  sldOsc(2).Value = ScopeSettings.bBrushSizeL
  sldOsc(3).Value = ScopeSettings.bFade
  sldOsc(4).Value = ScopeSettings.lPeakPause
  sldOsc(5).Value = ScopeSettings.lPeakDec
  lblOsc(0) = sldOsc(0).Value * 2 & " values"
  lblOsc(6) = sldOsc(1).Value & " pixels"
  lblOsc(8) = sldOsc(2).Value & " pixels"
  lblOsc(10) = sldOsc(5).Value
  If sldOsc(3).Value = 0 Then lblOsc(3) = "Disabled" Else lblOsc(3) = sldOsc(3).Value & " val/frm"
  If sldOsc(4).Value = 0 Then lblOsc(4) = "Disabled" Else lblOsc(4) = sldOsc(4).Value & " ms"
  cmbOscPeakCount.ListIndex = ScopeSettings.bPeakCount
  cmbOscPeaks.ListIndex = ScopeSettings.bPeaks
  optPeakDetail(ScopeSettings.bPeakDetail).Value = True
  lblOscColor(2).BackColor = ScopeSettings.lColorPeakR
  lblOscColor(2).ForeColor = InvertColor(ScopeSettings.lColorPeakR)
  lblOscColor(3).BackColor = ScopeSettings.lColorPeakL
  lblOscColor(3).ForeColor = InvertColor(ScopeSettings.lColorPeakL)
  chkOscPeakFall.Value = ScopeSettings.bFall
  'spectrum
  sldSpec(0).Value = SpectrumSettings.nZoom * 10
  sldSpec(1).Value = SpectrumSettings.iView
  lblSpec(1) = FormatNumber(sldSpec(0).Value / 10, 1)
  lblSpec(3) = sldSpec(1).Value
  lblSpecColor(0).BackColor = SpectrumSettings.lColorUp
  lblSpecColor(0).ForeColor = InvertColor(SpectrumSettings.lColorUp)
  lblSpecColor(1).BackColor = SpectrumSettings.lColorDn
  lblSpecColor(1).ForeColor = InvertColor(SpectrumSettings.lColorDn)
  lblSpecColor(2).BackColor = SpectrumSettings.lColorLine
  lblSpecColor(2).ForeColor = InvertColor(SpectrumSettings.lColorLine)
  lblSpecColor(3).BackColor = SpectrumSettings.lPeakColor
  lblSpecColor(3).ForeColor = InvertColor(SpectrumSettings.lPeakColor)
  optSpecDrawStyle(SpectrumSettings.bDrawStyle).Value = True
  sldSpec(2).Value = SpectrumSettings.bBarSize
  lblSpec(4) = sldSpec(2).Value & " pixels"
  sldSpec(3).Value = SpectrumSettings.bBrushSize
  lblSpec(6) = sldSpec(3).Value & " pixels"
  sldSpec(4).Value = SpectrumSettings.lPause
  If sldSpec(4).Value = 0 Then lblSpec(15) = "Disabled" Else lblSpec(15) = sldSpec(4).Value & " ms"
  sldSpec(5).Value = SpectrumSettings.nDec
  lblSpec(13) = sldSpec(5).Value
  sldSpec(6).Value = SpectrumSettings.lPeakPause
  If sldSpec(6).Value = 0 Then lblSpec(17) = "Disabled" Else lblSpec(17) = sldSpec(6).Value & " ms"
  sldSpec(7).Value = SpectrumSettings.lPeakDec
  lblSpec(19) = sldSpec(7).Value
  chkSpecFall.Value = SpectrumSettings.bFall
  chkSpecPeakFall.Value = SpectrumSettings.bPeakFall
  chkSpecPeaks.Value = SpectrumSettings.bPeaks
  sldSpec(8).Value = SpectrumSettings.bFade
  If sldSpec(8).Value = 0 Then lblSpec(21) = "Disabled" Else lblSpec(21) = sldSpec(8).Value & " val/frm"
  sldSpec(9).Value = SpectrumSettings.bCorrection
  If sldSpec(9).Value = 0 Then lblSpec(9) = "None" Else lblSpec(9) = sldSpec(9).Value
  'volume
  sldVol(0).Value = VolumeSettings.lPause
  If sldVol(0).Value = 0 Then lblVol(1) = "Disabled" Else lblVol(1) = sldVol(0).Value & " ms"
  sldVol(1).Value = VolumeSettings.nDec
  lblVol(3) = sldVol(1).Value
  sldVol(2).Value = VolumeSettings.bFade
  If sldVol(2).Value = 0 Then lblVol(5) = "Disabled" Else lblVol(5) = sldVol(2).Value & " val/frm"
  chkVolFall.Value = VolumeSettings.bFall
  optVolDrawStyle(VolumeSettings.bDrawStyle).Value = True
  lblVolColor(0).BackColor = VolumeSettings.lColorUp
  lblVolColor(0).ForeColor = InvertColor(VolumeSettings.lColorUp)
  lblVolColor(1).BackColor = VolumeSettings.lColorDn
  lblVolColor(1).ForeColor = InvertColor(VolumeSettings.lColorDn)
  'beat
  sldBeat(6).Value = BeatSettings.iDetectLow
  sldBeat(0).Value = BeatSettings.iDetectHigh
  sldBeat(1).Value = BeatSettings.nMulti * 10
  lblBeat(4) = FormatNumber(sldBeat(1).Value / 10, 1)
  sldBeat(2).Value = BeatSettings.nMin * 100
  lblBeat(6) = FormatNumber(sldBeat(2).Value / 100, 2)
  sldBeat(3).Value = BeatSettings.nRotMin * 100
  lblBeat(8) = FormatNumber(sldBeat(3).Value / 100, 2)
  sldBeat(4).Value = BeatSettings.nRotMove * 100
  lblBeat(10) = FormatNumber(sldBeat(4).Value / 100, 2)
  sldBeat(5).Value = BeatSettings.nRotSpeed * 100
  lblBeat(11) = FormatNumber(sldBeat(5).Value / 100, 2)
  sldBeat(7).Value = BeatSettings.bFade
  If sldBeat(7).Value = 0 Then lblBeat(14) = "Disabled" Else lblBeat(14) = sldBeat(7).Value & " val/frm"
  lblBeat(2) = "Detect values: " & sldBeat(6).Value & " - " & sldBeat(0).Value
End Sub

Private Sub lblOscColor_Click(index As Integer)
  On Error Resume Next
  Load frmColor
  Select Case index
    Case 0
      frmColor.Tag = ScopeSettings.lColorR
      frmColor.Show vbModal, Me
      ScopeSettings.lColorR = Val(frmColor.Tag)
      DeleteObject hPenRight
      hPenRight = CreatePen(0, ScopeSettings.bBrushSizeR, ScopeSettings.lColorR)
      hBrushSolidRight = CreateSolidBrush(ScopeSettings.lColorR)
    Case 1
      frmColor.Tag = ScopeSettings.lColorL
      frmColor.Show vbModal, Me
      ScopeSettings.lColorL = Val(frmColor.Tag)
      DeleteObject hPenLeft
      hPenLeft = CreatePen(0, ScopeSettings.bBrushSizeL, ScopeSettings.lColorL)
      hBrushSolidLeft = CreateSolidBrush(ScopeSettings.lColorL)
    Case 2
      frmColor.Tag = ScopeSettings.lColorPeakR
      frmColor.Show vbModal, Me
      ScopeSettings.lColorPeakR = Val(frmColor.Tag)
      DeleteObject hPenPeakRight
      hPenPeakRight = CreatePen(0, 1, ScopeSettings.lColorPeakR)
    Case 3
      frmColor.Tag = ScopeSettings.lColorPeakL
      frmColor.Show vbModal, Me
      ScopeSettings.lColorPeakL = Val(frmColor.Tag)
      DeleteObject hPenPeakLeft
      hPenPeakLeft = CreatePen(0, 1, ScopeSettings.lColorPeakL)
  End Select
  lblOscColor(index).BackColor = Val(frmColor.Tag)
  lblOscColor(index).ForeColor = InvertColor(Val(frmColor.Tag))
  frmColor.Tag = "Y"
  Unload frmColor
End Sub

Private Sub lblSpecColor_Click(index As Integer)
  On Error Resume Next
  Load frmColor
  Select Case index
    Case 0
      frmColor.Tag = SpectrumSettings.lColorUp
      frmColor.Show vbModal, Me
      SpectrumSettings.lColorUp = Val(frmColor.Tag)
      frmMain.DoGrad SpectrumSettings.lColorUp, SpectrumSettings.lColorDn
      cGradBar.Create 1, frmMain.picVis.ScaleHeight, frmMain.hdc
      cGradBar.BitBltFrom frmMain.p.hdc
    Case 1
      frmColor.Tag = SpectrumSettings.lColorDn
      frmColor.Show vbModal, Me
      SpectrumSettings.lColorDn = Val(frmColor.Tag)
      frmMain.DoGrad SpectrumSettings.lColorUp, SpectrumSettings.lColorDn
      cGradBar.Create 1, frmMain.picVis.ScaleHeight, frmMain.hdc
      cGradBar.BitBltFrom frmMain.p.hdc
    Case 2
      frmColor.Tag = SpectrumSettings.lColorLine
      frmColor.Show vbModal, Me
      SpectrumSettings.lColorLine = Val(frmColor.Tag)
      DeleteObject hPenLineSpec
      hPenLineSpec = CreatePen(0, SpectrumSettings.bBrushSize, SpectrumSettings.lColorLine)
    Case 3
      frmColor.Tag = SpectrumSettings.lPeakColor
      frmColor.Show vbModal, Me
      SpectrumSettings.lPeakColor = Val(frmColor.Tag)
      DeleteObject hPenPeakSpec
      hPenPeakSpec = CreatePen(0, 1, SpectrumSettings.lPeakColor)
  End Select
  lblSpecColor(index).BackColor = Val(frmColor.Tag)
  lblSpecColor(index).ForeColor = InvertColor(Val(frmColor.Tag))
  frmColor.Tag = "Y"
  Unload frmColor
End Sub

Private Sub lblVolColor_Click(index As Integer)
  On Error Resume Next
  Load frmColor
  Select Case index
    Case 0
      frmColor.Tag = VolumeSettings.lColorUp
      frmColor.Show vbModal, Me
      VolumeSettings.lColorUp = Val(frmColor.Tag)
    Case 1
      frmColor.Tag = VolumeSettings.lColorDn
      frmColor.Show vbModal, Me
      VolumeSettings.lColorDn = Val(frmColor.Tag)
  End Select
  frmMain.DoGrad VolumeSettings.lColorUp, VolumeSettings.lColorDn
  cGradVol.Create 2, frmMain.picVis.ScaleHeight, frmMain.hdc
  cGradVol.BitBltFrom frmMain.p.hdc
  lblVolColor(index).BackColor = Val(frmColor.Tag)
  lblVolColor(index).ForeColor = InvertColor(Val(frmColor.Tag))
  frmColor.Tag = "Y"
  Unload frmColor
End Sub

Private Sub optOscDet_Click(index As Integer)
  ScopeSettings.bDetail = index
End Sub

Private Sub optPeakDetail_Click(index As Integer)
  ScopeSettings.bPeakDetail = index
End Sub

Private Sub optSpecDrawStyle_Click(index As Integer)
  SpectrumSettings.bDrawStyle = index
End Sub

Private Sub optVolDrawStyle_Click(index As Integer)
  VolumeSettings.bDrawStyle = index
End Sub

Private Sub sldBeat_Scroll(index As Integer)
  On Error Resume Next
  If sldBeat(0).Value <= sldBeat(6).Value Then sldBeat(0).Value = sldBeat(6).Value + 1
  Select Case index
    Case 6
      BeatSettings.iDetectLow = sldBeat(6).Value
    Case 0
      BeatSettings.iDetectHigh = sldBeat(0).Value
    Case 1
      BeatSettings.nMulti = sldBeat(1).Value / 10
      lblBeat(4) = FormatNumber(sldBeat(1).Value / 10, 1)
    Case 2
      BeatSettings.nMin = sldBeat(2).Value / 100
      lblBeat(6) = FormatNumber(sldBeat(2).Value / 100, 2)
    Case 3
      BeatSettings.nRotMin = sldBeat(3).Value / 100
      lblBeat(8) = FormatNumber(sldBeat(3).Value / 100, 2)
    Case 4
      BeatSettings.nRotMove = sldBeat(4).Value / 100
      lblBeat(10) = FormatNumber(sldBeat(4).Value / 100, 2)
    Case 5
      BeatSettings.nRotSpeed = sldBeat(5).Value / 100
      lblBeat(11) = FormatNumber(sldBeat(5).Value / 100, 2)
    Case 7
      BeatSettings.bFade = sldBeat(7).Value
      If sldBeat(7).Value = 0 Then lblBeat(14) = "Disabled" Else lblBeat(14) = sldBeat(7).Value & " val/frm"
  End Select
  lblBeat(2) = "Detect values: " & sldBeat(6).Value & " - " & sldBeat(0).Value
End Sub

Private Sub sldOsc_Scroll(index As Integer)
  On Error Resume Next
  Select Case index
    Case 0
      ScopeSettings.bSkip = sldOsc(0).Value * 2
      lblOsc(0) = sldOsc(0).Value * 2 & " values"
    Case 1
      ScopeSettings.bBrushSizeR = sldOsc(1).Value
      lblOsc(6) = sldOsc(1).Value & " pixels"
      DeleteObject hPenRight
      hPenRight = CreatePen(0, ScopeSettings.bBrushSizeR, ScopeSettings.lColorR)
    Case 2
      ScopeSettings.bBrushSizeL = sldOsc(2).Value
      lblOsc(8) = sldOsc(2).Value & " pixels"
      DeleteObject hPenLeft
      hPenLeft = CreatePen(0, ScopeSettings.bBrushSizeL, ScopeSettings.lColorL)
    Case 3
      ScopeSettings.bFade = sldOsc(3).Value
      If sldOsc(3).Value = 0 Then lblOsc(3) = "Disabled" Else lblOsc(3) = sldOsc(3).Value & " val/frm"
    Case 4
      ScopeSettings.lPeakPause = sldOsc(4).Value
      If sldOsc(4).Value = 0 Then lblOsc(4) = "Disabled" Else lblOsc(4) = sldOsc(4).Value & " ms"
    Case 5
      ScopeSettings.lPeakDec = sldOsc(5).Value
      lblOsc(10) = sldOsc(5).Value
  End Select
End Sub

Private Sub sldSpec_Scroll(index As Integer)
  On Error Resume Next
  Select Case index
    Case 0
      SpectrumSettings.nZoom = sldSpec(0).Value / 10
      lblSpec(1) = FormatNumber(sldSpec(0).Value / 10, 1)
    Case 1
      SpectrumSettings.iView = sldSpec(1).Value
      lblSpec(3) = sldSpec(1).Value
    Case 2
      SpectrumSettings.bBarSize = sldSpec(2).Value
      lblSpec(4) = sldSpec(2).Value & " pixels"
    Case 3
      SpectrumSettings.bBrushSize = sldSpec(3).Value
      lblSpec(6) = sldSpec(3).Value & " pixels"
      DeleteObject hPenLineSpec
      hPenLineSpec = CreatePen(0, SpectrumSettings.bBrushSize, SpectrumSettings.lColorLine)
    Case 4
      SpectrumSettings.lPause = sldSpec(4).Value
      If sldSpec(4).Value = 0 Then lblSpec(15) = "Disabled" Else lblSpec(15) = sldSpec(4).Value & " ms"
    Case 5
      SpectrumSettings.nDec = sldSpec(5).Value
      lblSpec(13) = sldSpec(5).Value
    Case 6
      SpectrumSettings.lPeakPause = sldSpec(6).Value
      If sldSpec(6).Value = 0 Then lblSpec(17) = "Disabled" Else lblSpec(17) = sldSpec(6).Value & " ms"
    Case 7
      SpectrumSettings.lPeakDec = sldSpec(7).Value
      lblSpec(19) = sldSpec(7).Value
    Case 8
      SpectrumSettings.bFade = sldSpec(8).Value
      If sldSpec(8).Value = 0 Then lblSpec(21) = "Disabled" Else lblSpec(21) = sldSpec(8).Value & " val/frm"
    Case 9
      SpectrumSettings.bCorrection = sldSpec(9).Value
      If sldSpec(9).Value = 0 Then lblSpec(9) = "None" Else lblSpec(9) = sldSpec(9).Value
  End Select
End Sub

Private Sub sldVol_Scroll(index As Integer)
  On Error Resume Next
  Select Case index
    Case 0
      VolumeSettings.lPause = sldVol(0).Value
      If sldVol(0).Value = 0 Then lblVol(1) = "Disabled" Else lblVol(1) = sldVol(0).Value & " ms"
    Case 1
      VolumeSettings.nDec = sldVol(1).Value
      lblVol(3) = sldVol(1).Value
    Case 2
      VolumeSettings.bFade = sldVol(2).Value
      If sldVol(2).Value = 0 Then lblVol(5) = "Disabled" Else lblVol(5) = sldVol(2).Value & " val/frm"
  End Select
End Sub

Private Sub tabsOsc_Click()
  On Error Resume Next
  Dim X As Long
  For X = 0 To 1
    If X <> tabsOsc.SelectedItem.index - 1 Then
      frmOsc(X).Visible = False
    Else
      frmOsc(X).Visible = True
    End If
  Next
End Sub

Private Sub tabSpec_Click()
  On Error Resume Next
  Dim X As Long
  For X = 0 To 1
    If X <> tabSpec.SelectedItem.index - 1 Then
      frmSpec(X).Visible = False
    Else
      frmSpec(X).Visible = True
    End If
  Next
End Sub

Public Sub EnableDisable()
  On Error Resume Next
  Select Case cmbType.ListIndex
    Case 0
      sldOsc(1).Enabled = (cmbSubtype.ListIndex <> 2)
      sldOsc(2).Enabled = (cmbSubtype.ListIndex <> 2)
      lblOsc(5).Enabled = (cmbSubtype.ListIndex <> 2)
      lblOsc(6).Enabled = (cmbSubtype.ListIndex <> 2)
      lblOsc(7).Enabled = (cmbSubtype.ListIndex <> 2)
      lblOsc(8).Enabled = (cmbSubtype.ListIndex <> 2)
      cmbOscPeakCount.Enabled = (cmbSubtype.ListIndex <> 3)
      cmbOscPeaks.Enabled = (cmbSubtype.ListIndex <> 3)
      optPeakDetail(0).Enabled = (cmbSubtype.ListIndex <> 3)
      optPeakDetail(1).Enabled = (cmbSubtype.ListIndex <> 3)
      lblOscColor(3).Enabled = (cmbSubtype.ListIndex <> 3)
      lblOscColor(2).Enabled = (cmbSubtype.ListIndex <> 3)
      lblOsc(9).Enabled = (cmbSubtype.ListIndex <> 3)
      lblOsc(10).Enabled = (cmbSubtype.ListIndex <> 3)
      lblOsc(11).Enabled = (cmbSubtype.ListIndex <> 3)
      lblOsc(4).Enabled = (cmbSubtype.ListIndex <> 3)
      sldOsc(4).Enabled = (cmbSubtype.ListIndex <> 3)
      sldOsc(5).Enabled = (cmbSubtype.ListIndex <> 3)
      chkOscPeakFall.Enabled = (cmbSubtype.ListIndex <> 3)
    Case 1
      lblSpec(5).Enabled = (cmbSubtype.ListIndex = 1)
      lblSpec(4).Enabled = (cmbSubtype.ListIndex = 1)
      sldSpec(2).Enabled = (cmbSubtype.ListIndex = 1)
      lblSpec(6).Enabled = (cmbSubtype.ListIndex = 2)
      lblSpec(7).Enabled = (cmbSubtype.ListIndex = 2)
      sldSpec(3).Enabled = (cmbSubtype.ListIndex = 2)
      lblSpecColor(0).Enabled = (cmbSubtype.ListIndex <> 2)
      lblSpecColor(1).Enabled = (cmbSubtype.ListIndex <> 2)
      lblSpecColor(2).Enabled = (cmbSubtype.ListIndex = 2)
      optSpecDrawStyle(0).Enabled = (cmbSubtype.ListIndex <> 2)
      optSpecDrawStyle(1).Enabled = (cmbSubtype.ListIndex <> 2)
      lblSpec(14).Enabled = (cmbSubtype.ListIndex = 1)
      lblSpec(15).Enabled = (cmbSubtype.ListIndex = 1)
      sldSpec(4).Enabled = (cmbSubtype.ListIndex = 1)
      lblSpec(12).Enabled = (cmbSubtype.ListIndex = 1)
      lblSpec(13).Enabled = (cmbSubtype.ListIndex = 1)
      sldSpec(5).Enabled = (cmbSubtype.ListIndex = 1)
      chkSpecFall.Enabled = (cmbSubtype.ListIndex = 1)
      chkSpecPeaks.Enabled = (cmbSubtype.ListIndex = 1)
      lblSpecColor(3).Enabled = (cmbSubtype.ListIndex = 1)
      lblSpec(16).Enabled = (cmbSubtype.ListIndex = 1)
      lblSpec(17).Enabled = (cmbSubtype.ListIndex = 1)
      sldSpec(6).Enabled = (cmbSubtype.ListIndex = 1)
      lblSpec(18).Enabled = (cmbSubtype.ListIndex = 1)
      lblSpec(19).Enabled = (cmbSubtype.ListIndex = 1)
      sldSpec(7).Enabled = (cmbSubtype.ListIndex = 1)
      chkSpecPeakFall.Enabled = (cmbSubtype.ListIndex = 1)
    Case 2
      lblVolColor(0).Enabled = (cmbSubtype.ListIndex = 1)
      lblVolColor(1).Enabled = (cmbSubtype.ListIndex = 1)
      optVolDrawStyle(0).Enabled = (cmbSubtype.ListIndex = 1)
      optVolDrawStyle(1).Enabled = (cmbSubtype.ListIndex = 1)
      lblVol(0).Enabled = (cmbSubtype.ListIndex = 0)
      lblVol(1).Enabled = (cmbSubtype.ListIndex = 0)
      lblVol(2).Enabled = (cmbSubtype.ListIndex = 0)
      lblVol(3).Enabled = (cmbSubtype.ListIndex = 0)
      sldVol(0).Enabled = (cmbSubtype.ListIndex = 0)
      sldVol(1).Enabled = (cmbSubtype.ListIndex = 0)
      chkVolFall.Enabled = (cmbSubtype.ListIndex = 0)
    Case 3
  End Select
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
  On Error GoTo errh
  Dim cDlg As New clsCommonDialog
  
  Select Case Button.Key
    Case "save"
      With cDlg
        Set .Parent = Me
        .CancelError = True
        .Filter = "Visualization Presets (*.sap)|*.sap"
        .DialogTitle = "Save Preset"
        .InitDir = App.Path & "\vis"
        .Flags = cdlOFNOverwritePrompt
        .ShowSave
        SavePreset .FileName, Spectrum
        
      End With
    Case "load"
      frmPresetLoad.Show , frmMain
      
  End Select
errh:
End Sub
