VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStudio 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sound Studio"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      Height          =   3495
      Index           =   8
      Left            =   240
      TabIndex        =   162
      Top             =   480
      Width           =   6975
      Begin VB.CheckBox chkEffectOn 
         Caption         =   "En&able Waves Reverb"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   212
         Top             =   120
         Width           =   1935
      End
      Begin MSComctlLib.Slider sldWave 
         Height          =   495
         Index           =   0
         Left            =   600
         TabIndex        =   213
         ToolTipText     =   "Input gain of signal, in decibels (dB), in the range from -96 through 0. The default value is 0 dB."
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   100
         SmallChange     =   10
         Min             =   -96
         Max             =   0
         TickFrequency   =   10
      End
      Begin MSComctlLib.Slider sldWave 
         Height          =   495
         Index           =   1
         Left            =   600
         TabIndex        =   217
         ToolTipText     =   "Reverb mix, in dB, in the range from -96 through 0. The default value is 0 dB."
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   100
         SmallChange     =   10
         Min             =   -96
         Max             =   0
         TickFrequency   =   10
      End
      Begin MSComctlLib.Slider sldWave 
         Height          =   495
         Index           =   2
         Left            =   3960
         TabIndex        =   221
         ToolTipText     =   "Reverb time, in milliseconds, in the range from 1 through 3000. The default value is 1000. "
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   100
         SmallChange     =   10
         Min             =   1
         Max             =   3000
         SelStart        =   1000
         TickFrequency   =   100
         Value           =   1000
      End
      Begin MSComctlLib.Slider sldWave 
         Height          =   495
         Index           =   3
         Left            =   3960
         TabIndex        =   225
         ToolTipText     =   "High Frequency Reverb Time Ratio in the range from .001 through .999. The default value is 0.001."
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   100
         SmallChange     =   10
         Min             =   1
         Max             =   999
         SelStart        =   1
         TickFrequency   =   100
         Value           =   1
      End
      Begin VB.Label Label134 
         AutoSize        =   -1  'True
         Caption         =   "High Frequency Reverb Time Ratio:"
         Height          =   195
         Left            =   3600
         TabIndex        =   228
         Top             =   1440
         Width           =   2550
      End
      Begin VB.Label Label133 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.001"
         Height          =   195
         Left            =   3540
         TabIndex        =   227
         Top             =   1800
         Width           =   405
      End
      Begin VB.Label Label132 
         AutoSize        =   -1  'True
         Caption         =   "0.999"
         Height          =   195
         Left            =   6135
         TabIndex        =   226
         Top             =   1800
         Width           =   405
      End
      Begin VB.Label Label131 
         AutoSize        =   -1  'True
         Caption         =   "Reverb Time:"
         Height          =   195
         Left            =   3600
         TabIndex        =   224
         Top             =   600
         Width           =   960
      End
      Begin VB.Label Label130 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1 ms"
         Height          =   195
         Left            =   3615
         TabIndex        =   223
         Top             =   960
         Width           =   330
      End
      Begin VB.Label Label129 
         AutoSize        =   -1  'True
         Caption         =   "3000 ms"
         Height          =   195
         Left            =   6135
         TabIndex        =   222
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label128 
         AutoSize        =   -1  'True
         Caption         =   "Reverb Mix:"
         Height          =   195
         Left            =   240
         TabIndex        =   220
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label127 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "-96 dB"
         Height          =   195
         Left            =   120
         TabIndex        =   219
         Top             =   1800
         Width           =   465
      End
      Begin VB.Label Label126 
         AutoSize        =   -1  'True
         Caption         =   "0 dB"
         Height          =   195
         Left            =   2775
         TabIndex        =   218
         Top             =   1800
         Width           =   330
      End
      Begin VB.Label Label125 
         AutoSize        =   -1  'True
         Caption         =   "Input Gain:"
         Height          =   195
         Left            =   240
         TabIndex        =   216
         Top             =   600
         Width           =   780
      End
      Begin VB.Label Label124 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "-96 dB"
         Height          =   195
         Left            =   120
         TabIndex        =   215
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label123 
         AutoSize        =   -1  'True
         Caption         =   "0 dB"
         Height          =   195
         Left            =   2775
         TabIndex        =   214
         Top             =   960
         Width           =   330
      End
   End
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      Height          =   3495
      Index           =   3
      Left            =   240
      TabIndex        =   79
      Top             =   480
      Width           =   6975
      Begin VB.CheckBox chkEffectOn 
         Caption         =   "Enable D&istortion"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   80
         Top             =   120
         Width           =   1815
      End
      Begin MSComctlLib.Slider sldDistortion 
         Height          =   495
         Index           =   2
         Left            =   600
         TabIndex        =   84
         ToolTipText     =   "Center frequency of harmonic content addition, in the range from 100 through 8000. The default value is 4000 Hz."
         Top             =   2520
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   250
         SmallChange     =   50
         Min             =   100
         Max             =   8000
         SelStart        =   4000
         TickFrequency   =   500
         Value           =   4000
      End
      Begin MSComctlLib.Slider sldDistortion 
         Height          =   495
         Index           =   3
         Left            =   3960
         TabIndex        =   86
         ToolTipText     =   $"frmStudio.frx":0000
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   250
         SmallChange     =   50
         Min             =   100
         Max             =   8000
         SelStart        =   4000
         TickFrequency   =   500
         Value           =   4000
      End
      Begin MSComctlLib.Slider sldDistortion 
         Height          =   495
         Index           =   1
         Left            =   600
         TabIndex        =   82
         ToolTipText     =   "Percentage of distortion intensity, in the range in the range from 0 through 100. The default value is 50 percent."
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         SelStart        =   50
         TickFrequency   =   10
         Value           =   50
      End
      Begin MSComctlLib.Slider sldDistortion 
         Height          =   495
         Index           =   0
         Left            =   600
         TabIndex        =   81
         ToolTipText     =   "Amount of signal change after distortion, in the range from -60 through 0. The default value is 0 dB."
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Min             =   -60
         Max             =   0
         SelStart        =   -30
         TickFrequency   =   10
         Value           =   -30
      End
      Begin MSComctlLib.Slider sldDistortion 
         Height          =   495
         Index           =   4
         Left            =   3960
         TabIndex        =   88
         ToolTipText     =   "Filter cutoff for high-frequency harmonics attenuation, in the range from 100 through 8000. The default value is 4000 Hz."
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   250
         SmallChange     =   50
         Min             =   100
         Max             =   8000
         SelStart        =   4000
         TickFrequency   =   500
         Value           =   4000
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Post EQ Center Frequency:"
         Height          =   195
         Left            =   240
         TabIndex        =   100
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "100 hz"
         Height          =   195
         Left            =   120
         TabIndex        =   99
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "8000 hz"
         Height          =   195
         Left            =   2775
         TabIndex        =   98
         Top             =   2640
         Width           =   570
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "8000 hz"
         Height          =   195
         Left            =   6135
         TabIndex        =   97
         Top             =   1800
         Width           =   570
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "100 hz"
         Height          =   195
         Left            =   3480
         TabIndex        =   96
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "Pre Lowpass Cutoff:"
         Height          =   195
         Left            =   3600
         TabIndex        =   95
         Top             =   1440
         Width           =   1425
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "8000 hz"
         Height          =   195
         Left            =   6135
         TabIndex        =   94
         Top             =   960
         Width           =   570
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "100 hz"
         Height          =   195
         Left            =   3480
         TabIndex        =   93
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "Post EQ Bandwidth:"
         Height          =   195
         Left            =   3600
         TabIndex        =   92
         Top             =   600
         Width           =   1425
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "100%"
         Height          =   195
         Left            =   2775
         TabIndex        =   91
         Top             =   1800
         Width           =   390
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "0%"
         Height          =   195
         Left            =   360
         TabIndex        =   90
         Top             =   1800
         Width           =   210
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         Caption         =   "Edge:"
         Height          =   195
         Left            =   240
         TabIndex        =   89
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         Caption         =   "0 dB"
         Height          =   195
         Left            =   2775
         TabIndex        =   87
         Top             =   960
         Width           =   330
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "-60 dB"
         Height          =   195
         Left            =   120
         TabIndex        =   85
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         Caption         =   "Gain:"
         Height          =   195
         Left            =   240
         TabIndex        =   83
         Top             =   600
         Width           =   375
      End
   End
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      Height          =   3495
      Index           =   4
      Left            =   240
      TabIndex        =   101
      Top             =   480
      Width           =   6975
      Begin VB.CheckBox chkEffectOn 
         Caption         =   "Enable Ec&ho"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   102
         Top             =   120
         Width           =   1335
      End
      Begin VB.CheckBox chkEcho 
         Caption         =   "&Pan Swap"
         Height          =   255
         Left            =   720
         TabIndex        =   110
         ToolTipText     =   "Value that specifies whether to swap left and right delays with each successive echo."
         Top             =   2520
         Width           =   1095
      End
      Begin MSComctlLib.Slider sldEcho 
         Height          =   495
         Index           =   2
         Left            =   3960
         TabIndex        =   106
         ToolTipText     =   "Delay for left channel, in milliseconds, in the range from 1 through 2000. The default value is 333 ms."
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   100
         SmallChange     =   10
         Min             =   1
         Max             =   2000
         SelStart        =   333
         TickFrequency   =   250
         Value           =   333
      End
      Begin MSComctlLib.Slider sldEcho 
         Height          =   495
         Index           =   1
         Left            =   600
         TabIndex        =   104
         ToolTipText     =   "Percentage of output fed back into input, in the range from 0 through 100. The default value is 0."
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         TickFrequency   =   10
      End
      Begin MSComctlLib.Slider sldEcho 
         Height          =   495
         Index           =   0
         Left            =   600
         TabIndex        =   103
         ToolTipText     =   "Ratio of wet (processed) signal to dry (unprocessed) signal. Must be in the range from 0 through 100 (all wet)."
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         SelStart        =   50
         TickFrequency   =   10
         Value           =   50
      End
      Begin MSComctlLib.Slider sldEcho 
         Height          =   495
         Index           =   3
         Left            =   3960
         TabIndex        =   108
         ToolTipText     =   "Delay for right channel, in milliseconds, in the range from 1 through 2000. The default value is 333 ms."
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   100
         SmallChange     =   10
         Min             =   1
         Max             =   2000
         SelStart        =   333
         TickFrequency   =   250
         Value           =   333
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Feedback:"
         Height          =   195
         Left            =   240
         TabIndex        =   119
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         Caption         =   "0%"
         Height          =   195
         Left            =   360
         TabIndex        =   118
         Top             =   1800
         Width           =   210
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "100%"
         Height          =   195
         Left            =   2775
         TabIndex        =   117
         Top             =   1800
         Width           =   390
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "Left Delay:"
         Height          =   195
         Left            =   3600
         TabIndex        =   116
         Top             =   600
         Width           =   765
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         Caption         =   "1 ms"
         Height          =   195
         Left            =   3600
         TabIndex        =   115
         Top             =   960
         Width           =   330
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         Caption         =   "2000 ms"
         Height          =   195
         Left            =   6120
         TabIndex        =   114
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "Wet Dry Mix:"
         Height          =   195
         Left            =   240
         TabIndex        =   113
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "Dry"
         Height          =   195
         Left            =   360
         TabIndex        =   112
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Wet"
         Height          =   195
         Left            =   2775
         TabIndex        =   111
         Top             =   960
         Width           =   300
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "Right Delay:"
         Height          =   195
         Left            =   3600
         TabIndex        =   109
         Top             =   1440
         Width           =   870
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         Caption         =   "1 ms"
         Height          =   195
         Left            =   3600
         TabIndex        =   107
         Top             =   1800
         Width           =   330
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         Caption         =   "2000 ms"
         Height          =   195
         Left            =   6135
         TabIndex        =   105
         Top             =   1800
         Width           =   600
      End
   End
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      Height          =   3495
      Index           =   7
      Left            =   240
      TabIndex        =   161
      Top             =   480
      Width           =   6975
      Begin VB.CheckBox chkEffectOn 
         Caption         =   "Enable I3DL&2 Reverb"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   163
         Top             =   120
         Width           =   1935
      End
      Begin MSComctlLib.Slider sldI3DL2 
         Height          =   495
         Index           =   0
         Left            =   480
         TabIndex        =   164
         ToolTipText     =   "Attenuation of the room effect, in millibels (mB), in the range from -10000 to 0. The default value is -1000 mB."
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Min             =   -10000
         Max             =   0
         SelStart        =   -1000
         TickFrequency   =   1000
         Value           =   -1000
      End
      Begin MSComctlLib.Slider sldI3DL2 
         Height          =   495
         Index           =   1
         Left            =   480
         TabIndex        =   168
         ToolTipText     =   "Attenuation of the room high-frequency effect, in mB, in the range from -10000 to 0. The default value is 0 mB."
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Min             =   -10000
         Max             =   0
         TickFrequency   =   1000
      End
      Begin MSComctlLib.Slider sldI3DL2 
         Height          =   495
         Index           =   2
         Left            =   480
         TabIndex        =   172
         ToolTipText     =   $"frmStudio.frx":0091
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Max             =   1000
         TickFrequency   =   100
      End
      Begin MSComctlLib.Slider sldI3DL2 
         Height          =   495
         Index           =   3
         Left            =   480
         TabIndex        =   176
         ToolTipText     =   "Decay time, in seconds, in the range from .1 to 20. The default value is 1.49 seconds."
         Top             =   2640
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Min             =   10
         Max             =   2000
         SelStart        =   149
         TickFrequency   =   100
         Value           =   149
      End
      Begin MSComctlLib.Slider sldI3DL2 
         Height          =   495
         Index           =   4
         Left            =   2760
         TabIndex        =   180
         ToolTipText     =   $"frmStudio.frx":0139
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Min             =   100
         Max             =   2000
         SelStart        =   830
         TickFrequency   =   100
         Value           =   830
      End
      Begin MSComctlLib.Slider sldI3DL2 
         Height          =   495
         Index           =   8
         Left            =   5040
         TabIndex        =   184
         ToolTipText     =   $"frmStudio.frx":01C5
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         SelStart        =   11
         TickFrequency   =   10
         Value           =   11
      End
      Begin MSComctlLib.Slider sldI3DL2 
         Height          =   495
         Index           =   5
         Left            =   2760
         TabIndex        =   188
         ToolTipText     =   $"frmStudio.frx":0286
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Min             =   -10000
         Max             =   1000
         SelStart        =   -2602
         TickFrequency   =   1000
         Value           =   -2602
      End
      Begin MSComctlLib.Slider sldI3DL2 
         Height          =   495
         Index           =   6
         Left            =   2760
         TabIndex        =   192
         ToolTipText     =   $"frmStudio.frx":030E
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Max             =   300
         SelStart        =   7
         TickFrequency   =   20
         Value           =   7
      End
      Begin MSComctlLib.Slider sldI3DL2 
         Height          =   495
         Index           =   7
         Left            =   2760
         TabIndex        =   196
         ToolTipText     =   $"frmStudio.frx":039D
         Top             =   2640
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Min             =   -10000
         Max             =   2000
         SelStart        =   200
         TickFrequency   =   1000
         Value           =   200
      End
      Begin MSComctlLib.Slider sldI3DL2 
         Height          =   495
         Index           =   9
         Left            =   5040
         TabIndex        =   200
         ToolTipText     =   "Echo density in the late reverberation decay, in percent, in the range from 0 to 100. The default value is 100.0 percent."
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Max             =   1000
         SelStart        =   1000
         TickFrequency   =   100
         Value           =   1000
      End
      Begin MSComctlLib.Slider sldI3DL2 
         Height          =   495
         Index           =   10
         Left            =   5040
         TabIndex        =   204
         ToolTipText     =   "Modal density in the late reverberation decay, in percent, in the range from 0 to 100. The default value is 100.0 percent."
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Max             =   1000
         SelStart        =   1000
         TickFrequency   =   100
         Value           =   1000
      End
      Begin MSComctlLib.Slider sldI3DL2 
         Height          =   495
         Index           =   11
         Left            =   5040
         TabIndex        =   208
         ToolTipText     =   "Reference high frequency, in hertz, in the range from 20 to 20000. The default value is 5000 Hz."
         Top             =   2640
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Min             =   20
         Max             =   20000
         SelStart        =   5000
         TickFrequency   =   1000
         Value           =   5000
      End
      Begin VB.Label Label122 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "20000 hz"
         Height          =   195
         Left            =   6060
         TabIndex        =   211
         Top             =   3120
         Width           =   660
      End
      Begin VB.Label Label121 
         AutoSize        =   -1  'True
         Caption         =   "20 hz"
         Height          =   195
         Left            =   5160
         TabIndex        =   210
         Top             =   3120
         Width           =   390
      End
      Begin VB.Label Label120 
         Alignment       =   1  'Right Justify
         Caption         =   "HiFrq Ref:"
         Height          =   435
         Left            =   4560
         TabIndex        =   209
         Top             =   2640
         Width           =   465
      End
      Begin VB.Label Label119 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "100.0%"
         Height          =   195
         Left            =   6195
         TabIndex        =   207
         Top             =   2400
         Width           =   525
      End
      Begin VB.Label Label118 
         AutoSize        =   -1  'True
         Caption         =   "0.0%"
         Height          =   195
         Left            =   5160
         TabIndex        =   206
         Top             =   2400
         Width           =   345
      End
      Begin VB.Label Label117 
         Alignment       =   1  'Right Justify
         Caption         =   "Den sity:"
         Height          =   435
         Left            =   4560
         TabIndex        =   205
         Top             =   1920
         Width           =   465
      End
      Begin VB.Label Label116 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "100.0%"
         Height          =   195
         Left            =   6195
         TabIndex        =   203
         Top             =   1680
         Width           =   525
      End
      Begin VB.Label Label115 
         AutoSize        =   -1  'True
         Caption         =   "0.0%"
         Height          =   195
         Left            =   5160
         TabIndex        =   202
         Top             =   1680
         Width           =   345
      End
      Begin VB.Label Label114 
         Alignment       =   1  'Right Justify
         Caption         =   "Diffu sion:"
         Height          =   435
         Left            =   4560
         TabIndex        =   201
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label Label113 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "2000 mB"
         Height          =   195
         Left            =   3810
         TabIndex        =   199
         Top             =   3120
         Width           =   630
      End
      Begin VB.Label Label112 
         AutoSize        =   -1  'True
         Caption         =   "-10000 mB"
         Height          =   195
         Left            =   2880
         TabIndex        =   198
         Top             =   3120
         Width           =   765
      End
      Begin VB.Label Label111 
         Caption         =   "Revrb:"
         Height          =   315
         Left            =   2280
         TabIndex        =   197
         Top             =   2760
         Width           =   555
      End
      Begin VB.Label Label110 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.3 sec"
         Height          =   195
         Left            =   3915
         TabIndex        =   195
         Top             =   2400
         Width           =   525
      End
      Begin VB.Label Label109 
         AutoSize        =   -1  'True
         Caption         =   "0.0 sec"
         Height          =   195
         Left            =   2880
         TabIndex        =   194
         Top             =   2400
         Width           =   525
      End
      Begin VB.Label Label108 
         Alignment       =   1  'Right Justify
         Caption         =   "Reflec tions Delay:"
         Height          =   675
         Left            =   2280
         TabIndex        =   193
         Top             =   1800
         Width           =   465
      End
      Begin VB.Label Label107 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1000 mB"
         Height          =   195
         Left            =   3810
         TabIndex        =   191
         Top             =   1680
         Width           =   630
      End
      Begin VB.Label Label106 
         AutoSize        =   -1  'True
         Caption         =   "-10000 mB"
         Height          =   195
         Left            =   2880
         TabIndex        =   190
         Top             =   1680
         Width           =   765
      End
      Begin VB.Label Label105 
         Alignment       =   1  'Right Justify
         Caption         =   "Reflec tions:"
         Height          =   435
         Left            =   2280
         TabIndex        =   189
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label Label104 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.1 sec"
         Height          =   195
         Left            =   6195
         TabIndex        =   187
         Top             =   960
         Width           =   525
      End
      Begin VB.Label Label103 
         AutoSize        =   -1  'True
         Caption         =   "0.0 sec"
         Height          =   195
         Left            =   5160
         TabIndex        =   186
         Top             =   960
         Width           =   525
      End
      Begin VB.Label Label102 
         Alignment       =   1  'Right Justify
         Caption         =   "Revrb Delay:"
         Height          =   435
         Left            =   4560
         TabIndex        =   185
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label101 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "2.0"
         Height          =   195
         Left            =   4215
         TabIndex        =   183
         Top             =   960
         Width           =   225
      End
      Begin VB.Label Label100 
         AutoSize        =   -1  'True
         Caption         =   "0.1"
         Height          =   195
         Left            =   2880
         TabIndex        =   182
         Top             =   960
         Width           =   225
      End
      Begin VB.Label Label96 
         Alignment       =   1  'Right Justify
         Caption         =   "Decay HiFrq Ratio:"
         Height          =   555
         Left            =   2280
         TabIndex        =   181
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label95 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "20.0 sec."
         Height          =   195
         Left            =   1500
         TabIndex        =   179
         Top             =   3120
         Width           =   660
      End
      Begin VB.Label Label94 
         AutoSize        =   -1  'True
         Caption         =   "0.1 sec."
         Height          =   195
         Left            =   600
         TabIndex        =   178
         Top             =   3120
         Width           =   570
      End
      Begin VB.Label Label93 
         Alignment       =   1  'Right Justify
         Caption         =   "Decay Time:"
         Height          =   435
         Left            =   0
         TabIndex        =   177
         Top             =   2640
         Width           =   465
      End
      Begin VB.Label Label92 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "10.0"
         Height          =   195
         Left            =   1845
         TabIndex        =   175
         Top             =   2400
         Width           =   315
      End
      Begin VB.Label Label91 
         AutoSize        =   -1  'True
         Caption         =   "0.0"
         Height          =   195
         Left            =   600
         TabIndex        =   174
         Top             =   2400
         Width           =   225
      End
      Begin VB.Label Label90 
         Alignment       =   1  'Right Justify
         Caption         =   "Room Rolloff Factor:"
         Height          =   675
         Left            =   -90
         TabIndex        =   173
         Top             =   1800
         Width           =   585
      End
      Begin VB.Label Label89 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0 mB"
         Height          =   195
         Left            =   1800
         TabIndex        =   171
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label Label88 
         AutoSize        =   -1  'True
         Caption         =   "-10000 mB"
         Height          =   195
         Left            =   600
         TabIndex        =   170
         Top             =   1680
         Width           =   765
      End
      Begin VB.Label Label87 
         Alignment       =   1  'Right Justify
         Caption         =   "Room HiFrq:"
         Height          =   435
         Left            =   0
         TabIndex        =   169
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label Label85 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0 mB"
         Height          =   195
         Left            =   1800
         TabIndex        =   167
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label84 
         AutoSize        =   -1  'True
         Caption         =   "-10000 mB"
         Height          =   195
         Left            =   600
         TabIndex        =   166
         Top             =   960
         Width           =   765
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Room:"
         Height          =   195
         Left            =   0
         TabIndex        =   165
         Top             =   600
         Width           =   465
      End
   End
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      Caption         =   "Chorus"
      Height          =   3495
      Index           =   5
      Left            =   240
      TabIndex        =   120
      Top             =   480
      Width           =   6975
      Begin VB.ComboBox cmbFlanger 
         Height          =   315
         Index           =   1
         ItemData        =   "frmStudio.frx":0424
         Left            =   4560
         List            =   "frmStudio.frx":0437
         Style           =   2  'Dropdown List
         TabIndex        =   134
         ToolTipText     =   "Phase differential between left and right low-frequency oscillators."
         Top             =   2880
         Width           =   1695
      End
      Begin VB.ComboBox cmbFlanger 
         Height          =   315
         Index           =   0
         ItemData        =   "frmStudio.frx":047A
         Left            =   4560
         List            =   "frmStudio.frx":0484
         Style           =   2  'Dropdown List
         TabIndex        =   132
         ToolTipText     =   "Waveform of the low-frequency oscillator. . By default, the waveform is a sine."
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox chkEffectOn 
         Caption         =   "Enable Fl&anger"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   121
         Top             =   120
         Width           =   1455
      End
      Begin MSComctlLib.Slider sldFlanger 
         Height          =   495
         Index           =   2
         Left            =   600
         TabIndex        =   126
         ToolTipText     =   "Percentage of output signal to feed back into the effects input, in the range from -99 to 99. The default value is 0."
         Top             =   2520
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Min             =   -99
         Max             =   99
         TickFrequency   =   10
      End
      Begin MSComctlLib.Slider sldFlanger 
         Height          =   495
         Index           =   4
         Left            =   3960
         TabIndex        =   130
         ToolTipText     =   "Number of milliseconds the input is delayed before it is played back, in the range from 0 to 4. The default value is 0 ms."
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   2
         Max             =   20
         SelStart        =   5
         TickFrequency   =   2
         Value           =   5
      End
      Begin MSComctlLib.Slider sldFlanger 
         Height          =   495
         Index           =   3
         Left            =   3960
         TabIndex        =   128
         ToolTipText     =   "Frequency of the low-frequency oscillator, in the range from 0 to 10. The default value is 0."
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   2
         SelStart        =   3
         Value           =   3
      End
      Begin MSComctlLib.Slider sldFlanger 
         Height          =   495
         Index           =   1
         Left            =   600
         TabIndex        =   124
         ToolTipText     =   $"frmStudio.frx":0498
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         SelStart        =   25
         TickFrequency   =   10
         Value           =   25
      End
      Begin MSComctlLib.Slider sldFlanger 
         Height          =   495
         Index           =   0
         Left            =   600
         TabIndex        =   122
         ToolTipText     =   "Ratio of wet (processed) signal to dry (unprocessed) signal. Must be in the range from 0 through 100 (all wet). "
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         SelStart        =   50
         TickFrequency   =   10
         Value           =   50
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         Caption         =   "Wet Dry Mix:"
         Height          =   195
         Left            =   240
         TabIndex        =   145
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "Dry"
         Height          =   195
         Left            =   360
         TabIndex        =   144
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         Caption         =   "Wet"
         Height          =   195
         Left            =   2775
         TabIndex        =   143
         Top             =   960
         Width           =   300
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         Caption         =   "Depth:"
         Height          =   195
         Left            =   240
         TabIndex        =   142
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         Caption         =   "0%"
         Height          =   195
         Left            =   360
         TabIndex        =   141
         Top             =   1800
         Width           =   210
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         Caption         =   "100%"
         Height          =   195
         Left            =   2775
         TabIndex        =   140
         Top             =   1800
         Width           =   390
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         Caption         =   "Frequency:"
         Height          =   195
         Left            =   3600
         TabIndex        =   139
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   3720
         TabIndex        =   138
         Top             =   960
         Width           =   90
      End
      Begin VB.Label Label74 
         AutoSize        =   -1  'True
         Caption         =   "10"
         Height          =   195
         Left            =   6135
         TabIndex        =   137
         Top             =   960
         Width           =   180
      End
      Begin VB.Label Label75 
         AutoSize        =   -1  'True
         Caption         =   "Delay:"
         Height          =   195
         Left            =   3600
         TabIndex        =   136
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label Label76 
         AutoSize        =   -1  'True
         Caption         =   "0 ms"
         Height          =   195
         Left            =   3600
         TabIndex        =   135
         Top             =   1800
         Width           =   330
      End
      Begin VB.Label Label77 
         AutoSize        =   -1  'True
         Caption         =   "20 ms"
         Height          =   195
         Left            =   6135
         TabIndex        =   133
         Top             =   1800
         Width           =   420
      End
      Begin VB.Label Label78 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Phase:"
         Height          =   195
         Left            =   3840
         TabIndex        =   131
         Top             =   2925
         Width           =   495
      End
      Begin VB.Label Label79 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Waveform:"
         Height          =   195
         Left            =   3555
         TabIndex        =   129
         Top             =   2445
         Width           =   780
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "99%"
         Height          =   195
         Left            =   2775
         TabIndex        =   127
         Top             =   2640
         Width           =   300
      End
      Begin VB.Label Label81 
         AutoSize        =   -1  'True
         Caption         =   "-99%"
         Height          =   195
         Left            =   240
         TabIndex        =   125
         Top             =   2640
         Width           =   345
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         Caption         =   "Feedback:"
         Height          =   195
         Left            =   240
         TabIndex        =   123
         Top             =   2280
         Width           =   765
      End
   End
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      Caption         =   "Gargle"
      Height          =   3495
      Index           =   6
      Left            =   240
      TabIndex        =   146
      Top             =   480
      Width           =   6975
      Begin VB.CheckBox chkEffectOn 
         Caption         =   "Enable G&argle"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   147
         Top             =   120
         Width           =   1455
      End
      Begin VB.ComboBox cmbGargle 
         Height          =   315
         ItemData        =   "frmStudio.frx":055A
         Left            =   4560
         List            =   "frmStudio.frx":0564
         Style           =   2  'Dropdown List
         TabIndex        =   150
         ToolTipText     =   "Shape of the modulation wave."
         Top             =   960
         Width           =   1695
      End
      Begin MSComctlLib.Slider sldGargle 
         Height          =   495
         Left            =   600
         TabIndex        =   148
         ToolTipText     =   "Rate of modulation, in Hertz. Must be in the range from 1 through 1000."
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   100
         SmallChange     =   10
         Min             =   1
         Max             =   1000
         SelStart        =   500
         TickFrequency   =   100
         Value           =   500
      End
      Begin VB.Label Label86 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Waveform:"
         Height          =   195
         Left            =   3555
         TabIndex        =   153
         Top             =   1005
         Width           =   780
      End
      Begin VB.Label Label97 
         AutoSize        =   -1  'True
         Caption         =   "1000 hz"
         Height          =   195
         Left            =   2775
         TabIndex        =   152
         Top             =   960
         Width           =   570
      End
      Begin VB.Label Label98 
         AutoSize        =   -1  'True
         Caption         =   "1 hz"
         Height          =   195
         Left            =   240
         TabIndex        =   151
         Top             =   960
         Width           =   300
      End
      Begin VB.Label Label99 
         AutoSize        =   -1  'True
         Caption         =   "Rate of Modulation:"
         Height          =   195
         Left            =   240
         TabIndex        =   149
         Top             =   600
         Width           =   1395
      End
   End
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      Caption         =   "Chorus"
      Height          =   3495
      Index           =   2
      Left            =   240
      TabIndex        =   40
      Top             =   480
      Width           =   6975
      Begin VB.CheckBox chkEffectOn 
         Caption         =   "Enable C&ompressor"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   41
         Top             =   120
         Width           =   1815
      End
      Begin MSComctlLib.Slider sldCompressor 
         Height          =   495
         Index           =   2
         Left            =   600
         TabIndex        =   45
         ToolTipText     =   $"frmStudio.frx":057C
         Top             =   2520
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   250
         SmallChange     =   50
         Min             =   50
         Max             =   3000
         SelStart        =   50
         TickFrequency   =   250
         Value           =   50
      End
      Begin MSComctlLib.Slider sldCompressor 
         Height          =   495
         Index           =   4
         Left            =   3960
         TabIndex        =   49
         ToolTipText     =   "Compression ratio, in the range from 1 to 100. The default value is 10, which means 10:1 compression."
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Min             =   1
         Max             =   100
         SelStart        =   10
         TickFrequency   =   10
         Value           =   10
      End
      Begin MSComctlLib.Slider sldCompressor 
         Height          =   495
         Index           =   3
         Left            =   3960
         TabIndex        =   47
         ToolTipText     =   "Point at which compression begins, in decibels, in the range from -60 to 0. The default value is -10 dB. "
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         Min             =   -60
         Max             =   0
         SelStart        =   -10
         TickFrequency   =   10
         Value           =   -10
      End
      Begin MSComctlLib.Slider sldCompressor 
         Height          =   495
         Index           =   1
         Left            =   600
         TabIndex        =   43
         ToolTipText     =   "Time before compression reaches its full value, in the range from 0.01 to 500. The default value is 0.01 ms."
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   1000
         SmallChange     =   10
         Max             =   50000
         SelStart        =   1
         TickFrequency   =   5000
         Value           =   1
         TextPosition    =   1
      End
      Begin MSComctlLib.Slider sldCompressor 
         Height          =   495
         Index           =   0
         Left            =   600
         TabIndex        =   42
         ToolTipText     =   "Output gain of signal after compression, in the range from -60 to 60. The default value is 0 dB."
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Min             =   -60
         Max             =   60
         TickFrequency   =   10
      End
      Begin MSComctlLib.Slider sldCompressor 
         Height          =   495
         Index           =   5
         Left            =   3960
         TabIndex        =   51
         ToolTipText     =   $"frmStudio.frx":060D
         Top             =   2520
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   1
         Max             =   4
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Gain:"
         Height          =   195
         Left            =   240
         TabIndex        =   65
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "-60 dB"
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "60 dB"
         Height          =   195
         Left            =   2775
         TabIndex        =   63
         Top             =   960
         Width           =   420
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Attack:"
         Height          =   195
         Left            =   240
         TabIndex        =   62
         Top             =   1440
         Width           =   510
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "0 ms"
         Height          =   195
         Left            =   240
         TabIndex        =   61
         Top             =   1800
         Width           =   330
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "500 ms"
         Height          =   195
         Left            =   2775
         TabIndex        =   60
         Top             =   1680
         Width           =   510
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Threshold:"
         Height          =   195
         Left            =   3600
         TabIndex        =   59
         Top             =   600
         Width           =   750
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "-60 dB"
         Height          =   195
         Left            =   3480
         TabIndex        =   58
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "0 dB"
         Height          =   195
         Left            =   6135
         TabIndex        =   57
         Top             =   960
         Width           =   330
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Compression Ratio:"
         Height          =   195
         Left            =   3600
         TabIndex        =   56
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Left            =   3840
         TabIndex        =   55
         Top             =   1800
         Width           =   90
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "100"
         Height          =   195
         Left            =   6135
         TabIndex        =   54
         Top             =   1800
         Width           =   270
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "3000 ms"
         Height          =   195
         Left            =   2775
         TabIndex        =   53
         Top             =   2640
         Width           =   600
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "50 ms"
         Height          =   195
         Left            =   150
         TabIndex        =   52
         Top             =   2640
         Width           =   435
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "Release:"
         Height          =   195
         Left            =   240
         TabIndex        =   50
         Top             =   2280
         Width           =   630
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Predelay:"
         Height          =   195
         Left            =   3600
         TabIndex        =   48
         Top             =   2280
         Width           =   660
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "0 ms"
         Height          =   195
         Left            =   3600
         TabIndex        =   46
         Top             =   2640
         Width           =   330
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "4 ms"
         Height          =   195
         Left            =   6135
         TabIndex        =   44
         Top             =   2640
         Width           =   330
      End
   End
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      Caption         =   "Chorus"
      Height          =   3495
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   480
      Width           =   6975
      Begin VB.CheckBox chkEffectOn 
         Caption         =   "Enable &Chorus"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Width           =   1455
      End
      Begin VB.ComboBox cmbChorus 
         Height          =   315
         Index           =   0
         ItemData        =   "frmStudio.frx":06A0
         Left            =   4560
         List            =   "frmStudio.frx":06AA
         Style           =   2  'Dropdown List
         TabIndex        =   33
         ToolTipText     =   "Waveform of the low-frequency oscillator."
         Top             =   2400
         Width           =   1695
      End
      Begin VB.ComboBox cmbChorus 
         Height          =   315
         Index           =   1
         ItemData        =   "frmStudio.frx":06BE
         Left            =   4560
         List            =   "frmStudio.frx":06D1
         Style           =   2  'Dropdown List
         TabIndex        =   36
         ToolTipText     =   "Phase differential between left and right low-frequency oscillators."
         Top             =   2880
         Width           =   1695
      End
      Begin MSComctlLib.Slider sldChorus 
         Height          =   495
         Index           =   2
         Left            =   600
         TabIndex        =   24
         ToolTipText     =   "Percentage of output signal to feed back into the effects input, in the range from -99 to 99. The default value is 0."
         Top             =   2520
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Min             =   -99
         Max             =   99
         SelStart        =   33
         TickFrequency   =   10
         Value           =   33
      End
      Begin MSComctlLib.Slider sldChorus 
         Height          =   495
         Index           =   4
         Left            =   3960
         TabIndex        =   30
         ToolTipText     =   "Number of milliseconds the input is delayed before it is played back, in the range from 0 to 20. The default value is 0 ms."
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   2
         Max             =   20
         SelStart        =   5
         TickFrequency   =   2
         Value           =   5
      End
      Begin MSComctlLib.Slider sldChorus 
         Height          =   495
         Index           =   3
         Left            =   3960
         TabIndex        =   27
         ToolTipText     =   "Frequency of the low-frequency oscillator, in the range from 0 to 10. The default value is 0."
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   2
         SelStart        =   3
         Value           =   3
      End
      Begin MSComctlLib.Slider sldChorus 
         Height          =   495
         Index           =   1
         Left            =   600
         TabIndex        =   21
         ToolTipText     =   $"frmStudio.frx":0714
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         SelStart        =   25
         TickFrequency   =   10
         Value           =   25
      End
      Begin MSComctlLib.Slider sldChorus 
         Height          =   495
         Index           =   0
         Left            =   600
         TabIndex        =   18
         ToolTipText     =   "Ratio of wet (processed) signal to dry (unprocessed) signal. Must be in the range from 0 through 100 (all wet)."
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         SelStart        =   50
         TickFrequency   =   10
         Value           =   50
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Feedback:"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   2280
         Width           =   765
      End
      Begin VB.Label Label9 
         Caption         =   "-99%"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "99%"
         Height          =   255
         Left            =   2775
         TabIndex        =   38
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Waveform:"
         Height          =   195
         Left            =   3555
         TabIndex        =   31
         Top             =   2445
         Width           =   780
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Phase:"
         Height          =   195
         Left            =   3840
         TabIndex        =   34
         Top             =   2925
         Width           =   495
      End
      Begin VB.Label Label16 
         Caption         =   "20 ms"
         Height          =   255
         Left            =   6135
         TabIndex        =   37
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label Label15 
         Caption         =   "0 ms"
         Height          =   255
         Left            =   3600
         TabIndex        =   35
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Delay:"
         Height          =   195
         Left            =   3600
         TabIndex        =   28
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label Label13 
         Caption         =   "10"
         Height          =   255
         Left            =   6135
         TabIndex        =   32
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "0"
         Height          =   255
         Left            =   3720
         TabIndex        =   29
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Frequency:"
         Height          =   195
         Left            =   3600
         TabIndex        =   25
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label7 
         Caption         =   "100%"
         Height          =   255
         Left            =   2775
         TabIndex        =   26
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label Label6 
         Caption         =   "0%"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Depth:"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "Wet"
         Height          =   255
         Left            =   2775
         TabIndex        =   20
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Dry"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Wet Dry Mix:"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   915
      End
   End
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      Height          =   3495
      Index           =   9
      Left            =   240
      TabIndex        =   154
      Top             =   480
      Width           =   6975
      Begin VB.CommandButton cmdPitch 
         Caption         =   "&Reset"
         Height          =   375
         Left            =   5640
         TabIndex        =   158
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton optPitch 
         Caption         =   "View in percent."
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   157
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton optPitch 
         Caption         =   "View in hertz."
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   156
         Top             =   1320
         Value           =   -1  'True
         Width           =   1335
      End
      Begin MSComctlLib.Slider sldPitch 
         Height          =   495
         Left            =   120
         TabIndex        =   155
         Top             =   720
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   500
         SmallChange     =   10
         Min             =   100
         Max             =   500000
         SelStart        =   100
         TickFrequency   =   10000
         Value           =   100
         TextPosition    =   1
      End
      Begin VB.Label lblPitch 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "100 hertz"
         Height          =   195
         Left            =   6075
         TabIndex        =   160
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label83 
         AutoSize        =   -1  'True
         Caption         =   "Pitch:"
         Height          =   195
         Left            =   240
         TabIndex        =   159
         Top             =   480
         Width           =   405
      End
   End
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      Caption         =   "EQ"
      Height          =   3495
      Index           =   0
      Left            =   240
      TabIndex        =   66
      Top             =   480
      Width           =   6975
      Begin MSComctlLib.ImageList iml 
         Left            =   120
         Top             =   3240
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
               Picture         =   "frmStudio.frx":07CF
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmStudio.frx":0B23
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbr 
         Height          =   330
         Left            =   4080
         TabIndex        =   13
         Top             =   2880
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   582
         ButtonWidth     =   2249
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "iml"
         DisabledImageList=   "iml"
         HotImageList    =   "iml"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Load Preset"
               Key             =   "load"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Save Preset"
               Key             =   "save"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin VB.ListBox lstPreset 
         Height          =   2205
         Left            =   4080
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   2775
      End
      Begin VB.CheckBox chkEqualizer 
         Caption         =   "Enable Eq&ualizer"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   1575
      End
      Begin SimpleAmp.CtrlScroller scrEQ 
         Height          =   2250
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   150
         _ExtentX        =   873
         _ExtentY        =   873
         Vertical        =   -1  'True
         Snap            =   -1  'True
         Max             =   30
         Value           =   15
      End
      Begin SimpleAmp.CtrlScroller scrEQ 
         Height          =   2250
         Index           =   1
         Left            =   600
         TabIndex        =   3
         Top             =   600
         Width           =   150
         _ExtentX        =   873
         _ExtentY        =   873
         Vertical        =   -1  'True
         Snap            =   -1  'True
         Max             =   30
         Value           =   15
      End
      Begin SimpleAmp.CtrlScroller scrEQ 
         Height          =   2250
         Index           =   2
         Left            =   960
         TabIndex        =   4
         Top             =   600
         Width           =   150
         _ExtentX        =   873
         _ExtentY        =   873
         Vertical        =   -1  'True
         Snap            =   -1  'True
         Max             =   30
         Value           =   15
      End
      Begin SimpleAmp.CtrlScroller scrEQ 
         Height          =   2250
         Index           =   3
         Left            =   1320
         TabIndex        =   5
         Top             =   600
         Width           =   150
         _ExtentX        =   873
         _ExtentY        =   873
         Vertical        =   -1  'True
         Snap            =   -1  'True
         Max             =   30
         Value           =   15
      End
      Begin SimpleAmp.CtrlScroller scrEQ 
         Height          =   2250
         Index           =   4
         Left            =   1680
         TabIndex        =   6
         Top             =   600
         Width           =   150
         _ExtentX        =   873
         _ExtentY        =   873
         Vertical        =   -1  'True
         Snap            =   -1  'True
         Max             =   30
         Value           =   15
      End
      Begin SimpleAmp.CtrlScroller scrEQ 
         Height          =   2250
         Index           =   5
         Left            =   2040
         TabIndex        =   7
         Top             =   600
         Width           =   150
         _ExtentX        =   873
         _ExtentY        =   873
         Vertical        =   -1  'True
         Snap            =   -1  'True
         Max             =   30
         Value           =   15
      End
      Begin SimpleAmp.CtrlScroller scrEQ 
         Height          =   2250
         Index           =   6
         Left            =   2400
         TabIndex        =   8
         Top             =   600
         Width           =   150
         _ExtentX        =   873
         _ExtentY        =   873
         Vertical        =   -1  'True
         Snap            =   -1  'True
         Max             =   30
         Value           =   15
      End
      Begin SimpleAmp.CtrlScroller scrEQ 
         Height          =   2250
         Index           =   7
         Left            =   2760
         TabIndex        =   9
         Top             =   600
         Width           =   150
         _ExtentX        =   873
         _ExtentY        =   873
         Vertical        =   -1  'True
         Snap            =   -1  'True
         Max             =   30
         Value           =   15
      End
      Begin SimpleAmp.CtrlScroller scrEQ 
         Height          =   2250
         Index           =   8
         Left            =   3120
         TabIndex        =   10
         Top             =   600
         Width           =   150
         _ExtentX        =   873
         _ExtentY        =   873
         Vertical        =   -1  'True
         Snap            =   -1  'True
         Max             =   30
         Value           =   15
      End
      Begin SimpleAmp.CtrlScroller scrEQ 
         Height          =   2250
         Index           =   9
         Left            =   3480
         TabIndex        =   11
         Top             =   600
         Width           =   150
         _ExtentX        =   873
         _ExtentY        =   873
         Vertical        =   -1  'True
         Snap            =   -1  'True
         Max             =   30
         Value           =   15
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   3735
         Y1              =   1725
         Y2              =   1725
      End
      Begin VB.Label lblEQ 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "60"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   78
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label lblEQ 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "170"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   77
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label lblEQ 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "310"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   76
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label lblEQ 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "600"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   75
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label lblEQ 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1K"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   74
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label lblEQ 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "3K"
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   73
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label lblEQ 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "6K"
         Height          =   255
         Index           =   6
         Left            =   2280
         TabIndex        =   72
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label lblEQ 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "12K"
         Height          =   255
         Index           =   7
         Left            =   2640
         TabIndex        =   71
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label lblEQ 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "14K"
         Height          =   255
         Index           =   8
         Left            =   3000
         TabIndex        =   70
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label lblEQ 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "16K"
         Height          =   255
         Index           =   9
         Left            =   3360
         TabIndex        =   69
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "+"
         Height          =   255
         Left            =   3720
         TabIndex        =   68
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label20 
         Caption         =   "_"
         Height          =   255
         Left            =   3720
         TabIndex        =   67
         Top             =   2560
         Width           =   255
      End
   End
   Begin MSComctlLib.TabStrip tabs 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7011
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   10
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "E&qualizer"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Chorus"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "C&ompressor"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Distortion"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Echo"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Flanger"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Gargle"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "I&3DL2 Reverb"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Waves Reverb"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Pitch"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmStudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lDefPitch As Long

Private Sub chkEcho_Click()
  UpdateEcho
End Sub

Private Sub chkEffectOn_Click(index As Integer)
  UpdateFX
End Sub

Private Sub chkEqualizer_Click()
  Settings.EQon = CBool(chkEqualizer.Value)
  UpdateFX
End Sub

Private Sub cmbChorus_Click(index As Integer)
  UpdateChorus
End Sub

Private Sub cmbFlanger_Click(index As Integer)
  UpdateFlanger
End Sub

Private Sub cmbGargle_Click()
  UpdateGargle
End Sub

Private Sub cmdPitch_Click()
  If optPitch(0) Then
    sldPitch.Value = lDefPitch
  Else
    sldPitch.Value = 100
  End If
  sldPitch_Scroll
End Sub

Private Sub Form_Activate()
  AlwaysOnTop Me, Settings.OnTop
End Sub

Private Sub Form_Load()
  On Error Resume Next
  
  AlwaysOnTop Me, Settings.OnTop
  
  Dim x As Long
  tabs_Click
  optPitch_Click 0
  
  For x = 0 To 7
    chkEffectOn(x).Enabled = Settings.DXFXon
  Next x
  
  If Settings.DXFXon Then
    If Settings.EQon And EQHandle(0) = 0 Then UpdateFX
    chkEqualizer.Value = Abs(Settings.EQon)
  End If
  
  For x = 0 To 9
    Set scrEQ(x).Bar = LoadResPicture("EQBAR", vbResBitmap)
    Set scrEQ(x).BarOver = LoadResPicture("EQBAR", vbResBitmap)
    Set scrEQ(x).ScrollAfter = LoadResPicture("EQ", vbResBitmap)
    Set scrEQ(x).ScrollBefore = LoadResPicture("EQ", vbResBitmap)
    scrEQ(x).Max = 30
    scrEQ(x).Value = 30 - (EQValue(x) + 15)
  Next x
  
  RefreshList
  
  cmbChorus(0).ListIndex = 1
  cmbChorus(1).ListIndex = 2
  cmbFlanger(0).ListIndex = 1
  cmbFlanger(1).ListIndex = 2
  cmbGargle.ListIndex = 0
  Label24 = "Attack: (" & FormatNumber(Round(sldCompressor(1).Value / 100, 2), 2) & " ms)"
  
  ResetPitch
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim x As Long
  For x = 0 To 7
    chkEffectOn(x).Value = vbUnchecked
  Next x
  UpdateFX
End Sub

Private Sub lstPreset_DblClick()
  tbr_ButtonClick tbr.Buttons(1)
End Sub

Private Sub optPitch_Click(index As Integer)
  On Error Resume Next
  sldPitch.Visible = False
  Select Case index
    Case 0
      sldPitch.TickFrequency = 10000
      sldPitch.Min = 100
      sldPitch.Max = 500000
      sldPitch.LargeChange = 10000
      sldPitch.SmallChange = 100
      sldPitch.Value = Sound.StreamFrequency
    Case 1
      sldPitch.TickFrequency = 100
      sldPitch.Min = 1
      sldPitch.Max = 500
      sldPitch.LargeChange = 10
      sldPitch.SmallChange = 1
      sldPitch.Value = (Sound.StreamFrequency / lDefPitch) * 100
  End Select
  sldPitch.Visible = True
  sldPitch_Scroll
End Sub

Private Sub scrEQ_Change(index As Integer)
  On Error Resume Next
  EQValue(index) = (scrEQ(index).Value - 15) * -1
  
  UpdateEQ
End Sub

Private Sub sldChorus_Scroll(index As Integer)
  UpdateChorus
End Sub

Private Sub sldCompressor_Scroll(index As Integer)
  UpdateCompressor
  Label24 = "Attack: (" & FormatNumber(Round(sldCompressor(1).Value / 100, 2), 2) & " ms)"
End Sub

Private Sub sldDistortion_Scroll(index As Integer)
  UpdateDistortion
End Sub

Private Sub sldEcho_Scroll(index As Integer)
  UpdateEcho
End Sub

Private Sub sldFlanger_Scroll(index As Integer)
  UpdateFlanger
End Sub

Private Sub sldGargle_Scroll()
  UpdateGargle
End Sub

Private Sub sldI3DL2_Scroll(index As Integer)
  UpdateI3DL2
End Sub

Private Sub sldPitch_Click()
  sldPitch_Scroll
End Sub

Private Sub sldPitch_Scroll()
  On Error Resume Next
  If optPitch(0) Then
    Sound.StreamFrequency = sldPitch.Value
    lblPitch = sldPitch.Value & " hz"
  Else
    Sound.StreamFrequency = lDefPitch * (sldPitch.Value / 100)
    lblPitch = sldPitch.Value & "%"
  End If
End Sub

Private Sub sldWave_Scroll(index As Integer)
  UpdateWave
End Sub

Private Sub tabs_Click()
  On Error Resume Next
  Dim x As Long
  For x = 0 To 9
    If x <> tabs.SelectedItem.index - 1 Then
      frm(x).Visible = False
    Else
      frm(x).Visible = True
    End If
  Next
End Sub

Public Sub UpdateChorus()
  On Error Resume Next
  
  If DXFXHandle(0) <> 0 And Settings.DXFXon Then
    Call FSOUND_FX_SetChorus(DXFXHandle(0), sldChorus(0).Value, sldChorus(1).Value, sldChorus(2).Value, sldChorus(3).Value, cmbChorus(0).ListIndex, sldChorus(4).Value, cmbChorus(1).ListIndex)
  End If
End Sub

Public Sub UpdateCompressor()
  On Error Resume Next
  
  If DXFXHandle(1) <> 0 And Settings.DXFXon Then
    Call FSOUND_FX_SetCompressor(DXFXHandle(1), sldCompressor(0).Value, sldCompressor(1).Value / 100, sldCompressor(2).Value, sldCompressor(4).Value, sldCompressor(5).Value, sldCompressor(6).Value)
  End If
End Sub

Public Sub UpdateDistortion()
  On Error Resume Next
  
  If DXFXHandle(2) <> 0 And Settings.DXFXon Then
    Call FSOUND_FX_SetDistortion(DXFXHandle(2), sldDistortion(0).Value, sldDistortion(1).Value, sldDistortion(2).Value, sldDistortion(3).Value, sldDistortion(4).Value)
  End If
End Sub

Public Sub UpdateEcho()
  On Error Resume Next
  
  If DXFXHandle(3) <> 0 And Settings.DXFXon Then
    Call FSOUND_FX_SetEcho(DXFXHandle(3), sldEcho(0).Value, sldEcho(1).Value, sldEcho(2).Value, sldEcho(3).Value, chkEcho.Value)
  End If
End Sub

Public Sub UpdateFlanger()
  On Error Resume Next
  
  If DXFXHandle(4) <> 0 And Settings.DXFXon Then
    Call FSOUND_FX_SetFlanger(DXFXHandle(4), sldFlanger(0).Value, sldFlanger(1).Value, sldFlanger(2).Value, sldFlanger(3).Value, cmbFlanger(0).ListIndex, sldFlanger(4).Value, cmbFlanger(1).ListIndex)
  End If
End Sub

Public Sub UpdateGargle()
  On Error Resume Next
  
  If DXFXHandle(5) <> 0 And Settings.DXFXon Then
    Call FSOUND_FX_SetGargle(DXFXHandle(5), sldGargle.Value, cmbGargle.ListIndex)
  End If
End Sub

Public Sub UpdateI3DL2()
  On Error Resume Next
  
  If DXFXHandle(6) <> 0 And Settings.DXFXon Then
    Call FSOUND_FX_SetI3DL2Reverb(DXFXHandle(6), sldI3DL2(0).Value, sldI3DL2(1).Value, sldI3DL2(2).Value / 100, sldI3DL2(3).Value / 100, sldI3DL2(4).Value / 1000, sldI3DL2(5).Value, sldI3DL2(6).Value / 1000, sldI3DL2(7).Value, sldI3DL2(8).Value / 1000, sldI3DL2(9).Value / 1000, sldI3DL2(10).Value / 1000, sldI3DL2(11).Value)
  End If
End Sub

Public Sub UpdateWave()
  On Error Resume Next
  
  If DXFXHandle(7) <> 0 And Settings.DXFXon Then
    Call FSOUND_FX_SetWavesReverb(DXFXHandle(7), sldWave(0).Value, sldWave(1).Value, sldWave(2).Value, sldWave(3).Value / 1000)
  End If
End Sub

Public Sub ResetPitch()
  On Error Resume Next
  lDefPitch = Sound.StreamFrequency
  If optPitch(0) Then
    sldPitch.Value = lDefPitch
  Else
    sldPitch.Value = 100
  End If
  sldPitch_Scroll
End Sub

Public Sub RefreshList()
  On Error Resume Next
  Dim cFind As New clsFind, x As Long
  cFind.Find App.Path & "\eq", "*.seq"
  lstPreset.Clear
  For x = 1 To cFind.Count
    lstPreset.AddItem cFind(x).sName
  Next x
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
  On Error GoTo errh
  Dim cDlg As New clsCommonDialog
  
  Select Case Button.Key
    Case "save"
      With cDlg
        Set .Parent = Me
        .CancelError = True
        .Filter = "Equalizer Presets (*.seq)|*.seq"
        .DialogTitle = "Save Preset"
        .InitDir = App.Path & "\eq"
        .Flags = cdlOFNExplorer Or cdlOFNNoChangeDir Or cdlOFNOverwritePrompt
        .ShowSave
        SaveEQ .FileName
        RefreshList
      End With
    Case "load"
      If lstPreset.ListIndex > -1 Then
        LoadEQ App.Path & "\eq\" & lstPreset.List(lstPreset.ListIndex) & ".seq"
        Dim x As Byte
        For x = 0 To 9
          scrEQ(x).Value = 30 - (EQValue(x) + 15)
        Next x
        UpdateEQ
      End If
      
  End Select
errh:
End Sub
