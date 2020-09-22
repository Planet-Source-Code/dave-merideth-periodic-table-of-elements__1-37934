VERSION 5.00
Begin VB.Form form_main 
   Caption         =   "Periodic Table Of Elements"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_oxidation 
      Height          =   285
      Left            =   7920
      TabIndex        =   182
      Top             =   7440
      Width           =   615
   End
   Begin VB.TextBox txt_discovery 
      Height          =   285
      Left            =   7920
      TabIndex        =   169
      Top             =   7920
      Width           =   3495
   End
   Begin VB.TextBox txt_electronegativity 
      Height          =   285
      Left            =   7920
      TabIndex        =   168
      Top             =   6960
      Width           =   2295
   End
   Begin VB.TextBox txt_isotopes 
      Height          =   285
      Left            =   7920
      TabIndex        =   167
      Top             =   6480
      Width           =   495
   End
   Begin VB.TextBox txt_density 
      Height          =   285
      Left            =   7920
      TabIndex        =   166
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox txt_melt_boil 
      Height          =   285
      Left            =   7920
      TabIndex        =   165
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox txt_elec_shell 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3120
      TabIndex        =   164
      Top             =   7920
      Width           =   1455
   End
   Begin VB.TextBox txt_orbitals 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2640
      TabIndex        =   163
      Top             =   7440
      Width           =   1935
   End
   Begin VB.TextBox txt_atom_rad 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      TabIndex        =   162
      Top             =   6960
      Width           =   495
   End
   Begin VB.TextBox txt_atom_mass 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3600
      TabIndex        =   161
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox txt_name 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3360
      TabIndex        =   160
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox txt_atom_num 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3960
      TabIndex        =   159
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label lbl_hide 
      Caption         =   "Hide"
      Height          =   255
      Left            =   120
      TabIndex        =   183
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label lbl_oxidation 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Oxidation"
      Height          =   255
      Left            =   6960
      TabIndex        =   181
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label lbl_electronegativity 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Electronegativity"
      Height          =   255
      Left            =   6480
      TabIndex        =   180
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label lbl_atom_rad 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Atomic Radius"
      Height          =   255
      Left            =   4680
      TabIndex        =   179
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label lbl_isotopes 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Isotopes"
      Height          =   255
      Left            =   7080
      TabIndex        =   178
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label lbl_density 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Density"
      Height          =   255
      Left            =   7200
      TabIndex        =   177
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label lbl_discovery 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Discovery"
      Height          =   255
      Left            =   6960
      TabIndex        =   176
      Top             =   7920
      Width           =   855
   End
   Begin VB.Label lbl_melt_boil 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Melting/Boiling Points"
      Height          =   255
      Left            =   6120
      TabIndex        =   175
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label lbl_orbitals 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Orbitals"
      Height          =   255
      Left            =   4680
      TabIndex        =   174
      Top             =   7440
      Width           =   735
   End
   Begin VB.Label lbl_elec_shell 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Electron Shells"
      Height          =   255
      Left            =   4680
      TabIndex        =   173
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label lbl_atom_mass 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Atomic Mass"
      Height          =   255
      Left            =   4680
      TabIndex        =   172
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label lbl_name 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Name"
      Height          =   255
      Left            =   4680
      TabIndex        =   171
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label lbl_atom_num 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Atomic Number"
      Height          =   255
      Left            =   4680
      TabIndex        =   170
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label lbl_noble_gases 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "noble gases"
      Height          =   255
      Left            =   4920
      TabIndex        =   158
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lbl_other_metals 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "other metals"
      Height          =   255
      Left            =   4920
      TabIndex        =   157
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lbl_actinide_series 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "actinide series"
      Height          =   255
      Left            =   3360
      TabIndex        =   156
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lbl_lanthanide_series 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lanthanide series"
      Height          =   255
      Left            =   3360
      TabIndex        =   155
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lbl_transition_metals 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "transition metals"
      Height          =   255
      Left            =   4920
      TabIndex        =   154
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lbl_alkaline_earth_metals 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "alkaline earth metals"
      Height          =   255
      Left            =   3360
      TabIndex        =   153
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lbl_alkali_metals 
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "alkali metals"
      Height          =   255
      Left            =   3360
      TabIndex        =   152
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lbl_nonmetals 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "nonmetals"
      Height          =   255
      Left            =   4920
      TabIndex        =   151
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lbl_18 
      Alignment       =   2  'Center
      Caption         =   "18"
      Height          =   255
      Left            =   10920
      TabIndex        =   150
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lbl_17 
      Alignment       =   2  'Center
      Caption         =   "17"
      Height          =   255
      Left            =   10320
      TabIndex        =   149
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lbl_16 
      Alignment       =   2  'Center
      Caption         =   "16"
      Height          =   255
      Left            =   9720
      TabIndex        =   148
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lbl_15 
      Alignment       =   2  'Center
      Caption         =   "15"
      Height          =   255
      Left            =   9120
      TabIndex        =   147
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lbl_14 
      Alignment       =   2  'Center
      Caption         =   "14"
      Height          =   255
      Left            =   8520
      TabIndex        =   146
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lbl_13 
      Alignment       =   2  'Center
      Caption         =   "13"
      Height          =   255
      Left            =   7920
      TabIndex        =   145
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lbl_12 
      Alignment       =   2  'Center
      Caption         =   "12"
      Height          =   255
      Left            =   7320
      TabIndex        =   144
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl_11 
      Alignment       =   2  'Center
      Caption         =   "11"
      Height          =   255
      Left            =   6720
      TabIndex        =   143
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl_10 
      Alignment       =   2  'Center
      Caption         =   "10"
      Height          =   255
      Left            =   6120
      TabIndex        =   142
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl_9 
      Alignment       =   2  'Center
      Caption         =   "9"
      Height          =   255
      Left            =   5520
      TabIndex        =   141
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl_8 
      Alignment       =   2  'Center
      Caption         =   "8"
      Height          =   255
      Left            =   4920
      TabIndex        =   140
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl_7 
      Alignment       =   2  'Center
      Caption         =   "7"
      Height          =   255
      Left            =   4320
      TabIndex        =   139
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl_6 
      Alignment       =   2  'Center
      Caption         =   "6"
      Height          =   255
      Left            =   3720
      TabIndex        =   138
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl_5 
      Alignment       =   2  'Center
      Caption         =   "5"
      Height          =   255
      Left            =   3120
      TabIndex        =   137
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl_4 
      Alignment       =   2  'Center
      Caption         =   "4"
      Height          =   255
      Left            =   2520
      TabIndex        =   136
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl_3 
      Alignment       =   2  'Center
      Caption         =   "3"
      Height          =   255
      Left            =   1920
      TabIndex        =   135
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl_2 
      Alignment       =   2  'Center
      Caption         =   "2"
      Height          =   255
      Left            =   840
      TabIndex        =   134
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lbl_1 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Left            =   240
      TabIndex        =   133
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lbl_0 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   10920
      TabIndex        =   132
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lbl_VIIa 
      Alignment       =   2  'Center
      Caption         =   "VIIa"
      Height          =   255
      Left            =   10320
      TabIndex        =   131
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lbl_VIa 
      Alignment       =   2  'Center
      Caption         =   "VIa"
      Height          =   255
      Left            =   9720
      TabIndex        =   130
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lbl_Va 
      Alignment       =   2  'Center
      Caption         =   "Va"
      Height          =   255
      Left            =   9120
      TabIndex        =   129
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lbl_IVa 
      Alignment       =   2  'Center
      Caption         =   "IVa"
      Height          =   255
      Left            =   8520
      TabIndex        =   128
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lbl_IIIa 
      Alignment       =   2  'Center
      Caption         =   "IIIa"
      Height          =   255
      Left            =   7920
      TabIndex        =   127
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lbl_IIb 
      Alignment       =   2  'Center
      Caption         =   "IIb"
      Height          =   255
      Left            =   7320
      TabIndex        =   126
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lbl_Ib 
      Alignment       =   2  'Center
      Caption         =   "Ib"
      Height          =   255
      Left            =   6720
      TabIndex        =   125
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lbl_right 
      Alignment       =   2  'Center
      Caption         =   "------------|"
      Height          =   255
      Left            =   6120
      TabIndex        =   124
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lbl_VIIIb 
      Alignment       =   2  'Center
      Caption         =   "VIIIb"
      Height          =   255
      Left            =   5520
      TabIndex        =   123
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lbl_left 
      Alignment       =   2  'Center
      Caption         =   "|-------------"
      Height          =   255
      Left            =   4920
      TabIndex        =   122
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lbl_VIIb 
      Alignment       =   2  'Center
      Caption         =   "VIIb"
      Height          =   255
      Left            =   4320
      TabIndex        =   121
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lbl_VIb 
      Alignment       =   2  'Center
      Caption         =   "VIb"
      Height          =   255
      Left            =   3720
      TabIndex        =   120
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lbl_Vb 
      Alignment       =   2  'Center
      Caption         =   "Vb"
      Height          =   255
      Left            =   3120
      TabIndex        =   119
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lbl_IVb 
      Alignment       =   2  'Center
      Caption         =   "IVb"
      Height          =   255
      Left            =   2520
      TabIndex        =   118
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lbl_IIIb 
      Alignment       =   2  'Center
      Caption         =   "IIIb"
      Height          =   255
      Left            =   1920
      TabIndex        =   117
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lbl_IIa 
      Alignment       =   2  'Center
      Caption         =   "IIa"
      Height          =   255
      Left            =   840
      TabIndex        =   116
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lbl_Ia 
      Alignment       =   2  'Center
      Caption         =   "Ia"
      Height          =   255
      Left            =   240
      TabIndex        =   115
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lbl_Uuo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Uuo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      TabIndex        =   114
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lbl_Rn 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      TabIndex        =   113
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lbl_Xe 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Xe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      TabIndex        =   112
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lbl_Kr 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      TabIndex        =   111
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl_Ar 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      TabIndex        =   110
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lbl_Ne 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ne"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      TabIndex        =   109
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lbl_He 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "He"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      TabIndex        =   108
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lbl_At 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "At"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   107
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lbl_I 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   106
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lbl_Te 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Te"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   105
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lbl_Br 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Br"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   104
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl_Se 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Se"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   103
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl_As 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "As"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   102
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl_Cl 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cl"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   101
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lbl_S 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   100
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lbl_P 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   99
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lbl_Si 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Si"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   98
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lbl_F 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   97
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lbl_O 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   96
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lbl_N 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   95
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lbl_C 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   94
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lbl_B 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   93
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lbl_Uuh 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Uuh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   92
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lbl_Uuq 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Uuq"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   91
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lbl_Po 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Po"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   90
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lbl_Bi 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   89
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lbl_Pb 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pb"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   88
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lbl_Tl 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tl"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   87
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lbl_Sb 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sb"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   86
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lbl_Sn 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   85
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lbl_In 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "In"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   84
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lbl_Ge 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ge"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   83
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl_Ga 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ga"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   82
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl_Al 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Al"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   81
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lbl_Uub 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Uub"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   80
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lbl_Uuu 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Uuu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   79
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lbl_Uun 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Uun"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   78
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lbl_Mt 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   77
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lbl_Hs 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   76
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lbl_Bh 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   75
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lbl_Sg 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   74
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lbl_Db 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Db"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   73
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lbl_Rf 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rf"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   72
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lbl_Lr 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   71
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lbl_Hg 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   70
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lbl_Au 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Au"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   69
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lbl_Pt 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   68
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lbl_Ir 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   67
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lbl_Os 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Os"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   66
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lbl_Re 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Re"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   65
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lbl_W 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   64
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lbl_Ta 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   63
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lbl_Hf 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hf"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   62
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lbl_Lu 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   61
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lbl_Cd 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cd"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   60
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lbl_Ag 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ag"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   59
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lbl_Pd 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pd"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   58
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lbl_Rh 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   57
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lbl_Ru 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ru"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   56
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lbl_Tc 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   55
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lbl_Mo 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   54
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lbl_Nb 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nb"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   53
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lbl_Zr 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Zr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   52
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lbl_Y 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   51
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lbl_Zn 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Zn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   50
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl_Cu 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   49
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl_Ni 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ni"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   48
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl_Co 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Co"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   47
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl_Fe 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   46
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl_Mn 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   45
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl_Cr 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   44
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl_V 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   43
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl_Ti 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ti"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   42
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl_Sc 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   41
      Top             =   2160
      Width           =   615
   End
   Begin VB.Line Line6 
      X1              =   1560
      X2              =   2520
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line5 
      X1              =   1560
      X2              =   1560
      Y1              =   3840
      Y2              =   5040
   End
   Begin VB.Line Line4 
      X1              =   1440
      X2              =   1560
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line3 
      X1              =   1800
      X2              =   2520
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line2 
      X1              =   1800
      X2              =   1800
      Y1              =   3360
      Y2              =   4560
   End
   Begin VB.Line Line1 
      X1              =   1440
      X2              =   1800
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label lbl_No 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   40
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lbl_Md 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Md"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   39
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lbl_Fm 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   38
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lbl_Es 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Es"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   37
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lbl_Cf 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cf"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   36
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lbl_Bk 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bk"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   35
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lbl_Cm 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   34
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lbl_Am 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Am"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   33
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lbl_Pu 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   32
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lbl_Np 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Np"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   31
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lbl_U 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   30
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lbl_Pa 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   29
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lbl_Th 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Th"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   28
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lbl_Ac 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ac"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   27
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lbl_Yb 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Yb"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   26
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lbl_Tm 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   25
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lbl_Er 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Er"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   24
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lbl_Ho 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ho"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   23
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lbl_Dy 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   22
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lbl_Tb 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tb"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   21
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lbl_Gd 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Gd"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   20
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lbl_Eu 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Eu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   19
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lbl_Sm 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   18
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lbl_Pm 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   17
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lbl_Nd 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nd"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   16
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lbl_Pr 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   15
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lbl_Ce 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ce"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   14
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lbl_La 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "La"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   13
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lbl_Ra 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ra"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   12
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lbl_Ba 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ba"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   11
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lbl_Sr 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   10
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lbl_Ca 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ca"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   9
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl_Mg 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   8
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lbl_Be 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Be"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lbl_Fr 
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lbl_Cs 
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lbl_Rb 
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rb"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lbl_K 
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl_Na 
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Na"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lbl_Li 
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Li"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lbl_H 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "form_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public db As Database, rs As Recordset, FirstRow As Variant

Private Sub Form_Load()
  On Error Resume Next
  form_main.Height = 5790
  Set db = OpenDatabase((App.Path & "\ptoe.mdb"), , True)
  With db
    Set rs = .OpenRecordset("data_list")
      With rs
        .Index = "ID"
        FirstRow = .Bookmark
      End With
  End With
End Sub

Private Sub lbl_Ac_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 88
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Ag_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 46
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Al_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 12
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Am_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 94
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Ar_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 17
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_As_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 32
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_At_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 84
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Au_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 78
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_B_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 4
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Ba_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 55
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Be_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 3
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Bh_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 106
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Bi_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 82
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Bk_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 96
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Br_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 34
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_C_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 5
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Ca_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 19
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Cd_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 47
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Ce_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 57
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Cf_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 97
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Cl_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 16
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Cm_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 95
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Co_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 26
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Cr_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 23
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Cs_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 54
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Cu_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 28
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Db_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 104
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Dy_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 65
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Er_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 67
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Es_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 98
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Eu_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 62
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_F_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 8
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Fe_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 25
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Fm_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 99
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Fr_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 86
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Ga_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 30
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Gd_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 63
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Ge_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 31
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_H_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 0
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_He_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 1
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Hf_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 71
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Hg_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 79
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_hide_Click()
  form_main.Height = 5790
End Sub

Private Sub lbl_Ho_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 66
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Hs_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 107
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_I_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 52
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_In_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 48
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Ir_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 76
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_K_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 18
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Kr_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 35
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_La_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 56
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Li_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 2
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Lr_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 102
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Lu_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 70
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Md_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 100
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Mg_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 11
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Mn_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 24
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Mo_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 41
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Mt_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 108
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_N_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 6
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Na_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 10
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Nb_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 40
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Nd_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 59
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Ne_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 9
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Ni_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 27
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_No_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 101
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Np_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 92
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_O_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 7
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Os_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 75
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_P_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 14
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Pa_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 90
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Pb_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 81
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Pd_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 45
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Pm_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 60
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Po_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 83
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Pr_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 58
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Pt_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 77
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Pu_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 93
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Ra_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 87
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Rb_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 36
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Re_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 74
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Rf_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 103
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Rh_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 44
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Rn_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 85
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Ru_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 43
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_S_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 15
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Sb_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 50
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Sc_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 20
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Se_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 33
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Sg_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 105
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Si_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 13
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Sm_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 61
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Sn_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 49
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Sr_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 37
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Ta_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 72
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Tb_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 64
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Tc_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 42
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Te_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 51
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Th_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 89
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Ti_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 21
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Tl_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 80
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Tm_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 68
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_U_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 91
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Uub_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 112
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Uuh_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 0
    txt_atom_num.Text = "?"
    txt_name.Text = "?"
    txt_atom_mass.Text = "?"
    txt_atom_rad.Text = "?"
    txt_orbitals.Text = "?"
    txt_elec_shell.Text = "?"
    txt_melt_boil.Text = "?"
    txt_density.Text = "?"
    txt_isotopes.Text = "?"
    txt_electronegativity = "?"
    txt_oxidation.Text = "?"
    txt_discovery.Text = "?"
  End With
End Sub

Private Sub lbl_Uun_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 110
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Uuo_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 0
    txt_atom_num.Text = "?"
    txt_name.Text = "?"
    txt_atom_mass.Text = "?"
    txt_atom_rad.Text = "?"
    txt_orbitals.Text = "?"
    txt_elec_shell.Text = "?"
    txt_melt_boil.Text = "?"
    txt_density.Text = "?"
    txt_isotopes.Text = "?"
    txt_electronegativity = "?"
    txt_oxidation.Text = "?"
    txt_discovery.Text = "?"
  End With
End Sub

Private Sub lbl_Uuq_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 113
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Uuu_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 111
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_V_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 22
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_W_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 73
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Xe_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 53
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Y_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 38
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Yb_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 69
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Zn_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 29
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub

Private Sub lbl_Zr_Click()
  form_main.Height = 9090
  With rs
    .Bookmark = FirstRow
    .Move 39
    txt_atom_num.Text = !atom_num
    txt_name.Text = !Name
    txt_atom_mass.Text = !atom_mass
    txt_atom_rad.Text = !atom_rad
    txt_orbitals.Text = !orbitals
    txt_elec_shell.Text = !elec_shell
    txt_melt_boil.Text = !melt_boil
    txt_density.Text = !density
    txt_isotopes.Text = !isotopes
    txt_electronegativity = !electronegativity
    txt_oxidation.Text = !oxidation
    txt_discovery.Text = !discovery
  End With
End Sub
