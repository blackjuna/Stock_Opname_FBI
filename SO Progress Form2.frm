VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form Progress SO"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   20160
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   20160
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh Data"
      Height          =   495
      Left            =   10440
      TabIndex        =   61
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print Laporan"
      Height          =   495
      Left            =   10440
      TabIndex        =   60
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   10440
      Top             =   8280
   End
   Begin VB.Label ltittl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6225
      TabIndex        =   101
      Top             =   8400
      Width           =   1740
   End
   Begin VB.Label lcttl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8385
      TabIndex        =   100
      Top             =   8400
      Width           =   1545
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Total All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2280
      TabIndex        =   99
      Top             =   8400
      Width           =   855
   End
   Begin VB.Label ltrttl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4065
      TabIndex        =   98
      Top             =   8400
      Width           =   1740
   End
   Begin VB.Label lc_undefined 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15960
      TabIndex        =   97
      Top             =   4920
      Width           =   1740
   End
   Begin VB.Label lc_tembok 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15960
      TabIndex        =   96
      Top             =   4440
      Width           =   1740
   End
   Begin VB.Label lc_lantai2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15960
      TabIndex        =   95
      Top             =   3960
      Width           =   1740
   End
   Begin VB.Label lti_undefined 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14040
      TabIndex        =   94
      Top             =   4920
      Width           =   1740
   End
   Begin VB.Label lti_tembok 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14040
      TabIndex        =   93
      Top             =   4440
      Width           =   1740
   End
   Begin VB.Label lti_lantai2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14040
      TabIndex        =   92
      Top             =   3960
      Width           =   1740
   End
   Begin VB.Label ltr_undefined 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12120
      TabIndex        =   91
      Top             =   4920
      Width           =   1740
   End
   Begin VB.Label ltr_tembok 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12120
      TabIndex        =   90
      Top             =   4440
      Width           =   1740
   End
   Begin VB.Label ltr_lantai2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12120
      TabIndex        =   89
      Top             =   3960
      Width           =   1740
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "UNDEFINED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10440
      TabIndex        =   88
      Top             =   4920
      Width           =   1425
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "TEMBOK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10440
      TabIndex        =   87
      Top             =   4440
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "LANTAI 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10440
      TabIndex        =   86
      Top             =   3960
      Width           =   1035
   End
   Begin VB.Label lctb 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8280
      TabIndex        =   85
      Top             =   6120
      Width           =   1740
   End
   Begin VB.Label ltitb 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   84
      Top             =   6120
      Width           =   1740
   End
   Begin VB.Label ltrtb 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   83
      Top             =   6120
      Width           =   1740
   End
   Begin VB.Label lci 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8280
      TabIndex        =   82
      Top             =   5160
      Width           =   1740
   End
   Begin VB.Label ltii 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   81
      Top             =   5160
      Width           =   1740
   End
   Begin VB.Label ltri 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   80
      Top             =   5160
      Width           =   1740
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "AREA I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   79
      Top             =   5160
      Width           =   810
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Tag Blank"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   78
      Top             =   6120
      Width           =   1065
   End
   Begin VB.Label ltr_5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12120
      TabIndex        =   77
      Top             =   3000
      Width           =   1740
   End
   Begin VB.Label ltr_karantina 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12120
      TabIndex        =   76
      Top             =   3480
      Width           =   1740
   End
   Begin VB.Label lti_5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14040
      TabIndex        =   75
      Top             =   3000
      Width           =   1740
   End
   Begin VB.Label lti_karantina 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14040
      TabIndex        =   74
      Top             =   3480
      Width           =   1740
   End
   Begin VB.Label lc_5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15960
      TabIndex        =   73
      Top             =   3000
      Width           =   1740
   End
   Begin VB.Label lc_karantina 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15960
      TabIndex        =   72
      Top             =   3480
      Width           =   1740
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Area V"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10440
      TabIndex        =   71
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "KARANTINA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10440
      TabIndex        =   70
      Top             =   3480
      Width           =   1365
   End
   Begin VB.Label lch 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8400
      TabIndex        =   69
      Top             =   4680
      Width           =   1500
   End
   Begin VB.Label lcg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8400
      TabIndex        =   68
      Top             =   4200
      Width           =   1500
   End
   Begin VB.Label ltih 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   67
      Top             =   4680
      Width           =   1740
   End
   Begin VB.Label ltig 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   66
      Top             =   4200
      Width           =   1740
   End
   Begin VB.Label ltrh 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   65
      Top             =   4680
      Width           =   1740
   End
   Begin VB.Label ltrg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   64
      Top             =   4200
      Width           =   1740
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "AREA H"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   63
      Top             =   4680
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "AREA G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   62
      Top             =   4200
      Width           =   930
   End
   Begin VB.Label lbl_jam 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12000
      TabIndex        =   59
      Top             =   7440
      Width           =   2010
   End
   Begin VB.Label lbl_tanggal 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12000
      TabIndex        =   58
      Top             =   6840
      Width           =   2010
   End
   Begin VB.Label ltr_total_fbi2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12120
      TabIndex        =   57
      Top             =   6000
      Width           =   1740
   End
   Begin VB.Label lti_total_fbi2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14040
      TabIndex        =   56
      Top             =   6000
      Width           =   1740
   End
   Begin VB.Label lc_total_fbi2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   16080
      TabIndex        =   55
      Top             =   6000
      Width           =   1545
   End
   Begin VB.Line Line9 
      X1              =   10320
      X2              =   17760
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label Label111 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10440
      TabIndex        =   54
      Top             =   6000
      Width           =   1485
   End
   Begin VB.Line Line8 
      X1              =   10320
      X2              =   17760
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label Label110 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Area"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10560
      TabIndex        =   53
      Top             =   600
      Width           =   1305
   End
   Begin VB.Label Label109 
      AutoSize        =   -1  'True
      Caption         =   "% Complete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   15960
      TabIndex        =   52
      Top             =   600
      Width           =   1755
   End
   Begin VB.Label Label108 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Tag Input"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   14040
      TabIndex        =   51
      Top             =   600
      Width           =   1785
   End
   Begin VB.Label Label107 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Tag Released"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12120
      TabIndex        =   50
      Top             =   600
      Width           =   1770
   End
   Begin VB.Label Label106 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "PT. FAJAR BENUA INDONESIA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12510
      TabIndex        =   49
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label Label105 
      AutoSize        =   -1  'True
      Caption         =   "Area Sheet I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10440
      TabIndex        =   48
      Top             =   1080
      Width           =   1350
   End
   Begin VB.Label Label103 
      AutoSize        =   -1  'True
      Caption         =   "Area Sheet III"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10440
      TabIndex        =   47
      Top             =   2040
      Width           =   1500
   End
   Begin VB.Label Label102 
      AutoSize        =   -1  'True
      Caption         =   "Area Sheet V"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10440
      TabIndex        =   46
      Top             =   2520
      Width           =   1440
   End
   Begin VB.Line Line7 
      X1              =   10320
      X2              =   10320
      Y1              =   480
      Y2              =   6480
   End
   Begin VB.Line Line6 
      X1              =   10320
      X2              =   17760
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line5 
      X1              =   12000
      X2              =   12000
      Y1              =   480
      Y2              =   6480
   End
   Begin VB.Line Line4 
      X1              =   10320
      X2              =   17760
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line3 
      X1              =   13920
      X2              =   13920
      Y1              =   480
      Y2              =   6480
   End
   Begin VB.Label ltr_as5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12120
      TabIndex        =   45
      Top             =   2520
      Width           =   1740
   End
   Begin VB.Label ltr_as3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12120
      TabIndex        =   44
      Top             =   2040
      Width           =   1740
   End
   Begin VB.Label ltr_as1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12120
      TabIndex        =   43
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Label lti_as5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14040
      TabIndex        =   42
      Top             =   2520
      Width           =   1740
   End
   Begin VB.Label lti_as3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14040
      TabIndex        =   41
      Top             =   2040
      Width           =   1740
   End
   Begin VB.Label lti_as1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14040
      TabIndex        =   40
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Line Line2 
      X1              =   15840
      X2              =   15840
      Y1              =   480
      Y2              =   6480
   End
   Begin VB.Line Line1 
      X1              =   17760
      X2              =   17760
      Y1              =   480
      Y2              =   6480
   End
   Begin VB.Label lc_as5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15960
      TabIndex        =   39
      Top             =   2520
      Width           =   1740
   End
   Begin VB.Label lc_as3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15960
      TabIndex        =   38
      Top             =   2040
      Width           =   1740
   End
   Begin VB.Label lc_as1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15960
      TabIndex        =   37
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Label ltr_total_fbi 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   36
      Top             =   7440
      Width           =   1740
   End
   Begin VB.Label lti_total_fbi 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   35
      Top             =   7440
      Width           =   1740
   End
   Begin VB.Label lc_total_fbi 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8400
      TabIndex        =   34
      Top             =   7440
      Width           =   1545
   End
   Begin VB.Line Line18 
      X1              =   1680
      X2              =   10080
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Label Label74 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   33
      Top             =   7440
      Width           =   1845
   End
   Begin VB.Line Line17 
      X1              =   1680
      X2              =   10080
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Label Label73 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Area"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      TabIndex        =   32
      Top             =   600
      Width           =   1785
   End
   Begin VB.Label Label72 
      AutoSize        =   -1  'True
      Caption         =   "% Complete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8400
      TabIndex        =   31
      Top             =   600
      Width           =   1515
   End
   Begin VB.Label Label71 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Tag Input"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6240
      TabIndex        =   30
      Top             =   600
      Width           =   1785
   End
   Begin VB.Label Label70 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Tag Released"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      TabIndex        =   29
      Top             =   600
      Width           =   1770
   End
   Begin VB.Label Label69 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "PT. FAJAR BENUA INDONESIA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3870
      TabIndex        =   28
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label Label68 
      AutoSize        =   -1  'True
      Caption         =   "AREA A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   27
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label65 
      AutoSize        =   -1  'True
      Caption         =   "AREA B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   26
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label Label64 
      AutoSize        =   -1  'True
      Caption         =   "Area Sheet II"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10440
      TabIndex        =   25
      Top             =   1560
      Width           =   1425
   End
   Begin VB.Line Line16 
      X1              =   1680
      X2              =   1680
      Y1              =   480
      Y2              =   7920
   End
   Begin VB.Label Label61 
      AutoSize        =   -1  'True
      Caption         =   "AREA C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   24
      Top             =   2040
      Width           =   900
   End
   Begin VB.Label Label60 
      AutoSize        =   -1  'True
      Caption         =   "AREA D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   23
      Top             =   2520
      Width           =   915
   End
   Begin VB.Label Label59 
      Caption         =   "AREA E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1800
      TabIndex        =   22
      Top             =   3000
      Width           =   1560
   End
   Begin VB.Label Label57 
      AutoSize        =   -1  'True
      Caption         =   "AREA F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   21
      Top             =   3720
      Width           =   885
   End
   Begin VB.Line Line15 
      X1              =   1680
      X2              =   10080
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line14 
      X1              =   3840
      X2              =   3840
      Y1              =   480
      Y2              =   7920
   End
   Begin VB.Line Line13 
      X1              =   1680
      X2              =   10080
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line12 
      X1              =   6000
      X2              =   6000
      Y1              =   480
      Y2              =   7920
   End
   Begin VB.Label ltre 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   20
      Top             =   3000
      Width           =   1740
   End
   Begin VB.Label ltrd 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   19
      Top             =   2520
      Width           =   1740
   End
   Begin VB.Label ltrc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   18
      Top             =   2040
      Width           =   1740
   End
   Begin VB.Label ltr_as2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12120
      TabIndex        =   17
      Top             =   1560
      Width           =   1740
   End
   Begin VB.Label ltrb 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   16
      Top             =   1560
      Width           =   1740
   End
   Begin VB.Label ltra 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   15
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Label ltrf 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   14
      Top             =   3720
      Width           =   1740
   End
   Begin VB.Label ltie 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   13
      Top             =   3000
      Width           =   1740
   End
   Begin VB.Label ltid 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   12
      Top             =   2520
      Width           =   1740
   End
   Begin VB.Label ltic 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   11
      Top             =   2040
      Width           =   1740
   End
   Begin VB.Label lti_as2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14040
      TabIndex        =   10
      Top             =   1560
      Width           =   1740
   End
   Begin VB.Label ltib 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   9
      Top             =   1560
      Width           =   1740
   End
   Begin VB.Label ltia 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   8
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Label ltif 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   7
      Top             =   3720
      Width           =   1740
   End
   Begin VB.Line Line11 
      X1              =   8160
      X2              =   8160
      Y1              =   480
      Y2              =   7920
   End
   Begin VB.Line Line10 
      X1              =   10080
      X2              =   10080
      Y1              =   480
      Y2              =   7920
   End
   Begin VB.Label lce 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8280
      TabIndex        =   6
      Top             =   3000
      Width           =   1740
   End
   Begin VB.Label lcd 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8280
      TabIndex        =   5
      Top             =   2520
      Width           =   1740
   End
   Begin VB.Label lcc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8280
      TabIndex        =   4
      Top             =   2040
      Width           =   1740
   End
   Begin VB.Label lc_as2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15960
      TabIndex        =   3
      Top             =   1560
      Width           =   1740
   End
   Begin VB.Label lcb 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8280
      TabIndex        =   2
      Top             =   1560
      Width           =   1740
   End
   Begin VB.Label lca 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8280
      TabIndex        =   1
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Label lcf 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8400
      TabIndex        =   0
      Top             =   3720
      Width           =   1500
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub current_progress_jfi()
    'Gudang Area Sheet 1
    strsql = "select count(location) AS lti from tag_stock_opname_fbi where location = 'Area Sheet I' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_as1.Caption = Val(rs_so!lti)
    strsql = "select count(location) AS ltr from tag_stock_opname_fbi where location = 'Area Sheet I'"
    Set rs_so = conn.Execute(strsql)
    ltr_as1.Caption = Val(rs_so!ltr)
    If Val(rs_so!ltr) = 0 Then lc_as1.Caption = 0 Else _
        lc_as1.Caption = Round((Val(lti_as1.Caption) / Val(ltr_as1.Caption) * 100), 2)
    
    'Gudang Area Sheet II
    strsql = "select count(location) AS lti from tag_stock_opname_fbi where location = 'Area Sheet II' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_as2.Caption = Val(rs_so!lti)
    strsql = "select count(location) AS ltr from tag_stock_opname_fbi where location = 'Area Sheet II'"
    Set rs_so = conn.Execute(strsql)
    ltr_as2.Caption = Val(rs_so!ltr)
    If Val(rs_so!ltr) = 0 Then lc_as2.Caption = 0 Else _
        lc_as2.Caption = Round((Val(lti_as2.Caption) / Val(ltr_as2.Caption) * 100), 2)
    
    'Gudang Area Sheet III
    strsql = "select count(location) AS lti from tag_stock_opname_fbi where location = 'Area Sheet III' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_as3.Caption = Val(rs_so!lti)
    strsql = "select count(location) AS ltr from tag_stock_opname_fbi where location = 'Area Sheet III'"
    Set rs_so = conn.Execute(strsql)
    ltr_as3.Caption = Val(rs_so!ltr)
    If Val(rs_so!ltr) = 0 Then lc_as3.Caption = 0 Else _
        lc_as3.Caption = Round((Val(lti_as3.Caption) / Val(ltr_as3.Caption) * 100), 2)
        
    'Gudang Area Sheet V
    strsql = "select count(location) AS lti from tag_stock_opname_fbi where location = 'Area Sheet V' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_as5.Caption = Val(rs_so!lti)
    strsql = "select count(location) AS ltr from tag_stock_opname_fbi where location = 'Area Sheet V'"
    Set rs_so = conn.Execute(strsql)
    ltr_as5.Caption = Val(rs_so!ltr)
    If Val(rs_so!ltr) = 0 Then lc_as5.Caption = 0 Else _
        lc_as5.Caption = Round((Val(lti_as5.Caption) / Val(ltr_as5.Caption) * 100), 2)
        
    'Gudang Area V
    strsql = "select count(location) AS lti from tag_stock_opname_fbi where location = 'Area V' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_5.Caption = Val(rs_so!lti)
    strsql = "select count(location) AS ltr from tag_stock_opname_fbi where location = 'Area V'"
    Set rs_so = conn.Execute(strsql)
    ltr_5.Caption = Val(rs_so!ltr)
    If Val(rs_so!ltr) = 0 Then lc_5.Caption = 0 Else _
        lc_5.Caption = Round((Val(lti_5.Caption) / Val(ltr_5.Caption) * 100), 2)
    
    'Gudang Area KARANTINA
    strsql = "select count(location) AS lti from tag_stock_opname_fbi where location = 'KARANTINA' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_karantina.Caption = Val(rs_so!lti)
    strsql = "select count(location) AS ltr from tag_stock_opname_fbi where location = 'KARANTINA'"
    Set rs_so = conn.Execute(strsql)
    ltr_karantina.Caption = Val(rs_so!ltr)
    If Val(rs_so!ltr) = 0 Then lc_karantina.Caption = 0 Else _
        lc_karantina.Caption = Round((Val(lti_karantina.Caption) / Val(ltr_karantina.Caption) * 100), 2)
        
    'Gudang Area LANTAI 2
    strsql = "select count(location) AS lti from tag_stock_opname_fbi where location = 'LANTAI 2' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_lantai2.Caption = Val(rs_so!lti)
    strsql = "select count(location) AS ltr from tag_stock_opname_fbi where location = 'LANTAI 2'"
    Set rs_so = conn.Execute(strsql)
    ltr_lantai2.Caption = Val(rs_so!ltr)
    If Val(rs_so!ltr) = 0 Then lti_lantai2.Caption = 0 Else _
        lc_lantai2.Caption = Round((Val(lti_lantai2.Caption) / Val(ltr_lantai2.Caption) * 100), 2)
        
    'Gudang Area TEMBOK
    strsql = "select count(location) AS lti from tag_stock_opname_fbi where location = 'TEMBOK' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_tembok.Caption = Val(rs_so!lti)
    strsql = "select count(location) AS ltr from tag_stock_opname_fbi where location = 'TEMBOK'"
    Set rs_so = conn.Execute(strsql)
    ltr_tembok.Caption = Val(rs_so!ltr)
    If Val(rs_so!ltr) = 0 Then lc_tembok.Caption = 0 Else _
        lc_tembok.Caption = Round((Val(lti_tembok.Caption) / Val(ltr_tembok.Caption) * 100), 2)
        
    'Gudang Area UNDEFINED
    strsql = "select count(location) AS lti from tag_stock_opname_fbi where location = '' AND left(tag_no,2)<>'TB' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_undefined.Caption = Val(rs_so!lti)
    strsql = "select count(location) AS ltr from tag_stock_opname_fbi where location = '' AND left(tag_no,2)<>'TB'"
    Set rs_so = conn.Execute(strsql)
    ltr_undefined.Caption = Val(rs_so!ltr)
    If Val(rs_so!ltr) = 0 Then lti_undefined.Caption = 0 Else _
        lc_undefined.Caption = Round((Val(lti_undefined.Caption) / Val(ltr_undefined.Caption) * 100), 2)

    
    'Total
    ltr_total_fbi2.Caption = Val(ltr_as1.Caption) + Val(ltr_as2.Caption) + Val(ltr_as3.Caption) + Val(ltr_as5.Caption) _
        + Val(ltr_5.Caption) + Val(ltr_karantina.Caption) + Val(ltr_lantai2.Caption) + Val(ltr_tembok.Caption) + Val(ltr_undefined.Caption)
    lti_total_fbi2.Caption = Val(lti_as1.Caption) + Val(lti_as2.Caption) + Val(lti_as3.Caption) + Val(lti_as5.Caption) _
        + Val(lti_5.Caption) + Val(lti_karantina.Caption) + Val(lti_lantai2.Caption) + Val(lti_tembok.Caption) + Val(lti_undefined.Caption)
    If Val(lti_total_fbi2.Caption) = 0 Then lc_total_fbi2.Caption = 0 Else _
        lc_total_fbi2.Caption = Round(((Val(lti_total_fbi2) / Val(ltr_total_fbi2) * 100)), 2)
End Sub
Sub current_progress_tgs()
    'AREA A
    strsql = "select count(location) AS lti from tag_stock_opname_fbi where left(location,1)='A' and left(location,2) <> 'AR' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    ltia.Caption = Val(rs_so!lti)
    strsql = "select count(location) AS ltr from tag_stock_opname_fbi where left(location,1)='A' and left(location,2) <> 'AR'"
    Set rs_so = conn.Execute(strsql)
    ltra.Caption = Val(rs_so!ltr)
    If Val(rs_so!ltr) = 0 Then lca.Caption = 0 Else _
        lca.Caption = Round((Val(ltia.Caption) / Val(ltra.Caption) * 100), 2)
    
    'AREA B
    strsql = "select count(location) AS lti from tag_stock_opname_fbi where left(location,1)='B' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    ltib.Caption = Val(rs_so!lti)
    strsql = "select count(location) AS ltr from tag_stock_opname_fbi where left(location,1)='B'"
    Set rs_so = conn.Execute(strsql)
    ltrb.Caption = Val(rs_so!ltr)
    If Val(rs_so!ltr) = 0 Then lcb.Caption = 0 Else _
        lcb.Caption = Round((Val(ltib.Caption) / Val(ltrb.Caption) * 100), 2)
    
    'AREA C
    strsql = "select count(location) AS lti from tag_stock_opname_fbi where left(location,1)='C' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    ltic.Caption = Val(rs_so!lti)
    strsql = "select count(location) AS ltr from tag_stock_opname_fbi where left(location,1)='C'"
    Set rs_so = conn.Execute(strsql)
    ltrc.Caption = Val(rs_so!ltr)
    If Val(rs_so!ltr) = 0 Then lcc.Caption = 0 Else _
        lcc.Caption = Round((Val(ltic.Caption) / Val(ltrc.Caption) * 100), 2)
    
    'AREA D
    strsql = "select count(location) AS lti from tag_stock_opname_fbi where left(location,1)='D' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    ltid.Caption = Val(rs_so!lti)
    strsql = "select count(location) AS ltr from tag_stock_opname_fbi where left(location,1)='D'"
    Set rs_so = conn.Execute(strsql)
    ltrd.Caption = Val(rs_so!ltr)
    If Val(rs_so!ltr) = 0 Then lcd.Caption = 0 Else _
        lcd.Caption = Round((Val(ltid.Caption) / Val(ltrd.Caption) * 100), 2)
    
    'AREA E
    strsql = "select count(location) AS lti from tag_stock_opname_fbi where left(location,1)='E' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    ltie.Caption = Val(rs_so!lti)
    strsql = "select count(location) AS ltr from tag_stock_opname_fbi where left(location,1)='E'"
    Set rs_so = conn.Execute(strsql)
    ltre.Caption = Val(rs_so!ltr)
    If Val(rs_so!ltr) = 0 Then lce.Caption = 0 Else _
        lce.Caption = Round((Val(ltie.Caption) / Val(ltre.Caption) * 100), 2)
    
    'AREA F
    strsql = "select count(location) AS lti from tag_stock_opname_fbi where left(location,1)='F' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    ltif.Caption = Val(rs_so!lti)
    strsql = "select count(location) AS ltr from tag_stock_opname_fbi where left(location,1)='F'"
    Set rs_so = conn.Execute(strsql)
    ltrf.Caption = Val(rs_so!ltr)
    If Val(rs_so!ltr) = 0 Then lcf.Caption = 0 Else _
        lcf.Caption = Round((Val(ltif.Caption) / Val(ltrf.Caption) * 100), 2)
    
    'AREA G
    strsql = "select count(location) AS lti from tag_stock_opname_fbi where left(location,1)='G' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    ltig.Caption = Val(rs_so!lti)
    strsql = "select count(location) AS ltr from tag_stock_opname_fbi where left(location,1)='G'"
    Set rs_so = conn.Execute(strsql)
    ltrg.Caption = Val(rs_so!ltr)
    If Val(rs_so!ltr) = 0 Then lcg.Caption = 0 Else _
        lcg.Caption = Round((Val(ltig.Caption) / Val(ltrg.Caption) * 100), 2)
    
    'AREA H
    strsql = "select count(location) AS lti from tag_stock_opname_fbi where left(location,1)='H' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    ltih.Caption = Val(rs_so!lti)
    strsql = "select count(location) AS ltr from tag_stock_opname_fbi where left(location,1)='H'"
    Set rs_so = conn.Execute(strsql)
    ltrh.Caption = Val(rs_so!ltr)
    If Val(rs_so!ltr) = 0 Then lch.Caption = 0 Else _
        lch.Caption = Round((Val(ltih.Caption) / Val(ltrh.Caption) * 100), 2)
    
    'AREA I
    strsql = "select count(location) AS lti from tag_stock_opname_fbi where left(location,1)='I' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    ltii.Caption = Val(rs_so!lti)
    strsql = "select count(location) AS ltr from tag_stock_opname_fbi where left(location,1)='I'"
    Set rs_so = conn.Execute(strsql)
    ltri.Caption = Val(rs_so!ltr)
    If Val(rs_so!ltr) = 0 Then lci.Caption = 0 Else _
        lci.Caption = Round((Val(ltii.Caption) / Val(ltri.Caption) * 100), 2)
    
    'TAG BLANK
    strsql = "select count(tag_no) AS lti from tag_stock_opname_fbi where left(tag_no,2)='TB' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    ltitb.Caption = Val(rs_so!lti)
    strsql = "select count(tag_no) AS ltr from tag_stock_opname_fbi where left(tag_no,2)='TB'"
    Set rs_so = conn.Execute(strsql)
    ltrtb.Caption = Val(rs_so!ltr)
    If Val(rs_so!ltr) = 0 Then lctb.Caption = 0 Else _
        lctb.Caption = Round((Val(ltitb.Caption) / Val(ltrtb.Caption) * 100), 2)
    
    'Total
    ltr_total_fbi.Caption = Val(ltra.Caption) + Val(ltrb.Caption) + Val(ltrc.Caption) + Val(ltrd.Caption) + Val(ltre.Caption) + Val(ltrf.Caption) _
        + Val(ltrg.Caption) + Val(ltrh.Caption) + Val(ltri.Caption)
    'lti_total_tgs.Caption = Val(lti_ga.Caption) + Val(lti_gb.Caption) + Val(lti_gd.Caption) + Val(lti_gf.Caption) + Val(lti_pp.Caption) + Val(lti_mh.Caption) + Val(lti_gp.Caption) + Val(lti_ms.Caption) + Val(lti_ejm.Caption) + Val(lti_fh.Caption) + Val(lti_ejf.Caption)
    lti_total_fbi.Caption = Val(ltia.Caption) + Val(ltib.Caption) + Val(ltic.Caption) + Val(ltid.Caption) + Val(ltie.Caption) + Val(ltif.Caption) _
        + Val(ltig.Caption) + Val(ltih.Caption) + Val(ltii.Caption)
    If Val(lti_total_fbi.Caption) = 0 Then lc_total_fbi.Caption = 0 Else _
        lc_total_fbi.Caption = Round(((Val(lti_total_fbi) / Val(ltr_total_fbi) * 100)), 2)
    
End Sub


Private Sub Command1_Click()
    Printer.PaperSize = vbPRPSLegal
    Printer.Orientation = vbPRORLandscape
    Form2.PrintForm
End Sub

Private Sub Command2_Click()
Call current_progress_jfi
Call current_progress_tgs

End Sub

Private Sub Form_Load()
Timer1.Interval = 500
Timer1.Enabled = True
Call db
If rs_so.State = 1 Then rs_so.Close
rs_so.Open "Select * from tag_stock_opname_fbi", conn
Set rscompletion_slip = Nothing
Call current_progress_jfi
Call current_progress_tgs
    ltrttl.Caption = Val(ltr_total_fbi2.Caption) + Val(ltr_total_fbi.Caption)
    ltittl.Caption = Val(lti_total_fbi2.Caption) + Val(lti_total_fbi.Caption)
    lcttl.Caption = Round(((Val(ltittl.Caption) / Val(ltrttl.Caption)) * 100), 2)
lbl_tanggal = Date
End Sub

Private Sub Timer1_Timer()
lbl_jam.Caption = Format(Time, "hh:mm:ss")
End Sub
