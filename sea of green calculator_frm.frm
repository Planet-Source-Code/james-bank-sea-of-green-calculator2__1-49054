VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Sea of Green Calculator"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_cal 
      Caption         =   "Calculate"
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
      Left            =   1200
      TabIndex        =   22
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Frame Frame5 
      Caption         =   "Results"
      Height          =   1215
      Left            =   5160
      TabIndex        =   19
      Top             =   2160
      Width           =   4815
      Begin VB.TextBox result 
         Height          =   285
         Left            =   3600
         TabIndex        =   20
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Cost of lights per month"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   21
         Top             =   600
         Width           =   2295
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Flowering Room"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7680
      TabIndex        =   13
      Top             =   240
      Width           =   2295
      Begin VB.TextBox flow_watts 
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox flow_hours 
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label flow_watts_lbl 
         Caption         =   "watts of light"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label flow_hours_lbl 
         Caption         =   "hours per day"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Vegatative Room"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5160
      TabIndex        =   8
      Top             =   240
      Width           =   2295
      Begin VB.TextBox veg_hours 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox veg_watts 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.Label veg_hours_lbl 
         Caption         =   "hours per day"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label veg_watts_lbl 
         Caption         =   "watts of light"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Seedlings/Clones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   2295
      Begin VB.TextBox seed_watts 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox seed_hours 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label seed_watts_lbl 
         Caption         =   "watts of light"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label seed_hours_lbl 
         Caption         =   "hours per day"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cost of Electricty"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
      Begin VB.TextBox cost 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Example .09"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label cost_lbl 
         Caption         =   "cost per Kilowatt hour"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cal_Click()
result = ((seed_watts * seed_hours) + (veg_watts * veg_hours) + (flow_watts * flow_hours)) * 30 / 1000 * cost
End Sub

