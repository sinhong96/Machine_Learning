VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artificial Neural Networks - Backpropagation"
   ClientHeight    =   9405
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   15120
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0442
   ScaleHeight     =   9405
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLearningRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   43
      Text            =   "0.5"
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdRunData 
      Caption         =   "Run Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      TabIndex        =   41
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   12360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   40
      Text            =   "frmMain.frx":1D2234
      ToolTipText     =   "X1, X2, Target"
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox txtIterations 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   36
      Text            =   "1"
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdFP 
      Height          =   465
      Left            =   5280
      Picture         =   "frmMain.frx":1D2251
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   8400
      Width           =   2535
   End
   Begin VB.CommandButton cmdBP 
      Height          =   465
      Left            =   12720
      Picture         =   "frmMain.frx":1D2C8A
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   4800
      Width           =   2295
   End
   Begin VB.CommandButton cmdBiasH 
      Caption         =   "bias = 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   29
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdH2 
      Caption         =   ">"
      Height          =   435
      Left            =   8640
      TabIndex        =   26
      Top             =   6600
      Width           =   315
   End
   Begin VB.CommandButton cmdH1 
      Caption         =   ">"
      Height          =   435
      Left            =   8520
      TabIndex        =   25
      Top             =   2880
      Width           =   315
   End
   Begin VB.CommandButton cmdBias 
      Caption         =   ">"
      Height          =   555
      Left            =   3480
      TabIndex        =   24
      Top             =   7230
      Width           =   315
   End
   Begin VB.CommandButton cmdX2 
      Caption         =   ">"
      Height          =   555
      Left            =   3360
      TabIndex        =   23
      Top             =   4230
      Width           =   315
   End
   Begin VB.CommandButton cmdX1 
      Caption         =   ">"
      Height          =   555
      Left            =   3240
      TabIndex        =   22
      Top             =   1200
      Width           =   315
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   21
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtTarget 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   13320
      TabIndex        =   17
      Text            =   "1"
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox txtBias 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "1"
      Top             =   7230
      Width           =   1455
   End
   Begin VB.TextBox txtX2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   1920
      TabIndex        =   1
      Text            =   "1"
      Top             =   4230
      Width           =   1335
   End
   Begin VB.TextBox txtX1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   1800
      TabIndex        =   0
      Text            =   "1"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Image Image4 
      Height          =   225
      Left            =   9480
      Picture         =   "frmMain.frx":1D3579
      ToolTipText     =   "Transfer (Activation) Function"
      Top             =   8760
      Width           =   240
   End
   Begin VB.Image Image9 
      Height          =   1365
      Left            =   9840
      Picture         =   "frmMain.frx":1D3879
      Top             =   7920
      Width           =   1800
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Error Calculation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12600
      TabIndex        =   46
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Image Image8 
      Height          =   405
      Left            =   12240
      Picture         =   "frmMain.frx":1D5261
      Top             =   7920
      Width           =   2820
   End
   Begin VB.Image Image7 
      Height          =   420
      Left            =   12240
      Picture         =   "frmMain.frx":1D6953
      Top             =   7560
      Width           =   2640
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Weight Adjustment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12720
      TabIndex        =   45
      Top             =   8520
      Width           =   2055
   End
   Begin VB.Image Image6 
      Height          =   465
      Left            =   12720
      Picture         =   "frmMain.frx":1D7EAD
      ToolTipText     =   "LearningRate x Delta x Input"
      Top             =   8760
      Width           =   2025
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Learning Rate:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   44
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblMSE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MSE"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   10920
      TabIndex        =   42
      ToolTipText     =   "Mean Square Error"
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "An Introduction to Data Mining by Dr. Saed Sayad "
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   9000
      Width           =   4455
   End
   Begin VB.Label lblCounter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   9480
      TabIndex        =   38
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Iterations:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9480
      TabIndex        =   37
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblDelta2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delta2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   7320
      TabIndex        =   33
      ToolTipText     =   "Y2 * (1-Y2) * (W23 * Delta3)"
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label lblDelta1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delta1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   7200
      TabIndex        =   32
      ToolTipText     =   "Y1 * (1-Y1) * (W13 * Delta3)"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblDelta3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Delta3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   8520
      TabIndex        =   31
      ToolTipText     =   "Y * (1-Y) * (Target-Y)"
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label lblW33 
      BackStyle       =   0  'Transparent
      Caption         =   "W33"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11520
      TabIndex        =   30
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Summation && Activation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   28
      Top             =   7320
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Target"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13320
      TabIndex        =   27
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   7920
      Picture         =   "frmMain.frx":1D8F07
      ToolTipText     =   "Transfer (Activation) Function"
      Top             =   6120
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   11280
      Picture         =   "frmMain.frx":1D920B
      ToolTipText     =   "Transfer (Activation) Function"
      Top             =   4560
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   7800
      Picture         =   "frmMain.frx":1D950B
      ToolTipText     =   "Transfer (Activation) Function"
      Top             =   2400
      Width           =   225
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "bias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   20
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "X2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   19
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "X1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   18
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblY 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   13320
      TabIndex        =   16
      Top             =   4140
      Width           =   1455
   End
   Begin VB.Label lblSum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10560
      TabIndex        =   15
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lblW23 
      BackStyle       =   0  'Transparent
      Caption         =   "W23"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9720
      TabIndex        =   14
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label lblW13 
      BackStyle       =   0  'Transparent
      Caption         =   "W13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9600
      TabIndex        =   13
      Top             =   2580
      Width           =   1455
   End
   Begin VB.Label lblY2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7320
      TabIndex        =   12
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label lblSum2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sum2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7320
      TabIndex        =   11
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label lblY1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7320
      TabIndex        =   10
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblSum1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sum1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7200
      TabIndex        =   9
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblW32 
      BackStyle       =   0  'Transparent
      Caption         =   "W32"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4800
      TabIndex        =   8
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label lblW31 
      BackStyle       =   0  'Transparent
      Caption         =   "W31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4800
      TabIndex        =   7
      Top             =   6450
      Width           =   1455
   End
   Begin VB.Label lblW22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "W22"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3540
      TabIndex        =   6
      Top             =   5190
      Width           =   1335
   End
   Begin VB.Label lblW21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "W21"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3540
      TabIndex        =   5
      Top             =   3450
      Width           =   1395
   End
   Begin VB.Label lblW12 
      BackStyle       =   0  'Transparent
      Caption         =   "W12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4740
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblW11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "W11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4680
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Image Image5 
      Height          =   1080
      Left            =   11160
      Picture         =   "frmMain.frx":1D980F
      Top             =   5400
      Width           =   360
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pStop As Boolean
Private LEARNING_RATE As Double
Private Weights(1 To 3, 1 To 3) As Double
Private pdicNet As New Dictionary
Private Sub cmdBias_Click()
        
    On Error GoTo EH
    
    'Declare
    Dim x As Double
    Dim w1 As Double
    Dim w2 As Double
    Dim s1 As Double
    Dim s2 As Double
    Dim y1 As Double
    Dim y2 As Double
    
    '
    x = Val(txtBias.Text)
    w1 = Val(lblW31.Caption)
    w2 = Val(lblW32.Caption)
    s1 = Val(lblSum1.Caption)
    s2 = Val(lblSum2.Caption)
    
    '
    s1 = s1 + x * w1
    s2 = s2 + x * w2
    
    '
    y1 = Activation(s1)
    y2 = Activation(s2)
    
    '
    lblSum1.Caption = Format(s1, "0.00000")
    lblSum2.Caption = Format(s2, "0.00000")
    lblY1.Caption = Format(y1, "0.00000")
    lblY2.Caption = Format(y2, "0.00000")
    
EH:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdBiasH_Click()
     
    On Error GoTo EH
    
    'Declare
    Dim x As Double
    Dim w As Double
    Dim s As Double
    Dim y As Double
    
    '
    x = 1
    w = Val(lblW33.Caption)
    s = Val(lblSum.Caption)
    
    '
    s = s + x * w
    
    '
    y = Activation(s)
    
    '
    lblSum.Caption = Format(s, "0.00000")
    lblY.Caption = Format(y, "0.00000")
    
EH:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdBP_Click()
        
    On Error GoTo EH
    
    'Declare
    Dim Deltas(3) As Double
    Dim y As Double
    Dim y1 As Double
    Dim y2 As Double
    Dim Target As Double
    Dim v1 As Double
    Dim v2 As Double
    Dim i As Long
    
    '
    Target = Val(txtTarget.Text)
    
    '
    y = Val(lblY.Caption)
    y1 = Val(lblY1.Caption)
    y2 = Val(lblY2.Caption)
    
    'Delta
    Deltas(3) = y * (1 - y) * (Target - y)
    Deltas(2) = y2 * (1 - y2) * Weights(3, 3) * Deltas(3)
    Deltas(1) = y1 * (1 - y1) * Weights(2, 3) * Deltas(3)
    
    '
    lblDelta3.Caption = Format(Deltas(3), "0.00000")
    lblDelta2.Caption = Format(Deltas(2), "0.00000")
    lblDelta1.Caption = Format(Deltas(1), "0.00000")
    
    'Now, alter the weights accordingly.
    v1 = Val(txtX1.Text)
    v2 = Val(txtX2.Text)
    For i = 1 To 3
        'Change the values for the output layer, if necessary.
        If i = 3 Then
            v1 = y1
            v2 = y2
        End If
        Weights(1, i) = Weights(1, i) + LEARNING_RATE * 1 * Deltas(i) 'Bias
        Weights(2, i) = Weights(2, i) + LEARNING_RATE * v1 * Deltas(i)
        Weights(3, i) = Weights(3, i) + LEARNING_RATE * v2 * Deltas(i)
    Next i
    
    '
    lblW11.Caption = Format(Weights(2, 1), "0.00000")
    lblW12.Caption = Format(Weights(2, 2), "0.00000")
    lblW21.Caption = Format(Weights(3, 1), "0.00000")
    lblW22.Caption = Format(Weights(3, 2), "0.00000")
    lblW31.Caption = Format(Weights(1, 1), "0.00000")
    lblW32.Caption = Format(Weights(1, 2), "0.00000")
    '
    lblW13.Caption = Format(Weights(2, 3), "0.00000")
    lblW23.Caption = Format(Weights(3, 3), "0.00000")
    lblW33.Caption = Format(Weights(1, 3), "0.00000")
    
EH:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdFP_Click()
        
    On Error GoTo EH
    
    '
    cmdX1.Value = True
    cmdX2.Value = True
    cmdBias.Value = True
    
    '
    cmdH1.Value = True
    cmdH2.Value = True
    cmdBiasH.Value = True
    
EH:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdH1_Click()
        
    On Error GoTo EH
    
    'Declare
    Dim x As Double
    Dim w As Double
    Dim s As Double
    Dim y As Double
    
    '
    x = Val(lblY1.Caption)
    w = Val(lblW13.Caption)
    s = Val(lblSum.Caption)
    
    '
    s = x * w
    
    '
    y = Activation(s)
    
    '
    lblSum.Caption = Format(s, "0.00000")
    lblY.Caption = Format(y, "0.00000")
    
EH:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdH2_Click()
     
    On Error GoTo EH
    
    'Declare
    Dim x As Double
    Dim w As Double
    Dim s As Double
    Dim y As Double
    
    '
    x = Val(lblY2.Caption)
    w = Val(lblW23.Caption)
    s = Val(lblSum.Caption)
    
    '
    s = s + x * w
    
    '
    y = Activation(s)
    
    '
    lblSum.Caption = Format(s, "0.00000")
    lblY.Caption = Format(y, "0.00000")
    
EH:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdReset_Click()
        
    On Error GoTo EH
    
    '
    Reset
    
    '
    lblW11.Caption = Format(Weights(2, 1), "0.00000")
    lblW12.Caption = Format(Weights(2, 2), "0.00000")
    lblW21.Caption = Format(Weights(3, 1), "0.00000")
    lblW22.Caption = Format(Weights(3, 2), "0.00000")
    lblW31.Caption = Format(Weights(1, 1), "0.00000")
    lblW32.Caption = Format(Weights(1, 2), "0.00000")
    '
    lblW13.Caption = Format(Weights(2, 3), "0.00000")
    lblW23.Caption = Format(Weights(3, 3), "0.00000")
    lblW33.Caption = Format(Weights(1, 3), "0.00000")
    
    '
    lblSum1.Caption = "Sum1"
    lblSum2.Caption = "Sum2"
    lblSum.Caption = "Sum"
    lblY1.Caption = "Y1"
    lblY2.Caption = "Y2"
    lblY.Caption = "Y"
    
    '
    lblDelta3.Caption = "Delta3"
    lblDelta2.Caption = "Delta2"
    lblDelta1.Caption = "Delta1"
    
    '
    lblCounter.Caption = ""
    lblMSE.Caption = "MSE"
    
EH:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
End Sub

Private Sub RunOne()
        
    On Error GoTo EH
    
    'Declare
    Dim i As Long
    Dim n As Long
    Dim X1 As Double
    Dim X2 As Double
    Dim Target As Double
    Dim mse As Double
    Dim output
    
    '
    pStop = False
    lblCounter.Caption = 0
    lblMSE.Caption = "MSE"
    
    '
    n = Val(txtIterations.Text)
    If n < 1 Then
        n = 1
        txtIterations.Text = 1
    End If
    
    '
    X1 = Val(txtX1.Text)
    X2 = Val(txtX2.Text)
    Target = Val(txtTarget.Text)
    
    '
    mse = 0
    For i = 1 To n
        output = Train(X1, X2, Target)
        ShowNet X1, X2, Target
        lblCounter.Caption = i
        mse = mse + (Target - output) ^ 2
        lblMSE.Caption = Format(Sqr(mse / i), "0.000")
        DoEvents
        If pStop = True Then Exit For
    Next i
    
EH:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
End Sub

Private Sub Test()

    Dim i As Long

    
    '
    For i = 1 To 3000
        Train 0, 0, 0
        ShowNet 0, 0, 0
        Train 0, 1, 1
        ShowNet 0, 1, 1
        Train 1, 0, 1
        ShowNet 1, 0, 1
        Train 1, 1, 0
        ShowNet 1, 1, 0
    Next i

    Print CCur(Run(0, 0))
    Print CCur(Run(0, 1))
    Print CCur(Run(1, 0))
    Print CCur(Run(1, 1))

End Sub

Private Sub cmdRunData_Click()
        
    On Error GoTo EH
    
    'Declare
    Dim i As Long
    Dim k As Long
    Dim n As Long
    Dim m As Long
    Dim X1 As Double
    Dim X2 As Double
    Dim Target As Double
    Dim mse As Double
    Dim data, rec, output
    
    '
    data = Split(txtData.Text, vbNewLine)
    
    '
    pStop = False
    lblCounter.Caption = 0
    lblMSE.Caption = "MSE"
    
    '
    n = Val(txtIterations.Text)
    If n < 1 Then
        n = 1
        txtIterations.Text = 1
    End If
    
    '
    mse = 0: m = 0
    For i = 1 To n
        For k = 0 To UBound(data)
            If Trim(data(k)) <> "" Then
                rec = Split(data(k), ",")
                X1 = Val(rec(0))
                X2 = Val(rec(1))
                Target = Val(rec(2))
                '
                output = Train(X1, X2, Target)
                ShowNet X1, X2, Target
                lblCounter.Caption = i
                '
                mse = mse + (Target - output) ^ 2
                m = m + 1
                '
                DoEvents
                If pStop = True Then Exit For
            End If
        Next k
        lblMSE.Caption = Format(Sqr(mse / m), "0.000")
        If pStop = True Then Exit For
    Next i
    
EH:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdX1_Click()
        
    On Error GoTo EH
    
    'Declare
    Dim x As Double
    Dim w1 As Double
    Dim w2 As Double
    Dim s1 As Double
    Dim s2 As Double
    Dim y1 As Double
    Dim y2 As Double
    
    '
    x = Val(txtX1.Text)
    w1 = Val(lblW11.Caption)
    w2 = Val(lblW12.Caption)
    s1 = Val(lblSum1.Caption)
    s2 = Val(lblSum2.Caption)
    
    '
    s1 = x * w1
    s2 = x * w2
    
    '
    y1 = Activation(s1)
    y2 = Activation(s2)
    
    '
    lblSum1.Caption = Format(s1, "0.00000")
    lblSum2.Caption = Format(s2, "0.00000")
    lblY1.Caption = Format(y1, "0.00000")
    lblY2.Caption = Format(y2, "0.00000")
    
EH:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdX2_Click()
    
    On Error GoTo EH
    
    'Declare
    Dim x As Double
    Dim w1 As Double
    Dim w2 As Double
    Dim s1 As Double
    Dim s2 As Double
    Dim y1 As Double
    Dim y2 As Double
    
    '
    x = Val(txtX2.Text)
    w1 = Val(lblW21.Caption)
    w2 = Val(lblW22.Caption)
    s1 = Val(lblSum1.Caption)
    s2 = Val(lblSum2.Caption)
    
    '
    s1 = s1 + x * w1
    s2 = s2 + x * w2
    
    '
    y1 = Activation(s1)
    y2 = Activation(s2)
    
    '
    lblSum1.Caption = Format(s1, "0.00000")
    lblSum2.Caption = Format(s2, "0.00000")
    lblY1.Caption = Format(y1, "0.00000")
    lblY2.Caption = Format(y2, "0.00000")
    
EH:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
End Sub

Private Sub ShowNet(X1, X2, Target)
    
    '
    txtX1.Text = X1
    txtX2.Text = X2
    txtTarget.Text = Target
    
    '
    lblW11.Caption = Format(Weights(2, 1), "0.00000")
    lblW12.Caption = Format(Weights(2, 2), "0.00000")
    lblW21.Caption = Format(Weights(3, 1), "0.00000")
    lblW22.Caption = Format(Weights(3, 2), "0.00000")
    lblW31.Caption = Format(Weights(1, 1), "0.00000")
    lblW32.Caption = Format(Weights(1, 2), "0.00000")
    '
    lblW13.Caption = Format(Weights(2, 3), "0.00000")
    lblW23.Caption = Format(Weights(3, 3), "0.00000")
    lblW33.Caption = Format(Weights(1, 3), "0.00000")
    
    '
    lblSum1.Caption = Format(pdicNet("sum1"), "0.00000")
    lblSum2.Caption = Format(pdicNet("sum2"), "0.00000")
    lblY1.Caption = Format(pdicNet("net1"), "0.00000")
    lblY2.Caption = Format(pdicNet("net2"), "0.00000")
    
    '
    lblSum.Caption = Format(pdicNet("sum3"), "0.00000")
    lblY.Caption = Format(pdicNet("net3"), "0.00000")
    
    '
    lblDelta3.Caption = Format(pdicNet("delta3"), "0.00000")
    lblDelta2.Caption = Format(pdicNet("delta2"), "0.00000")
    lblDelta1.Caption = Format(pdicNet("delta1"), "0.00000")
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        pStop = True
    End If
    
End Sub

Public Function Reset()
    
    'Declare
    Dim i As Integer
    Dim j As Integer
    
    '
    Randomize
    
    '
    For i = 1 To 3
        For j = 1 To 3
            Weights(i, j) = (Rnd * 2) - 1
        Next j
    Next i
    
End Function


Public Function Train(Input1 As Double, Input2 As Double, Target As Double)
    
    'Declare
    Dim Deltas(1 To 3) As Double
    Dim sum1 As Double
    Dim sum2 As Double
    Dim sum3 As Double
    Dim net1 As Double
    Dim net2 As Double
    Dim net3 As Double
    Dim v1 As Double
    Dim v2 As Double
    Dim i As Integer
    Dim out As Double
    
    'hidden layer
    sum1 = Weights(1, 1) + Input1 * Weights(2, 1) + Input2 * Weights(3, 1)
    net1 = Activation(sum1)
    sum2 = Weights(1, 2) + Input1 * Weights(2, 2) + Input2 * Weights(3, 2)
    net2 = Activation(sum2)
    
    'output layer
    sum3 = Weights(1, 3) + net1 * Weights(2, 3) + net2 * Weights(3, 3)
    net3 = Activation(sum3)
    
    'calculate the deltas for the two layers (back-propagation)
    Deltas(3) = net3 * (1 - net3) * (Target - net3)
    Deltas(2) = net2 * (1 - net2) * (Weights(3, 3)) * (Deltas(3))
    Deltas(1) = net1 * (1 - net1) * (Weights(2, 3)) * (Deltas(3))
    
    'update weights
    v1 = Input1
    v2 = Input2
    For i = 1 To 3
        'Change the values for the output layer, if necessary.
        If i = 3 Then
            v1 = net1
            v2 = net2
        End If
        Weights(1, i) = Weights(1, i) + LEARNING_RATE * 1 * Deltas(i) 'Bias
        Weights(2, i) = Weights(2, i) + LEARNING_RATE * v1 * Deltas(i)
        Weights(3, i) = Weights(3, i) + LEARNING_RATE * v2 * Deltas(i)
    Next i
    
    '
    With pdicNet
        .Item("sum1") = sum1
        .Item("sum2") = sum2
        .Item("sum3") = sum3
        .Item("net1") = net1
        .Item("net2") = net2
        .Item("net3") = net3
        .Item("delta1") = Deltas(1)
        .Item("delta2") = Deltas(2)
        .Item("delta3") = Deltas(3)
    End With
    
    '
    DoEvents
    Train = net3
    
End Function


Public Function Run(Input1 As Double, Input2 As Double)

    'Declare
    Dim net1 As Double
    Dim net2 As Double
    Dim net3 As Double
    
    '
    net1 = Activation(Weights(1, 1) + Input1 * Weights(2, 1) + Input2 * Weights(3, 1))
    net2 = Activation(Weights(1, 2) + Input1 * Weights(2, 2) + Input2 * Weights(3, 2))
    net3 = Activation(Weights(1, 3) + net1 * Weights(2, 3) + net2 * Weights(3, 3))
    
    '
    Run = net3
    
End Function


Public Function Activation(Value As Double)
    Activation = (1 / (1 + Exp(Value * -1)))
End Function

Private Sub Form_Load()

    '
    LEARNING_RATE = Val(txtLearningRate.Text)
    If LEARNING_RATE < 0.01 Then
        LEARNING_RATE = 0.01
        txtLearningRate.Text = 0.01
    End If
    
    '
    With pdicNet
        .Add "sum1", 0
        .Add "sum2", 0
        .Add "sum3", 0
        .Add "net1", 0
        .Add "net2", 0
        .Add "net3", 0
        .Add "delta1", 0
        .Add "delta2", 0
        .Add "delta3", 0
    End With
    
End Sub

