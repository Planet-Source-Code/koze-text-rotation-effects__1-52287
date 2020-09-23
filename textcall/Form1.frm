VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10605
   ClientLeft      =   3600
   ClientTop       =   2310
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   ScaleHeight     =   10605
   ScaleWidth      =   11925
   Begin VB.Frame Frame4 
      Caption         =   "Eliptic text"
      Height          =   2535
      Left            =   6960
      TabIndex        =   19
      Top             =   7800
      Width           =   3615
      Begin VB.HScrollBar HScroll12 
         Height          =   255
         Left            =   960
         Max             =   100
         Min             =   1
         TabIndex        =   29
         Top             =   1680
         Value           =   1
         Width           =   2415
      End
      Begin VB.HScrollBar HScroll11 
         Height          =   255
         Left            =   960
         Max             =   100
         Min             =   1
         TabIndex        =   28
         Top             =   1320
         Value           =   1
         Width           =   2415
      End
      Begin VB.HScrollBar HScroll10 
         Height          =   255
         Left            =   960
         Max             =   360
         TabIndex        =   22
         Top             =   600
         Value           =   360
         Width           =   2415
      End
      Begin VB.HScrollBar HScroll9 
         Height          =   255
         Left            =   960
         Max             =   200
         TabIndex        =   21
         Top             =   240
         Width           =   2415
      End
      Begin VB.HScrollBar HScroll8 
         Height          =   255
         Left            =   960
         Max             =   180
         Min             =   1
         TabIndex        =   20
         Top             =   960
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label Label11 
         Caption         =   "Bfact"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Afact"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Deg"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Rad"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Arc"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Wave Text"
      Height          =   2655
      Left            =   240
      TabIndex        =   12
      Top             =   7800
      Width           =   2775
      Begin VB.HScrollBar HScroll6 
         Height          =   255
         Left            =   720
         Max             =   1800
         TabIndex        =   15
         Top             =   1080
         Value           =   50
         Width           =   1695
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2535
         Left            =   2520
         Max             =   100
         TabIndex        =   14
         Top             =   120
         Width           =   255
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   0
         Max             =   100
         TabIndex        =   13
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Curves"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rotated Text"
      Height          =   1095
      Left            =   3240
      TabIndex        =   7
      Top             =   7800
      Width           =   3255
      Begin VB.HScrollBar HScroll5 
         Height          =   255
         Left            =   720
         Max             =   50
         TabIndex        =   10
         Top             =   720
         Value           =   25
         Width           =   2415
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   255
         Left            =   720
         Max             =   360
         TabIndex        =   8
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Length"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Radious"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Circle Text"
      Height          =   1335
      Left            =   3240
      TabIndex        =   2
      Top             =   9120
      Width           =   3255
      Begin VB.HScrollBar HScroll7 
         Height          =   255
         Left            =   720
         Max             =   180
         Min             =   1
         TabIndex        =   17
         Top             =   960
         Value           =   1
         Width           =   2415
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   720
         Max             =   200
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         Left            =   720
         Max             =   360
         TabIndex        =   3
         Top             =   600
         Value           =   360
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Deg"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Arc"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Rad"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   7440
      Width           =   11295
   End
   Begin VB.PictureBox Picture1 
      FillColor       =   &H000000FF&
      FillStyle       =   2  'Horizontal Line
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   240
      ScaleHeight     =   469
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   733
      TabIndex        =   0
      Top             =   240
      Width           =   11055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub HScroll1_Scroll()
wavetext Picture1, (Picture1.ScaleWidth / 2) - (HScroll1.Value * (Len(Text1.text) / 2)), Picture1.ScaleHeight / 2, HScroll1.Value, HScroll6.Value, VScroll1.Value, Text1.text
End Sub

Private Sub HScroll10_Scroll()
eliptictext Picture1, Picture1.ScaleWidth / 2, Picture1.ScaleHeight / 2, HScroll9.Value, HScroll8.Value, HScroll11.Value, HScroll12.Value, HScroll10.Value, Text1.text

End Sub

Private Sub HScroll11_Scroll()
eliptictext Picture1, Picture1.ScaleWidth / 2, Picture1.ScaleHeight / 2, HScroll9.Value, HScroll8.Value, HScroll11.Value, HScroll12.Value, HScroll10.Value, Text1.text

End Sub

Private Sub HScroll12_Scroll()
eliptictext Picture1, Picture1.ScaleWidth / 2, Picture1.ScaleHeight / 2, HScroll9.Value, HScroll8.Value, HScroll11.Value, HScroll12.Value, HScroll10.Value, Text1.text

End Sub

Private Sub HScroll2_Scroll()
circletext Picture1, Picture1.ScaleWidth / 2, Picture1.ScaleHeight / 2, HScroll2.Value, HScroll7.Value, HScroll3.Value, Text1.text

End Sub

Private Sub HScroll3_Scroll()
circletext Picture1, Picture1.ScaleWidth / 2, Picture1.ScaleHeight / 2, HScroll2.Value, HScroll7.Value, HScroll3.Value, Text1.text

End Sub

Private Sub HScroll4_Scroll()
rotateText Picture1, Picture1.ScaleWidth / 2, Picture1.ScaleHeight / 2, HScroll4.Value, HScroll5.Value, Text1
End Sub

Private Sub HScroll5_Scroll()
rotateText Picture1, Picture1.ScaleWidth / 2, Picture1.ScaleHeight / 2, HScroll4.Value, HScroll5.Value, Text1

End Sub

Private Sub HScroll6_Scroll()
wavetext Picture1, (Picture1.ScaleWidth / 2) - (HScroll1.Value * (Len(Text1.text) / 2)), Picture1.ScaleHeight / 2, HScroll1.Value, HScroll6.Value, VScroll1.Value, Text1.text

End Sub

Private Sub HScroll7_Scroll()
circletext Picture1, Picture1.ScaleWidth / 2, Picture1.ScaleHeight / 2, HScroll2.Value, HScroll7.Value, HScroll3.Value, Text1.text

End Sub

Private Sub HScroll8_Scroll()
eliptictext Picture1, Picture1.ScaleWidth / 2, Picture1.ScaleHeight / 2, HScroll9.Value, HScroll8.Value, HScroll11.Value, HScroll12.Value, HScroll10.Value, Text1.text

End Sub

Private Sub HScroll9_Scroll()
eliptictext Picture1, Picture1.ScaleWidth / 2, Picture1.ScaleHeight / 2, HScroll9.Value, HScroll8.Value, HScroll11.Value, HScroll12.Value, HScroll10.Value, Text1.text
End Sub

Private Sub VScroll1_Scroll()
wavetext Picture1, (Picture1.ScaleWidth / 2) - (HScroll1.Value * (Len(Text1.text) / 2)), Picture1.ScaleHeight / 2, HScroll1.Value, HScroll6.Value, VScroll1.Value, Text1.text

End Sub

