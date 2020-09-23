VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDI Form"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9870
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Height          =   5625
      Left            =   0
      ScaleHeight     =   5565
      ScaleWidth      =   1665
      TabIndex        =   0
      Top             =   0
      Width           =   1725
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Min"
         Height          =   615
         Left            =   240
         TabIndex        =   2
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Max"
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Toggle Button States"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_blnClose As Boolean
Private m_blnMin As Boolean
Private m_blnMax As Boolean


Private Sub Command1_Click()
    m_blnMax = Not m_blnMax
    EnableMaxButton MDIForm1.hWnd, m_blnMax
    EnableMaxButton Form1.hWnd, m_blnMax
    EnableMaxButton Form2.hWnd, m_blnMax
End Sub

Private Sub Command2_Click()
    m_blnMin = Not m_blnMin
    EnableMinButton MDIForm1.hWnd, m_blnMin
    EnableMinButton Form1.hWnd, m_blnMin
    EnableMinButton Form2.hWnd, m_blnMin
End Sub

Private Sub Command3_Click()
    m_blnClose = Not m_blnClose
    EnableCloseButton MDIForm1.hWnd, m_blnClose
    EnableCloseButton Form1.hWnd, m_blnClose
    EnableCloseButton Form2.hWnd, m_blnClose
End Sub

Private Sub MDIForm_Load()
    m_blnClose = True
    m_blnMin = True
    m_blnMax = True
    
    Form1.Show vbModeless, Me
    Form2.Show
End Sub
