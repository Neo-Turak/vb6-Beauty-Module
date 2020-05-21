VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   4515
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4065
      Left            =   180
      TabIndex        =   2
      Top             =   540
      Width           =   4110
      Begin VB.PictureBox Picture1 
         Height          =   3525
         Left            =   270
         ScaleHeight     =   3465
         ScaleWidth      =   3420
         TabIndex        =   3
         Top             =   360
         Width           =   3480
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   13
            Top             =   360
            Width           =   1050
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   330
            Index           =   0
            Left            =   90
            TabIndex        =   12
            Top             =   720
            Width           =   1050
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   285
            Index           =   1
            Left            =   90
            TabIndex        =   11
            Top             =   90
            Width           =   1050
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   330
            Index           =   1
            Left            =   90
            TabIndex        =   10
            Top             =   1080
            Width           =   1050
         End
         Begin VB.ListBox List1 
            Height          =   1680
            Left            =   90
            TabIndex        =   9
            Top             =   1530
            Width           =   1230
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Left            =   1530
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   360
            Width           =   1680
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Left            =   1530
            TabIndex        =   7
            Text            =   "Combo1"
            Top             =   675
            Width           =   1680
         End
         Begin VB.DriveListBox Drive1 
            Height          =   300
            Left            =   1530
            TabIndex        =   6
            Top             =   1035
            Width           =   1680
         End
         Begin VB.DirListBox Dir1 
            Height          =   720
            Left            =   1530
            TabIndex        =   5
            Top             =   1440
            Width           =   1680
         End
         Begin VB.FileListBox File1 
            Height          =   810
            Left            =   1530
            TabIndex        =   4
            Top             =   2340
            Width           =   1635
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Label1"
            Height          =   180
            Left            =   1530
            TabIndex        =   14
            Top             =   90
            Width           =   540
         End
      End
   End
   Begin VB.CommandButton QQStyleCmd 
      Caption         =   "还原外观"
      Height          =   375
      Index           =   1
      Left            =   2700
      TabIndex        =   1
      Top             =   90
      Width           =   1545
   End
   Begin VB.CommandButton QQStyleCmd 
      Caption         =   "更改外观"
      Height          =   375
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub QQStyleCmd_Click(Index As Integer)
    If Index = 0 Then
        Attach Me.hWnd
    Else
        Detach Me.hWnd
    End If
End Sub
