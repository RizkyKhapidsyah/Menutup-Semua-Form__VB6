VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menutup Semua Form"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Tutup"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buka Form"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   2160
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()  'Buka form lainnya
   Form2.Show
   Form3.Show
End Sub

Private Sub Command2_Click()   'Tutup semua form yang 'ada, termasuk Form1
Dim Form As Form               '(Sebenarnya, hal ini 'sama dengan 'End')
   For Each Form In Forms
       Unload Form
       Set Form = Nothing      'Bersihkan memori yang digunakan sebelumnya
   Next Form
End Sub

