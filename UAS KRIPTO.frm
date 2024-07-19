VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080FF80&
   Caption         =   "Form1"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Kosongkan Text"
      Height          =   495
      Left            =   8640
      TabIndex        =   11
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Deskrip"
      Height          =   375
      Left            =   8880
      TabIndex        =   10
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   8760
      TabIndex        =   9
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   8760
      TabIndex        =   8
      Top             =   1440
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Kosongkan Text"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enskrip"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Top             =   2040
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "OUTPUT TEXT"
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "INPUT TEXT"
      Height          =   255
      Left            =   6960
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "OUTPUTTEXT"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "INPUT TEXT"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Enkrip, Output, Inputan As String
Dim Panjang_Input As Integer
Inputan = Text1.Text
Panjang_Input = Len(Text1.Text)

For i = 1 To Panjang_Input
Enkrip = Mid(Inputan, i, 1) 'ambil karakter input
Enkrip = Asc(Enkrip) 'ubah karakter ke ascii
Enkrip = (Enkrip + 20) - 43 'key + 20 - 30 (bisa gunakan yg lain)
Enkrip = Chr(Enkrip) ' ubah kembali ke karakter
Output = Output & Enkrip
Next i

Text2.Text = Output ' tampilkan enkrip
End Sub

Private Sub Command3_Click()
Dim Deskrip, Output, Inputan As String
Dim Panjang_Input As Integer
Inputan = Text3.Text
Panjang_Input = Len(Text1.Text)

For i = 1 To Panjang_Input
Deskrip = Mid(Inputan, i, 1) 'ambil karakter input
Deskrip = Asc(Deskrip) 'ubah karakter ke ascii
Deskrip = (Deskrip - 20) + 43 'key + 20 - 30 (bisa gunakan yg lain)
Deskrip = Chr(Deskrip) ' ubah kembali ke karakter
Output = Output & Deskrip
Next i

Text4.Text = Output ' tampilkan enkrip
End Sub

Private Sub Command2_Click() 'tombol deskripsi
Text1 = ""
Text2 = ""
Dim Dekrip, Output, Inputan As String
Dim Panjang_Input, Pesan As Integer
Inputan = Text3.Text
Panjang_Input = Len(Text3.Text)

For i = 1 To Panjang_Input
Deskrip = Mid(Inputan, i, 1) 'ambil karakter input
Deskrip = Asc(Deskrip) 'ubah karakter ke ascii
Deskrip = (Deskrip - 20) + 43 'key + 20 - 30 (bisa gunakan yg lain)
Deskrip = Chr(Deskrip) ' ubah kembali ke karakter
Output = Output & Deskrip
Next i

Text4.Text = Output
End Sub

Private Sub Command4_Click()
Text3 = ""
Text4 = ""
End Sub

