VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "h"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "KELUAR"
      Height          =   375
      Left            =   3960
      TabIndex        =   16
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   15
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ULANGI"
      Height          =   375
      Left            =   1680
      TabIndex        =   14
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2640
      TabIndex        =   13
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2640
      TabIndex        =   12
      Top             =   3600
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2640
      TabIndex        =   11
      Top             =   2760
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   1920
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   1920
      Width           =   255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2640
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Data"
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "Biasa"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Total  "
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Jual Beli"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Nominal"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Jenis"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Provider"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "BAMZ CELL"
      Height          =   255
      Left            =   2640
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Combo1.Text = ""
Combo2.Text = ""
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Command2_Click()
If Option1.Value = True Then
MsgBox "Selamat Anda Berhasil Melakukan pembeli Pulsa Biasa" & Combo1.Text & "Sebesar" & Text2.Text
Else
If Option2.Value = True Then
If Combo2.Text = 10000 Then
MsgBox "Selamat Anda Berhasil Melakukan pembeli Pulsa Data" & Combo1.Text & "Sebesar" & Text2.Text & "Untuk 5 GB (14 Hari)"
Else
If Combo2.Text = 20000 Then
MsgBox "Selamat Anda Berhasil Melakukan pembeli Pulsa Data" & Combo1.Text & "Sebesar" & Text2.Text & " Untuk 10 GB (30 Hari)"
Else
If Combo2.Text = 25000 Then
MsgBox "Selamat Anda Berhasil Melakukan pembeli Pulsa Data" & Combo1.Text & "Sebesar" & Text2.Text & "Untuk 15 GB (30 Hari)"
Else
If Combo2.Text = 50000 Then
MsgBox "Selamat Anda Berhasil Melakukan pembeli Pulsa Data" & Combo1.Text & "Sebesar" & Text2.Text & "Untuk 35 GB (30 Hari)"
Else
MsgBox "Selamat Anda Berhasil Melakukan pembeli Pulsa Data" & Combo1.Text & "Sebesar" & Text2.Text & "Untuk 70 GB (30 Hari)"
End If
End If
End If
End If
End If
End If
End Sub


Private Sub Command3_Click()
End
End Sub


Private Sub Form_Load()
Combo1.AddItem "TELKOMSEL"
Combo1.AddItem "INDOSAT"
Combo1.AddItem "XL"
Combo2.AddItem "10000"
Combo2.AddItem "20000"
Combo2.AddItem "25000"
Combo2.AddItem "50000"
Combo2.AddItem "100000"
End Sub


Private Sub Text1_Change()
Text2.Text = Val(Combo2.Text) * Val(Text1.Text) + 5000
End Sub

