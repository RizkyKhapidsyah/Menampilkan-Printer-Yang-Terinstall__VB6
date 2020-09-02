VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menampilkan Printer yang Terinstall"
   ClientHeight    =   3090
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1200
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ptr As Printer   'Deklarasi objek printer

Private Sub Form_Load()
  If cboPrinter.ListCount = 0 Then
     For Each Ptr In Printers
         cboPrinter.AddItem Ptr.DeviceName
     Next
  End If
  
  cboPrinter.Text = Printer.DeviceName  'Tampilkan
  'printer defaultnya.
End Sub

