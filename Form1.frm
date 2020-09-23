VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   450
      Left            =   3480
      TabIndex        =   3
      Top             =   1575
      Width           =   1005
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1065
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   375
      Width           =   2460
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Properties"
      Height          =   315
      Left            =   3600
      TabIndex        =   0
      Top             =   375
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "Printer: "
      Height          =   240
      Left            =   75
      TabIndex        =   2
      Top             =   375
      Width           =   885
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim PrintDef As PRINTER_DEFAULTS
Dim f As Boolean
Dim r As Long
Dim pHwnd As Long

PrintDef.DesiredAccess = PRINTER_ACCESS_USE

'Get the printer handle
r = OpenPrinter(Combo1.Text & vbNullString, pHwnd, PrintDef)
'Show the dialog box
f = PrinterProperties(hwnd, pHwnd)
'close the printer
ClosePrinter pHwnd

End Sub


Private Sub Command2_Click()
Unload Me

End Sub


Private Sub Form_Load()
'Add all the available printers to the combo box
For i% = 0 To Printers.Count - 1
    Combo1.AddItem Printers(i%).DeviceName
Next i%

'Display the default printer
For i% = 0 To Combo1.ListCount - 1
If Combo1.List(i%) = Printer.DeviceName Then
Combo1.ListIndex = i%
Exit For
End If
Next i%

End Sub


Private Sub Form_Unload(Cancel As Integer)
End

End Sub


