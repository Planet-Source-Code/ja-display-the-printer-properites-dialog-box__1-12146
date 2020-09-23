Attribute VB_Name = "modPrinterProperties"
'Required constants
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32
Public Const PRINTER_ACCESS_USE = &H8

'Requires Types
Public Type ACL
        AclRevision As Byte
        Sbz1 As Byte
        AclSize As Integer
        AceCount As Integer
        Sbz2 As Integer
End Type

Public Type SECURITY_DESCRIPTOR
        Revision As Byte
        Sbz1 As Byte
        Control As Long
        Owner As Long
        Group As Long
        Sacl As ACL
        Dacl As ACL
End Type

Public Type DEVMODE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Long
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

Public Type PRINTER_INFO_2
        pServerName As String
        pPrinterName As String
        pShareName As String
        pPortName As String
        pDriverName As String
        pComment As String
        pLocation As String
        pDevMode As DEVMODE
        pSepFile As String
        pPrintProcessor As String
        pDatatype As String
        pParameters As String
        pSecurityDescriptor As SECURITY_DESCRIPTOR
        Attributes As Long
        Priority As Long
        DefaultPriority As Long
        StartTime As Long
        UntilTime As Long
        Status As Long
        cJobs As Long
        AveragePPM As Long
End Type

Public Type PRINTER_DEFAULTS
        pDatatype As String
        pDevMode As DEVMODE
        DesiredAccess As Long
End Type

'Required Functions
Public Declare Function PrinterProperties Lib "winspool.drv" (ByVal hwnd As Long, ByVal hPrinter As Long) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long

