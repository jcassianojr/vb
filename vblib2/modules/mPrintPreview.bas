Attribute VB_Name = "mPrintPreview"
Option Explicit


'Needed for print preview:
Public Property Get Printer() As Printer
  Set Printer = PrinterReplacement
End Property

Public Property Set Printer(nPrinter As Printer)
  Set PrinterReplacement = nPrinter
End Property

