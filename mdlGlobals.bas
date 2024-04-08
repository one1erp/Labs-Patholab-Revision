Attribute VB_Name = "mdlGlobals"
Option Explicit

Public con As ADODB.Connection
Public Const DEFINE_DEBUG = False

Public Function nte(e As Variant) As Variant
5320      nte = IIf(IsNull(e), "", e)
End Function

Public Function ntz(e As Variant) As Variant
5330      ntz = IIf(IsNull(e), 0, e)
End Function
