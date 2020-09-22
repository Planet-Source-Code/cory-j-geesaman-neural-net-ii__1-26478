VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEngram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A Previous Memory."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarData As Variant
Private Bonds As Collection

Public Property Let Data(ByVal vData As Variant)
    mvarData = vData
End Property

Public Property Set Data(ByVal vData As Variant)
    Set mvarData = vData
End Property

Public Property Get Data() As Variant
    If IsObject(mvarData) Then
        Set Data = mvarData
    Else
        Data = mvarData
    End If
End Property

Public Sub ChangeBondStrength(tEngram As clsEngram, BondStrength As Long)
If Bonds.Count > 0 Then
i = 1
Do
If Bonds(i).Engram Is tEngram Then Bonds(i).Strength = BondStrength
i = i + 1
Loop Until i > Bonds.Count
End If
End Sub

Public Function GetBondStrength(tEngram As clsEngram)
If Bonds.Count > 0 Then
i = 1
Do
If Bonds(i).Engram Is tEngram Then GetBondStrength = Bonds(i).Strength
i = i + 1
Loop Until i > Bonds.Count
End If
End Function

Public Sub RemoveConnection(tEngram As clsEngram)
If Bonds.Count > 0 Then
i = 1
Do
If Bonds(i).Engram Is tEngram Then Bonds.Remove i
i = i + 1
Loop Until i > Bonds.Count
End If
End Sub

Public Sub AddConnection(tEngram As clsEngram, BondStrength As Long)
Dim a As typeBond
Set a = New typeBond
Set a.Engram = tEngram
a.Strength = BondStrength
Bonds.Add a
End Sub

Public Function GetStrongestBond(IndexOfMe)
Dim e, i
If Bonds.Count > 0 Then
e = 1
i = 1
Do
If Bonds(i).Strength > Bonds(e).Strength Then e = i
i = i + 1
Loop Until i > Bonds.Count
End If
If e < IndexOfMe Or IndexOfMe = 0 Then GetStrongestBond = e Else GetStrongestBond = e + 1
End Function

Public Function CountConnections()
CountConnections = Bonds.Count
End Function

Private Sub Class_Initialize()
Set Bonds = New Collection
End Sub
