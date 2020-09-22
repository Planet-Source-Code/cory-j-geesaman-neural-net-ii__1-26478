VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "typeBond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarEngram As clsEngram 'local copy
Private mvarStrength As Long 'local copy
Public Property Let Strength(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Strength = 5
    mvarStrength = vData
End Property


Public Property Get Strength() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Strength
    Strength = mvarStrength
End Property



Public Property Set Engram(ByVal vData As clsEngram)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Engram = Form1
    Set mvarEngram = vData
End Property


Public Property Get Engram() As clsEngram
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Engram
    Set Engram = mvarEngram
End Property



