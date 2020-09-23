VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls2DVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Copyright (c) 2003 Richard Hayden. All Rights Reserved.
'
'You may use any code contained in this project for NON-commercial gain
'as long as credit is given to the author (Richard Hayden).
'
'Thanks for downloading my code!

Private mvarx As Double 'local copy
Private mvary As Double 'local copy
Public Property Let y(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.y = 5
    mvary = vData
End Property


Public Property Get y() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.y
    y = mvary
End Property



Public Property Let x(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.x = 5
    mvarx = vData
End Property


Public Property Get x() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.x
    x = mvarx
End Property



