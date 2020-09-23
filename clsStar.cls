VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStar"
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

'local variable(s) to hold property value(s)
Private mvarPosition As cls2DVector 'local copy
Private mvarVelocity As cls2DVector 'local copy
'local variable(s) to hold property value(s)
Private mvarMass As Double 'local copy
Public Property Let Mass(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Mass = 5
    mvarMass = vData
End Property

Public Property Get Mass() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Mass
    Mass = mvarMass
End Property

Public Property Set Velocity(ByVal vData As cls2DVector)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Velocity = Form1
    Set mvarVelocity = vData
End Property

Public Property Get Velocity() As cls2DVector
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Velocity
    Set Velocity = mvarVelocity
End Property

Public Property Set Position(ByVal vData As cls2DVector)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Position = Form1
    Set mvarPosition = vData
End Property


Public Property Get Position() As cls2DVector
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Position
    Set Position = mvarPosition
End Property

Private Sub Class_Initialize()
    Set mvarPosition = New cls2DVector
    Set mvarVelocity = New cls2DVector
End Sub
