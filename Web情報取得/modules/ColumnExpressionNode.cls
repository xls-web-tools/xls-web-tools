VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnExpressionNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pKind As String
Private pTextValue As String
Private pOperatorText As String
Private pLeftNode As ColumnExpressionNode
Private pMiddleNode As ColumnExpressionNode
Private pRightNode As ColumnExpressionNode

Public Property Get Kind() As String
    Kind = pKind
End Property
Public Property Let Kind(ByVal NewValue As String)
    pKind = NewValue
End Property

Public Property Get TextValue() As String
    TextValue = pTextValue
End Property
Public Property Let TextValue(ByVal NewValue As String)
    pTextValue = NewValue
End Property

Public Property Get OperatorText() As String
    OperatorText = pOperatorText
End Property
Public Property Let OperatorText(ByVal NewValue As String)
    pOperatorText = NewValue
End Property

Public Property Get LeftNode() As ColumnExpressionNode
    Set LeftNode = pLeftNode
End Property
Public Property Set LeftNode(ByVal NewValue As ColumnExpressionNode)
    Set pLeftNode = NewValue
End Property

Public Property Get MiddleNode() As ColumnExpressionNode
    Set MiddleNode = pMiddleNode
End Property
Public Property Set MiddleNode(ByVal NewValue As ColumnExpressionNode)
    Set pMiddleNode = NewValue
End Property

Public Property Get RightNode() As ColumnExpressionNode
    Set RightNode = pRightNode
End Property
Public Property Set RightNode(ByVal NewValue As ColumnExpressionNode)
    Set pRightNode = NewValue
End Property
