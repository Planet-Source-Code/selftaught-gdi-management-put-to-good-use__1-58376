VERSION 5.00
Begin VB.UserControl UserControl1 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3960
   PropertyPages   =   "UserControl1.ctx":0000
   ScaleHeight     =   2100
   ScaleWidth      =   3960
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iSupportFontPropPage

Private WithEvents moFont As cFont
Attribute moFont.VB_VarHelpID = -1

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Private Const PROP_Font = "Font"

Private Sub iSupportFontPropPage_AddFonts(ByVal o As ppFont)
    o.ShowProps "Font"
End Sub

Private Function iSupportFontPropPage_GetAmbientFont() As stdole.Font
    Set iSupportFontPropPage_GetAmbientFont = Ambient.Font
End Function


Private Sub moFont_Changed()
    pFontChanged
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If StrComp(PROP_Font, PropertyName) = 0 Then moFont.OnAmbientFontChanged Ambient.Font
End Sub

Private Sub UserControl_InitProperties()
    Set moFont = Font_CreateDefault(Ambient.Font)
    pFontChanged
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set moFont = Font_Read(PropBag, PROP_Font, Ambient.Font)
    pFontChanged
End Sub

Private Sub UserControl_Terminate()
    Set moFont = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Font_Write moFont, PropBag, PROP_Font
End Sub


Public Property Get Font() As cFont
    Set Font = moFont
End Property

Public Property Set Font(ByVal oNew As cFont)
    If oNew Is Nothing Then Set oNew = Font_CreateDefault(Ambient.Font)
    Set moFont = oNew
    pFontChanged
End Property

Private Sub pFontChanged()
    Dim lhFont
    lhFont = moFont.GetHandle()
    If lhFont Then
        Cls
        Dim lhFontOld As Long
        lhFontOld = SelectObject(hdc, lhFont)
        TextOut hdc, 10, 10, "This is a test string!", 22&
        SelectObject hdc, lhFontOld
        moFont.ReleaseHandle lhFont
    End If
End Sub
