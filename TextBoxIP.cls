VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextBoxIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' VBA TextBoxIP
' Copyright (C) 2021  Philipp C. Ruedinger
' https://github.com/pcr-coding/vba-textboxip
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.

'@IgnoreModule ProcedureNotUsed, EmptyCaseBlock, HungarianNotation
'@Folder("Libraries.TextBoxIP")
Option Explicit

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents m_TextBoxIP As MSForms.TextBox
Attribute m_TextBoxIP.VB_VarHelpID = -1
Private m_EditMode As Boolean

'@Ignore WriteOnlyProperty
Public Property Set TextBox(ByVal TextBoxControl As MSForms.TextBox)
    Set m_TextBoxIP = TextBoxControl
    m_TextBoxIP.Font.Name = "Consolas"
End Property

' add dots automatically each block of 3
Private Sub m_TextBoxIP_Change()
    If m_EditMode Then Exit Sub
    
    With m_TextBoxIP
        Dim AmountOfDots As Long
        AmountOfDots = Len(.Value) - Len(Replace$(.Value, ".", vbNullString))

        If Len(.Value) = 3 And Right$(.Value, 1) <> "." Then
            m_EditMode = True
            .Value = .Value & "."
            m_EditMode = False
        ElseIf AmountOfDots > 0 And AmountOfDots < 3 And Right$(.Value, 1) <> "." Then
            Dim Pos As Long
            Pos = InStrRev(.Value, ".")

            Dim RightPart As String
            RightPart = Mid$(.Value, Pos + 1)
            If Len(RightPart) = 3 And Right$(RightPart, 1) <> "." Then
                m_EditMode = True
                .Value = .Value & "."
                m_EditMode = False
            End If
        ElseIf AmountOfDots > 3 Then
            m_EditMode = True
            Do While AmountOfDots > 3
                .Value = Left$(.Value, InStrRev(.Value, ".") - 1)
                AmountOfDots = Len(.Value) - Len(Replace$(.Value, ".", vbNullString))
            Loop
            m_EditMode = False
        End If
        
        ' if more dots then 3 cut off from the end (can happen on insert)
        m_EditMode = True
        Do While AmountOfDots > 3
            .Value = Left$(.Value, InStrRev(.Value, ".") - 1)
            AmountOfDots = Len(.Value) - Len(Replace$(.Value, ".", vbNullString))
        Loop
        m_EditMode = False
        
        ' validate each block to be <= 255
        Dim Blocks() As String
        Blocks = Split(.Value, ".")

        Dim iBlock As Long
        For iBlock = LBound(Blocks) To UBound(Blocks)
            If IsNumeric(Blocks(iBlock)) Then
                If Blocks(iBlock) > 255 Then Blocks(iBlock) = 255
            Else
                Blocks(iBlock) = Val(Blocks(iBlock))
                If Blocks(iBlock) > 255 Then Blocks(iBlock) = 255
                If Blocks(iBlock) = 0 Then Blocks(iBlock) = vbNullString
            End If
        Next iBlock
        
        m_EditMode = True
        .Value = Join(Blocks, ".")
        m_EditMode = False
    End With
End Sub


' allow only valid ip chars and control input in 4 groups devided by dots
Private Sub m_TextBoxIP_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Debug.Print KeyCode
    Dim SelPos As Long
    Dim wf As WorksheetFunction
    Set wf = Application.WorksheetFunction
    
    With m_TextBoxIP
        ' get block of cursor
        Dim CursorPos As Long
        CursorPos = Application.WorksheetFunction.Max(1, .SelStart)
    
        Dim BlockStart As Long
        BlockStart = InStrRev(.Value, ".", CursorPos) + 1
        If BlockStart <= 0 Then BlockStart = 1
        
        Dim BlockEnd As Long
        BlockEnd = InStr(CursorPos + 1, .Value, ".") - 1
        If BlockEnd <= 0 Then BlockEnd = Len(.Value)
        
        Dim Block As String
        Block = Mid$(.Value, BlockStart, BlockEnd - BlockStart + 1)
                
        Select Case KeyCode.Value
            Case vbKey0 To vbKey9, vbKeyNumpad0 To vbKeyNumpad9:
                If Len(Block) >= 3 And .SelLength = 0 Then
                    KeyCode.Value = 0
                End If
                
            Case vbKeyLeft, vbKeyRight:
            Case vbKeyInsert:
            Case vbKeyDelete:
                If .SelLength <> Len(.Value) Then
                    If .SelStart < BlockStart - 1 Or .SelStart + .SelLength > BlockEnd Then
                        KeyCode.Value = 0
                        Exit Sub
                    End If
                    
                    If .SelStart < Len(.Value) Then
                        If Mid$(.Value, .SelStart + 1, 1) = "." And Mid$(.Value, .SelStart, 1) <> "." Then
                            m_EditMode = True
                            SelPos = .SelStart
                            
                            .Value = Left$(.Value, .SelStart + 1) & Mid$(.Value, .SelStart + 3)
                            .SelStart = SelPos + 1
    
                            KeyCode.Value = 0
                            m_EditMode = False
                        End If
                    End If
                End If

            Case vbKeyHome, vbKeyEnd:
            Case vbKeyReturn, vbKeyEscape:

            Case vbKeyTab, vbKeyCapital:
            Case 190: 'dot
                Dim AmountOfDots As Long
                AmountOfDots = Len(.Value) - Len(Replace$(.Value, ".", vbNullString))
                
                ' allow only 3 dots and if there is no dot left cursor
                If Mid$(.Value, .SelStart, 1) = "." Or AmountOfDots >= 3 Then
                    KeyCode.Value = 0
                End If
                
            Case vbKeyBack:
                If .SelLength <> Len(.Value) Then
                    If .SelStart < BlockStart - 1 Or .SelStart + .SelLength > BlockEnd Then
                        KeyCode.Value = 0
                        Exit Sub
                    End If
                    
                    If .SelStart > 0 And .SelLength = 0 Then
                        If Mid$(.Value, .SelStart, 1) = "." And Mid$(.Value, .SelStart + 1, 1) <> "." Then
                            m_EditMode = True
                            
                            SelPos = .SelStart
                            If .SelStart = Len(.Value) Then
                                .Value = Left$(.Value, wf.Max(0, .SelStart - 2)) & Mid$(.Value, .SelStart + 1)
                            Else
                                .Value = Left$(.Value, wf.Max(0, .SelStart - 2)) & Mid$(.Value, .SelStart)
                                .SelStart = wf.Max(0, SelPos - 2)
                            End If
                            KeyCode.Value = 0
                            m_EditMode = False
                        End If
                    End If
                End If
                
            Case vbKeyV, vbKeyC: ' allow copy paste
                If Not Shift = 2 Then ' allow olny Ctrl+C and Ctrl+V
                    KeyCode.Value = 0
                End If
                
            Case Else: ' disallow all other keys that are not handled above
                KeyCode.Value = 0
                Exit Sub
        End Select
    End With
End Sub

