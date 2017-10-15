Attribute VB_Name = "MyOutlineModule"
' 変数宣言を義務付ける
Option Explicit

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const SleepWaitMs As Integer = 0 ' 開発時に値を大きくすると四角のスキャンの様子が見えやすい

Private Type Rect
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type

Private Function RangeToRect(r1 As Range, ByRef r2 As Rect)
' Range が表現する範囲を Rect に保存する
    r2.Left = r1(1).Column
    r2.Top = r1(1).Row
    r2.Right = r1(r1.Count).Column
    r2.Bottom = r1(r1.Count).Row
End Function

Private Sub SelectRect(r As Rect)
' Rect が表現する範囲を選択する
    Range(Cells(r.Top, r.Left), Cells(r.Bottom, r.Right)).Select
    DoEvents
    Sleep SleepWaitMs
End Sub

Private Function CompareRect(r1 As Rect, r2 As Rect) As Boolean
' Rect の範囲が同じならゼロを返す
    If r1.Left = r2.Left And r1.Top = r2.Top And r1.Right = r2.Right And r1.Bottom = r2.Bottom Then
        CompareRect = 0
    Else
        CompareRect = 1
    End If
End Function

Private Function DebugPrintRect(r As Rect)
' Rect の範囲をイミディエイトに出力する
    Dim Msg As String
    Msg = "Rect: (" & r.Left & ", " & r.Top & "), (" & r.Right & ", " & r.Bottom & ")"
    Debug.Print Msg
End Function

Private Sub EraseBorder(r As Range)
'
' 範囲内の罫線を消す
'
    With r
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
End Sub

Private Sub DrawOutsideBorder(r As Range, Optional LineStyle = xlContinuous, Optional Weight = xlThin)
' 罫線（外枠）を描く
    With r
        With .Borders(xlEdgeLeft)
            .LineStyle = LineStyle
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = Weight
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = LineStyle
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = Weight
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = LineStyle
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = Weight
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = LineStyle
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = Weight
        End With
    End With
End Sub

Private Sub DrawInsideBorder(r As Range, Optional LineStyle = xlContinuous, Optional Weight = xlThin)
' 罫線（格子）を描く
    DrawOutsideBorder r, LineStyle, Weight
    With r
        With .Borders(xlInsideVertical)
            .LineStyle = LineStyle
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = Weight
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = LineStyle
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = Weight
        End With
    End With
End Sub

Private Function CellIsEmpty(Left As Integer, Top As Integer) As Boolean
' セルが空白か調べる
    CellIsEmpty = Cells(Top, Left).Value = ""
End Function

Private Function CellIsNotEmpty(Left As Integer, Top As Integer) As Boolean
' セルが空白ではないか調べる
    CellIsNotEmpty = Not CellIsEmpty(Left, Top)
End Function

Private Function SplitRectTop(ParentRect As Rect, LastRect As Rect, Left As Integer, ByRef Top As Integer) As Boolean
' Topを計算する
    Dim t As Integer
    
    If LastRect.Bottom = 0 Then
        Debug.Assert CellIsNotEmpty(ParentRect.Left, ParentRect.Top)
        SplitRectTop = True
        Top = ParentRect.Top
    Else
        t = LastRect.Bottom + 1
        If t > ParentRect.Bottom Then
            SplitRectTop = False
            Exit Function
        End If
        If CellIsNotEmpty(Left, t) Then
            SplitRectTop = True
            Top = t
            Exit Function
        End If
    
        Do While t <= ParentRect.Bottom
            If CellIsNotEmpty(Left, t + 1) Then
                Exit Do
            End If
            t = t + 1
        Loop
        
        If t > ParentRect.Bottom Then
            SplitRectTop = False
            Exit Function
        End If
        
        SplitRectTop = True
        Top = t
    End If
End Function

Private Function SplitRectBottom(ParentRect As Rect, LastRect As Rect, Left As Integer, Top As Integer, ByRef Bottom As Integer) As Boolean
' Bottom を計算する
    Dim b As Integer

    If LastRect.Top = 0 Then
        Debug.Assert CellIsNotEmpty(ParentRect.Left, ParentRect.Top)
        
        b = Top
        
        Do While b <= ParentRect.Bottom
            If CellIsNotEmpty(Left, b + 1) Then
                Exit Do
            End If
            b = b + 1
        Loop
        
        If b > ParentRect.Bottom Then
            SplitRectBottom = True
            Bottom = ParentRect.Bottom
            Exit Function
        End If
        
        SplitRectBottom = True
        Bottom = b
    Else
        b = LastRect.Bottom + 1
        If LastRect.Bottom + 1 > ParentRect.Bottom Then
            SplitRectBottom = False
            Exit Function
        End If
        
        Do While b <= ParentRect.Bottom
            If CellIsNotEmpty(Left, b + 1) Then
                Exit Do
            End If
            b = b + 1
        Loop
        
        If b > ParentRect.Bottom Then
            SplitRectBottom = True
            Bottom = ParentRect.Bottom
            Exit Function
        End If
        
        SplitRectBottom = True
        Bottom = b
    End If
End Function

Private Function FindSplitRect(ParentRect As Rect, LastRect As Rect, ByRef ResultRect As Rect) As Boolean
' 四角を水平に分割する
    Dim r As Rect
    
    ' 範囲左上に値があることを確かめる
    Debug.Assert CellIsNotEmpty(ParentRect.Left, ParentRect.Top)
    
    ' 座標を決める
    r.Left = ParentRect.Left
    If Not SplitRectTop(ParentRect, LastRect, r.Left, r.Top) Then
        FindSplitRect = False
        Exit Function
    End If
    If Not SplitRectBottom(ParentRect, LastRect, r.Left, r.Top, r.Bottom) Then
        FindSplitRect = False
        Exit Function
    End If
    r.Right = ParentRect.Right
    
    ' 結果が ParentRange と同じなら見つからない判定とする
    If CompareRect(ParentRect, r) = 0 Then
        FindSplitRect = False
        Exit Function
    End If
    
    ' 結果を格納する
    FindSplitRect = True
    ResultRect = r

End Function

Private Function FindSubRect(ParentRect As Rect, ByRef ResultRect As Rect, DataHeadLeft As Integer) As Boolean
' 四角の幅を狭めつつ、より小さな四角を探す
    Dim r As Rect
    
    r.Right = ParentRect.Right
    r.Bottom = ParentRect.Bottom
    
    For r.Left = ParentRect.Left To r.Right
        For r.Top = ParentRect.Top To r.Bottom
            If r.Left = ParentRect.Left And r.Top = ParentRect.Top Then
                ' do nothing
            ElseIf CellIsNotEmpty(r.Left, r.Top) Then
                FindSubRect = True
                ResultRect = r
                Exit Function
            End If
        Next
    Next
    FindSubRect = False
End Function

Private Sub DrawRectRecursive(ParentRect As Rect, DataHeadLeft As Integer)
' 再帰的に四角を描く
    Dim y As Integer
    Dim Found As Boolean
    Dim LastRect As Rect
    Dim SplitRect As Rect
    Dim SubRect As Rect

    ' 範囲左上に値があることを確かめる
    Debug.Assert CellIsNotEmpty(ParentRect.Left, ParentRect.Top)

    ' 水平に四角を分割する
    Do While FindSplitRect(ParentRect, LastRect, SplitRect)
        ' 外枠を描く
        SelectRect SplitRect
        DrawOutsideBorder Selection
    
        ' より小さな四角を探す
        If FindSubRect(SplitRect, SubRect, DataHeadLeft) Then
            SelectRect SubRect
            If SubRect.Left < DataHeadLeft Then
                DrawOutsideBorder Selection
                DrawRectRecursive SubRect, DataHeadLeft
            Else
                DrawInsideBorder Selection
            End If
            
        End If
    
        LastRect = SplitRect
    Loop
    If LastRect.Top = 0 Then
        ' 外枠を描かず、より小さな四角を探す
        If FindSubRect(ParentRect, SubRect, DataHeadLeft) Then
            SelectRect SubRect
            If SubRect.Left < DataHeadLeft Then
                DrawOutsideBorder Selection
                DrawRectRecursive SubRect, DataHeadLeft
            Else
                DrawInsideBorder Selection
            End If
        End If
    End If
End Sub

Private Sub Main()
    Dim HeadWidth As Integer
    Dim CatHeadWidth As Integer
    Dim DataHeadLeft As Integer
    Dim r As Rect
    
    Debug.Assert TypeName(Selection) = "Range"  ' セル範囲か確かめる
    Debug.Assert Selection.Areas.Count = 1      ' セル範囲が１つであることを確かめる
    Debug.Assert Selection.Count > 1            ' 複数セルを選択していることを確かめる
    
    ' カテゴリ列数を決める
    RangeToRect Selection, r
    HeadWidth = r.Right - r.Left + 1
    CatHeadWidth = Val(InputBox("分類列の列数を入力してください。\nデータ列は罫線を格子で描きます。", "分類列とデータ列の境界", HeadWidth))
    Debug.Assert CatHeadWidth > 0
    Debug.Assert CatHeadWidth <= HeadWidth
    DataHeadLeft = r.Left + CatHeadWidth
    
    ' 選択範囲内の罫線を消す
    EraseBorder Selection
    
    ' 外枠を描画する
    DrawOutsideBorder Selection
    
    ' 再帰的に描画する
    DrawRectRecursive r, DataHeadLeft
    
    ' もとの選択範囲を選択しなおす
    SelectRect r

End Sub

Public Sub MyOutline_選択範囲をいい感じに罫線を引く()
'Public Sub MyOutline_Excelで表計算しにくくなる罫線を引く()
    Main
End Sub
