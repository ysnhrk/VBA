'仕様書作成：編集中のブックのモジュール・プロシージャの一覧を名前順で出力
    '※実行前にExcel[オプション]>[セキュリティセンター]>[マクロの設定]
    '   >[VBAプロジェクトオブジェクトモデルへのアクセスを信頼する]をONにする
Sub Dev_MakeDevDoc()

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet: Set ws = Worksheets.Add     'ワークシートを挿入
    Dim lp_line As Long       'プロシージャ内の行番号
    Dim r As Long: r = 1      '書出先ワークシートの行番号
    Dim procName As String    'プロシージャ名
    Dim lastCol: lastCol = 5  '最終列
    
    'タイトル行
    ws.Cells(r, 1) = "モジュール名"
    ws.Cells(r, 2) = "プロシージャ名"
    ws.Cells(r, 3) = "コメント1行目"
    ws.Cells(r, 4) = "宣言文全体"
    ws.Cells(r, 5) = "コメント全体"
    Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = rgbMidnightBlue
    Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Font.Color = rgbWhite
    Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Font.Bold = True
    
    'VBComponents Collectionの添字をComponentの名前順でバブルソート
    Dim i, j, swap, comp
    Dim comps: Set comps = wb.VBProject.VBComponents
    Dim compsOrder: ReDim compsOrder(1 To comps.Count)
    For i = 1 To comps.Count
        compsOrder(i) = i
    Next i
    For i = 1 To comps.Count
        For j = comps.Count To i Step -1
            If comps(compsOrder(i)).name > comps(compsOrder(j)).name Then
                swap = compsOrder(i)
                compsOrder(i) = compsOrder(j)
                compsOrder(j) = swap
            End If
        Next j
    Next i
    
    'プロジェクト内のコンポーネントについてループ
    For i = 1 To comps.Count
        
        Set comp = comps(compsOrder(i))
        
        'コンポーネント名の書出
        If comp.name <> ws.name Then
            r = r + 1
            ws.Cells(r, 1) = comp.name
            Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = rgbPaleTurquoise
        End If
        
        With comp.CodeModule
            procName = ""
            'モジュール内のコードの各行についてループ
            For lp_line = 1 To .CountOfLines
                
                '新しいプロシージャに入ったら
                If procName <> .ProcOfLine(lp_line, 0) Then
                    
                    'プロシージャ名の設定
                    procName = .ProcOfLine(lp_line, 0)
                    r = r + 1
                    
                    'プロシージャ名
                    ws.Cells(r, 2) = procName
                     
                    'コメント
                    Dim commentLineCount, bodyLine
                    bodyLine = .ProcBodyLine(procName, 0)
                    commentLineCount = bodyLine - lp_line
                    If commentLineCount > 0 Then
                        ws.Cells(r, 3) = .Lines(lp_line, 1)                 'コメント1行目
                        ws.Cells(r, 5) = .Lines(lp_line, commentLineCount)  'コメント全体
                    End If
                    
                    '関数の宣言
                    Dim str: str = ""
                    Dim k: k = 0
                    Do
                        str = str & LTrim(.Lines(bodyLine + k, 1))
                        If Right(str, 2) <> " _" Then Exit Do
                        str = Left(str, Len(str) - 1)                       '行末の_を除く
                        k = k + 1
                    Loop
                    ws.Cells(r, 4) = str                                    '宣言全体
                    If str Like "Private *" Then
                        Range(ws.Cells(r, 2), ws.Cells(r, lastCol)).Interior.Color = rgbGainsboro
                    End If
                End If
            Next lp_line
        End With
    Next i
    
    Cells.WrapText = False
    Cells.VerticalAlignment = xlTop
    Columns(1).ColumnWidth = 12
    Columns(2).AutoFit
    Columns(3).AutoFit
    Columns(4).ColumnWidth = 12
    Columns(5).ColumnWidth = 12
    Rows.AutoFit
  
  MsgBox ws.name & "にプロシージャ一覧を作成しました。"

End Sub
