```vbs
Option Explicit

Sub test()
       Range("A1:J1") = "hello"
       'vba では同じオブジェクトの集まりをコレクションという
       'オブジェクトの指定は コレクション(オブジェクト)で指定 Workbooks("book1.xlsx") など
       'range オブジェクトのプロパティとしてオブジェクトを取得する物がある
       '.Interior など Range("A1").Interior.color = rgbLightBlue などとInteriorオブジェクトのcolorプロパティに値を設定する
       'method => .Select, .Copy, .Deleteなど
       'return object のメソッドもある Worksheets.Add とかはWorksheetオブジェクトを返すので
       'Worksheets.Add.Namae = "new_sheet"などと設定可能
       '引数が必要なメソッドは .メソッド 引数1, 引数２...のように指定 Range("A1").Insert xlShiftDown
       '複数引数がある場合は .Insert Shift:=xlShiftDownのように名前付き引数にするとよい
       'メソッドがオブジェクトを返す際の引数の指定は()内に記述する
       '.specialCells()は引数に指定した条件を満たすセルのオブジェクトを返す
       
       Range("A1:A5").SpecialCells(xlCellTypeBlanks).Interior.Color = rgbRed
       
       'プロパティに引数を指定するものもある .End(Direction)は終端のセルのオブジェクトを返す
       Range("A1").End(xlToRight).Select
       
       
       'データ型
       'String, Integer, Single(Float), Date, Object, Variant, Boolean
       
       'オブジェクトを変数に入れる場合は
       Dim rng As Range 'とオブジェクトの種類を指定してset=で格納 参照型
       Set rng = Range("A12")
       rng.Value = "xxx"
       Set rng = Nothing  '変数を使用し終えたら参照を外す
       '(Subが終了した時点で自動的に解除されるが、毎回書いておくとメモリ圧迫の心配がなくなる)
       
       '定数の宣言
       Const TAX As Single = 0.1
       
       '配列の宣言
       Dim arr(5) As Integer 'index0-5の６つの配列が作成される
       arr(0) = 100
       
       '動的配列
       Dim arr2() As Variant '最大インデックスを指定せず宣言
       ReDim arr2(1) ' 要素数を変更
       arr2(0) = "John"
       arr2(1) = "Lisa"
       ReDim Preserve arr2(2) '現在の要素をリセットせずにインデックスを変更
       arr2(2) = "Lucas"

       
       '演算子
       '10^3 = Math.pow(10, 3), 10\3 = Math.floor(10/3), 10 Mod 3 = 10%3
       
       '日付リテラル
       Range("A15").Value = #6/21/1989#
       Range("B15").Value = #11:45:00 PM#
       
       'オブジェクトの比較
       Set rng = Range("A12")
       MsgBox rng Is Range("A12") '参照が同じであればTrue
       MsgBox rng Is Nothing 'どこも参照していなければTrue
       
       'ワイルドカードとLike演算子
       Range("A16").Value = "Hello world"
       MsgBox Range("A16") Like "Hello*"
       
       'For Each ステートメント ー＞ コレクションや配列のイテレーション
       '各要素を格納する変数はObject, Variant型にする
       For Each rng In Range("A18:C20")
                rng.Value = rng.Address
       Next
       
       Dim elem As Variant
       For Each elem In Array("tokyo", "new york", "london")
            MsgBox "array element: " & elem
       Next

        
    
End Sub
```

```vbs

Option Explicit

Sub test()
       Range("A1:J1") = "hello"
       'vba では同じオブジェクトの集まりをコレクションという
       'オブジェクトの指定は コレクション(オブジェクト)で指定 Workbooks("book1.xlsx") など
       'range オブジェクトのプロパティとしてオブジェクトを取得する物がある
       '.Interior など Range("A1").Interior.color = rgbLightBlue などとInteriorオブジェクトのcolorプロパティに値を設定する
       'method => .Select, .Copy, .Deleteなど
       'return object のメソッドもある Worksheets.Add とかはWorksheetオブジェクトを返すので
       'Worksheets.Add.Namae = "new_sheet"などと設定可能
       '引数が必要なメソッドは .メソッド 引数1, 引数２...のように指定 Range("A1").Insert xlShiftDown
       '複数引数がある場合は .Insert Shift:=xlShiftDownのように名前付き引数にするとよい
       'メソッドがオブジェクトを返す際の引数の指定は()内に記述する
       '.specialCells()は引数に指定した条件を満たすセルのオブジェクトを返す
       
       Range("A1:A5").SpecialCells(xlCellTypeBlanks).Interior.Color = rgbRed
       
       'プロパティに引数を指定するものもある .End(Direction)は終端のセルのオブジェクトを返す
       Range("A1").End(xlToRight).Select
       
       
       'データ型
       'String, Integer, Single(Float), Date, Object, Variant, Boolean
       
       'オブジェクトを変数に入れる場合は
       Dim rng As Range 'とオブジェクトの種類を指定してset=で格納 参照型
       Set rng = Range("A12")
       rng.Value = "xxx"
       Set rng = Nothing  '変数を使用し終えたら参照を外す
       '(Subが終了した時点で自動的に解除されるが、毎回書いておくとメモリ圧迫の心配がなくなる)
       
       '定数の宣言
       Const TAX As Single = 0.1
       
       '配列の宣言
       Dim arr(5) As Integer 'index0-5の６つの配列が作成される
       arr(0) = 100
       
       '動的配列
       Dim arr2() As Variant '最大インデックスを指定せず宣言
       ReDim arr2(1) ' 要素数を変更
       arr2(0) = "John"
       arr2(1) = "Lisa"
       ReDim Preserve arr2(2) '現在の要素をリセットせずにインデックスを変更
       arr2(2) = "Lucas"

       
       '演算子
       '10^3 = Math.pow(10, 3), 10\3 = Math.floor(10/3), 10 Mod 3 = 10%3
       
       '日付リテラル
       Range("A15").Value = #6/21/1989#
       Range("B15").Value = #11:45:00 PM#
       
       'オブジェクトの比較
       Set rng = Range("A12")
       MsgBox rng Is Range("A12") '参照が同じであればTrue
       MsgBox rng Is Nothing 'どこも参照していなければTrue
       
       'ワイルドカードとLike演算子
       Range("A16").Value = "Hello world"
       MsgBox Range("A16") Like "Hello*"
       
       'For Each ステートメント ー＞ コレクションや配列のイテレーション
       '各要素を格納する変数はObject, Variant型にする
       For Each rng In Range("A18:C20")
                rng.Value = rng.Address
       Next
       
       Dim elem As Variant
       For Each elem In Array("tokyo", "new york", "london")
            MsgBox "array element: " & elem
       Next
End Sub

'プロシージャの呼び出し
Sub sayHallo()
    Call greet("good morning")
End Sub

Sub greet(greeting As String)
        MsgBox greeting & " Mr. Williams."
End Sub

Sub dbg()
    Dim i As Integer
    i = 1
    
    Do While Cells(18, i).Value <> ""
        Debug.Print i & " : " & Cells(18, i).Value
        i = i + 1
    Loop
    
End Sub

Sub operateWroksheet()
    MsgBox Selection.Address
    MsgBox ActiveCell
End Sub

Sub getCells()
        Cells.Clear
        Range(Cells(1, 1), Cells(3, 3)).Interior.Color = vbGreen
        
        Rows(3).Select
        
        Dim rng As Range
        Set rng = Range("A5:C10")
        rng.Cells(1, 1).Select 'rng範囲の(1,1)すなわちA5が取得される
End Sub

Sub readCsvFile()
        Cells.Clear
        Const workSpaceDir As String = "C:\Users\egwat\OneDrive\egwatsWorkSpace"
        Dim targetFile As String, filePath As String
        Dim max_n As Long
        
        '2次元配列の宣言
        Dim arr() As Variant

        filePath = workSpaceDir & "\VBA\data\*.csv"
        Debug.Print filePath
        targetFile = Dir(filePath)
        max_n = CreateObject("Scripting.FilesySystemObject").OpenTextFile(targetFile, 8).Line 'ファイルの行数取得
        Debug.Print max_n
        
        ReDim ary(max_n - 1, 2) As Variant
        
        'csvファイルを配列へ格納する
        Open file For Input As #1
End Sub

Sub operateFiles()
        'Open file
        'Applicationオブジェクト.dialogs(index)
        
        'dialogsを使った開き方
        
        Application.Dialogs(xlDialogOpen).Show 'エクスプローラーみたいな奴が開く
        Application.Dialogs(xlDialogSaveAs).Show '名前を付けて保存
        
        'open ステートメントを使用
    

End Sub



ーーーーーー
Option Explicit

Sub test()
       Range("A1:J1") = "hello"
       'vba では同じオブジェクトの集まりをコレクションという
       'オブジェクトの指定は コレクション(オブジェクト)で指定 Workbooks("book1.xlsx") など
       'range オブジェクトのプロパティとしてオブジェクトを取得する物がある
       '.Interior など Range("A1").Interior.color = rgbLightBlue などとInteriorオブジェクトのcolorプロパティに値を設定する
       'method => .Select, .Copy, .Deleteなど
       'return object のメソッドもある Worksheets.Add とかはWorksheetオブジェクトを返すので
       'Worksheets.Add.Namae = "new_sheet"などと設定可能
       '引数が必要なメソッドは .メソッド 引数1, 引数２...のように指定 Range("A1").Insert xlShiftDown
       '複数引数がある場合は .Insert Shift:=xlShiftDownのように名前付き引数にするとよい
       'メソッドがオブジェクトを返す際の引数の指定は()内に記述する
       '.specialCells()は引数に指定した条件を満たすセルのオブジェクトを返す
       
       Range("A1:A5").SpecialCells(xlCellTypeBlanks).Interior.Color = rgbRed
       
       'プロパティに引数を指定するものもある .End(Direction)は終端のセルのオブジェクトを返す
       Range("A1").End(xlToRight).Select
       
       
       'データ型
       'String, Integer, Single(Float), Date, Object, Variant, Boolean
       
       'オブジェクトを変数に入れる場合は
       Dim rng As Range 'とオブジェクトの種類を指定してset=で格納 参照型
       Set rng = Range("A12")
       rng.Value = "xxx"
       Set rng = Nothing  '変数を使用し終えたら参照を外す
       '(Subが終了した時点で自動的に解除されるが、毎回書いておくとメモリ圧迫の心配がなくなる)
       
       '定数の宣言
       Const TAX As Single = 0.1
       
       '配列の宣言
       Dim arr(5) As Integer 'index0-5の６つの配列が作成される
       arr(0) = 100
       
       '動的配列
       Dim arr2() As Variant '最大インデックスを指定せず宣言
       ReDim arr2(1) ' 要素数を変更
       arr2(0) = "John"
       arr2(1) = "Lisa"
       ReDim Preserve arr2(2) '現在の要素をリセットせずにインデックスを変更
       arr2(2) = "Lucas"

       
       '演算子
       '10^3 = Math.pow(10, 3), 10\3 = Math.floor(10/3), 10 Mod 3 = 10%3
       
       '日付リテラル
       Range("A15").Value = #6/21/1989#
       Range("B15").Value = #11:45:00 PM#
       
       'オブジェクトの比較
       Set rng = Range("A12")
       MsgBox rng Is Range("A12") '参照が同じであればTrue
       MsgBox rng Is Nothing 'どこも参照していなければTrue
       
       'ワイルドカードとLike演算子
       Range("A16").Value = "Hello world"
       MsgBox Range("A16") Like "Hello*"
       
       'For Each ステートメント ー＞ コレクションや配列のイテレーション
       '各要素を格納する変数はObject, Variant型にする
       For Each rng In Range("A18:C20")
                rng.Value = rng.Address
       Next
       
       Dim elem As Variant
       For Each elem In Array("tokyo", "new york", "london")
            MsgBox "array element: " & elem
       Next
End Sub

'プロシージャの呼び出し
Sub sayHallo()
    Call greet("good morning")
End Sub

Sub greet(greeting As String)
        MsgBox greeting & " Mr. Williams."
End Sub

Sub dbg()
    Dim i As Integer
    i = 1
    
    Do While Cells(18, i).Value <> ""
        Debug.Print i & " : " & Cells(18, i).Value
        i = i + 1
    Loop
    
End Sub

Sub operateWroksheet()
    MsgBox Selection.Address
    MsgBox ActiveCell
End Sub

Sub getCells()
        Cells.Clear
        Range(Cells(1, 1), Cells(3, 3)).Interior.Color = vbGreen
        
        Rows(3).Select
        
        Dim rng As Range
        Set rng = Range("A5:C10")
        rng.Cells(1, 1).Select 'rng範囲の(1,1)すなわちA5が取得される
End Sub

Sub readCsvFile()
        Debug.Print "----------------- readCsvFile-----------------------"
        Cells.Clear
        Const workSpaceDir As String = "C:\Users\egwat\OneDrive\egwatsWorkSpace"
        Dim targetFile As String, filePath As String
        Dim i As Long
        Dim buf As String
        
        '2次元配列の宣言
        Dim arr As Variant

        filePath = workSpaceDir & "\VBA\data\*.csv"
        Debug.Print filePath
        targetFile = workSpaceDir & "\VBA\data\" & Dir(filePath)
        Debug.Print targetFile
        
        i = 1
        'open file
        Open targetFile For Input As #1
        Do While Not EOF(1)
                Debug.Print "-------------" & i & "---------------"
                Line Input #1, buf
                Debug.Print buf
                arr = Split(buf, ",")
                Debug.Print LBound(arr, 1)
                Debug.Print UBound(arr, 1)
                
               ' For i = LBound(arr, 1) To UBound(arr, 1)
               '         Debug.Print arr(i, 1)
               ' Next i
                
               ' For i = LBound(arr, 1) To UBound(arr, 1)
               '         Cells(i + 1, 1) = arr(i + 1, 1)
               ' Next i
               
               i = i + 1
        Loop
        
        Close #1
                
        
        
End Sub

Sub operateFiles()

        'Excelブックを開くには .Open()メソッドを使えばよいが、テキストファイルを開くときにはOpen ステートメントを使用する
        'Open file
        'Applicationオブジェクト.dialogs(index)
        
        'dialogsを使った開き方
        
        'Application.Dialogs(xlDialogOpen).Show 'エクスプローラーみたいな奴が開く
        'Application.Dialogs(xlDialogSaveAs).Show '名前を付けて保存
        
        'open ステートメントを使用
        
        
        'ここからopenステートメント
        'Open {fliename} For {開き方} As #{ファイル番号}
        '存在しないファイルを指定するとエラーになるのでDir()で確認してから
        
        MsgBox ThisWorkbook.Path
        
        Dim buf As String
        Open ThisWorkbook.Path & "/data/sample_csv_data.csv" For Input As #1
                Line Input #1, buf
                MsgBox buf
        Close #1

End Sub

Sub sampleArray()
        Dim sample_array() As Variant
        sample_array() = Range("A1:C27")
        MsgBox sample_array(3, 2)
End Sub






```

