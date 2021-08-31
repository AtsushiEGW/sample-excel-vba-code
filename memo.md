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


