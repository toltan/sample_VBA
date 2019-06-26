Attribute VB_Name = "Module2"
Public Sub DataScraping() 'webページから情報を取得 2016/08/11
Dim nextro As Integer

Dim objE As InternetExplorer
Dim htmlDoc As HTMLDocument, htmlDoc2 As HTMLDocument, htmlDoc3 As HTMLDocument
Dim el As IHTMLElement, fl As IHTMLElement, fl2 As IHTMLElement
Dim colTR, colTH, colTD, h2tag As IHTMLElementCollection
Dim col2L, colScript, colRotation, colNoNo As IHTMLElementCollection
Dim urls As New Collection
Dim urcol As Variant, counter As Variant, varCounter As Variant
Dim urlNo As Integer, nextUrl As Integer
Dim a, b, c, col, ro As Integer

Set objE = CreateObject("Internetexplorer.Application") 'IEオブジェクトの生成
nextro = 2
urlNo = 4
nextUrl = 0

Worksheets("scraiping").Range(Cells(1, 2), Cells(Rows.Count, Columns.Count)).ClearContents
Application.ScreenUpdating = False '動作を非表示にする
objE.navigate "http://********************************" 'サイトURL

Do While objE.Busy = True Or objE.readyState < READYSTATE_COMPLETE '更新可能になるまでDoEventsを繰り返す
    DoEvents
Loop

Set htmlDoc = objE.document

For Each el In htmlDoc.Links 'ページのHTMLのaタグを収集
    urls.Add (el.href) '変数urlsにリンクURLを加えていく
Next el

For Each urcol In urls
    
    If objE.LocationURL = "http://daidata.goraggio.com/100185/list/?type=3&bt=21.30&f=2#HeaderWrapper&hist_num=1" Then
        objE.Visible = True 'スクレイピングが終わったらウィンドウを表示し、完了メッセージを表示
        MsgBox "scraiping complete!", vbInformation
        Exit Sub
    End If
        
    objE.navigate urls(urlNo + nextUrl) & "&hist_num=1"
    
    Do While objE.Busy = True Or objE.readyState < READYSTATE_COMPLETE '待ち
        DoEvents
    Loop
    'もしページがでなかったら処理をしないで次のページへ
    If Not objE.LocationURL Like "http://daidata.goraggio.com/100185/unit_list/?model*" Then
        GoTo err
    End If
        
    Do While objE.Busy = True Or objE.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
    
    Set htmlDoc2 = objE.document
    Set colTR = htmlDoc2.getElementsByTagName("tr") '現在のページのtrタグを取得
    Set colTH = htmlDoc2.getElementsByTagName("th") '現在のページのthタグを取得
    Set h2tag = htmlDoc2.getElementsByTagName("strong") '現在のページのstrongタグを取得
    
    col = 0
    Worksheets("scraiping").Cells(nextro - 1, 2).Value = h2tag(0).innerText
        
        For Each fl In colTR 'tr毎のデータを取得
            Set colTD = fl.getElementsByTagName("td") 'tr毎にtdを取り出す。
            On Error Resume Next
            a = 0
            col = col + 1
            ro = nextro
            
            For c = 0 To 4 '一つのtrにtdは9つあるため(正確には0番目に空のtdがある)
            
                Select Case c '1,2,3,4をシートに転記
                    Case c = 0, 1, 2, 3, 4 '
                    Worksheets("scraiping").Cells(ro, col).Value = colTD(a).innerText
                    ro = ro + 1 '行を1つずらす
                End Select
        
            a = a + 1 '次のtd要素
            Next c

        Next fl
        
    objE.navigate urls(urlNo + nextUrl) & "&hist_num=1&disp=2&graph=1"
    
    Do While objE.Busy = True Or objE.readyState < READYSTATE_COMPLETE '待ち
        DoEvents
    Loop
    
    Set htmlDoc3 = objE.document
    Set col2L = htmlDoc3.getElementById("Main-Contents")
    Set colScript = col2L.getElementsByTagName("script")
    Set colRotation = htmlDoc3.getElementsByClassName("Text-Green today")
    Set colNoNo = htmlDoc3.getElementsByClassName("Radius-Slot")
    col = 1
    a = 0
    
        For Each fl2 In colRotation
            col = col + 1
            Worksheets("scraiping").Cells(ro, col).Value = colRotation(a).innerHTML
            a = a + 1
        Next
        
    ro = ro + 1 '行を1つずらす
    col = 1
    a = 0
        
    For Each fl2 In colRotation
        col = col + 1
        counter = Split(colScript(a).innerHTML, Chr(10))
        counter = Split(counter(5), "]")
        counter = Split(counter(UBound(counter) - 2), ",")
        Worksheets("scraiping").Cells(ro, col).Value = counter(UBound(counter))
        a = a + 1
    Next
    
    nextro = ro + 2
    
err:
    nextUrl = nextUrl + 1
    Application.Wait Now + TimeValue("00:00:05")
    DoEvents '待ち
    
Next urcol

Application.ScreenUpdating = True '動作を表示する

End Sub



