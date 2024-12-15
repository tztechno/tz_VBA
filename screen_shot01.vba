Sub CaptureScreenshot()
    Dim ie As Object
    Dim ws As Worksheet
    Dim fileName As String
    
    ' Internet Explorer を起動
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.Navigate "https://www.example.com" ' 対象URL
    
    ' ページのロードを待機
    Do While ie.Busy Or ie.ReadyState <> 4
        DoEvents
    Loop

    ' スクリーンショットを撮る
    Application.SendKeys "{PRTSC}" ' Print Screen キーを送信
    
    ' スクリーンショットを貼り付ける
    Set ws = ThisWorkbook.Sheets(1)
    ws.PasteSpecial Format:="Bitmap"
    
    ' Internet Explorer を閉じる
    ie.Quit
    Set ie = Nothing
End Sub
