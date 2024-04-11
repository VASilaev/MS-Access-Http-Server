Attribute VB_Name = "ServerProcessor"
Option Compare Database
Const nPort = 5678

Dim objServerProcess


Public Sub StartServer()
  'Запуск сервера
  Set WshShell = CreateObject("WScript.Shell")
  Set objServerProcess = WshShell.Exec("""" & CurrentProject.Path & "\AutoHotkeyU32.exe"" " & _
               """" & CurrentProject.Path & "\AccessServer.ahk"" " & _
               """" & CurrentProject.Path & "\" & CurrentProject.Name & """ " & _
               nPort)
End Sub

Public Sub StopServer()
  'Остановка сервера
  If Not IsEmpty(objServerProcess) Then
    objServerProcess.Terminate
    objServerProcess = Empty
  End If
End Sub

'прямой запрос в Access
Public Function QR(context)
  'Функция обрабатывает конечную точку /access/qr
  Dim svQuery
  'Параметры получаем следующим образом
  svQuery = context("sys$queries")(0)("q")
  
  'Здесь производим обработку. В идеале быстро куда то положили и забыли
  Debug.Print svQuery, context("sys$pathWithParam")
  CurrentDb.Execute "insert into tRequest (sQuery) values ('" & Replace(svQuery, "'", "''") & "')"
  
  'Нужно вернуть HTML страничку.
  QR = "<html><body>OK</body></html>"
  
  'По умолчанию все страничку возвращаются с ответом 200, если нужно изменить статус:
  'context("sys$status") = 500
  
  'Для передачи дополнительных заголовков использую следующий метод
  'context("sys$response")(0)("headers")("Content-type") = "text/html"
End Function

