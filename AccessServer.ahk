#Persistent
#SingleInstance, force
SetBatchLines, -1

#include <ActiveScript>

;Ведение лога
WriteLog(text) {
 FileAppend, % A_NowUTC ": " text "`n", % A_ScriptDir "/logfile.txt"
}

WriteLog("Подключаем JS для работы c URL")
java := new ActiveScript("JScript")
java.eval("
(
//Словарь перекодировки Win1251 в UTF-8
  var win1251toutf8 = {
    '`%A5'`: '`%D2`%90', '`%a5'`: '`%D2`%90', '`%A8'`: '`%D0`%81', '`%a8'`: '`%D0`%81', '`%AF'`: '`%D0`%87', '`%af'`: '`%D0`%87', '`%B2'`: '`%D0`%86', '`%b2'`: '`%D0`%86',
    '`%B3'`: '`%D1`%96', '`%b3'`: '`%D1`%96', '`%B4'`: '`%D2`%91', '`%b4'`: '`%D2`%91', '`%B8'`: '`%D1`%91', '`%b8'`: '`%D1`%91', '`%BF'`: '`%D1`%97', '`%bf'`: '`%D1`%97',
    '`%C0'`: '`%D0`%90', '`%c0'`: '`%D0`%90', '`%C1'`: '`%D0`%91', '`%c1'`: '`%D0`%91', '`%C2'`: '`%D0`%92', '`%c2'`: '`%D0`%92', '`%C3'`: '`%D0`%93', '`%c3'`: '`%D0`%93',
    '`%C4'`: '`%D0`%94', '`%c4'`: '`%D0`%94', '`%C5'`: '`%D0`%95', '`%c5'`: '`%D0`%95', '`%C6'`: '`%D0`%96', '`%c6'`: '`%D0`%96', '`%C7'`: '`%D0`%97', '`%c7'`: '`%D0`%97',
    '`%C8'`: '`%D0`%98', '`%c8'`: '`%D0`%98', '`%C9'`: '`%D0`%99', '`%c9'`: '`%D0`%99', '`%CA'`: '`%D0`%9A', '`%ca'`: '`%D0`%9A', '`%CB'`: '`%D0`%9B', '`%cb'`: '`%D0`%9B',
    '`%CC'`: '`%D0`%9C', '`%cc'`: '`%D0`%9C', '`%CD'`: '`%D0`%9D', '`%cd'`: '`%D0`%9D', '`%CE'`: '`%D0`%9E', '`%ce'`: '`%D0`%9E', '`%CF'`: '`%D0`%9F', '`%cf'`: '`%D0`%9F',
    '`%D0'`: '`%D0`%A0', '`%d0'`: '`%D0`%A0', '`%D1'`: '`%D0`%A1', '`%d1'`: '`%D0`%A1', '`%D2'`: '`%D0`%A2', '`%d2'`: '`%D0`%A2', '`%D3'`: '`%D0`%A3', '`%d3'`: '`%D0`%A3',
    '`%D4'`: '`%D0`%A4', '`%d4'`: '`%D0`%A4', '`%D5'`: '`%D0`%A5', '`%d5'`: '`%D0`%A5', '`%D6'`: '`%D0`%A6', '`%d6'`: '`%D0`%A6', '`%D7'`: '`%D0`%A7', '`%d7'`: '`%D0`%A7',
    '`%D8'`: '`%D0`%A8', '`%d8'`: '`%D0`%A8', '`%D9'`: '`%D0`%A9', '`%d9'`: '`%D0`%A9', '`%DA'`: '`%D0`%AA', '`%da'`: '`%D0`%AA', '`%DB'`: '`%D0`%AB', '`%db'`: '`%D0`%AB',
    '`%DC'`: '`%D0`%AC', '`%dc'`: '`%D0`%AC', '`%DD'`: '`%D0`%AD', '`%dd'`: '`%D0`%AD', '`%DE'`: '`%D0`%AE', '`%de'`: '`%D0`%AE', '`%DF'`: '`%D0`%AF', '`%df'`: '`%D0`%AF',
    '`%E0'`: '`%D0`%B0', '`%e0'`: '`%D0`%B0', '`%E1'`: '`%D0`%B1', '`%e1'`: '`%D0`%B1', '`%E2'`: '`%D0`%B2', '`%e2'`: '`%D0`%B2', '`%E3'`: '`%D0`%B3', '`%e3'`: '`%D0`%B3',
    '`%E4'`: '`%D0`%B4', '`%e4'`: '`%D0`%B4', '`%E5'`: '`%D0`%B5', '`%e5'`: '`%D0`%B5', '`%E6'`: '`%D0`%B6', '`%e6'`: '`%D0`%B6', '`%E7'`: '`%D0`%B7', '`%e7'`: '`%D0`%B7',
    '`%E8'`: '`%D0`%B8', '`%e8'`: '`%D0`%B8', '`%E9'`: '`%D0`%B9', '`%e9'`: '`%D0`%B9', '`%EA'`: '`%D0`%BA', '`%ea'`: '`%D0`%BA', '`%EB'`: '`%D0`%BB', '`%eb'`: '`%D0`%BB',
    '`%EC'`: '`%D0`%BC', '`%ec'`: '`%D0`%BC', '`%ED'`: '`%D0`%BD', '`%ed'`: '`%D0`%BD', '`%EE'`: '`%D0`%BE', '`%ee'`: '`%D0`%BE', '`%EF'`: '`%D0`%BF', '`%ef'`: '`%D0`%BF',
    '`%F0'`: '`%D1`%80', '`%f0'`: '`%D1`%80', '`%F1'`: '`%D1`%81', '`%f1'`: '`%D1`%81', '`%F2'`: '`%D1`%82', '`%f2'`: '`%D1`%82', '`%F3'`: '`%D1`%83', '`%f3'`: '`%D1`%83',
    '`%F4'`: '`%D1`%84', '`%f4'`: '`%D1`%84', '`%F5'`: '`%D1`%85', '`%f5'`: '`%D1`%85', '`%F6'`: '`%D1`%86', '`%f6'`: '`%D1`%86', '`%F7'`: '`%D1`%87', '`%f7'`: '`%D1`%87',
    '`%F8'`: '`%D1`%88', '`%f8'`: '`%D1`%88', '`%F9'`: '`%D1`%89', '`%f9'`: '`%D1`%89', '`%FA'`: '`%D1`%8A', '`%fa'`: '`%D1`%8A', '`%FB'`: '`%D1`%8B', '`%fb'`: '`%D1`%8B',
    '`%FC'`: '`%D1`%8C', '`%fc'`: '`%D1`%8C', '`%FD'`: '`%D1`%8D', '`%fd'`: '`%D1`%8D', '`%FE'`: '`%D1`%8E', '`%fe'`: '`%D1`%8E', '`%FF'`: '`%D1`%8F', '`%ff'`: '`%D1`%8F'
  }`;
      
//Раскодирует URI в нормальный текст. Если не удалось расшифровать текст сразу то предпринимается попытка преобразовать кодировку win1251 в utf-8 начиная с последнего компонента пути.  
function DecodeURI(s){
  try{
    return decodeURIComponent(s.replace(/\+/g, ' '))`;
  } catch (err) {
    s = s.replace(/\+/g, ' ')`;
    var a = s.split('/')`;      
    for (var i = a.length - 1`;i>=0`;i--) {
      a[i] = a[i].replace(/`%[\da-fA-F]{2}/g, function($0){return win1251toutf8[$0] || $0})
      try{return decodeURIComponent(a.join('/'))`;} catch (err) {void(0)`;}                
    }      
    return s`;            
  }
}`; 
//Кодирует ссылку в URI
function EncodeURI(s){return encodeURIComponent(s).replace(/'/g,""`%27"")`;}`;
)")

java["WriteLog"] := Func("WriteLog")  

class Uri
{
    Decode(str) {
        global java
        str := java.DecodeURI(str)
        Return, str
    }

    Encode(str) {
        global java
        str := java.EncodeURI(str)
        Return, str 
    }
}

WriteLog("Подключаем VBS для создания линка на Access")

vb := new ActiveScript("VBScript")

vb.Exec("
(
Dim  Access, CurrentContext, GlobalContext, ObjJava, ObjFSO, nDbgLevel, sPageContent, sCurrentPath, sLastFile, sBlockCode
Set ObjFSO = CreateObject(""Scripting.FileSystemObject"")
Set GlobalContext = CreateObject(""Scripting.Dictionary""): GlobalContext.CompareMode = 1

function dbg(v, l)
  if l < 0 then nDbgLevel = nDbgLevel + l
  if nDbgLevel < 0 then nDbgLevel = 0
  Dim Line: For each Line in split(v, vbCrLf): ObjJava.WriteLog(Space(nDbgLevel) & Line): Next
  if l > 0 then nDbgLevel = nDbgLevel + l
End function

Function AssignJava(JavaScript): set ObjJava = JavaScript: end function

Function InitAccess(spPath):  set Access = GetObject(spPath): end function

Function HTML(byVal spText, spTag, additional)
  spText = replace(replace(replace(replace(replace(replace(spText, ""&"", ""&amp;""), ""<"", ""&lt;""), "">"", ""&gt;""), """""""", ""&quot;""), vbcrlf, vbcr), vblf, vbcr)
  if spTag & """" <> """" then 
    dim sLine: for each sLine in split(spText, vbcr): HTML = HTML & ""<"" & spTag & additional & "">"" & sLine & ""</"" & spTag & "">"" & vbcr:  next
  else 
    HTML = replace(spText,vbcr,""</br>"")
  end if  
end function

Function ErrorPage(byref tpdic, sError, sContext)
  tpdic(""sys$status"") = 500
  ErrorPage = ""<html><head><meta http-equiv=""""Content-Type"""" content=""""text/html; charset=utf-8""""></head><body>"" & ""Error = "" & HTML(sError,"""")
  if  sContext <> """" then  ErrorPage = ErrorPage & ""<pre><code>"" & sContext & ""</pre></code>""
  ErrorPage = ErrorPage & ""</body></html>""  
end function

function fs_base(a): fs_base = ObjFSO.GetBaseName(a): end Function

function InitContext(req, res)
  dim localContext
  set localContext = CreateObject(""Scripting.Dictionary"")
  localContext.CompareMode = 1
  localContext(""sys$status"") = 200
  localContext(""sys$queries"") = array(req.queries)
  localContext(""sys$request"") = array(req)
  localContext(""sys$pathWithParam"") = req.pathWithParam
  localContext(""sys$response"") = array(res)
  set InitContext = localContext
end function

Function FinalizeContext(context)  
  FinalizeContext = cInt(context(""sys$status""))
  context(""sys$queries"") = array()
  context(""sys$request"") = array()
  context.removeall
  set context = nothing    
end Function

function ProcessRequest(req, res, server)
  dim localContext
  set localContext = InitContext(req, res)
  on error resume next: err.clear  
  s = Access.run (fs_base(req.path), localContext)    
  if err.number > 0 then
    dim sError: s = err.description
    on error goto 0
    s = ErrorPage(localContext, s, """")
  end if
  on error goto 0     
  res.SetBodyText(s)
  ProcessRequest = FinalizeContext(localContext) 
end function 

function Include(byval sPath)
  dim  content, i, tmp, LocalContent, CurContent, PrevsCurrentPath
  sPath = replace(sPath,""/"",""\""): if left(sPath,2) = "".\"" then sPath = sCurrentPath & mid(sPath,2)
  
  sLastFile = sPath
  PrevsCurrentPath = sCurrentPath: sCurrentPath = ObjFSO.GetParentFolderName(sPath)
  content = split(ObjFSO.OpenTextFile(sPath).ReadAll,""<`%"")  
  write content(0)
  for i = 1 to ubound(content)
    LocalContent = sPageContent:  sPageContent = """": tmp = split(content(i),""`%>"")
    sBlockCode = tmp(0)
    execute sBlockCode
    sBlockCode = """"
    CurContent = sPageContent: sPageContent = LocalContent: Write CurContent: CurContent = """": LocalContent = """"
    if uBound(tmp) > 0 then Write tmp(1)
  next
  content = Empty
  sCurrentPath = PrevsCurrentPath
end function

function ProcessVBFile(req, res, server)
  set CurrentContext = InitContext(req, res)
  on error resume next: err.clear  
  Include (req.queries(""path""))
  if err.number > 0 then
    dim sError: sPageContent = err.Source & "": "" & err.description & vbCrLf & ""Source: "" & sLastFile
    on error goto 0
    sPageContent = ErrorPage(CurrentContext, sPageContent, sBlockCode)
  end if
  on error goto 0      
  res.SetBodyText(sPageContent): sPageContent = """"    
  ProcessVBFile = FinalizeContext(CurrentContext)   
  set CurrentContext = Nothing
end function


sub Write(Content)
  sPageContent = sPageContent & Content
end sub

)")

WriteLog("Связываем JS, VBS и Access")
vb.AssignJava(java)
Param=%1%
WriteLog("Connect to Access: " . param)

vb.InitAccess(param)

SplitPath param,,StaticFile

;Дальше идет настройка путей
;paths - Содержит адреса которые обрабатываются как есть 1-й приоритет
paths := {}
paths["/"] := [Func("NotFound"), ""]
paths["404"] := [Func("NotFound"), ""]

;paths - Специальные расширения файлов для обработки в special у которых особый обработчик. 
;Массив состоит из 3-х параметров, Функция обработчик, дополнительные параметры, объект с экстра параметрами,
;Экстра параметры копируются в queries. Например Object("language","pgsql")
;Так же при обработки специального расширения в queries будет добавлена переменная path, содержащая путь до файла (реального местоположения)

;special - Обработчик путей по маске 
special := {}
special["i)^/access/"] := [Func("ProcessRequest"),""]

WriteLog("Static File from: " . StaticFile . "\static\")
specialExt := {}
specialExt["vb"] := [Func("ProcessVBFile"), "", ""]
special["i)^/static/"] := [Func("Resource"),[StaticFile,specialExt]]

;Конфигурация сервера
server := new HttpServer()
server.LoadMimes(A_ScriptDir . "/mime.types")
server.SetPaths(paths)
server.SetSpecial(special)

;Стартуем сервер
Param=%2%
Param:=0+Param
WriteLog("Start server on Port: " . param)
server.Serve(Param)
return 

NotFound(ByRef req, ByRef res, ByRef server, Dummy) {
;Заглушка для страницы "Не найдено"
  req.queries["path"]:= req.path
  global StaticFile
  if (FileExist(StaticFile . "\static\404.html")) {
    server.ServeFile(res, StaticFile . "\static\404.html")
  } else {
    res.SetBodyText("Page not found")
  }
  res.status := 404   
}


Resource(ByRef req, ByRef res, ByRef server, param) {
;Отдает локальные файлы
;#param param - Массив дополнительных параметров. 1-й элемент база на локальном компьютере, 2-й Ассоциативный массив с специальными расширениями

  global vb
  base := param[1]
  file := base . req.path  
  WriteLog("Process file " file)
  if (InStr( FileExist(file), "D")) {
    NotFound(req, res, server, file) 
    return
  } else {
    if (not FileExist(file)) {
      NotFound(req, res, server, file)
    } else {
      ;Специальные расширения
      if (param[2]) {      
        SplitPath, file,,, ext
        func := param[2][ext]
        if (func) {
          req.queries["path"] := file
          extraParam := func[3]
          if (extraParam) {
            for key, value in extraParam
               req.queries[key] := value
          }
          func[1].(req, res, server, func[2])
          return
        }
      }    
      server.ServeFile(res, file)
      ;dummy := vb.CheckExtraHeaders(res, file)
      res.status := 200  
    }
  }
}

ProcessRequest(ByRef req, ByRef res, ByRef server, dummy) {
;Передает обработку в ProcessRequest
;#param dummy - Массив дополнительных параметров. Не используется 
  global vb
  v := vb.ProcessRequest(req, res, server)
  if(v){
    res.status := v 
  } else {
    NotFound(req, res, server, "")
  }
}

ProcessVBFile(ByRef req, ByRef res, ByRef server, dummy) {
;Передает обработку в ProcessVBFile
;#param dummy - Массив дополнительных параметров. Не используется 
  global vb
  v := vb.ProcessVBFile(req, res, server)
  if(v){
    res.status := v 
  } else {
    NotFound(req, res, server, "")
  }
}

#include <AHKsock>
#include <AHKhttp>
