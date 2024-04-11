<HTML>
<head><meta http-equiv="Content-Type" content="text/HTML; charset=utf-8">
<style>
 td {
  border:1px solid;   
 }
</style>

</head>
  <body>
    <H2>Структура файлов</H2>
    
    <dl>
    <dt>AutoHotkeyU32.exe</dt><dd> Интерпретатор файлов AHK</dd>
    <dt>AccessServer.ahk</dt><dd> Основной файл сценария настройки Web сервера</dd>
    <dt>mime.types</dt><dd> Ассоциации расширений файлов с MIME</dd>
    <dt>\lib\ActiveScript.ahk</dt><dd> Библиотека AHK позволяющая внедрять скрипты VBS и JS</dd>
    <dt>\lib\AHKsock.ahk</dt><dd> Библиотека AHK для работы с сокетами, нужна для AHKhttp</dd>
    <dt>\lib\AHKhttp.ahk</dt><dd> Библиотека HTTP сервера</dd>
    <dt>\static\404.html</dt><dd> Заглушка для "Страница не найдена"</dd>
    <dt>\static\*</dt><dd> Файлы расположенные в этой папке будут доступны по адресу /static/*</dd>
    </dl>
    
  <H2>Настройка сервера</H2>
  
  <p>Порт по которому будет приниматься запросы задается в Модуле ServerProcessor, константа nPort</p>
    
  <p>Остальные настройки производятся в файле AccessServer.ahk</p>

  <p>Сервер может обрабатывать два вида запросов точный и по регулярному выражению. Сначала проверяется список точных соответствий, затем по регулярному выражению, если соответствий не найдено, то выдается страница /404.</p>
  
  <p>Точные запросы собираются в переменную <code>paths</code>  и затем передаются в объект сервера <code>server.SetPaths(paths)</code></p>
  
  <p>Paths представляет собой ассоциативный массив, где ключ представляет собой точный путь до обрабатываемого ресурса, а значение является массивом из двух значений. Первым элементом идет ссылка на функцию AHK - обработчик запроса на уровне AHK. Второй параметр (экстра параметр) передается как четвертый параметр в функцию обработчик.</p>
  
  <p>Запросы по регулярному выражению собираются в переменную <code>special</code>  и затем передаются в объект сервера <code>server.SetSpecial(special)</code></p>

  <p>special представляет собой ассоциативный массив, где ключ представляет собой регулярное выражение, используемое как шаблон пути до обрабатываемого ресурса, а значение является массивом из двух значений. Первым элементом идет ссылка на функцию AHK - обработчик запроса на уровне AHK. Второй параметр (экстра параметр) передается как четвертый параметр в функцию обработчик.</p>
     
  <h3>Преднастроенные функции обработчики</h3>
  
  <h4>NotFound</h4>
  
  <p>Заглушка для "Страница не найдена". При наличии файла /static/404.html, выводит его. Если файла нет, то выводит фразу "Page not found"</p>   
        
  <h4>Resource</h4>
  
  <p>Передает файл с диска клиенту. При настройке экстра параметр должен содержать массив из двух элементов. 
  
  <ul>
  <li>Первый элемент это базовый путь, который дописывается в начала запроса. Например базовый путь "с:\temp", запрошен был файл "/static/404.html", клиенту будет передан файл "с:\temp\static\404.html".</li> 
  <li>Второй элемент - это обработчик специальных расширений файлов, который представляет из себя ассоциативный массив, где в качестве ключа расширение файла (в нижнем регистр), а значение массив из 3-х элементов</li>
    <ul>  
      <li>Ссылка на функцию AHK - обработчик запроса на уровне AHK</li>
      <li>Значение передаваемое как 4-й параметр обработчика</li>
      <li>Ассоциативный массив дополнительных параметров, который копируется в <code>req.queries</code></li>
    </ul>
  </ul></p>   
        
  <h4>ProcessRequest</h4>
          
  <p>Вызывает функцию в Access, а ее результат отдает клиенту. Имя Функции берется как последняя компонента пути запроса. Например для <code>/access/qr</code> должна быть функция с именем QR. Функция должна быть с единственным параметром типа Variant. Через данный параметр передается контекст выполнения.</p>
  
<pre><code>Public Function QR(context)
  'Функция обрабатывает конечную точку /access/qr
  Dim svQuery
  'Параметры получаем следующим образом
  svQuery = context("sys$queries")(0)("s")
  
  'Здесь производим обработку. В идеале быстро куда то положили и забыли
  Debug.Print svQuery, context("sys$pathWithParam")
  CurrentDb.Execute "insert into tRequest (sQuery) values ('" & Replace(svQuery, "'", "''") & "')"
  
  'Нужно вернуть HTML страничку.
  QR = "<html><body>OK</body></html>"
  
  'По умолчанию все страничку возвращаются с ответом 200, если нужно изменить статус:
  'context("sys$status") = 500
  
  'Для передачи дополнительных заголовков использую следующий метод
  'context("sys$response")(0)("headers")("Content-type") = "text/html"
End Function</code></pre>  
  
  <h4>ProcessVBFile</h4>
          
  <p>Возвращает содержимое динамической страницы. Сильно урезанная технология ASP.</p>
  
    <p>Представляет собой обычный HTML файл в кодировке Win1251 (но клиенту файл уйдет в формате UTF-8, не забудьте указать кодировку в заголовке страницы). Файл может содержать скрипты.</br>
    
    <p>Скрипт помеченный &lt;% и %&gt; выполняется сервером.</br>
    
    <p>Для вывода текста из скрипта используйте функцию Write.</br>
      
    <p>Пример вывода строки: 
    <table style="width:100%"><tr><td style="width:50%">
<pre><code>Write "Hello World"</code></pre>    
    </td><td style="width:50%">
      <%Write "Hello World"%>
    </td></tr></table></p>
    
    <p>Для доступа к объектной модель Access.Application связанной БД используйте глобальную переменную Access: 
    <table style="width:100%"><tr><td style="width:50%">
<pre><code>Write Access.CurrentProject.Path & "\" & Access.CurrentProject.Name</code></pre>    
    </td><td style="width:50%">
      <%Write Access.CurrentProject.Path & "\" & Access.CurrentProject.Name%>
    </td></tr></table></p>
        
    <p>Можно выводить содержимое таблицы: </br>
    
    <table style="width:100%"><tr><td style="width:50%">
<pre><code>With Access.CurrentDb.OpenRecordset("select * from tRequest")
Do While Not .EOF
  Write "&lt;tr>&lt;td>" & .Fields("id").Value & "&lt;/td>&lt;td>" & _
        .Fields("dCreateDateTime").Value & "&lt;/td>&lt;td>" & _
        .Fields("sQuery").Value & "&lt;/td>&lt;/tr>"
  .MoveNext
Loop
.Close
End With  </code></pre>    
    </td><td style="width:50%">
         <table>
    <%
      With Access.CurrentDb.OpenRecordset("select * from tRequest")
        Do While Not .EOF
          Write "<tr><td>" & .Fields("id").Value & "</td><td>" & .Fields("dCreateDateTime").Value & "</td><td>" & .Fields("sQuery").Value & "</td></tr>"
          .MoveNext
        Loop
        .Close
      End With    
    %>
    </table>
    </td></tr></table></p>
        
</br></br>
    
    <p>Для сохранения переменных между блоками используйте CurrentContext <%Write "Var = " & CurrentContext("var")%>, содержимое не сохраняется между отрисовками страниц.
    <table style="width:100%"><tr><td style="width:50%">
<pre><code>CurrentContext("var") = CurrentContext("var") + 1
...
Write "Var = " & CurrentContext("var")
</code></pre>    
    </td><td style="width:50%">
      <%CurrentContext("var") = CurrentContext("var") + 1%>
      <%Write "Var = " & CurrentContext("var")%>
    </td></tr></table></p>    
    
    </br></br>
    
    <p>Для сохранения данных между отрисовками используйте GlobalContext: 
    
    <table style="width:100%"><tr><td style="width:50%">
<pre><code>GlobalContext("Counter") =  GlobalContext("Counter") + 1
Write "Счетчик запуска страницы = " & GlobalContext("Counter")</code></pre>    
    </td><td style="width:50%">
<%
GlobalContext("Counter") =  GlobalContext("Counter") + 1
Write "Счетчик запуска страницы = " & GlobalContext("Counter")%>
    </td></tr></table></p>   
        
     </br></br>
    
    <p>В движке используются так же еще переменные: ObjJava, ObjFSO, nDbgLevel, sPageContent, sCurrentPath, sLastFile, sBlockCode. Изменение их содержимого может привести к не предсказуемым последствиям.</br></br>
    
    <p>Для вывода строки содержащей спецсимволы HTML используйте функцию HTML.
    <table style="width:100%"><tr><td style="width:50%">
<pre><code>Write html("Мне нужны эти <угловые скобки>", "", "")</code></pre>    
    </td><td style="width:50%">
<%Write html("Мне нужны эти <угловые скобки>", "", "")%>
    </td></tr></table></p>       
    
     </br></br>
    
    <p>Если вторым параметром передать имя тега, то каждая строка будет завернута в этот тег, иначе символы перевода строк заменяются на BR:
    <table style="width:100%"><tr><td style="width:50%">
<pre><code>Write html("<Параграф 1>" & vbCr & "<Параграф 1>", "p", "")</code></pre>    
    </td><td style="width:50%">
<%Write html("<Параграф 1>" & vbCr & "<Параграф 1>", "p", "")%>
    </td></tr></table></p>        
    
    
    
    <p>Третьим параметром передается дополнительные атрибуты открывающего тега
    <table style="width:100%"><tr><td style="width:50%">
<pre><code>Write html("<Блок 1>" & vbCr & "<Блок 1>", "span", " style=""background-color: #0090f0;""")</code></pre>    
    </td><td style="width:50%">
<%Write html("<Блок 1>" & vbCr & "<Блок 1>", "span", " style=""background-color: #0090f0;""")%>
    </td></tr></table></p>  
    
        
    </br></br>
    
    <p>Функция Include интегрирует в текущий файл содержимое указанного. Для указания относительного пути начинайте имя с ".\". Относительный путь считается от текущего обрабатываемого файла.

    <table style="width:100%"><tr><td style="width:50%">
<pre><code>Include ".\IncludeSample.inc"</code></pre>    
    </td><td style="width:50%">
<%Include ".\IncludeSample.inc"%>
    </td></tr></table></p>      
        
    <p>Отладка только принтами. Для вывода в лог в файле AccessServer.ahk раскоменнтируйте запись в лог файл :
    <pre><code>
;Ведение лога
WriteLog(text) {
 FileAppend, % A_NowUTC ": " text "`n", % A_ScriptDir "/logfile.txt"
}       
    </code></pre>
    
    <p>Затем можно использовать функцию dbg. Принимает два параметра Текст, Изменения уровня.
    
    <pre><code>
'Для вывода на текущем уровне
dbg "Var = [" & var & "]", 0

'Если нужно структурировать вывод 
dbg "Block Entry", +2
'Весь последующий вывод будет сдвинут на два пробела
...
dbg "Block List 1", 0
...
dbg "Block List 2", 0
...
'При выходе из блока нужно уменьшить число пробелов
dbg "Block Exit", -2
    </code></pre>
        
  </body>
</HTML>