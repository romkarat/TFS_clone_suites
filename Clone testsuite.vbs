Dim textInput,textInputBox
Dim collection,teamproject,path_tcm,path_vbs
Dim suiteid,destinationsuiteid
Dim dash_position
Dim commands
Dim TextOutput,TextError,TextOutputString,TextErrorString
Dim write_log,filename_log

'переменные, выставляются индивидуально для проекта:
'путь к серверу TFS
    collection="http://<server_name>:<port>/tfs/IT_Systems/"
'название проекта
    teamproject="YOUR_PROJECT_NAME"
'путь к утилите TCM.exe
    path_tcm="C:\Program Files\Microsoft Visual Studio 12.0\Common7\IDE"
'флаг записи в лог: Y-да/N-нет
    write_log="Y"
'имя файла лога с расширением
    filename_log="clone_testsuite.log"

'диалоговое окно со строкой ввода
    textInputBox="Проект: " & teamproject & vbCrLf & vbCrLf & "Введите ID тест-сьютов через дефис 'Источник-Цель'"
    textInput=InputBox(textInputBox,"Клонирование набора тестов (test-suite) в TFS","123456-654321")

'обработка событий
If textInput="" Then
    'если строка ввода пустая или нажата кнопка "Cancel" или кнопка закрытия окна, то ничего не делаем - выход
Else
    'извлечение введенных ID из строки ввода
    dash_position=InStr(1,textInput,"-",vbTextCompare)
    suiteid=mid(textInput, 1, dash_position-1)
    destinationsuiteid=mid(textInput, dash_position+1, len(textInput))

    'запуск команды клонирования в консоли cmd (ключ /k после cmd.exe - не закрывать окно консоли после выполнения, ключ /c - закрывать)
    commands="cmd.exe /c tcm suites /clone /collection:" & collection & " /teamproject:" & teamproject & " /suiteid:" & suiteid & " /destinationsuiteid:" & destinationsuiteid
    Set WshShell = CreateObject("WScript.Shell")
    path_vbs = WshShell.CurrentDirectory
    WshShell.CurrentDirectory = path_tcm
    Set WshExec = WshShell.Exec(commands)
    
    'вывод сообщения о клонировании в зависимости от результата выполнения
    Set TextOutput = WshExec.StdOut
    Set TextError = WshExec.StdErr
    TextOutputString = TextOutput.ReadAll
    TextErrorString = TextError.ReadAll
    If TextErrorString = vbNullString Then
        Str = vbNullString & "Ход выполнения (сообщение из консоли):"
        message = "Клонирован тест-сьют с ID:" & suiteid & " в тест-сьют с ID:" & destinationsuiteid
        messageForScreen = MsgBox (Str & TextOutputString & vbCrLf & "Проект: " & teamproject & vbCrLf & "Сервер: " & collection & vbCrLf & vbCrLf & message, vbOKOnly+vbInformation,"Сообщение")
        messageForLog = "INFO - " & TextOutputString & " " & message
    Else
        Str = vbNullString & "Ошибка выполнения (сообщение из консоли):" & vbCrLf
        message = "Ошибка клонирования тест-сьюта с ID:" & suiteid & " в тест-сьют с ID:" & destinationsuiteid
        messageForScreen = MsgBox (Str & TextErrorString & vbCrLf & "Проект: " & teamproject & vbCrLf & "Сервер: " & collection & vbCrLf & vbCrLf & message, vbOKOnly+vbExclamation,"Сообщение")
        messageForLog = "ERROR - " & TextErrorString & " " & message
    End If
End If

'логирование в файл хода выполнения клонирования (создается в папке со скриптом)
If write_log="Y" Then
    Dim FSO, File
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    '8 - режим открытия файла: 1 - Только для чтения; 2 - Для записи, если файл существовал, то его содержимое теряется; 8 - Дописывать в конец файла.
    'True/False - cоздать файл, если он не существует (True), в противном случае False.
    Set File = FSO.OpenTextFile(path_vbs & "\" & filename_log, 8, True)
	
    'Формирование сообщения для лога и его преобразование в одну строку, удалением символов перевода строки
    TextForLog = Now & " [Проект:" & teamproject & ", Сервер:" & collection & "] " & messageForLog
    TextForLog = Replace(TextForLog, vbCrLf, "")
	
    'непосредственно запись сообщения в файл и его закрытие
    File.WriteLine(TextForLog)
    File.Close
Else
    'если флаг write_log="N" (или любое другое значение), то лог не пишем, т.е. ничего не делаем - выход
End If
