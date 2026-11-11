Attribute VB_Name = "OutlookAttachmentExporter"
Option Explicit

' ==========================================================
' OutlookAttachmentExporter.bas
'
' Назначение:
' 1) Следит за подпапкой во "Входящих".
' 2) Сохраняет нужные файлы-вложения из писем на локальный диск.
' 3) НЕ удаляет письма, НЕ переносит письма, НЕ меняет тему.
' 4) Ставит цветные категории:
'    - Выгружено          = зеленая
'    - Нет файлов         = желтая
'    - Нет нужных файлов  = желтая
'    - Ошибка выгрузки    = красная
' 5) Ставит внутреннюю метку, чтобы не создавать дубли при повторном запуске.
'
' Куда вставлять:
' Outlook -> Alt+F11 -> ThisOutlookSession -> вставить весь код.
'
' Как запустить:
' 1) Изменить НАСТРОЙКИ ниже.
' 2) Сохранить VBA-проект.
' 3) Перезапустить Outlook.
' 4) Для ручной обработки всей папки запустить макрос:
'    ExportAllAttachmentsFromWatchFolder
' ==========================================================

Private WithEvents WatchFolderItems As Outlook.Items

' ==========================================================
' НАСТРОЙКИ ПОЛЬЗОВАТЕЛЯ
' ==========================================================

' Подпапка внутри "Входящих", за которой следим.
Private Const OUTLOOK_SUBFOLDER_NAME As String = "Для выгрузки"

' Локальная папка, куда сохраняем файлы-вложения.
Private Const SAVE_TO_FOLDER As String = "C:\MailExport\Test\"

' Какие расширения сохранять.
' Пустая строка "" = сохранять все вложения.
' Рекомендуется оставить только рабочие форматы, чтобы не сохранять картинки из подписей.
Private Const ALLOWED_EXTENSIONS As String = "xlsx;xls;xlsm;csv;pdf;zip;txt;doc;docx"

' Внутренняя техническая метка письма.
' Нужна, чтобы не сохранять одни и те же вложения повторно.
Private Const PROCESSED_FLAG As String = "Файлы выгружены"

' Цветовые категории Outlook.
Private Const CAT_EXPORTED As String = "Выгружено"
Private Const CAT_NO_ATTACHMENTS As String = "Нет файлов"
Private Const CAT_NO_ALLOWED_FILES As String = "Нет нужных файлов"
Private Const CAT_ERROR As String = "Ошибка выгрузки"

' ==========================================================
' АВТОЗАПУСК ПРИ СТАРТЕ OUTLOOK
' ==========================================================

Private Sub Application_Startup()
    On Error GoTo ErrHandler

    Dim ns As Outlook.NameSpace
    Dim inbox As Outlook.Folder
    Dim watchFolder As Outlook.Folder

    EnsureExportCategories

    Set ns = Application.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(olFolderInbox)
    Set watchFolder = inbox.Folders(OUTLOOK_SUBFOLDER_NAME)

    Set WatchFolderItems = watchFolder.Items

    Exit Sub

ErrHandler:
    MsgBox "Не удалось подключиться к папке '" & OUTLOOK_SUBFOLDER_NAME & "'." & vbCrLf & _
           "Проверь, что такая подпапка есть внутри 'Входящих'." & vbCrLf & _
           "Ошибка: " & Err.Description, vbExclamation
End Sub

' ==========================================================
' АВТООБРАБОТКА: письмо добавлено в отслеживаемую папку
' ==========================================================

Private Sub WatchFolderItems_ItemAdd(ByVal Item As Object)
    On Error GoTo ErrHandler

    If TypeName(Item) <> "MailItem" Then Exit Sub

    Dim mail As Outlook.MailItem
    Set mail = Item

    ProcessOneMail mail

    Exit Sub

ErrHandler:
    On Error Resume Next
    If TypeName(Item) = "MailItem" Then
        AddCategory Item, CAT_ERROR
    End If
    MsgBox "Ошибка автоматической выгрузки: " & Err.Description, vbExclamation
End Sub

' ==========================================================
' РУЧНАЯ ОБРАБОТКА: пройти по всем письмам в папке
' Используй после перезапуска компьютера/Outlook или для старых писем.
' ==========================================================

Public Sub ExportAllAttachmentsFromWatchFolder()
    On Error GoTo ErrHandler

    Dim ns As Outlook.NameSpace
    Dim inbox As Outlook.Folder
    Dim watchFolder As Outlook.Folder
    Dim item As Object
    Dim mail As Outlook.MailItem
    Dim processedCount As Long

    EnsureExportCategories
    EnsureFolderExists SAVE_TO_FOLDER

    Set ns = Application.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(olFolderInbox)
    Set watchFolder = inbox.Folders(OUTLOOK_SUBFOLDER_NAME)

    For Each item In watchFolder.Items
        If TypeName(item) = "MailItem" Then
            Set mail = item
            ProcessOneMail mail
            processedCount = processedCount + 1
        End If
    Next item

    MsgBox "Готово. Проверено писем: " & processedCount & vbCrLf & _
           "Папка Outlook: " & OUTLOOK_SUBFOLDER_NAME & vbCrLf & _
           "Папка сохранения: " & SAVE_TO_FOLDER, vbInformation

    Exit Sub

ErrHandler:
    MsgBox "Ошибка ручной выгрузки: " & Err.Description, vbExclamation
End Sub

' ==========================================================
' ОСНОВНАЯ ЛОГИКА ОБРАБОТКИ ОДНОГО ПИСЬМА
' ==========================================================

Private Sub ProcessOneMail(ByVal mail As Outlook.MailItem)
    On Error GoTo ErrHandler

    Dim savedCount As Long

    ' Уже обработано раньше — не создаем дубли.
    If IsAlreadyProcessed(mail) Then Exit Sub

    ' Вложений нет.
    If mail.Attachments.Count = 0 Then
        AddCategory mail, CAT_NO_ATTACHMENTS
        MarkAsProcessed mail
        Exit Sub
    End If

    EnsureFolderExists SAVE_TO_FOLDER

    savedCount = SaveAllowedAttachments(mail)

    If savedCount > 0 Then
        AddCategory mail, CAT_EXPORTED
        MarkAsProcessed mail
    Else
        AddCategory mail, CAT_NO_ALLOWED_FILES
        MarkAsProcessed mail
    End If

    Exit Sub

ErrHandler:
    On Error Resume Next
    AddCategory mail, CAT_ERROR
    ' ВАЖНО: при ошибке НЕ ставим MarkAsProcessed,
    ' чтобы можно было повторить выгрузку позже.
End Sub

' ==========================================================
' СОХРАНЕНИЕ ВЛОЖЕНИЙ
' ==========================================================

Private Function SaveAllowedAttachments(ByVal mail As Outlook.MailItem) As Long
    Dim att As Outlook.Attachment
    Dim ext As String
    Dim filePath As String
    Dim savedCount As Long

    savedCount = 0

    For Each att In mail.Attachments
        ext = GetFileExtension(att.FileName)

        If IsAllowedExtension(ext) Then
            filePath = SAVE_TO_FOLDER & _
                       Format(mail.ReceivedTime, "yyyy-mm-dd_hh-nn-ss") & "_" & _
                       CleanFileName(att.FileName)

            filePath = GetUniqueFilePath(filePath)
            att.SaveAsFile filePath

            savedCount = savedCount + 1
        End If
    Next att

    SaveAllowedAttachments = savedCount
End Function

Private Function IsAllowedExtension(ByVal ext As String) As Boolean
    Dim allowed As String

    allowed = LCase(Trim(ALLOWED_EXTENSIONS))
    ext = LCase(Trim(ext))

    ' Если список пустой — сохраняем все вложения.
    If allowed = "" Then
        IsAllowedExtension = True
        Exit Function
    End If

    IsAllowedExtension = (InStr(1, ";" & allowed & ";", ";" & ext & ";", vbTextCompare) > 0)
End Function

Private Function GetFileExtension(ByVal fileName As String) As String
    Dim p As Long

    p = InStrRev(fileName, ".")

    If p = 0 Then
        GetFileExtension = ""
    Else
        GetFileExtension = Mid(fileName, p + 1)
    End If
End Function

' ==========================================================
' ЗАЩИТА ОТ ПОВТОРНОЙ ВЫГРУЗКИ
' ==========================================================

Private Function IsAlreadyProcessed(ByVal mail As Outlook.MailItem) As Boolean
    On Error GoTo SafeExit

    Dim prop As Outlook.UserProperty
    Set prop = mail.UserProperties.Find(PROCESSED_FLAG)

    If prop Is Nothing Then
        IsAlreadyProcessed = False
    Else
        IsAlreadyProcessed = CBool(prop.Value)
    End If

    Exit Function

SafeExit:
    IsAlreadyProcessed = False
End Function

Private Sub MarkAsProcessed(ByVal mail As Outlook.MailItem)
    On Error Resume Next

    Dim prop As Outlook.UserProperty
    Set prop = mail.UserProperties.Find(PROCESSED_FLAG)

    If prop Is Nothing Then
        ' True в конце — поле можно добавить в колонки представления Outlook.
        Set prop = mail.UserProperties.Add(PROCESSED_FLAG, olYesNo, True)
    End If

    prop.Value = True
    mail.Save
End Sub

' ==========================================================
' КАТЕГОРИИ И ЦВЕТА OUTLOOK
' ==========================================================

Private Sub EnsureExportCategories()
    On Error Resume Next

    Dim cats As Outlook.Categories
    Set cats = Application.Session.Categories

    EnsureCategory cats, CAT_EXPORTED, olCategoryColorGreen
    EnsureCategory cats, CAT_NO_ATTACHMENTS, olCategoryColorYellow
    EnsureCategory cats, CAT_NO_ALLOWED_FILES, olCategoryColorYellow
    EnsureCategory cats, CAT_ERROR, olCategoryColorRed
End Sub

Private Sub EnsureCategory(ByVal cats As Outlook.Categories, ByVal catName As String, ByVal catColor As OlCategoryColor)
    On Error Resume Next

    Dim cat As Outlook.Category
    Set cat = cats.Item(catName)

    If cat Is Nothing Then
        cats.Add catName, catColor
    Else
        cat.Color = catColor
    End If
End Sub

Private Sub AddCategory(ByVal mail As Outlook.MailItem, ByVal catName As String)
    On Error Resume Next

    If mail.Categories = "" Then
        mail.Categories = catName
    ElseIf InStr(1, mail.Categories, catName, vbTextCompare) = 0 Then
        mail.Categories = mail.Categories & ";" & catName
    End If

    mail.Save
End Sub

' ==========================================================
' ФАЙЛОВАЯ СИСТЕМА И БЕЗОПАСНЫЕ ИМЕНА
' ==========================================================

Private Sub EnsureFolderExists(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(folderPath) Then Exit Sub

    CreateFoldersRecursive folderPath
End Sub

Private Sub CreateFoldersRecursive(ByVal folderPath As String)
    Dim fso As Object
    Dim parentPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(folderPath) Then Exit Sub

    parentPath = fso.GetParentFolderName(folderPath)

    If parentPath <> "" Then
        If Not fso.FolderExists(parentPath) Then
            CreateFoldersRecursive parentPath
        End If
    End If

    fso.CreateFolder folderPath
End Sub

Private Function CleanFileName(ByVal s As String) As String
    Dim badChars As Variant
    Dim ch As Variant

    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbCr, vbLf, vbTab)

    For Each ch In badChars
        s = Replace(s, ch, "_")
    Next ch

    s = Trim(s)

    If Len(s) = 0 Then s = "file"
    If Len(s) > 140 Then s = Left(s, 140)

    CleanFileName = s
End Function

Private Function GetUniqueFilePath(ByVal fullPath As String) As String
    Dim fso As Object
    Dim folderPath As String
    Dim baseName As String
    Dim ext As String
    Dim candidate As String
    Dim n As Long

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(fullPath) Then
        GetUniqueFilePath = fullPath
        Exit Function
    End If

    folderPath = fso.GetParentFolderName(fullPath)
    baseName = fso.GetBaseName(fullPath)
    ext = fso.GetExtensionName(fullPath)

    n = 1

    Do
        candidate = folderPath & "\" & baseName & " (" & n & ")"
        If ext <> "" Then candidate = candidate & "." & ext

        If Not fso.FileExists(candidate) Then
            GetUniqueFilePath = candidate
            Exit Function
        End If

        n = n + 1
    Loop
End Function
