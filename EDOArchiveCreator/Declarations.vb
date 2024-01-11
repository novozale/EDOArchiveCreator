Module Declarations
    Public DiadocApiURL As String                               '--WEB адрес Diadoc API
    Public DeveloperKey As String                               '--ключ разработчика
    Public MyLogin As String                                    '--логин в Diadoc
    Public MyPassword As String                                 '--пароль в Diadoc
    Public MyToken As String                                    '--токен для работы
    Public MyboxId As String                                    '--ID бокса нашей компании
    Public MyArchivePath As String                              '--путь к архиву
    Public MyOverwriteFlag As Integer                           '--перезаписывать существующие файлы в архиве (1) или нет (0)
    Public MyINN As String                                      '--ИНН компании
    Public MyKPP As String                                      '--КПП компании
    Public MyCompanyName As String                              '--Название компании
    Public MyDocType As String                                  '--тип документа (СФ, Счет, Накладная, акт и т.д.)
    Public MyDocNum As String                                   '--номер документа
    Public MyFileFullPath As String                             '--путь архива
    Public MyFileFullPathName As String                         '--путь + имя файла архива
    Public MyNumberDownloadTries As Integer                     '--количество попыток скачивания файла с Диадока
End Module
