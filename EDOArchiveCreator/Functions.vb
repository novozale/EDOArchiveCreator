Imports System.Net.Mail
Imports Diadoc.Api
Imports Diadoc.Api.Cryptography
Imports System.Net.Http
Imports System.IO

Module Functions
    Public Sub SendMyReminder(ByVal Subject As String, ByVal MyWrkString As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// отправка сообщения о ошибке по почте в ИТ
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Try
            Dim smtp As SmtpClient = New SmtpClient(My.Settings.SMTPService)
            Dim msg As New MailMessage

            msg.To.Add(My.Settings.MessageTo)
            If Trim(My.Settings.MessageCC) <> "" Then
                msg.CC.Add(My.Settings.MessageCC)
            End If
            msg.From = New MailAddress(My.Settings.MessageFrom)
            msg.Subject = Subject
            msg.Body = MyWrkString
            smtp.Send(msg)
        Catch ex As Exception
            EventLog.WriteEntry("EDOArchiveCreator", ex.Message)
        End Try
    End Sub

    Public Sub GetFullArchive()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// получение и сохранение электронного архива за промежуток времени
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim Crypt As WinApiCrypt
        Dim MyApi As DiadocApi
        Dim MyCurrDate As Date

        Try
            Crypt = New WinApiCrypt
            MyApi = New DiadocApi(Declarations.DeveloperKey, Declarations.DiadocApiURL, Crypt)
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("EDOArchiveCreator", ex.Message)
                SendMyReminder("Ошибка утилиты получения архива ЭДО", ex.Message)
            End If
            Exit Sub
        End Try

        Declarations.MyToken = GetMyToken(MyApi)
        If Declarations.MyToken <> "" Then
            MyCurrDate = MyDateFrom
            While MyCurrDate < MyDateTo
                GetOneDayArchive(MyApi, MyCurrDate)
                MyCurrDate = DateAdd(DateInterval.Day, 1, MyCurrDate)
            End While
        End If
    End Sub

    Public Function GetMyToken(ByRef MyApi As DiadocApi) As String
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// получение токена для работы с API
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyTokenLogin As String

        MyTokenLogin = ""
        Try
            MyTokenLogin = MyApi.Authenticate(Declarations.MyLogin, Declarations.MyPassword)
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("EDOArchiveCreator", ex.Message)
                SendMyReminder("Ошибка утилиты получения архива ЭДО", ex.Message)
            End If
        End Try
        GetMyToken = MyTokenLogin
    End Function

    Public Sub GetOneDayArchive(ByRef MyApi As DiadocApi, MyCurrDate As Date)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// получение и сохранение электронного архива за одни сутки
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Console.WriteLine("==========" + Format(MyCurrDate, "yyyy_MM_dd") + "============")
        Console.WriteLine("+++++++++Входящие+++++++++++")
        GetOneDayIncomingDocs(MyApi, MyCurrDate)
        Console.WriteLine("+++++++++Исходящие+++++++++++")
        GetOneDayOutcomingDocs(MyApi, MyCurrDate)
    End Sub

    Public Sub GetOneDayIncomingDocs(ByRef MyApi As DiadocApi, MyCurrDate As Date)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// получение и сохранение входящих документов за одни сутки
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyDocList As Diadoc.Api.Proto.Documents.DocumentList
        Dim MyDocument As Diadoc.Api.Proto.Documents.Document
        Dim MyDF As Diadoc.Api.DocumentsFilter
        Dim MyContragent As Diadoc.Api.Proto.Organization
        Dim MyAfterIndexKey As String = ""
        Dim Mylength As Integer = 0

        MyDF = New DocumentsFilter
        MyDF.FilterCategory = "Any.InboundFinished"
        MyDF.TimestampFrom = MyCurrDate
        MyDF.TimestampTo = DateAdd(DateInterval.Day, 1, MyCurrDate)
        MyDF.BoxId = Declarations.MyboxId

        '----получение списка документов на дату
        Try
            MyDocList = MyApi.GetDocuments(Declarations.MyToken, MyDF)
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("EDOArchiveCreator", ex.Message)
                SendMyReminder("Ошибка утилиты получения архива ЭДО", ex.Message)
            End If
        End Try

        While MyDocList.Documents.Count <> 0
            '----в цикле просматриваем все документы
            For i As Integer = 0 To MyDocList.Documents.Count - 1
                MyDocument = MyDocList.Documents(i)

                '---получение параметров документа
                MyAfterIndexKey = MyDocument.IndexKey
                '---тип документа (СФ, Счет, Накладная, акт и т.д.)
                Declarations.MyDocType = GetMyDocumentType(MyDocument)
                '---номер документа
                Declarations.MyDocNum = RemoveNotCorrectChars(MyDocument.DocumentNumber)
                If Declarations.MyDocNum = Nothing Or Trim(Declarations.MyDocNum) = "" Then
                    Declarations.MyDocNum = RemoveNotCorrectChars(MyDocument.EntityId)
                End If
                Try
                    MyContragent = MyApi.GetOrganizationByBoxId(MyDocument.CounteragentBoxId)
                    '---ИНН
                    Declarations.MyINN = MyContragent.Inn
                    '---КПП
                    Declarations.MyKPP = MyContragent.Kpp
                    '---название компании
                    Declarations.MyCompanyName = RemoveNotCorrectChars(MyContragent.FullName)
                    Mylength = Declarations.MyCompanyName.Length
                    If Mylength > 105 Then
                        Declarations.MyCompanyName = Left(Declarations.MyCompanyName, 103) & ".."
                    End If
                    '---полный путь + имя файла
                    Declarations.MyFileFullPath = Declarations.MyArchivePath + "Входящие" + "\" + Declarations.MyCompanyName + " ИНН " + Declarations.MyINN + " КПП " + Declarations.MyKPP + "\"
                    Declarations.MyFileFullPath = Declarations.MyFileFullPath + Format(MyCurrDate, "yyyy_MM_dd") + "\" + Declarations.MyDocType
                    Declarations.MyFileFullPathName = Declarations.MyFileFullPath + "\" + Declarations.MyDocNum + ".zip"
                    Console.WriteLine("-----" + Declarations.MyCompanyName + " ИНН " + Declarations.MyINN + " КПП " + Declarations.MyKPP + " документ " + Declarations.MyDocType + " N " + Declarations.MyDocNum)
                Catch ex As Exception
                    '---полный путь + имя файла
                    Declarations.MyFileFullPath = Declarations.MyArchivePath + "Входящие" + "\" + MyDocument.CounteragentBoxId + "\"
                        Declarations.MyFileFullPath = Declarations.MyFileFullPath + Format(MyCurrDate, "yyyy_MM_dd") + "\" + Declarations.MyDocType
                        Declarations.MyFileFullPathName = Declarations.MyFileFullPath + "\" + Declarations.MyDocNum + ".zip"
                        Console.WriteLine("-----" + MyDocument.CounteragentBoxId + " документ " + Declarations.MyDocType + " N " + Declarations.MyDocNum)
                    End Try
                    If Declarations.MyOverwriteFlag = 0 Then         '--не перезаписывать существующие документы
                    '---проверка - может, уже есть в архиве
                    If MyDocumentExist(Declarations.MyFileFullPathName) = True Then
                    Else        '---если нет в архиве - записываем
                        SaveMyDocument(MyApi, MyDocument)
                    End If
                Else
                    SaveMyDocument(MyApi, MyDocument)
                End If
            Next i

            MyDF.AfterIndexKey = MyAfterIndexKey
            Try
                MyDocList = MyApi.GetDocuments(Declarations.MyToken, MyDF)
            Catch ex As Exception
                If My.Settings.MyDebug = "YES" Then
                    EventLog.WriteEntry("EDOArchiveCreator", ex.Message)
                    SendMyReminder(Format(MyCurrDate, "yyyy_MM_dd") + "Ошибка утилиты получения архива ЭДО", ex.Message)
                End If
            End Try
        End While
    End Sub

    Public Sub GetOneDayOutcomingDocs(ByRef MyApi As DiadocApi, MyCurrDate As Date)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// получение и сохранение исходящих документов за одни сутки
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyDocList As Diadoc.Api.Proto.Documents.DocumentList
        Dim MyDocument As Diadoc.Api.Proto.Documents.Document
        Dim MyDF As Diadoc.Api.DocumentsFilter
        Dim MyContragent As Diadoc.Api.Proto.Organization
        Dim MyAfterIndexKey As String = ""
        Dim MyLength As Integer = 0

        MyDF = New DocumentsFilter
        MyDF.FilterCategory = "Any.OutboundFinished"
        MyDF.TimestampFrom = MyCurrDate
        MyDF.TimestampTo = DateAdd(DateInterval.Day, 1, MyCurrDate)
        MyDF.BoxId = Declarations.MyboxId


        '----получение списка документов на дату
        Try
            MyDocList = MyApi.GetDocuments(Declarations.MyToken, MyDF)
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("EDOArchiveCreator", ex.Message)
                SendMyReminder("Ошибка утилиты получения архива ЭДО", ex.Message)
            End If
        End Try

        While MyDocList.Documents.Count <> 0
            '----в цикле просматриваем все документы
            For i As Integer = 0 To MyDocList.Documents.Count - 1
                MyDocument = MyDocList.Documents(i)

                '---получение параметров документа
                MyAfterIndexKey = MyDocument.IndexKey
                '---тип документа (СФ, Счет, Накладная, акт и т.д.)
                Declarations.MyDocType = GetMyDocumentType(MyDocument)
                '---номер документа
                Declarations.MyDocNum = RemoveNotCorrectChars(MyDocument.DocumentNumber)
                If Declarations.MyDocNum = Nothing Or Trim(Declarations.MyDocNum) = "" Then
                    Declarations.MyDocNum = RemoveNotCorrectChars(MyDocument.EntityId)
                End If
                Try
                    MyContragent = MyApi.GetOrganizationByBoxId(MyDocument.CounteragentBoxId)
                    '---ИНН
                    Declarations.MyINN = MyContragent.Inn
                    '---КПП
                    Declarations.MyKPP = MyContragent.Kpp
                    '---название компании
                    Declarations.MyCompanyName = RemoveNotCorrectChars(MyContragent.FullName)
                    MyLength = Declarations.MyCompanyName.Length
                    If MyLength > 105 Then
                        Declarations.MyCompanyName = Left(Declarations.MyCompanyName, 103) & ".."
                    End If
                    '---полный путь + имя файла
                    Declarations.MyFileFullPath = Declarations.MyArchivePath + "Исходящие" + "\" + Declarations.MyCompanyName + " ИНН " + Declarations.MyINN + " КПП " + Declarations.MyKPP + "\"
                    Declarations.MyFileFullPath = Declarations.MyFileFullPath + Format(MyCurrDate, "yyyy_MM_dd") + "\" + Declarations.MyDocType
                    Declarations.MyFileFullPathName = Declarations.MyFileFullPath + "\" + Declarations.MyDocNum + ".zip"
                    Console.WriteLine("-----" + Declarations.MyCompanyName + " ИНН " + Declarations.MyINN + " КПП " + Declarations.MyKPP + " документ " + Declarations.MyDocType + " N " + Declarations.MyDocNum)
                Catch ex As Exception
                    '---полный путь + имя файла
                    Declarations.MyFileFullPath = Declarations.MyArchivePath + "Исходящие" + "\" + MyDocument.CounteragentBoxId + "\"
                    Declarations.MyFileFullPath = Declarations.MyFileFullPath + Format(MyCurrDate, "yyyy_MM_dd") + "\" + Declarations.MyDocType
                    Declarations.MyFileFullPathName = Declarations.MyFileFullPath + "\" + Declarations.MyDocNum + ".zip"
                    Console.WriteLine("-----" + MyDocument.CounteragentBoxId + " документ " + Declarations.MyDocType + " N " + Declarations.MyDocNum)
                End Try

                If Declarations.MyOverwriteFlag = 0 Then         '--не перезаписывать существующие документы
                    '---проверка - может, уже есть в архиве
                    If MyDocumentExist(Declarations.MyFileFullPathName) = True Then
                    Else        '---если нет в архиве - записываем
                        SaveMyDocument(MyApi, MyDocument)
                    End If
                Else
                    SaveMyDocument(MyApi, MyDocument)
                End If
            Next i

            MyDF.AfterIndexKey = MyAfterIndexKey
            Try
                MyDocList = MyApi.GetDocuments(Declarations.MyToken, MyDF)
            Catch ex As Exception
                If My.Settings.MyDebug = "YES" Then
                    EventLog.WriteEntry("EDOArchiveCreator", ex.Message)
                    SendMyReminder(Format(MyCurrDate, "yyyy_MM_dd") + "Ошибка утилиты получения архива ЭДО", ex.Message)
                End If
            End Try
        End While
    End Sub

    Public Function GetMyDocumentType(ByRef MyDocument As Diadoc.Api.Proto.Documents.Document) As String
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// получение типа документа
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Select Case MyDocument.DocumentType
            Case Proto.DocumentType.AcceptanceCertificate
                GetMyDocumentType = "AcceptanceCertificate"
            Case Proto.DocumentType.CertificateRegistry
                GetMyDocumentType = "CertificateRegistry"
            Case Proto.DocumentType.Contract
                GetMyDocumentType = "Contract"
            Case Proto.DocumentType.Invoice
                GetMyDocumentType = "Invoice"
            Case Proto.DocumentType.InvoiceCorrection
                GetMyDocumentType = "InvoiceCorrection"
            Case Proto.DocumentType.InvoiceCorrectionRevision
                GetMyDocumentType = "InvoiceCorrectionRevision"
            Case Proto.DocumentType.InvoiceRevision
                GetMyDocumentType = "InvoiceRevision"
            Case Proto.DocumentType.Nonformalized
                GetMyDocumentType = "Nonformalized"
            Case Proto.DocumentType.PriceList
                GetMyDocumentType = "PriceList"
            Case Proto.DocumentType.PriceListAgreement
                GetMyDocumentType = "PriceListAgreement"
            Case Proto.DocumentType.PriceListAgreement
                GetMyDocumentType = "PriceListAgreement"
            Case Proto.DocumentType.ProformaInvoice
                GetMyDocumentType = "ProformaInvoice"
            Case Proto.DocumentType.ReconciliationAct
                GetMyDocumentType = "ReconciliationAct"
            Case Proto.DocumentType.ServiceDetails
                GetMyDocumentType = "ServiceDetails"
            Case Proto.DocumentType.SupplementaryAgreement
                GetMyDocumentType = "SupplementaryAgreement"
            Case Proto.DocumentType.Torg12
                GetMyDocumentType = "Torg12"
            Case Proto.DocumentType.Torg13
                GetMyDocumentType = "Torg13"
            Case Proto.DocumentType.TrustConnectionRequest
                GetMyDocumentType = "TrustConnectionRequest"
            Case Proto.DocumentType.UniversalTransferDocument
                GetMyDocumentType = "UniversalTransferDocument"
            Case Proto.DocumentType.XmlAcceptanceCertificate
                GetMyDocumentType = "XmlAcceptanceCertificate"
            Case Proto.DocumentType.XmlTorg12
                GetMyDocumentType = "XmlTorg12"
            Case Else
                GetMyDocumentType = "NotDefined"
        End Select
    End Function

    Public Function RemoveNotCorrectChars(MyString As String) As String
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Замена некорректных (запрещенных в названиях файлов) символов нижним подчеркиванием
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim NotCorrectChars As String

        NotCorrectChars = "*|\:<>?/"
        For i As Integer = 0 To Len(NotCorrectChars) - 1
            MyString = Replace(MyString, NotCorrectChars(i), "_")
        Next
        MyString = Replace(MyString, """", "_")
        MyString = Replace(MyString, Chr(9), "")
        RemoveNotCorrectChars = MyString
    End Function

    Public Function MyDocumentExist(MyFileFullPathName As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка - существует ли файл с названием MyFileFullPathName
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyDocumentExist = File.Exists(MyFileFullPathName)
    End Function

    Public Sub SaveMyDocument(ByRef MyApi As DiadocApi, ByRef MyDocument As Diadoc.Api.Proto.Documents.Document)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Сохранение документа в локальный архив
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyZip As Diadoc.Api.Proto.Documents.IDocumentZipGenerationResult
        Dim MyBytes As Byte()

        '---подготовка архива на сервере
        MyZip = MyApi.GenerateDocumentZip(Declarations.MyToken, Declarations.MyboxId, MyDocument.MessageId, MyDocument.EntityId, True)
        For i = 1 To Declarations.MyNumberDownloadTries
            If MyZip.ZipFileNameOnShelf = "" Then   '---не готово - ждем
                System.Threading.Thread.Sleep((MyZip.RetryAfter + 1) * 1000)
            Else                                    '---готово - переходим к сохранению
                Exit For
            End If
            Try
                MyZip = MyApi.GenerateDocumentZip(Declarations.MyToken, Declarations.MyboxId, MyDocument.MessageId, MyDocument.EntityId, True)
            Catch ex1 As Exception
                MyZip.ZipFileNameOnShelf = ""
            End Try
        Next

        If MyZip.ZipFileNameOnShelf = "" Then       '---архив так и не подготовился
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("EDOArchiveCreator", "Ошибка подготовки архива " + " Направление " + CStr(MyDocument.DocumentDirection) + " Полное имя " + Declarations.MyFileFullPathName)
                SendMyReminder("Ошибка утилиты получения архива ЭДО", "Ошибка подготовки архива " + " Направление " + CStr(MyDocument.DocumentDirection) + " Полное имя " + Declarations.MyFileFullPathName)
            End If
        Else                                        '---архив готов - сохраняем
            Try
                MyBytes = MyApi.GetFileFromShelf(Declarations.MyToken, MyZip.ZipFileNameOnShelf)
            Catch ex As Exception
                If My.Settings.MyDebug = "YES" Then
                    EventLog.WriteEntry("EDOArchiveCreator", ex.Message)
                    SendMyReminder(" Полное имя " + Declarations.MyFileFullPathName + " Ошибка утилиты получения архива ЭДО", ex.Message)
                End If
            End Try

            Try
                If Directory.Exists(Declarations.MyFileFullPath) = False Then
                    Directory.CreateDirectory(Declarations.MyFileFullPath)
                End If
                File.WriteAllBytes(Declarations.MyFileFullPathName, MyBytes)
            Catch ex As Exception
                If My.Settings.MyDebug = "YES" Then
                    EventLog.WriteEntry("EDOArchiveCreator", ex.Message)
                    SendMyReminder(" Полное имя " + Declarations.MyFileFullPathName + " Ошибка утилиты получения архива ЭДО", ex.Message)
                End If
            End Try
        End If
    End Sub
End Module
