Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Net
Imports System.Net.Sockets
Imports System.Text

Namespace BarCodeScanner_USB_HID_KEYB_vb

    'ЗДЕСЬ НАДО ЗАДАТЬ ПРОГИД И ГУИД
    <ComVisible(True), Guid("9840BED5-950E-4679-B4DD-F32E8C86200C"), ProgId("AddIn.BarCodeScanner_USB_HID_KEYB_vb"), ClassInterface(ClassInterfaceType.AutoDispatch)>
    Public Class BarCodeScanner_USB_HID_KEYB_vb
        Implements IInitDone
        Implements ILanguageExtender

        Const c_AddinName As String = "BarCodeScanner_USB_HID_KEYB_vb"

#Region "Переменные"
        Dim strMessageBoxIcon As String
        Dim strMessageBoxButtons As String
        Dim IntCountOpenDevices As Integer
        Dim strTableRegDevices As String
        Dim WithEvents Timer1 As New Timers.Timer()
        Dim WithEvents Timer2 As New Timers.Timer()
        Dim TcpListener1 As TcpListener = New TcpListener(IPAddress.Any, 1077)
        Dim TcpListener1_connected As Boolean = False
        Dim TcpListener1_SrtBuld As StringBuilder = New StringBuilder(200)
        Dim TextImport1 As [String] = ""
#End Region

#Region "IInitDone implementation"
        Public Sub New() ' Обязательно для COM инициализации
        End Sub

        Private Sub Init(<MarshalAs(UnmanagedType.IDispatch)> ByVal pConnection As Object) Implements IInitDone.Init
            '  Сохранить pConnection для дальнейшего
            strMessageBoxButtons = "OK" : strMessageBoxIcon = "None"

            If (True) Then ' (V7Data.V7Object.ToString <> "System.__ComObject")
                V7Data.V7Object = pConnection
                Timer1.Interval = 1777 : Timer1.AutoReset = True : AddHandler Timer1.Elapsed, AddressOf Timer1_Tick
                Timer2.Interval = 77 : Timer2.AutoReset = True : AddHandler Timer2.Elapsed, AddressOf Timer2_Tick
            End If

            If (Not Timer1.Enabled) Then Timer1.Start() 'как бы внешнее событие запускает второй таймер

            'MessageBox.Show("Init" & vbNewLine & c_AddinName, "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        End Sub

        Private Sub Done() Implements IInitDone.Done
            TcpListener1_connected = False
            If TcpListener1 IsNot Nothing Then TcpListener1.Stop()
            'Timer1.Stop() : Timer1.Dispose()
            'Timer2.Stop() : Timer2.Dispose()
            'MessageBox.Show("Done" & vbNewLine & c_AddinName, "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        End Sub

        Private Sub GetInfo(ByRef pInfo() As Object) Implements IInitDone.GetInfo
            pInfo.SetValue("2000", 0)
            'MessageBox.Show("GetInfo" & vbNewLine & c_AddinName, "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        End Sub
        Sub RegisterExtensionAs(ByRef bstrExtensionName As String) Implements ILanguageExtender.RegisterExtensionAs
            bstrExtensionName = c_AddinName
        End Sub

#End Region

        Public Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Elapsed
            Timer1.Stop()
            If Not TcpListener1_connected Then
                TcpListener1_connected = True : TcpListener1.Start()
                While TcpListener1_connected
                    Dim TcpClient1 As TcpClient = TcpListener1.AcceptTcpClient() : Dim NetworkStream1 As NetworkStream = TcpClient1.GetStream()
                    TcpListener1_SrtBuld.Length = 0 : Dim LastDataEventDate As DateTime = DateTime.Now : Dim idle_event As Double = 0
                    While (TcpClient1.Connected) AndAlso (idle_event <= 777.777) 'ждем почти секунду пока не придет весь запрос
                        idle_event = (DateTime.Now - LastDataEventDate).TotalMilliseconds
                        If (TcpClient1.Available > 0) Then
                            Dim PacketBytes As [Byte]() = New [Byte](4096) {}
                            Dim PacketLength1 As Integer = NetworkStream1.Read(PacketBytes, 0, Math.Min(TcpClient1.Available, PacketBytes.Length))
                            TcpListener1_SrtBuld.Append(System.Text.Encoding.ASCII.GetString(PacketBytes, 0, PacketLength1))
                        End If
                    End While
                    TextImport1 = TcpListener1_SrtBuld.ToString() : TcpListener1_SrtBuld.Length = 0
                    If TextImport1.Substring(0, "GET /favicon.ico".Length) = "GET /favicon.ico" Then
                        'Dim test As var = New System.Resources.Extensions.DeserializingResourceReader("")
                        'Dim rm As New ResourceManager("TcpListener1", Assembly.GetExecutingAssembly())
                        'Dim myIcon As Icon = (DirectCast((rm.GetObject("icon64.ico")), Icon))
                        'Dim ms1 As New MemoryStream()
                        'myIcon.Save(ms1)
                        'Dim Buffer As Byte() = ms1.ToArray()
                        'If TcpClient1.Connected Then
                        '    NetworkStream1.Write(Buffer, 0, Buffer.Length)
                        'End If
                    Else
                        Dim Code As Integer = 200 : Dim CodeStr As String = Code.ToString() + " " + (DirectCast(Code, HttpStatusCode)).ToString()
                        Dim Html As String = "<html><body><h1>" + CodeStr + "</h1>" + TextImport1.Replace(vbCr & vbLf, "<BR>" & vbCr & vbLf) + "</body></html>"
                        Dim Str As String = "HTTP/1.1 " + CodeStr + vbLf & "Content-type: text/html" & vbLf & "Content-Length:" + Html.Length.ToString() + vbLf & vbLf + Html
                        Dim Buffer As Byte() = Encoding.ASCII.GetBytes(Str)
                        If (TcpClient1.Connected) Then
                            NetworkStream1.Write(Buffer, 0, Buffer.Length)
                        End If
                    End If
                    NetworkStream1.Close() : TcpClient1.Close()
                    Timer2.Start()
                End While
            End If 'If Not TcpListener1_connected Then
        End Sub
        Public Sub Timer2_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer2.Elapsed
            Timer2.Enabled = False
            strTableRegDevices = "" & DateTime.Now.ToString() & vbNewLine & TextImport1
            'Dim ei As ExcepInfo
            'ei.wCode = 1006 'Вид пиктограммы
            'ei.bstrDescription = "" + DateTime.Now.ToString()
            'ei.bstrSource = c_AddinName
            'V7Data.ErrorLog.AddError("", ei) : Throw New System.Exception("An exception has occurred.")
            'MsgBox("" + DateTime.Now.ToString())
            'MessageBox.Show("Timer2_Tick = до" & vbNewLine & c_AddinName, "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            V7Data.AsyncEvent.ExternalEvent("s1", "s2", "s3")
            'MessageBox.Show("Timer2_Tick = по" & vbNewLine & c_AddinName, "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            'V7Data.StatusLine.SetStatusLine("" + DateTime.Now.ToString())
            'System.Threading.Thread.Sleep(1000) 'Делаем паузу 1 сек
            'V7Data.StatusLine.ResetStatusLine()
            strTableRegDevices = "" & DateTime.Now.ToString() & " жир"
            'Timer2 .Enabled = True
        End Sub


#Region "Свойства"
        Enum Props
            'Числовые идентификаторы свойств нашей внешней компоненты

            propMessageBoxIcon = 0  'Пиктограмма для MessageBox'а
            propMessageBoxButtons = 1 'Кнопки для MessageBox'a
            propCountOpenDevices = 2
            propTableRegDevices = 3
            LastProp = 4

        End Enum

        Sub FindProp(ByVal bstrPropName As String, ByRef plPropNum As Integer) Implements ILanguageExtender.FindProp
            'Здесь 1С ищет числовой идентификатор свойства по его текстовому имени

            Select Case bstrPropName

                Case "MessageBoxIcon", "ПиктограммаПредупреждения"
                    plPropNum = Props.propMessageBoxIcon

                Case "MessageBoxButtons", "КнопкиПредупреждения"
                    plPropNum = Props.propMessageBoxButtons

                Case "CountOpenDevices", "КоличествоОткрытыхУстройств"
                    plPropNum = Props.propCountOpenDevices

                Case "TableRegDevices", "ТаблицаЗаписейРегистра"
                    plPropNum = Props.propTableRegDevices

                Case Else
                    plPropNum = -1
            End Select
        End Sub

        Sub GetPropVal(ByVal lPropNum As Integer, ByRef pvarPropVal As Object) Implements ILanguageExtender.GetPropVal
            'Здесь 1С узнает значения свойств 

            pvarPropVal = Nothing
            Select Case lPropNum

                Case Props.propMessageBoxIcon
                    pvarPropVal = strMessageBoxIcon

                Case Props.propMessageBoxButtons
                    pvarPropVal = strMessageBoxButtons

                Case Props.propCountOpenDevices
                    pvarPropVal = IntCountOpenDevices

                Case Props.propTableRegDevices
                    pvarPropVal = strTableRegDevices

            End Select
        End Sub

        Sub SetPropVal(ByVal lPropNum As Integer, ByRef varPropVal As Object) Implements ILanguageExtender.SetPropVal
            'Здесь 1С изменяет значения свойств 

            Select Case lPropNum

                Case Props.propMessageBoxIcon
                    strMessageBoxIcon = CType(varPropVal, String)

                Case Props.propMessageBoxButtons
                    strMessageBoxButtons = CType(varPropVal, String)

                Case Props.propCountOpenDevices
                    IntCountOpenDevices = CInt(varPropVal)

                Case Props.propTableRegDevices
                    strTableRegDevices = CType(varPropVal, String)

            End Select
        End Sub

        Sub IsPropReadable(ByVal lPropNum As Integer, ByRef pboolPropRead As Boolean) Implements ILanguageExtender.IsPropReadable
            'Здесь 1С узнает, какие свойства доступны для чтения

            pboolPropRead = True ' Все свойства доступны для чтения
        End Sub

        Sub IsPropWritable(ByVal lPropNum As Integer, ByRef pboolPropWrite As Boolean) Implements ILanguageExtender.IsPropWritable
            'Здесь 1С узнает, какие свойства доступны для записи

            pboolPropWrite = True ' Все свойства доступны для записи
        End Sub

        Sub GetNProps(ByRef plProps As Integer) Implements ILanguageExtender.GetNProps
            'Здесь 1С получает количество доступных из ВК свойств

            plProps = Props.LastProp
        End Sub

        Sub GetPropName(ByVal lPropNum As Integer, ByVal lPropAlias As Integer, ByRef pbstrPropName As String) Implements ILanguageExtender.GetPropName
            'Здесь 1С (теоретически) узнает имя свойства по его идентификатору. lPropAlias - номер псевдонима

            pbstrPropName = ""
        End Sub

#End Region

#Region "Методы"

        Enum Methods
            'Числовые идентификаторы методов (процедур или функций) нашей внешней компоненты

            methMessageBoxShow = 0 'Предупреждение, но с возможностью задавать пиктограмму и заголовок окна
            methExternalEvent = 1 'Генерирует внешнее событие (1С перехватывает его в процедуре ОбработкаВнешнегоСобытия())
            methShowErrorLog = 2 ' Показываем ошибочное сообщение 
            methStatusLine = 3 ' Показываем сообщение в строке состояния
            LastMethod = 4
        End Enum


        Sub FindMethod(ByVal bstrMethodName As String, ByRef plMethodNum As Integer) Implements ILanguageExtender.FindMethod
            'Здесь 1С получает числовой идентификатор метода (процедуры или функции) по имени (названию) процедуры или функции

            plMethodNum = -1
            Select Case bstrMethodName
                Case "MessageBoxShow", "Предупреждение"
                    plMethodNum = Methods.methMessageBoxShow
                Case "ExternalEvent", "ВнешнееСобытие"
                    plMethodNum = Methods.methExternalEvent
                Case "ShowErrorLog", "Сообщить"
                    plMethodNum = Methods.methShowErrorLog
                Case "StatusLine", "Состояние"
                    plMethodNum = Methods.methStatusLine

            End Select
        End Sub

        Sub GetNParams(ByVal lMethodNum As Integer, ByRef plParams As Integer) Implements ILanguageExtender.GetNParams
            'Здесь 1С получает количество параметров у метода (процедуры или функции)

            Select Case lMethodNum
                Case Methods.methMessageBoxShow
                    plParams = 2
                Case Methods.methExternalEvent
                    plParams = 3
                Case Methods.methShowErrorLog
                    plParams = 1
                Case Methods.methStatusLine
                    plParams = 1
            End Select


        End Sub

        Sub GetNMethods(ByRef plMethods As Integer) Implements ILanguageExtender.GetNMethods
            plMethods = Methods.LastMethod
        End Sub


        Sub GetMethodName(ByVal lMethodNum As Integer, ByVal lMethodAlias As Integer, ByRef pbstrMethodName As String) Implements ILanguageExtender.GetMethodName
            'Здесь 1С (теоретически) получает имя метода по его идентификатору. lMethodAlias - номер синонима.

            pbstrMethodName = ""
        End Sub

        Sub GetParamDefValue(ByVal lMethodNum As Integer, ByVal lParamNum As Integer, ByRef pvarParamDefValue As Object) Implements ILanguageExtender.GetParamDefValue
            'Здесь 1С получает значения параметров процедуры или функции по умолчанию

            pvarParamDefValue = Nothing 'Нет значений по умолчанию
        End Sub

        Sub HasRetVal(ByVal lMethodNum As Integer, ByRef pboolRetValue As Boolean) Implements ILanguageExtender.HasRetVal
            'Здесь 1С узнает, возвращает ли метод значение (т.е. является процедурой или функцией)

            pboolRetValue = True  'Все методы у нас будут функциями (т.е. будут возвращать значение). 
        End Sub

        Sub CallAsProc(ByVal lMethodNum As Integer, ByRef paParams As System.Array) Implements ILanguageExtender.CallAsProc
            'Здесь внешняя компонента выполняет код процедур. А процедур у нас нет.

        End Sub



        Sub CallAsFunc(ByVal lMethodNum As Integer, ByRef pvarRetValue As Object, ByRef paParams As System.Array) Implements ILanguageExtender.CallAsFunc

            'Здесь внешняя компонента выполняет код функций.

            Dim icon As MessageBoxIcon


            pvarRetValue = 0 'Возвращаемое значение метода для 1С


            Select Case lMethodNum 'Порядковый номер метода

#Region "Methods.methMessageBoxShow"

                Case Methods.methMessageBoxShow
                    'Реализуем метод MessageBoxShow внешней компоненты

                    icon = MessageBoxIcon.None

                    'Преобразовываем текстовое описание значка в MessageBoxIcon.ххх

                    Select Case strMessageBoxIcon
                        Case "Asterisk"
                            icon = MessageBoxIcon.Asterisk
                        Case "Error"
                            icon = MessageBoxIcon.Error
                        Case "Exclamation"
                            icon = MessageBoxIcon.Exclamation
                        Case "Hand"
                            icon = MessageBoxIcon.Hand
                        Case "Information"
                            icon = MessageBoxIcon.Information
                        Case "None"
                            icon = MessageBoxIcon.None
                        Case "Question"
                            icon = MessageBoxIcon.Question
                        Case "Stop"
                            icon = MessageBoxIcon.Stop
                        Case "Warning"
                            icon = MessageBoxIcon.Warning
                    End Select


                    'Преобразовываем текстовое описание кнопок в MessageBoxButtons.ххх
                    Dim butt As MessageBoxButtons
                    butt = MessageBoxButtons.OK
                    Select Case strMessageBoxButtons
                        Case "AbortRetryIgnore"
                            butt = MessageBoxButtons.AbortRetryIgnore
                        Case "OK"
                            butt = MessageBoxButtons.OK
                        Case "OKCancel"
                            butt = MessageBoxButtons.OKCancel
                        Case "RetryCancel"
                            butt = MessageBoxButtons.RetryCancel
                        Case "YesNo"
                            butt = MessageBoxButtons.YesNo
                        Case "YesNoCancel"
                            butt = MessageBoxButtons.YesNoCancel
                    End Select



                    Dim res As DialogResult
                    Dim strMessageBoxText As String
                    Dim strMessageBoxHeader As String
                    Dim strDialogResult As String


                    strMessageBoxText = CType(paParams.GetValue(0), String)
                    'Получаем первый параметр нашей функции - текст предупреждения

                    strMessageBoxHeader = CType(paParams.GetValue(1), String)
                    'Получаем второй параметр нашей функции - заголовок предупреждения



                    'Показываем диалоговое окно MessageBox.Show

                    res = MessageBox.Show(
                          strMessageBoxText,
                          strMessageBoxHeader,
                          butt,
                          icon
                        )


                    'Преобразовываем результат из DialogResult.ххх в текстовую строку

                    Select Case res
                        Case DialogResult.Abort
                            strDialogResult = "Abort"
                        Case DialogResult.Cancel
                            strDialogResult = "Cancel"
                        Case DialogResult.Ignore
                            strDialogResult = "Ignore"
                        Case DialogResult.No
                            strDialogResult = "No"
                        Case DialogResult.None
                            strDialogResult = "None"
                        Case DialogResult.OK
                            strDialogResult = "OK"
                        Case DialogResult.Retry
                            strDialogResult = "Retry"
                        Case DialogResult.Yes
                            strDialogResult = "Yes"
                        Case Else
                            strDialogResult = ""
                    End Select

                    pvarRetValue = strDialogResult 'Возвращаемое значение

                    'Сбрасываем все свойства в исходное состояние

                    strMessageBoxButtons = "OK"
                    strMessageBoxIcon = "None"


#End Region ' Methods.methMessageBoxShow


            '//////////////////////////////////////////////////////////
                Case Methods.methExternalEvent  'Реализуем метод для генерации внешнего события
                    Dim s1 As String
                    Dim s2 As String
                    Dim s3 As String
                    s1 = CType(paParams.GetValue(0), String)
                    s2 = CType(paParams.GetValue(1), String)
                    s3 = CType(paParams.GetValue(2), String)

                    V7Data.AsyncEvent.ExternalEvent(s1, s2, s3)

            '//////////////////////////////////////////////////////////
                Case Methods.methShowErrorLog  'Реализуем метод для показа сообщения об ошибке

                    Dim s1 As String
                    s1 = CType(paParams.GetValue(0), String)

                    Dim ei As ExcepInfo
                    ei.wCode = 1006 'Вид пиктограммы
                    ei.bstrDescription = s1
                    ei.bstrSource = c_AddinName
                    V7Data.ErrorLog.AddError("", ei) : Throw New System.Exception("An exception has occurred.")


            '//////////////////////////////////////////////////////////
                Case Methods.methStatusLine 'Реализуем тестовый метод для изменения строки состояния

                    Dim s1 As String
                    s1 = CType(paParams.GetValue(0), String)
                    V7Data.StatusLine.SetStatusLine(s1)
                    System.Threading.Thread.Sleep(1000) 'Делаем паузу 1 сек
                    V7Data.StatusLine.ResetStatusLine()

            End Select

        End Sub

#End Region



    End Class

End Namespace

