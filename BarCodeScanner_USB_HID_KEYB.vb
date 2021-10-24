Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Net
Imports System.Net.Sockets
Imports System.Text

Namespace BarCodeScanner_USB_HID_KEYB_vb

    '����� ���� ������ ������ � ����
    <ComVisible(True), Guid("9840BED5-950E-4679-B4DD-F32E8C86200C"), ProgId("AddIn.BarCodeScanner_USB_HID_KEYB_vb"), ClassInterface(ClassInterfaceType.AutoDispatch)>
    Public Class BarCodeScanner_USB_HID_KEYB_vb
        Implements IInitDone
        Implements ILanguageExtender

        Const c_AddinName As String = "BarCodeScanner_USB_HID_KEYB_vb"

#Region "����������"
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
        Public Sub New() ' ����������� ��� COM �������������
        End Sub

        Private Sub Init(<MarshalAs(UnmanagedType.IDispatch)> ByVal pConnection As Object) Implements IInitDone.Init
            '  ��������� pConnection ��� �����������
            strMessageBoxButtons = "OK" : strMessageBoxIcon = "None"

            If (True) Then ' (V7Data.V7Object.ToString <> "System.__ComObject")
                V7Data.V7Object = pConnection
                Timer1.Interval = 1777 : Timer1.AutoReset = True : AddHandler Timer1.Elapsed, AddressOf Timer1_Tick
                Timer2.Interval = 77 : Timer2.AutoReset = True : AddHandler Timer2.Elapsed, AddressOf Timer2_Tick
            End If

            If (Not Timer1.Enabled) Then Timer1.Start() '��� �� ������� ������� ��������� ������ ������

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
                    While (TcpClient1.Connected) AndAlso (idle_event <= 777.777) '���� ����� ������� ���� �� ������ ���� ������
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
            'ei.wCode = 1006 '��� �����������
            'ei.bstrDescription = "" + DateTime.Now.ToString()
            'ei.bstrSource = c_AddinName
            'V7Data.ErrorLog.AddError("", ei) : Throw New System.Exception("An exception has occurred.")
            'MsgBox("" + DateTime.Now.ToString())
            'MessageBox.Show("Timer2_Tick = ��" & vbNewLine & c_AddinName, "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            V7Data.AsyncEvent.ExternalEvent("s1", "s2", "s3")
            'MessageBox.Show("Timer2_Tick = ��" & vbNewLine & c_AddinName, "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            'V7Data.StatusLine.SetStatusLine("" + DateTime.Now.ToString())
            'System.Threading.Thread.Sleep(1000) '������ ����� 1 ���
            'V7Data.StatusLine.ResetStatusLine()
            strTableRegDevices = "" & DateTime.Now.ToString() & " ���"
            'Timer2 .Enabled = True
        End Sub


#Region "��������"
        Enum Props
            '�������� �������������� ������� ����� ������� ����������

            propMessageBoxIcon = 0  '����������� ��� MessageBox'�
            propMessageBoxButtons = 1 '������ ��� MessageBox'a
            propCountOpenDevices = 2
            propTableRegDevices = 3
            LastProp = 4

        End Enum

        Sub FindProp(ByVal bstrPropName As String, ByRef plPropNum As Integer) Implements ILanguageExtender.FindProp
            '����� 1� ���� �������� ������������� �������� �� ��� ���������� �����

            Select Case bstrPropName

                Case "MessageBoxIcon", "�������������������������"
                    plPropNum = Props.propMessageBoxIcon

                Case "MessageBoxButtons", "��������������������"
                    plPropNum = Props.propMessageBoxButtons

                Case "CountOpenDevices", "���������������������������"
                    plPropNum = Props.propCountOpenDevices

                Case "TableRegDevices", "����������������������"
                    plPropNum = Props.propTableRegDevices

                Case Else
                    plPropNum = -1
            End Select
        End Sub

        Sub GetPropVal(ByVal lPropNum As Integer, ByRef pvarPropVal As Object) Implements ILanguageExtender.GetPropVal
            '����� 1� ������ �������� ������� 

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
            '����� 1� �������� �������� ������� 

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
            '����� 1� ������, ����� �������� �������� ��� ������

            pboolPropRead = True ' ��� �������� �������� ��� ������
        End Sub

        Sub IsPropWritable(ByVal lPropNum As Integer, ByRef pboolPropWrite As Boolean) Implements ILanguageExtender.IsPropWritable
            '����� 1� ������, ����� �������� �������� ��� ������

            pboolPropWrite = True ' ��� �������� �������� ��� ������
        End Sub

        Sub GetNProps(ByRef plProps As Integer) Implements ILanguageExtender.GetNProps
            '����� 1� �������� ���������� ��������� �� �� �������

            plProps = Props.LastProp
        End Sub

        Sub GetPropName(ByVal lPropNum As Integer, ByVal lPropAlias As Integer, ByRef pbstrPropName As String) Implements ILanguageExtender.GetPropName
            '����� 1� (������������) ������ ��� �������� �� ��� ��������������. lPropAlias - ����� ����������

            pbstrPropName = ""
        End Sub

#End Region

#Region "������"

        Enum Methods
            '�������� �������������� ������� (�������� ��� �������) ����� ������� ����������

            methMessageBoxShow = 0 '��������������, �� � ������������ �������� ����������� � ��������� ����
            methExternalEvent = 1 '���������� ������� ������� (1� ������������� ��� � ��������� ������������������������())
            methShowErrorLog = 2 ' ���������� ��������� ��������� 
            methStatusLine = 3 ' ���������� ��������� � ������ ���������
            LastMethod = 4
        End Enum


        Sub FindMethod(ByVal bstrMethodName As String, ByRef plMethodNum As Integer) Implements ILanguageExtender.FindMethod
            '����� 1� �������� �������� ������������� ������ (��������� ��� �������) �� ����� (��������) ��������� ��� �������

            plMethodNum = -1
            Select Case bstrMethodName
                Case "MessageBoxShow", "��������������"
                    plMethodNum = Methods.methMessageBoxShow
                Case "ExternalEvent", "��������������"
                    plMethodNum = Methods.methExternalEvent
                Case "ShowErrorLog", "��������"
                    plMethodNum = Methods.methShowErrorLog
                Case "StatusLine", "���������"
                    plMethodNum = Methods.methStatusLine

            End Select
        End Sub

        Sub GetNParams(ByVal lMethodNum As Integer, ByRef plParams As Integer) Implements ILanguageExtender.GetNParams
            '����� 1� �������� ���������� ���������� � ������ (��������� ��� �������)

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
            '����� 1� (������������) �������� ��� ������ �� ��� ��������������. lMethodAlias - ����� ��������.

            pbstrMethodName = ""
        End Sub

        Sub GetParamDefValue(ByVal lMethodNum As Integer, ByVal lParamNum As Integer, ByRef pvarParamDefValue As Object) Implements ILanguageExtender.GetParamDefValue
            '����� 1� �������� �������� ���������� ��������� ��� ������� �� ���������

            pvarParamDefValue = Nothing '��� �������� �� ���������
        End Sub

        Sub HasRetVal(ByVal lMethodNum As Integer, ByRef pboolRetValue As Boolean) Implements ILanguageExtender.HasRetVal
            '����� 1� ������, ���������� �� ����� �������� (�.�. �������� ���������� ��� ��������)

            pboolRetValue = True  '��� ������ � ��� ����� ��������� (�.�. ����� ���������� ��������). 
        End Sub

        Sub CallAsProc(ByVal lMethodNum As Integer, ByRef paParams As System.Array) Implements ILanguageExtender.CallAsProc
            '����� ������� ���������� ��������� ��� ��������. � �������� � ��� ���.

        End Sub



        Sub CallAsFunc(ByVal lMethodNum As Integer, ByRef pvarRetValue As Object, ByRef paParams As System.Array) Implements ILanguageExtender.CallAsFunc

            '����� ������� ���������� ��������� ��� �������.

            Dim icon As MessageBoxIcon


            pvarRetValue = 0 '������������ �������� ������ ��� 1�


            Select Case lMethodNum '���������� ����� ������

#Region "Methods.methMessageBoxShow"

                Case Methods.methMessageBoxShow
                    '��������� ����� MessageBoxShow ������� ����������

                    icon = MessageBoxIcon.None

                    '��������������� ��������� �������� ������ � MessageBoxIcon.���

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


                    '��������������� ��������� �������� ������ � MessageBoxButtons.���
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
                    '�������� ������ �������� ����� ������� - ����� ��������������

                    strMessageBoxHeader = CType(paParams.GetValue(1), String)
                    '�������� ������ �������� ����� ������� - ��������� ��������������



                    '���������� ���������� ���� MessageBox.Show

                    res = MessageBox.Show(
                          strMessageBoxText,
                          strMessageBoxHeader,
                          butt,
                          icon
                        )


                    '��������������� ��������� �� DialogResult.��� � ��������� ������

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

                    pvarRetValue = strDialogResult '������������ ��������

                    '���������� ��� �������� � �������� ���������

                    strMessageBoxButtons = "OK"
                    strMessageBoxIcon = "None"


#End Region ' Methods.methMessageBoxShow


            '//////////////////////////////////////////////////////////
                Case Methods.methExternalEvent  '��������� ����� ��� ��������� �������� �������
                    Dim s1 As String
                    Dim s2 As String
                    Dim s3 As String
                    s1 = CType(paParams.GetValue(0), String)
                    s2 = CType(paParams.GetValue(1), String)
                    s3 = CType(paParams.GetValue(2), String)

                    V7Data.AsyncEvent.ExternalEvent(s1, s2, s3)

            '//////////////////////////////////////////////////////////
                Case Methods.methShowErrorLog  '��������� ����� ��� ������ ��������� �� ������

                    Dim s1 As String
                    s1 = CType(paParams.GetValue(0), String)

                    Dim ei As ExcepInfo
                    ei.wCode = 1006 '��� �����������
                    ei.bstrDescription = s1
                    ei.bstrSource = c_AddinName
                    V7Data.ErrorLog.AddError("", ei) : Throw New System.Exception("An exception has occurred.")


            '//////////////////////////////////////////////////////////
                Case Methods.methStatusLine '��������� �������� ����� ��� ��������� ������ ���������

                    Dim s1 As String
                    s1 = CType(paParams.GetValue(0), String)
                    V7Data.StatusLine.SetStatusLine(s1)
                    System.Threading.Thread.Sleep(1000) '������ ����� 1 ���
                    V7Data.StatusLine.ResetStatusLine()

            End Select

        End Sub

#End Region



    End Class

End Namespace

