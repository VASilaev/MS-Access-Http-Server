AHKsock_Listen(sPort, sFunction = False) {
    If (sktListen := AHKsock_Sockets("GetSocketFromNamePort", A_Space, sPort)) {
        If Not sFunction {
            AHKsock_Close(sktListen)
        } Else If (sFunction = "()") {
            Return AHKsock_Sockets("GetFunction", sktListen)
        } Else If (sFunction <> AHKsock_Sockets("GetFunction", sktListen))
            AHKsock_Sockets("SetFunction", sktListen, sFunction)
        Return
    }
    If Not IsFunc(sFunction)
        Return 2
    If (i := AHKsock_Startup())
        Return (i = 1) ? 3
                       : 4
    VarSetCapacity(aiHints, 16 + 4 * A_PtrSize, 0)
    NumPut(1, aiHints,  0, "Int")
    NumPut(2, aiHints,  4, "Int")
    NumPut(1, aiHints,  8, "Int")
    NumPut(6, aiHints, 12, "Int")
    iResult := DllCall("Ws2_32\GetAddrInfo", "Ptr", 0, "Ptr", &sPort, "Ptr", &aiHints, "Ptr*", aiResult)
    If (iResult != 0) Or ErrorLevel {
        ErrorLevel := ErrorLevel ? ErrorLevel : iResult
        Return 5
    }
    sktListen := -1
    sktListen := DllCall("Ws2_32\socket", "Int", NumGet(aiResult+0, 04, "Int")
                                        , "Int", NumGet(aiResult+0, 08, "Int")
                                        , "Int", NumGet(aiResult+0, 12, "Int"), "Ptr")
    If (sktListen = -1) Or ErrorLevel {
        sErrorLevel := ErrorLevel ? ErrorLevel : AHKsock_LastError()
        DllCall("Ws2_32\FreeAddrInfo", "Ptr", aiResult)
        ErrorLevel := sErrorLevel
        Return 6
    }
    iResult := DllCall("Ws2_32\bind", "Ptr", sktListen, "Ptr", NumGet(aiResult+0, 16 + 2 * A_PtrSize), "Int", NumGet(aiResult+0, 16, "Ptr"))
    If (iResult = -1) Or ErrorLevel {
        sErrorLevel := ErrorLevel ? ErrorLevel : AHKsock_LastError()
        DllCall("Ws2_32\closesocket",  "Ptr", sktListen)
        DllCall("Ws2_32\FreeAddrInfo", "Ptr", aiResult)
        ErrorLevel := sErrorLevel
        Return 7
    }
    DllCall("Ws2_32\FreeAddrInfo", "Ptr", aiResult)
    AHKsock_Sockets("Add", sktListen, A_Space, A_Space, sPort, sFunction)
    If AHKsock_RegisterAsyncSelect(sktListen) {
        sErrorLevel := ErrorLevel
        DllCall("Ws2_32\closesocket", "Ptr", sktListen)
        AHKsock_Sockets("Delete", sktListen)
        ErrorLevel := sErrorLevel
        Return 8
    }
    iResult := DllCall("Ws2_32\listen", "Ptr", sktListen, "Int", 0x7FFFFFFF)
    If (iResult = -1) Or ErrorLevel {
        sErrorLevel := ErrorLevel ? ErrorLevel : AHKsock_LastError()
        DllCall("Ws2_32\closesocket", "Ptr", sktListen)
        AHKsock_Sockets("Delete", sktListen)
        ErrorLevel := sErrorLevel
        Return 9
    }
}

AHKsock_Connect(sName, sPort, sFunction) {
    Static aiResult, iPointer, bProcessing, iMessage
    Static sCurName, sCurPort, sCurFunction, sktConnect
    If (Not sName And Not sPort And Not sFunction)
        Return bProcessing
    If bProcessing And (sFunction != iMessage) {
        ErrorLevel := sCurName A_Tab sCurPort
        Return 1
    } Else If bProcessing {
        If (i := sPort >> 16) {
            DllCall("Ws2_32\closesocket", "Ptr", sktConnect)
            iPointer := NumGet(iPointer+0, 16 + 3 * A_PtrSize)
            If (iPointer = 0) {
                DllCall("Ws2_32\FreeAddrInfo", "Ptr", aiResult)
                bProcessing := False
                ErrorLevel := i
                AHKsock_RaiseError(1)
                If IsFunc(sCurFunction)
                    %sCurFunction%("CONNECTED", -1, sCurName, 0, sCurPort)
                Return
            }
        } Else {
            sIP := DllCall("Ws2_32\inet_ntoa", "UInt", NumGet(NumGet(iPointer+0, 16 + 2 * A_PtrSize)+4, 0, "UInt"), "AStr")
            DllCall("Ws2_32\FreeAddrInfo", "Ptr", aiResult)
            AHKsock_Sockets("Add", sktConnect, sCurName, sIP, sCurPort, sCurFunction)
            bProcessing := False
            Critical
            If AHKsock_RegisterAsyncSelect(sktConnect) {
                sErrorLevel := ErrorLevel ? ErrorLevel : AHKsock_LastError()
                DllCall("Ws2_32\closesocket", "Ptr", sktConnect)
                AHKsock_Sockets("Delete", sktConnect)
                ErrorLevel := sErrorLevel
                AHKsock_RaiseError(2)
                If IsFunc(sCurFunction)
                    %sCurFunction%("CONNECTED", -1, sCurName, 0, sCurPort)
            } Else If IsFunc(sCurFunction)
                %sCurFunction%("CONNECTED", sktConnect, sCurName, sIP, sCurPort)
            Return
        }
    } Else {
        If Not IsFunc(sFunction)
            Return 2
        bProcessing := True
        sCurName := sName
        sCurPort := sPort
        sCurFunction := sFunction
        If (i := AHKsock_Startup()) {
            bProcessing := False
            Return (i = 1) ? 3
                           : 4
        }
        VarSetCapacity(aiHints, 16 + 4 * A_PtrSize, 0)
        NumPut(2, aiHints,  4, "Int")
        NumPut(1, aiHints,  8, "Int")
        NumPut(6, aiHints, 12, "Int")
        iResult := DllCall("Ws2_32\GetAddrInfo", "Ptr", &sName, "Ptr", &sPort, "Ptr", &aiHints, "Ptr*", aiResult)
        If (iResult != 0) Or ErrorLevel {
            ErrorLevel := ErrorLevel ? ErrorLevel : iResult
            bProcessing := False
            Return 5
        }
        iPointer := aiResult
    }
    sktConnect := DllCall("Ws2_32\socket", "Int", NumGet(iPointer+0, 04, "Int")
                                         , "Int", NumGet(iPointer+0, 08, "Int")
                                         , "Int", NumGet(iPointer+0, 12, "Int"), "Ptr")
    If (sktConnect = 0xFFFFFFFF) Or ErrorLevel {
        sErrorLevel := ErrorLevel ? ErrorLevel : AHKsock_LastError()
        DllCall("Ws2_32\FreeAddrInfo", "Ptr", aiResult)
        bProcessing := False
        ErrorLevel := sErrorLevel
        If (sFunction = iMessage) {
            AHKsock_RaiseError(3)
            If IsFunc(sCurFunction)
                %sCurFunction%("CONNECTED", -1)
        }
        Return 6
    }
    iMessage := AHKsock_Settings("Message") + 1
    If AHKsock_RegisterAsyncSelect(sktConnect, 16, "AHKsock_Connect", iMessage) {
        sErrorLevel := ErrorLevel
        DllCall("Ws2_32\FreeAddrInfo", "Ptr", aiResult)
        DllCall("Ws2_32\closesocket",  "Ptr", sktConnect)
        bProcessing := False
        ErrorLevel := sErrorLevel
        If (sFunction = iMessage) {
            AHKsock_RaiseError(4)
            If IsFunc(sCurFunction)
                %sCurFunction%("CONNECTED", -1)
        }
        Return 7
    }
    iResult := DllCall("Ws2_32\connect", "Ptr", sktConnect, "Ptr", NumGet(iPointer+0, 16 + 2 * A_PtrSize), "Int", NumGet(iPointer+0, 16))
    If ErrorLevel Or ((iResult = -1) And (AHKsock_LastError() != 10035)) {
        sErrorLevel := ErrorLevel ? ErrorLevel : AHKsock_LastError()
        DllCall("Ws2_32\FreeAddrInfo", "Ptr", aiResult)
        DllCall("Ws2_32\closesocket",  "Ptr", sktConnect)
        bProcessing := False
        ErrorLevel := sErrorLevel
        If (sFunction = iMessage) {
            AHKsock_RaiseError(5)
            If IsFunc(sCurFunction)
                %sCurFunction%("CONNECTED", -1)
        }
        Return 8
    }
}

AHKsock_Send(iSocket, ptrData = 0, iLength = 0) {
    If Not AHKsock_Sockets("Index", iSocket)
        Return -4
    If Not AHKsock_Startup(1)
        Return -1
    If Not AHKsock_Sockets("GetSend", iSocket)
        Return -5
    iSendResult := DllCall("Ws2_32\send", "Ptr", iSocket, "Ptr", ptrData, "Int", iLength, "Int", 0)
    If (iSendResult = -1) And ((iErr := AHKsock_LastError()) = 10035) {
        AHKsock_Sockets("SetSend", iSocket, False)
        Return -2
    } Else If (iSendResult = -1) Or ErrorLevel {
        ErrorLevel := ErrorLevel ? ErrorLevel : iErr
        Return -3
    } Else Return iSendResult
}

AHKsock_ForceSend(iSocket, ptrData, iLength) {
    If Not AHKsock_Startup(1)
        Return -1
    If Not AHKsock_Sockets("Index", iSocket)
        Return -4
    If A_IsCritical
        Return -5
    Thread, Priority, 0
    If ((iMaxChunk := AHKsock_SockOpt(iSocket, "SO_SNDBUF")) = -1)
        Return -6
    If (iMaxChunk <= 1) {
        Loop {
            While Not AHKsock_Sockets("GetSend", iSocket)
                Sleep -1
            Loop {
                If ((iSendResult := AHKsock_Send(iSocket, ptrData, iLength)) < 0) {
                    If (iSendResult = -2)
                        Break
                    Else Return iSendResult
                } Else {
                    If (iSendResult < iLength)
                        ptrData += iSendResult, iLength -= iSendResult
                    Else Return
                }
            }
        }
    } Else {
        iMaxChunk -= 1
        Loop {
            While Not AHKsock_Sockets("GetSend", iSocket)
                Sleep -1
            If (iLength < iMaxChunk) {
                Loop {
                    If ((iSendResult := AHKsock_Send(iSocket, ptrData, iLength)) < 0) {
                        If (iSendResult = -2)
                            Break
                        Else Return iSendResult
                    } Else {
                        If (iSendResult < iLength)
                            ptrData += iSendResult, iLength -= iSendResult
                        Else Return
                    }
                }
            } Else {
                If ((iSendResult := AHKsock_Send(iSocket, ptrData, iMaxChunk)) < 0) {
                    If (iSendResult = -2)
                        Continue
                    Else Return iSendResult
                } Else ptrData += iSendResult, iLength -= iSendResult
            }
        }
    }
}

AHKsock_Close(iSocket = -1, iTimeout = 5000) {
    If Not AHKsock_Startup(1)
        Return
    If (iSocket = -1) {
        If Not AHKsock_Sockets() {
            DllCall("Ws2_32\WSACleanup")
            AHKsock_Startup(2)
            Return
        }
        iStartClose := A_TickCount
        Loop % AHKsock_Sockets()
            AHKsock_ShutdownSocket(AHKsock_Sockets("GetSocketFromIndex", A_Index))
        If Not A_ExitReason {
            A_IsCriticalOld := A_IsCritical
            Critical, Off
            Thread, Priority, 0
            While (AHKsock_Sockets()) And (A_TickCount - iStartClose < iTimeout)
                Sleep, -1
            Critical, %A_IsCriticalOld%
        }
        DllCall("Ws2_32\WSACleanup")
        AHKsock_Startup(2)
    } Else If AHKsock_ShutdownSocket(iSocket)
        Return 1
}

AHKsock_GetAddrInfo(sHostName, ByRef sIPList, bOne = False) {
    If (i := AHKsock_Startup())
        Return i
    VarSetCapacity(aiHints, 16 + 4 * A_PtrSize, 0)
    NumPut(2, aiHints,  4, "Int")
    NumPut(1, aiHints,  8, "Int")
    NumPut(6, aiHints, 12, "Int")
    iResult := DllCall("Ws2_32\GetAddrInfo", "Ptr", &sHostName, "Ptr", 0, "Ptr", &aiHints, "Ptr*", aiResult)
    If (iResult = 11001)
        Return 3
    Else If (iResult != 0) Or ErrorLevel {
        ErrorLevel := ErrorLevel ? ErrorLevel : iResult
        Return 4
    }
    If bOne
        sIPList := DllCall("Ws2_32\inet_ntoa", "UInt", NumGet(NumGet(aiResult+0, 16 + 2 * A_PtrSize)+4, 0, "UInt"), "AStr")
    Else {
        iPointer := aiResult, sIPList := ""
        While iPointer {
            s := DllCall("Ws2_32\inet_ntoa", "UInt", NumGet(NumGet(iPointer+0, 16 + 2 * A_PtrSize)+4, 0, "UInt"), "AStr")
            iPointer := NumGet(iPointer+0, 16 + 3 * A_PtrSize)
            sIPList .=  s (iPointer ? "`n" : "")
        }
    }
    DllCall("Ws2_32\FreeAddrInfo", "Ptr", aiResult)
}

AHKsock_GetNameInfo(sIP, ByRef sHostName, sPort = 0, ByRef sService = "") {
    If (i := AHKsock_Startup())
        Return i
    iIP := DllCall("Ws2_32\inet_addr", "AStr", sIP, "UInt")
    If (iIP = 0 Or iIP = 0xFFFFFFFF)
        Return 3
    VarSetCapacity(tSockAddr, 16, 0)
    NumPut(2,   tSockAddr, 0, "Short")
    NumPut(iIP, tSockAddr, 4, "UInt")
    If sPort
        NumPut(DllCall("Ws2_32\htons", "UShort", sPort, "UShort"), tSockAddr, 2, "UShort")
    VarSetCapacity(sHostName, 1025 * 2, 0)
    If sPort
        VarSetCapacity(sService, 32 * 2, 0)
    iResult := DllCall("Ws2_32\GetNameInfoW", "Ptr", &tSockAddr, "Int", 16, "Str", sHostName, "UInt", 1025 * 2
                                           , sPort ? "Str" : "UInt", sPort ? sService : 0, "UInt", 32 * 2, "Int", 0)
    If (iResult != 0) Or ErrorLevel {
        ErrorLevel := ErrorLevel ? ErrorLevel : DllCall("Ws2_32\WSAGetLastError")
        Return 4
    }
}

AHKsock_SockOpt(iSocket, sOption, iValue = -1) {
    VarSetCapacity(iOptVal, iOptValLength := 4, 0)
    If (iValue <> -1)
        NumPut(iValue, iOptVal, 0, "UInt")
    If (sOption = "SO_KEEPALIVE") {
        intLevel := 0xFFFF
        intOptName := 0x0008
    } Else If (sOption = "SO_SNDBUF") {
        intLevel := 0xFFFF
        intOptName := 0x1001
    } Else If (sOption = "SO_RCVBUF") {
        intLevel := 0xFFFF
        intOptName := 0x1002
    } Else If (sOption = "TCP_NODELAY") {
        intLevel := 6
        intOptName := 0x0001
    }
    If (iValue = -1) {
        iResult := DllCall("Ws2_32\getsockopt", "Ptr", iSocket, "Int", intLevel, "Int", intOptName
                                              , "UInt*", iOptVal, "Int*", iOptValLength)
        If (iResult = -1) Or ErrorLevel {
            ErrorLevel := ErrorLevel ? ErrorLevel : AHKsock_LastError()
            Return -1
        } Else Return iOptVal
    } Else {
        iResult := DllCall("Ws2_32\setsockopt", "Ptr", iSocket, "Int", intLevel, "Int", intOptName
                                              , "Ptr", &iOptVal, "Int",  iOptValLength)
        If (iResult = -1) Or ErrorLevel {
            ErrorLevel := ErrorLevel ? ErrorLevel : AHKsock_LastError()
            Return -2
        }
    }
}

AHKsock_Startup(iMode = 0) {
    Static bAlreadyStarted
    If (iMode = 2)
        bAlreadyStarted := False
    Else If (iMode = 1)
        Return bAlreadyStarted
    Else If Not bAlreadyStarted {
        VarSetCapacity(wsaData, A_PtrSize = 4 ? 400 : 408, 0)
        iResult := DllCall("Ws2_32\WSAStartup", "UShort", 0x0202, "Ptr", &wsaData)
        If (iResult != 0) Or ErrorLevel {
            ErrorLevel := ErrorLevel ? ErrorLevel : iResult
            Return 1
        }
        If (NumGet(wsaData, 2, "UShort") < 0x0202) {
            DllCall("Ws2_32\WSACleanup")
            ErrorLevel := "The Winsock DLL does not support version 2.2."
            Return 2
        }
        bAlreadyStarted := True
    }
}

AHKsock_ShutdownSocket(iSocket) {
    sName := AHKsock_Sockets("GetName", iSocket)
    If (sName != A_Space) {
        iResult := DllCall("Ws2_32\shutdown", "Ptr", iSocket, "Int", 1)
        If (iResult = -1) Or ErrorLevel {
            sErrorLevel := ErrorLevel ? ErrorLevel : AHKsock_LastError()
            DllCall("Ws2_32\closesocket", "Ptr", iSocket)
            AHKsock_Sockets("Delete", iSocket)
            ErrorLevel := sErrorLevel
            Return 1
        }
        AHKsock_Sockets("SetShutdown", iSocket)
    } Else {
        DllCall("Ws2_32\closesocket", "Ptr", iSocket)
        AHKsock_Sockets("Delete", iSocket)
    }
}

AHKsock_RegisterAsyncSelect(iSocket, fFlags = 43, sFunction = "AHKsock_AsyncSelect", iMsg = 0) {
    Static hwnd := False
    If Not hwnd {
        A_DetectHiddenWindowsOld := A_DetectHiddenWindows
        DetectHiddenWindows, On
        WinGet, hwnd, ID, % "ahk_pid " DllCall("GetCurrentProcessId") " ahk_class AutoHotkey"
        DetectHiddenWindows, %A_DetectHiddenWindowsOld%
    }
    iMsg := iMsg ? iMsg : AHKsock_Settings("Message")
    If (OnMessage(iMsg) <> sFunction)
        OnMessage(iMsg, sFunction)
    iResult := DllCall("Ws2_32\WSAAsyncSelect", "Ptr", iSocket, "Ptr", hwnd, "UInt", iMsg, "Int", fFlags)
    If (iResult = -1) Or ErrorLevel {
        ErrorLevel := ErrorLevel ? ErrorLevel : AHKsock_LastError()
        Return 1
    }
}

AHKsock_AsyncSelect(wParam, lParam) {
    Critical
    If Not AHKsock_Sockets("Index", wParam)
        Return
    iEvent := lParam & 0xFFFF, iErrorCode := lParam >> 16
    If (iEvent = 1) {
        If iErrorCode {
            ErrorLevel := iErrorCode
            AHKsock_RaiseError(6, wParam)
            Return
        }
        VarSetCapacity(bufReceived, bufReceivedLength := AHKsock_Settings("Buffer"), 0)
        iResult := DllCall("Ws2_32\recv", "UInt", wParam, "Ptr", &bufReceived, "Int", bufReceivedLength, "Int", 0)
        If (iResult > 0) {
            VarSetCapacity(bufReceived, -1)
            If IsFunc(sFunc := AHKsock_Sockets("GetFunction", wParam))
                %sFunc%("RECEIVED", wParam, AHKsock_Sockets("GetName", wParam)
                                          , AHKsock_Sockets("GetAddr", wParam)
                                          , AHKsock_Sockets("GetPort", wParam), bufReceived, iResult)
        } Else If ErrorLevel Or ((iResult = -1) And Not ((iErrorCode := AHKsock_LastError()) = 10035)) {
            ErrorLevel := ErrorLevel ? ErrorLevel : iErrorCode
            AHKsock_RaiseError(7, wParam)
            iResult = -1
        }
        Return iResult
    } Else If (iEvent = 2) {
        If iErrorCode {
            ErrorLevel := iErrorCode
            AHKsock_RaiseError(8, wParam)
            Return
        }
        AHKsock_Sockets("SetSend", wParam, True)
        If Not AHKsock_Sockets("GetShutdown", wParam)
            If IsFunc(sFunc := AHKsock_Sockets("GetFunction", wParam))
                %sFunc%("SEND", wParam, AHKsock_Sockets("GetName", wParam)
                                      , AHKsock_Sockets("GetAddr", wParam)
                                      , AHKsock_Sockets("GetPort", wParam))
    } Else If (iEvent = 8) {
        If iErrorCode {
            ErrorLevel := iErrorCode
            AHKsock_RaiseError(9, wParam)
            Return
        }
        VarSetCapacity(tSockAddr, tSockAddrLength := 16, 0)
        sktClient := DllCall("Ws2_32\accept", "Ptr", wParam, "Ptr", &tSockAddr, "Int*", tSockAddrLength)
        If (sktClient = -1) And ((iErrorCode := AHKsock_LastError()) = 10035)
            Return
        Else If (sktClient = -1) Or ErrorLevel {
            ErrorLevel := ErrorLevel ? ErrorLevel : iErrorCode
            AHKsock_RaiseError(10, wParam)
            Return
        }
        sName := ""
        sAddr := DllCall("Ws2_32\inet_ntoa", "UInt", NumGet(tSockAddr, 4, "UInt"), "AStr")
        sPort := AHKsock_Sockets("GetPort", wParam)
        sFunc := AHKsock_Sockets("GetFunction", wParam)
        AHKsock_Sockets("Add", sktClient, sName, sAddr, sPort, sFunc)
        iResult := DllCall("Ws2_32\listen", "Ptr", wParam, "Int", 0x7FFFFFFF)
        If (iResult = -1) Or ErrorLevel {
            sErrorLevel := ErrorLevel ? ErrorLevel : AHKsock_LastError()
            DllCall("Ws2_32\closesocket", "Ptr", wParam)
            AHKsock_Sockets("Delete", wParam)
            ErrorLevel := sErrorLevel
            AHKsock_RaiseError(12, wParam)
            Return
        }
        If IsFunc(sFunc)
            %sFunc%("ACCEPTED", sktClient, sName, sAddr, sPort)
    } Else If (iEvent = 32) {
        While (AHKsock_AsyncSelect(wParam, 1) > 0)
            Sleep, -1
        If Not AHKsock_Sockets("GetShutdown", wParam) {
            If IsFunc(sFunc := AHKsock_Sockets("GetFunction", wParam))
                %sFunc%("SENDLAST", wParam, AHKsock_Sockets("GetName", wParam)
                                          , AHKsock_Sockets("GetAddr", wParam)
                                          , AHKsock_Sockets("GetPort", wParam))
            If AHKsock_ShutdownSocket(wParam) {
                AHKsock_RaiseError(13, wParam)
                Return
            }
        }
        DllCall("Ws2_32\closesocket", "Ptr", wParam)
        sFunc := AHKsock_Sockets("GetFunction", wParam)
        sName := AHKsock_Sockets("GetName", wParam)
        sAddr := AHKsock_Sockets("GetAddr", wParam)
        sPort := AHKsock_Sockets("GetPort", wParam)
        AHKsock_Sockets("Delete", wParam)
        If IsFunc(sFunc)
            %sFunc%("DISCONNECTED", wParam, sName, sAddr, sPort)
    }
}

AHKsock_Sockets(sAction = "Count", iSocket = "", sName = "", sAddr = "", sPort = "", sFunction = "") {
    Static
    Static aSockets0 := 0
    Static iLastSocket := 0xFFFFFFFF
    Local i, ret, A_IsCriticalOld
    A_IsCriticalOld := A_IsCritical
    Critical
    If (sAction = "Count") {
        ret := aSockets0
    } Else If (sAction = "Add") {
        aSockets0 += 1
        aSockets%aSockets0%_Sock := iSocket
        aSockets%aSockets0%_Name := sName
        aSockets%aSockets0%_Addr := sAddr
        aSockets%aSockets0%_Port := sPort
        aSockets%aSockets0%_Func := sFunction
        aSockets%aSockets0%_Shutdown := False
        aSockets%aSockets0%_Send := False
    } Else If (sAction = "Delete") {
        i := (iSocket = iLastSocket)
        ? iLastSocketIndex
        : AHKsock_Sockets("Index", iSocket)
        If i {
            iLastSocket := 0xFFFF
            If (i < aSockets0) {
                aSockets%i%_Sock := aSockets%aSockets0%_Sock
                aSockets%i%_Name := aSockets%aSockets0%_Name
                aSockets%i%_Addr := aSockets%aSockets0%_Addr
                aSockets%i%_Port := aSockets%aSockets0%_Port
                aSockets%i%_Func := aSockets%aSockets0%_Func
                aSockets%i%_Shutdown := aSockets%aSockets0%_Shutdown
                aSockets%i%_Send := aSockets%aSockets0%_Send
            }
            aSockets0 -= 1
        }
    } Else If (sAction = "GetName") {
        i := (iSocket = iLastSocket)
        ? iLastSocketIndex
        : AHKsock_Sockets("Index", iSocket)
        ret := aSockets%i%_Name
    } Else If (sAction = "GetAddr") {
        i := (iSocket = iLastSocket)
        ? iLastSocketIndex
        : AHKsock_Sockets("Index", iSocket)
        ret := aSockets%i%_Addr
    } Else If (sAction = "GetPort") {
        i := (iSocket = iLastSocket)
        ? iLastSocketIndex
        : AHKsock_Sockets("Index", iSocket)
        ret := aSockets%i%_Port
    } Else If (sAction = "GetFunction") {
        i := (iSocket = iLastSocket)
        ? iLastSocketIndex
        : AHKsock_Sockets("Index", iSocket)
        ret := aSockets%i%_Func
    } Else If (sAction = "SetFunction") {
        i := (iSocket = iLastSocket)
        ? iLastSocketIndex
        : AHKsock_Sockets("Index", iSocket)
        aSockets%i%_Func := sName
    } Else If (sAction = "GetSend") {
        i := (iSocket = iLastSocket)
        ? iLastSocketIndex
        : AHKsock_Sockets("Index", iSocket)
        ret := aSockets%i%_Send
    } Else If (sAction = "SetSend") {
        i := (iSocket = iLastSocket)
        ? iLastSocketIndex
        : AHKsock_Sockets("Index", iSocket)
        aSockets%i%_Send := sName
    } Else If (sAction = "GetShutdown") {
        i := (iSocket = iLastSocket)
        ? iLastSocketIndex
        : AHKsock_Sockets("Index", iSocket)
        ret := aSockets%i%_Shutdown
    } Else If (sAction = "SetShutdown") {
        i := (iSocket = iLastSocket)
        ? iLastSocketIndex
        : AHKsock_Sockets("Index", iSocket)
        aSockets%i%_Shutdown := True
    } Else If (sAction = "GetSocketFromNamePort") {
        Loop % aSockets0 {
            If (aSockets%A_Index%_Name = iSocket)
            And (aSockets%A_Index%_Port = sName) {
                ret := aSockets%A_Index%_Sock
                Break
            }
        }
    } Else If (sAction = "GetSocketFromIndex") {
        ret := aSockets%iSocket%_Sock
    } Else If (sAction = "Index") {
        Loop % aSockets0 {
            If (aSockets%A_Index%_Sock = iSocket) {
                iLastSocketIndex := A_Index, iLastSocket := iSocket
                ret := A_Index
                Break
            }
        }
    }
    Critical %A_IsCriticalOld%
    Return ret
}

AHKsock_LastError() {
    Return DllCall("Ws2_32\WSAGetLastError")
}

AHKsock_ErrorHandler(sFunction = """") {
    Static sCurrentFunction
    If (sFunction = """")
        Return sCurrentFunction
    Else sCurrentFunction := sFunction
}

AHKsock_RaiseError(iError, iSocket = -1) {
    If IsFunc(sFunc := AHKsock_ErrorHandler())
        %sFunc%(iError, iSocket)
}

AHKsock_Settings(sSetting, sValue = "") {
    Static iMessage := 0x8000
    Static iBuffer := 65536
    If (sSetting = "Message") {
        If Not sValue
            Return iMessage
        Else iMessage := (sValue = "Reset") ? 0x8000 : sValue
    } Else If (sSetting = "Buffer") {
        If Not sValue
            Return iBuffer
        Else iBuffer := (sValue = "Reset") ? 65536 : sValue
    }
}
