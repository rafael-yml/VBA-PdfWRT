VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PdfWRT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Windows.Data.Pdf WinRT - renders PDF pages to PNG, no external tools or Kernel32.
' vtable offsets: LoadFromStreamAsync=8, GetPage=6, get_PageCount=7,
'   RenderToStreamAsync=6, RenderToStreamWithOptionsAsync=7, set_DestinationWidth=9
' Based on UWPOCR by DanysysTeam; modified by rafael-yml to use msvcrt.dll.

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

#If VBA7 Then
    Private Declare PtrSafe Sub MemoryCopy Lib "msvcrt.dll" Alias "memcpy" (ByVal Destination As LongPtr, ByVal source As LongPtr, ByVal Length As LongPtr)
    Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As LongPtr, ByRef retVAR As Variant) As Long
    Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal OleStringCLSID As LongPtr, ByRef cGUID As Any) As Long
    Private Declare PtrSafe Function RoGetActivationFactory Lib "Combase.dll" (ByVal activatableClassId As LongPtr, ByRef riid As GUID, ByRef factory As LongPtr) As Long
    Private Declare PtrSafe Function WindowsCreateString Lib "Combase.dll" (ByVal sourceString As LongPtr, ByVal Length As Long, ByRef hString As LongPtr) As Long
    Private Declare PtrSafe Function WindowsDeleteString Lib "Combase.dll" (ByVal sourceString As LongPtr) As Long
    Private Declare PtrSafe Function WindowsGetStringRawBuffer Lib "Combase.dll" (ByVal hString As LongPtr, ByRef Length As Long) As LongPtr
    Private Declare PtrSafe Function CreateRandomAccessStreamOnFile Lib "Shcore.dll" (ByVal pFilePath As LongPtr, ByVal Access As Long, ByRef riid As GUID, ByRef ppv As LongPtr) As Long
    Private Declare PtrSafe Function RoActivateInstance Lib "Combase.dll" (ByVal activatableClassId As LongPtr, ByRef instance As LongPtr) As Long
#Else
    Private Enum LongPtr
        [_]
    End Enum
    Private Declare Sub MemoryCopy Lib "msvcrt.dll" Alias "memcpy" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
    Private Declare Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
    Private Declare Function CLSIDFromString Lib "ole32" (ByVal OleStringCLSID As Long, ByRef cGUID As Any) As Long
    Private Declare Function RoGetActivationFactory Lib "Combase.dll" (ByVal activatableClassId As Long, ByRef riid As GUID, ByRef factory As Long) As Long
    Private Declare Function WindowsCreateString Lib "Combase.dll" (ByVal sourceString As Long, ByVal Length As Long, ByRef hString As Long) As Long
    Private Declare Function WindowsDeleteString Lib "Combase.dll" (ByVal sourceString As LongPtr) As Long
    Private Declare Function WindowsGetStringRawBuffer Lib "Combase.dll" (ByVal hString As Long, ByRef Length As Long) As Long
    Private Declare Function CreateRandomAccessStreamOnFile Lib "Shcore.dll" (ByVal pFilePath As Long, ByVal Access As Long, ByRef riid As GUID, ByRef ppv As Long) As Long
    Private Declare Function RoActivateInstance Lib "Combase.dll" (ByVal activatableClassId As Long, ByRef instance As Long) As Long
#End If

#If Win64 Then
    Private Const PTR_SIZE As Long = 8
#Else
    Private Const PTR_SIZE As Long = 4
#End If

Private Const CC_STDCALL = 4

Private pIPdfDocumentFactory As LongPtr
Private IsInitialized As Boolean

Public Function RenderPDFToImages(ByVal pdfPath As String, _
                                   ByVal outputFolder As String, _
                                   Optional ByVal widthPx As Long = 2480) As Boolean
    If Not IsInitialized Then Initialize
    If Not IsInitialized Then
        Err.Raise 1, "PDFWRT.RenderPDFToImages", "Failed to initialize Windows.Data.Pdf WinRT factory"
    End If

    Dim pStream As LongPtr
    pStream = CreateReadStreamOnFile(pdfPath)
    If pStream = 0 Then
        Err.Raise 2, "PDFWRT.RenderPDFToImages", "Failed to open PDF stream: " & pdfPath
    End If

    Dim pAsync As LongPtr  ' IPdfDocumentStatics.LoadFromStreamAsync (offset 8)
    Call vtblCall(pIPdfDocumentFactory, 8 * PTR_SIZE, vbLong, CC_STDCALL, pStream, VarPtr(pAsync))
    If pAsync = 0 Then
        Err.Raise 3, "PDFWRT.RenderPDFToImages", "LoadFromStreamAsync returned null"
    End If

    Dim pGUID As GUID
    Dim pIAsyncInfo As LongPtr
    Call CLSIDFromString(StrPtr("{00000036-0000-0000-C000-000000000046}"), pGUID)
    Call vtblCall(pAsync, 0 * PTR_SIZE, vbLong, CC_STDCALL, VarPtr(pGUID), VarPtr(pIAsyncInfo))
    If pIAsyncInfo = 0 Then
        Err.Raise 4, "PDFWRT.RenderPDFToImages", "IAsyncInfo QI failed on document load"
    End If

    Dim Status As Long
    Dim lRetry As Long
    Do
        DoEvents
        Call vtblCall(pIAsyncInfo, 7 * PTR_SIZE, vbLong, CC_STDCALL, VarPtr(Status))
        lRetry = lRetry + 1
        If lRetry > 500000 Then
            Err.Raise 9999, "PDFWRT.RenderPDFToImages", "Timeout waiting for document load"
        End If
    Loop Until Status <> 0

    Dim pDoc As LongPtr  ' GetResults (offset 8)
    Call vtblCall(pAsync, 8 * PTR_SIZE, vbLong, CC_STDCALL, VarPtr(pDoc))
    If pDoc = 0 Then
        Err.Raise 5, "PDFWRT.RenderPDFToImages", "PdfDocument null after load - file may be corrupt or password protected"
    End If

    Dim pageCount As Long  ' get_PageCount (offset 7)
    Call vtblCall(pDoc, 7 * PTR_SIZE, vbLong, CC_STDCALL, VarPtr(pageCount))

    If Dir(outputFolder, vbDirectory) = "" Then MkDir outputFolder

    Dim pOpts As LongPtr
    pOpts = CreateRenderOptions(widthPx)

    Dim i As Long
    For i = 0 To pageCount - 1

        Dim pPage As LongPtr  ' GetPage (offset 6)
        Call vtblCall(pDoc, 6 * PTR_SIZE, vbLong, CC_STDCALL, CLng(i), VarPtr(pPage))
        If pPage = 0 Then
            GoTo NextPage
        End If

        ' CreateRandomAccessStreamOnFile requires the target file to pre-exist.
        Dim outPath As String
        outPath = outputFolder & "\page_" & Format(i + 1, "000") & ".png"
        Dim fNum As Integer
        fNum = FreeFile
        Open outPath For Binary As #fNum
        Close #fNum

        Dim pOutStream As LongPtr
        pOutStream = CreateWriteStreamOnFile(outPath)
        If pOutStream = 0 Then
            GoTo NextPage
        End If

        Dim pAsyncRender As LongPtr
        If pOpts <> 0 Then
            ' IPdfPage.RenderToStreamWithOptionsAsync - offset 7
            Call vtblCall(pPage, 7 * PTR_SIZE, vbLong, CC_STDCALL, pOutStream, pOpts, VarPtr(pAsyncRender))
        Else
            ' IPdfPage.RenderToStreamAsync - offset 6 (fallback, lower quality)
            Call vtblCall(pPage, 6 * PTR_SIZE, vbLong, CC_STDCALL, pOutStream, VarPtr(pAsyncRender))
        End If

        If pAsyncRender <> 0 Then
            Dim pAsyncInfo2 As LongPtr
            Call vtblCall(pAsyncRender, 0 * PTR_SIZE, vbLong, CC_STDCALL, VarPtr(pGUID), VarPtr(pAsyncInfo2))
            If pAsyncInfo2 <> 0 Then
                Dim Status2 As Long
                Dim lRetry2 As Long
                Do
                    DoEvents
                    Call vtblCall(pAsyncInfo2, 7 * PTR_SIZE, vbLong, CC_STDCALL, VarPtr(Status2))
                    lRetry2 = lRetry2 + 1
                    If lRetry2 > 500000 Then
                        GoTo NextPage
                    End If
                Loop Until Status2 <> 0
            End If
        End If

NextPage:
    Next i

    RenderPDFToImages = True
End Function

Private Sub Initialize()
    pIPdfDocumentFactory = RoGetActivationIFactory( _
        "Windows.Data.Pdf.PdfDocument", _
        "{433A0B5F-C007-4788-90F2-08143D922599}")

    If pIPdfDocumentFactory <> 0 Then
        IsInitialized = True
    End If
End Sub

Private Function CreateRenderOptions(ByVal widthPx As Long) As LongPtr
    Dim hClass As LongPtr
    Dim pOpts As LongPtr
    Call WindowsCreateString(StrPtr("Windows.Data.Pdf.PdfPageRenderOptions"), _
                              Len("Windows.Data.Pdf.PdfPageRenderOptions"), hClass)
    Call RoActivateInstance(hClass, pOpts)
    WindowsDeleteString hClass

    If pOpts <> 0 Then
        Call vtblCall(pOpts, 9 * PTR_SIZE, vbLong, CC_STDCALL, CLng(widthPx))  ' set_DestinationWidth (offset 9)
    End If

    CreateRenderOptions = pOpts
End Function

Private Function CreateReadStreamOnFile(ByVal filePath As String) As LongPtr
    Dim pGUID As GUID
    Call CLSIDFromString(StrPtr("{905A0FE1-BC53-11DF-8C49-001E4FC686DA}"), pGUID)
    Call CreateRandomAccessStreamOnFile(StrPtr(filePath), 0, pGUID, CreateReadStreamOnFile)
End Function

Private Function CreateWriteStreamOnFile(ByVal filePath As String) As LongPtr
    Dim pGUID As GUID
    Call CLSIDFromString(StrPtr("{905A0FE1-BC53-11DF-8C49-001E4FC686DA}"), pGUID)
    Call CreateRandomAccessStreamOnFile(StrPtr(filePath), 1, pGUID, CreateWriteStreamOnFile)
End Function

Private Function RoGetActivationIFactory(ByVal activatableClassId As String, ByVal sIID As String) As LongPtr
    Dim hString As LongPtr
    Call WindowsCreateString(StrPtr(activatableClassId), Len(activatableClassId), hString)
    If hString <> 0 Then
        Dim pGUID As GUID
        Dim pFactory As LongPtr
        Call CLSIDFromString(StrPtr(sIID), pGUID)
        Call RoGetActivationFactory(hString, pGUID, pFactory)
        If pFactory Then
            RoGetActivationIFactory = pFactory
            Exit Function
        End If
        WindowsDeleteString hString
    End If
End Function

Private Function vtblCall( _
        ByVal InterfacePointer As LongPtr, _
        ByVal VTableOffset As Long, _
        ByVal FunctionReturnType As Long, _
        ByVal CallConvention As Long, _
        ParamArray FunctionParameters() As Variant) As Variant

    Dim vParamPtr() As LongPtr
    Dim hr As Long

    If InterfacePointer = 0& Or VTableOffset < 0& Then Exit Function
    If Not (FunctionReturnType And &HFFFF0000) = 0& Then Exit Function

    Dim pIndex As Long, pCount As Long
    Dim vParamType() As Integer
    Dim vRtn As Variant, vParams() As Variant

    vParams() = FunctionParameters()
    pCount = Abs(UBound(vParams) - LBound(vParams) + 1&)
    If pCount = 0& Then
        ReDim vParamPtr(0 To 0)
        ReDim vParamType(0 To 0)
    Else
        ReDim vParamPtr(0 To pCount - 1&)
        ReDim vParamType(0 To pCount - 1&)
        For pIndex = 0& To pCount - 1&
            vParamPtr(pIndex) = VarPtr(vParams(pIndex))
            vParamType(pIndex) = VarType(vParams(pIndex))
        Next
    End If

    hr = DispCallFunc(InterfacePointer, VTableOffset, CallConvention, FunctionReturnType, pCount, _
                      vParamType(0), vParamPtr(0), vRtn)
    If hr = 0& Then
        vtblCall = vRtn
    Else
        Err.Raise Number:=vbObjectError + hr, source:="PDFWRT.vtblCall", _
                  Description:="DispCallFunc failed with HRESULT: " & Hex$(hr)
    End If
End Function
