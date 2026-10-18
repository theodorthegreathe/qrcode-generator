Attribute VB_Name = "QRCode"
'===============================================================================
' QRCode - offline QR code generator for Microsoft Excel
'
' Pure VBA: no add-ins, no internet, no Windows API declarations, no references
' to set. Works in 32- and 64-bit Office.
'
' Made for Czech/Slovak "QR Platba" payment strings in the SPD format
' (SPD*1.0*ACC:CZ...*AM:...*CC:CZK*...), but it encodes any text.
'
' HOW TO USE
'   1. Select the cell that holds the SPD string - or the cell with your
'      online link formula, e.g.
'        =HYPERLINK("https://quickchart.io/qr?text=" & ENCODEURL(H88); "Generate QR Code")
'      For such a cell the macro evaluates the argument of ENCODEURL(...), so
'      the online link and the offline picture always encode the same text.
'   2. Run the macro InsertQRCode (keyboard shortcut Ctrl+Shift+Q).
'   The QR picture is placed right next to the selected cell. Running the macro
'   again replaces that picture. Several cells may be selected at once.
'
' OTHER ENTRY POINTS
'   DeleteQRCodes                      removes the QR pictures of the selected
'                                      cells (select all to remove every one)
'   AddQRCode(text, sheet, left, top)  inserts a QR picture from your own code
'   SaveQRCodeBmp(text, path)          writes the QR code to a .bmp file
'   QRMatrix(text)                     returns the module matrix (1 = dark)
'
' ENCODING
'   ISO/IEC 18004 (model 2), versions 1-40. Alphanumeric mode when the text
'   allows it (typical SPD strings do, giving the smallest code), otherwise
'   byte mode with UTF-8. Error correction level M by default, as recommended
'   by the QR Platba specification.
'
' Written with qrcode-generator by Kazuhiko Arase (MIT License, (c) 2009)
' as a reference.
'===============================================================================
Option Explicit

'--- Settings ------------------------------------------------------------------
Private Const QR_SIZE_PT As Double = 100    ' picture side in points (72 pt = 1 inch, 100 pt = 35 mm)
Private Const QR_ECC_LEVEL As String = "M"  ' error correction: L (7%), M (15%), Q (25%), H (30%)
Private Const QR_QUIET_ZONE As Long = 4     ' white margin in modules (the standard requires 4)
Private Const QR_PX_PER_MODULE As Long = 8  ' bitmap resolution; the picture is scaled to QR_SIZE_PT
Private Const QR_GAP_PT As Double = 3       ' gap between the cell and the picture
Private Const QR_NAME_PREFIX As String = "QR_"

'--- QR code tables ------------------------------------------------------------
Private Const ALNUM_CHARS As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ $%*+-./:"

' Error correction codewords per block, versions 1..40
Private Const ECC_PER_BLOCK_L As String = "7,10,15,20,26,18,20,24,30,18,20,24,26,30,22,24,28,30,28,28,28,28,30,30,26,28,30,30,30,30,30,30,30,30,30,30,30,30,30,30"
Private Const ECC_PER_BLOCK_M As String = "10,16,26,18,24,16,18,22,22,26,30,22,22,24,24,28,28,26,26,26,26,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28"
Private Const ECC_PER_BLOCK_Q As String = "13,22,18,26,18,24,18,22,20,24,28,26,24,20,30,24,28,28,26,30,28,30,30,30,30,28,30,30,30,30,30,30,30,30,30,30,30,30,30,30"
Private Const ECC_PER_BLOCK_H As String = "17,28,22,16,22,28,26,26,24,28,24,28,22,24,24,30,28,28,26,28,30,24,30,30,30,30,30,30,30,30,30,30,30,30,30,30,30,30,30,30"

' Number of error correction blocks, versions 1..40
Private Const NUM_BLOCKS_L As String = "1,1,1,1,1,2,2,2,2,4,4,4,4,4,6,6,6,6,7,8,8,9,9,10,12,12,12,13,14,15,16,17,18,19,19,20,21,22,24,25"
Private Const NUM_BLOCKS_M As String = "1,1,1,2,2,4,4,4,5,5,5,8,9,9,10,10,11,13,14,16,17,17,18,20,21,23,25,26,28,29,31,33,35,37,38,40,43,45,47,49"
Private Const NUM_BLOCKS_Q As String = "1,1,2,2,4,4,6,6,8,8,8,10,12,16,12,17,16,18,21,20,23,23,25,27,29,34,34,35,38,40,43,45,48,51,53,56,59,62,65,68"
Private Const NUM_BLOCKS_H As String = "1,1,2,4,4,4,5,6,8,8,11,11,16,16,18,16,19,21,25,25,25,34,30,32,35,37,40,42,45,48,51,54,57,60,63,66,70,74,77,81"

' Galois field GF(256) tables for Reed-Solomon
Private gfExp(0 To 511) As Long
Private gfLog(0 To 255) As Long
Private gfReady As Boolean

'===============================================================================
' Macros
'===============================================================================

' Inserts a QR code next to every selected cell (Ctrl+Shift+Q).
Public Sub InsertQRCode()
Attribute InsertQRCode.VB_Description = "Insert an offline QR code next to each selected cell (SPD string, or a HYPERLINK/ENCODEURL formula)."
Attribute InsertQRCode.VB_ProcData.VB_Invoke_Func = "Q\n14"
    Dim target As Range, cell As Range, text As String
    Dim done As Long, failed As String

    If TypeName(Selection) <> "Range" Then
        MsgBox "Select the cell with the payment (SPD) string first.", vbExclamation, "QR code"
        Exit Sub
    End If
    Set target = Intersect(Selection, Selection.Worksheet.UsedRange)

    Application.ScreenUpdating = False
    If Not target Is Nothing Then
        For Each cell In target.Cells
            ' For merged cells only the top-left cell holds the value
            If cell.Address = cell.MergeArea.Cells(1, 1).Address Then
                text = PayloadFromCell(cell)
                If Len(text) > 0 Then
                    On Error Resume Next
                    AddQRCodeNextTo cell.MergeArea, text
                    If Err.Number <> 0 Then
                        failed = failed & vbLf & cell.Address(False, False) & ": " & Err.Description
                    Else
                        done = done + 1
                    End If
                    On Error GoTo 0
                End If
            End If
        Next
    End If
    Application.ScreenUpdating = True

    If Len(failed) > 0 Then
        MsgBox "Could not create the QR code for:" & failed, vbExclamation, "QR code"
    ElseIf done = 0 Then
        MsgBox "The selected cell is empty - nothing to encode.", vbExclamation, "QR code"
    End If
End Sub

' Removes the QR pictures that belong to the selected cells.
Public Sub DeleteQRCodes()
Attribute DeleteQRCodes.VB_Description = "Remove the QR codes created for the selected cells."
    Dim ws As Worksheet, i As Long, nm As String, src As Range

    If TypeName(Selection) <> "Range" Then Exit Sub
    Set ws = Selection.Worksheet
    For i = ws.Shapes.Count To 1 Step -1
        nm = ws.Shapes(i).Name
        If Left$(nm, Len(QR_NAME_PREFIX)) = QR_NAME_PREFIX Then
            Set src = Nothing
            On Error Resume Next
            Set src = ws.Range(Mid$(nm, Len(QR_NAME_PREFIX) + 1))
            On Error GoTo 0
            If Not src Is Nothing Then
                If Not Intersect(src, Selection) Is Nothing Then ws.Shapes(i).Delete
            End If
        End If
    Next
End Sub

'===============================================================================
' Public helpers
'===============================================================================

' Inserts a QR code picture into a worksheet and returns the shape.
Public Function AddQRCode(ByVal text As String, ByVal ws As Worksheet, _
                          ByVal leftPt As Double, ByVal topPt As Double, _
                          Optional ByVal sizePt As Double = QR_SIZE_PT, _
                          Optional ByVal shapeName As String = "", _
                          Optional ByVal eccLevel As String = QR_ECC_LEVEL) As Shape
    Dim path As String, shp As Shape, errNum As Long, errDesc As String

    path = TempBmpPath()
    SaveQRCodeBmp text, path, eccLevel

    On Error Resume Next
    Set shp = ws.Shapes.AddPicture(path, msoFalse, msoTrue, leftPt, topPt, sizePt, sizePt)
    errNum = Err.Number: errDesc = Err.Description
    On Error GoTo 0
    Kill path
    If errNum <> 0 Then Err.Raise errNum, "AddQRCode", errDesc

    shp.LockAspectRatio = msoTrue
    shp.Placement = xlMove
    shp.AlternativeText = text
    If Len(shapeName) > 0 Then shp.Name = shapeName
    Set AddQRCode = shp
End Function

' Writes the QR code of a text to a monochrome .bmp file.
Public Sub SaveQRCodeBmp(ByVal text As String, ByVal path As String, _
                         Optional ByVal eccLevel As String = QR_ECC_LEVEL, _
                         Optional ByVal pxPerModule As Long = QR_PX_PER_MODULE)
    Dim m() As Byte
    m = QRMatrix(text, eccLevel)
    WriteBmp m, QR_QUIET_ZONE, pxPerModule, path
End Sub

' Encodes a text as a QR code. Returns a square matrix m(x, y), 1 = dark module,
' without the quiet zone.
Public Function QRMatrix(ByVal text As String, Optional ByVal eccLevel As String = QR_ECC_LEVEL) As Byte()
    Dim ecc As Long, alnum As Boolean, utf8() As Byte, nChars As Long
    Dim ver As Long, capBits As Long, ccBits As Long, dataBits As Long
    Dim bits() As Byte, nb As Long, i As Long, n As Long, padByte As Long
    Dim dataCw() As Byte, allCw() As Byte, size As Long
    Dim m() As Byte, fn() As Boolean
    Dim mask As Long, bestMask As Long, pen As Long, bestPen As Long

    Select Case UCase$(Trim$(eccLevel))
        Case "L": ecc = 0
        Case "M": ecc = 1
        Case "Q": ecc = 2
        Case "H": ecc = 3
        Case Else: Err.Raise 5, "QRMatrix", "Error correction level must be L, M, Q or H."
    End Select

    ' Choose the mode and the smallest version that fits
    alnum = IsAlphanumeric(text)
    If alnum Then
        nChars = Len(text)
        dataBits = 11 * (nChars \ 2) + 6 * (nChars Mod 2)
    Else
        nChars = Utf8Encode(text, utf8)
        dataBits = 8 * nChars
    End If
    For ver = 1 To 40
        ccBits = CharCountBits(ver, alnum)
        capBits = DataCodewords(ver, ecc) * 8
        If 4 + ccBits + dataBits <= capBits Then Exit For
    Next
    If ver > 40 Then Err.Raise vbObjectError + 1, "QRMatrix", "The text is too long for a QR code."

    ' Bit stream: mode, character count, data, terminator, padding
    ReDim bits(0 To capBits - 1)
    If alnum Then AppendBits bits, nb, 2, 4 Else AppendBits bits, nb, 4, 4
    AppendBits bits, nb, nChars, ccBits
    If alnum Then
        For i = 1 To nChars - 1 Step 2
            AppendBits bits, nb, AlnumIndex(Mid$(text, i, 1)) * 45 + AlnumIndex(Mid$(text, i + 1, 1)), 11
        Next
        If nChars Mod 2 = 1 Then AppendBits bits, nb, AlnumIndex(Right$(text, 1)), 6
    Else
        For i = 0 To nChars - 1
            AppendBits bits, nb, utf8(i), 8
        Next
    End If
    n = capBits - nb
    If n > 4 Then n = 4
    AppendBits bits, nb, 0, n
    AppendBits bits, nb, 0, (8 - nb Mod 8) Mod 8
    padByte = &HEC
    Do While nb < capBits
        AppendBits bits, nb, padByte, 8
        padByte = padByte Xor &HFD   ' alternates &HEC and &H11
    Loop

    ReDim dataCw(0 To capBits \ 8 - 1)
    For i = 0 To capBits - 1
        If bits(i) = 1 Then dataCw(i \ 8) = dataCw(i \ 8) Or Pow2(7 - (i Mod 8))
    Next
    allCw = AddEccAndInterleave(dataCw, ver, ecc)

    ' Draw the symbol and pick the mask with the lowest penalty
    size = ver * 4 + 17
    ReDim m(0 To size - 1, 0 To size - 1)
    ReDim fn(0 To size - 1, 0 To size - 1)
    DrawFunctionPatterns m, fn, ver, ecc
    DrawCodewords m, fn, allCw

    bestPen = -1
    For mask = 0 To 7
        ApplyMask m, fn, mask
        DrawFormatBits m, fn, ecc, mask
        pen = Penalty(m)
        If bestPen < 0 Or pen < bestPen Then bestPen = pen: bestMask = mask
        ApplyMask m, fn, mask   ' XOR again to undo
    Next
    ApplyMask m, fn, bestMask
    DrawFormatBits m, fn, ecc, bestMask

    QRMatrix = m
End Function

'===============================================================================
' Excel glue
'===============================================================================

Private Sub AddQRCodeNextTo(ByVal area As Range, ByVal text As String)
    Dim nm As String
    nm = QR_NAME_PREFIX & area.Cells(1, 1).Address(False, False)
    On Error Resume Next
    area.Worksheet.Shapes(nm).Delete
    On Error GoTo 0
    AddQRCode text, area.Worksheet, area.Left + area.Width + QR_GAP_PT, area.Top, QR_SIZE_PT, nm
End Sub

' The text to encode: for a formula containing ENCODEURL(...) the evaluated
' argument, otherwise the cell value.
Private Function PayloadFromCell(ByVal c As Range) As String
    Dim f As String, p As Long, arg As String, v As Variant, s As String

    If c.HasFormula Then
        f = c.Formula
        p = InStr(1, f, "ENCODEURL(", vbTextCompare)
        If p > 0 Then
            arg = FirstArgument(f, p + Len("ENCODEURL("))
            If Len(arg) > 0 Then
                On Error Resume Next
                v = c.Worksheet.Evaluate(arg)
                If Not IsError(v) Then s = CStr(v)
                On Error GoTo 0
                If Len(s) > 0 Then
                    PayloadFromCell = s
                    Exit Function
                End If
            End If
        End If
    End If

    v = c.Value
    If Not IsError(v) Then PayloadFromCell = CStr(v)
End Function

' Text from startPos up to the matching closing parenthesis.
Private Function FirstArgument(ByVal f As String, ByVal startPos As Long) As String
    Dim i As Long, depth As Long, inQuotes As Boolean, ch As String
    For i = startPos To Len(f)
        ch = Mid$(f, i, 1)
        If ch = """" Then
            inQuotes = Not inQuotes
        ElseIf Not inQuotes Then
            If ch = "(" Then
                depth = depth + 1
            ElseIf ch = ")" Then
                If depth = 0 Then
                    FirstArgument = Mid$(f, startPos, i - startPos)
                    Exit Function
                End If
                depth = depth - 1
            End If
        End If
    Next
End Function

Private Function TempBmpPath() As String
    Randomize
    TempBmpPath = Environ$("TEMP") & "\xlqr_" & Format$(Now, "yyyymmddhhnnss") & "_" & _
                  Format$(Int(Rnd * 1000000), "000000") & ".bmp"
End Function

' 1-bit BMP, white background, black modules, with quiet zone.
Private Sub WriteBmp(m() As Byte, ByVal quiet As Long, ByVal scale_ As Long, ByVal path As String)
    Dim size As Long, w As Long, stride As Long, buf() As Byte, rowBuf() As Byte
    Dim mx As Long, my As Long, px As Long, r As Long, k As Long, offset As Long
    Dim bitVal(0 To 7) As Byte, f As Integer

    For k = 0 To 7
        bitVal(k) = Pow2(7 - k)
    Next
    size = UBound(m, 1) + 1
    w = (size + 2 * quiet) * scale_
    stride = ((w + 31) \ 32) * 4

    ReDim buf(0 To 62 + stride * w - 1)
    buf(0) = 66: buf(1) = 77                  ' "BM"
    PutLong buf, 2, UBound(buf) + 1           ' file size
    PutLong buf, 10, 62                       ' pixel data offset
    PutLong buf, 14, 40                       ' BITMAPINFOHEADER size
    PutLong buf, 18, w                        ' width
    PutLong buf, 22, w                        ' height (rows stored bottom-up)
    buf(26) = 1                               ' planes
    buf(28) = 1                               ' bits per pixel
    PutLong buf, 34, stride * w               ' image size
    PutLong buf, 38, 11811                    ' 300 dpi
    PutLong buf, 42, 11811
    PutLong buf, 46, 2                        ' colours used
    PutLong buf, 50, 2
    buf(54) = 255: buf(55) = 255: buf(56) = 255   ' palette 0 = white, 1 = black

    For my = 0 To size - 1
        ReDim rowBuf(0 To stride - 1)
        For mx = 0 To size - 1
            If m(mx, my) = 1 Then
                For px = (mx + quiet) * scale_ To (mx + quiet + 1) * scale_ - 1
                    rowBuf(px \ 8) = rowBuf(px \ 8) Or bitVal(px Mod 8)
                Next
            End If
        Next
        For r = 0 To scale_ - 1
            offset = 62 + (w - 1 - ((my + quiet) * scale_ + r)) * stride
            For k = 0 To stride - 1
                buf(offset + k) = rowBuf(k)
            Next
        Next
    Next

    If Len(Dir$(path)) > 0 Then Kill path
    f = FreeFile
    Open path For Binary Access Write As #f
    Put #f, 1, buf
    Close #f
End Sub

Private Sub PutLong(buf() As Byte, ByVal offset As Long, ByVal v As Long)
    buf(offset) = v And 255
    buf(offset + 1) = (v \ 256) And 255
    buf(offset + 2) = (v \ 65536) And 255
    buf(offset + 3) = (v \ 16777216) And 255
End Sub

'===============================================================================
' Data encoding
'===============================================================================

Private Function IsAlphanumeric(ByVal s As String) As Boolean
    Dim i As Long
    For i = 1 To Len(s)
        If InStr(1, ALNUM_CHARS, Mid$(s, i, 1), vbBinaryCompare) = 0 Then Exit Function
    Next
    IsAlphanumeric = True
End Function

Private Function AlnumIndex(ByVal ch As String) As Long
    AlnumIndex = InStr(1, ALNUM_CHARS, ch, vbBinaryCompare) - 1
End Function

' UTF-8 bytes of s in out(); returns the byte count.
Private Function Utf8Encode(ByVal s As String, out() As Byte) As Long
    Dim i As Long, n As Long, c As Long, c2 As Long

    ReDim out(0 To 3 * Len(s))
    i = 1
    Do While i <= Len(s)
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        If c >= &HD800& And c <= &HDBFF& And i < Len(s) Then   ' surrogate pair
            c2 = AscW(Mid$(s, i + 1, 1)) And &HFFFF&
            If c2 >= &HDC00& And c2 <= &HDFFF& Then
                c = &H10000 + (c - &HD800&) * 1024 + (c2 - &HDC00&)
                i = i + 1
            End If
        End If
        If c < &H80& Then
            out(n) = c: n = n + 1
        ElseIf c < &H800& Then
            out(n) = &HC0 Or (c \ 64)
            out(n + 1) = &H80 Or (c And &H3F)
            n = n + 2
        ElseIf c < &H10000 Then
            out(n) = &HE0 Or (c \ 4096)
            out(n + 1) = &H80 Or ((c \ 64) And &H3F)
            out(n + 2) = &H80 Or (c And &H3F)
            n = n + 3
        Else
            out(n) = &HF0 Or (c \ 262144)
            out(n + 1) = &H80 Or ((c \ 4096) And &H3F)
            out(n + 2) = &H80 Or ((c \ 64) And &H3F)
            out(n + 3) = &H80 Or (c And &H3F)
            n = n + 4
        End If
        i = i + 1
    Loop
    Utf8Encode = n
End Function

Private Sub AppendBits(bits() As Byte, nb As Long, ByVal value As Long, ByVal length As Long)
    Dim i As Long
    For i = length - 1 To 0 Step -1
        bits(nb) = (value \ Pow2(i)) And 1
        nb = nb + 1
    Next
End Sub

Private Function CharCountBits(ByVal ver As Long, ByVal alnum As Boolean) As Long
    If alnum Then
        If ver <= 9 Then
            CharCountBits = 9
        ElseIf ver <= 26 Then
            CharCountBits = 11
        Else
            CharCountBits = 13
        End If
    Else
        If ver <= 9 Then CharCountBits = 8 Else CharCountBits = 16
    End If
End Function

'===============================================================================
' Capacity and error correction
'===============================================================================

Private Function EccPerBlock(ByVal ver As Long, ByVal ecc As Long) As Long
    EccPerBlock = CLng(Split(Choose(ecc + 1, ECC_PER_BLOCK_L, ECC_PER_BLOCK_M, ECC_PER_BLOCK_Q, ECC_PER_BLOCK_H), ",")(ver - 1))
End Function

Private Function NumBlocks(ByVal ver As Long, ByVal ecc As Long) As Long
    NumBlocks = CLng(Split(Choose(ecc + 1, NUM_BLOCKS_L, NUM_BLOCKS_M, NUM_BLOCKS_Q, NUM_BLOCKS_H), ",")(ver - 1))
End Function

' Modules available for data + ECC (after all function patterns).
Private Function RawDataModules(ByVal ver As Long) As Long
    Dim r As Long, na As Long
    r = (16 * ver + 128) * ver + 64
    If ver >= 2 Then
        na = ver \ 7 + 2
        r = r - (25 * na - 10) * na + 55
        If ver >= 7 Then r = r - 36
    End If
    RawDataModules = r
End Function

Private Function DataCodewords(ByVal ver As Long, ByVal ecc As Long) As Long
    DataCodewords = RawDataModules(ver) \ 8 - EccPerBlock(ver, ecc) * NumBlocks(ver, ecc)
End Function

Private Function AddEccAndInterleave(dataCw() As Byte, ByVal ver As Long, ByVal ecc As Long) As Byte()
    Dim nBlk As Long, eccLen As Long, rawCw As Long, numShort As Long, shortLen As Long
    Dim b As Long, i As Long, k As Long, n As Long
    Dim dStart() As Long, dLen() As Long, eccBytes() As Long, divisor() As Long, remd() As Long
    Dim out() As Byte

    nBlk = NumBlocks(ver, ecc)
    eccLen = EccPerBlock(ver, ecc)
    rawCw = RawDataModules(ver) \ 8
    numShort = nBlk - rawCw Mod nBlk
    shortLen = rawCw \ nBlk
    divisor = RSDivisor(eccLen)

    ReDim dStart(0 To nBlk - 1)
    ReDim dLen(0 To nBlk - 1)
    ReDim eccBytes(0 To nBlk - 1, 0 To eccLen - 1)
    For b = 0 To nBlk - 1
        dStart(b) = k
        dLen(b) = shortLen - eccLen
        If b >= numShort Then dLen(b) = dLen(b) + 1
        remd = RSRemainder(dataCw, dStart(b), dLen(b), divisor)
        For i = 0 To eccLen - 1
            eccBytes(b, i) = remd(i)
        Next
        k = k + dLen(b)
    Next

    ReDim out(0 To rawCw - 1)
    For i = 0 To shortLen - eccLen
        For b = 0 To nBlk - 1
            If i < dLen(b) Then out(n) = dataCw(dStart(b) + i): n = n + 1
        Next
    Next
    For i = 0 To eccLen - 1
        For b = 0 To nBlk - 1
            out(n) = eccBytes(b, i): n = n + 1
        Next
    Next
    AddEccAndInterleave = out
End Function

Private Sub InitGF()
    Dim i As Long, x As Long
    x = 1
    For i = 0 To 254
        gfExp(i) = x
        gfLog(x) = i
        x = x * 2
        If x >= 256 Then x = x Xor &H11D
    Next
    For i = 255 To 511
        gfExp(i) = gfExp(i - 255)
    Next
    gfReady = True
End Sub

Private Function GFMul(ByVal a As Long, ByVal b As Long) As Long
    If a = 0 Or b = 0 Then Exit Function
    If Not gfReady Then InitGF
    GFMul = gfExp(gfLog(a) + gfLog(b))
End Function

' Generator polynomial coefficients (highest degree first, leading 1 omitted).
Private Function RSDivisor(ByVal degree As Long) As Long()
    Dim res() As Long, i As Long, j As Long, root As Long
    ReDim res(0 To degree - 1)
    res(degree - 1) = 1
    root = 1
    For i = 0 To degree - 1
        For j = 0 To degree - 1
            res(j) = GFMul(res(j), root)
            If j + 1 < degree Then res(j) = res(j) Xor res(j + 1)
        Next
        root = GFMul(root, 2)
    Next
    RSDivisor = res
End Function

Private Function RSRemainder(data() As Byte, ByVal start As Long, ByVal length As Long, divisor() As Long) As Long()
    Dim res() As Long, degree As Long, k As Long, j As Long, factor As Long
    degree = UBound(divisor) + 1
    ReDim res(0 To degree - 1)
    For k = start To start + length - 1
        factor = data(k) Xor res(0)
        For j = 0 To degree - 2
            res(j) = res(j + 1)
        Next
        res(degree - 1) = 0
        For j = 0 To degree - 1
            res(j) = res(j) Xor GFMul(divisor(j), factor)
        Next
    Next
    RSRemainder = res
End Function

'===============================================================================
' Symbol drawing (m(x, y): x = column, y = row)
'===============================================================================

Private Sub SetFn(m() As Byte, fn() As Boolean, ByVal x As Long, ByVal y As Long, ByVal dark As Boolean)
    If dark Then m(x, y) = 1 Else m(x, y) = 0
    fn(x, y) = True
End Sub

Private Sub DrawFunctionPatterns(m() As Byte, fn() As Boolean, ByVal ver As Long, ByVal ecc As Long)
    Dim size As Long, i As Long, j As Long, na As Long, pos() As Long

    size = UBound(m, 1) + 1
    For i = 0 To size - 1                       ' timing patterns
        SetFn m, fn, 6, i, (i Mod 2 = 0)
        SetFn m, fn, i, 6, (i Mod 2 = 0)
    Next
    DrawFinder m, fn, 3, 3
    DrawFinder m, fn, size - 4, 3
    DrawFinder m, fn, 3, size - 4

    If ver >= 2 Then
        pos = AlignmentPositions(ver)
        na = UBound(pos) + 1
        For i = 0 To na - 1
            For j = 0 To na - 1
                ' skip the three corners occupied by finder patterns
                If Not ((i = 0 And j = 0) Or (i = 0 And j = na - 1) Or (i = na - 1 And j = 0)) Then
                    DrawAlignment m, fn, pos(i), pos(j)
                End If
            Next
        Next
    End If

    DrawFormatBits m, fn, ecc, 0   ' reserves the area; real bits are drawn after masking
    DrawVersion m, fn, ver
End Sub

Private Sub DrawFinder(m() As Byte, fn() As Boolean, ByVal cx As Long, ByVal cy As Long)
    Dim size As Long, dx As Long, dy As Long, x As Long, y As Long, d As Long
    size = UBound(m, 1) + 1
    For dy = -4 To 4
        For dx = -4 To 4
            x = cx + dx: y = cy + dy
            If x >= 0 And x < size And y >= 0 And y < size Then
                d = Abs(dx)
                If Abs(dy) > d Then d = Abs(dy)
                SetFn m, fn, x, y, (d <> 2 And d <> 4)
            End If
        Next
    Next
End Sub

Private Sub DrawAlignment(m() As Byte, fn() As Boolean, ByVal cx As Long, ByVal cy As Long)
    Dim dx As Long, dy As Long, d As Long
    For dy = -2 To 2
        For dx = -2 To 2
            d = Abs(dx)
            If Abs(dy) > d Then d = Abs(dy)
            SetFn m, fn, cx + dx, cy + dy, (d <> 1)
        Next
    Next
End Sub

Private Function AlignmentPositions(ByVal ver As Long) As Long()
    Dim na As Long, stp As Long, p As Long, i As Long, pos() As Long
    na = ver \ 7 + 2
    stp = ((ver * 8 + na * 3 + 5) \ (na * 4 - 4)) * 2
    ReDim pos(0 To na - 1)
    pos(0) = 6
    p = ver * 4 + 17 - 7
    For i = na - 1 To 1 Step -1
        pos(i) = p
        p = p - stp
    Next
    AlignmentPositions = pos
End Function

Private Sub DrawFormatBits(m() As Byte, fn() As Boolean, ByVal ecc As Long, ByVal mask As Long)
    Dim size As Long, data As Long, r As Long, bits As Long, i As Long

    size = UBound(m, 1) + 1
    data = Choose(ecc + 1, 1, 0, 3, 2) * 8 + mask    ' format codes: L=01 M=00 Q=11 H=10
    r = data
    For i = 1 To 10
        r = (r * 2) Xor ((r \ 512) * &H537)
    Next
    bits = ((data * 1024) Or r) Xor &H5412

    For i = 0 To 5                                    ' copy around the top-left finder
        SetFn m, fn, 8, i, Bit(bits, i)
    Next
    SetFn m, fn, 8, 7, Bit(bits, 6)
    SetFn m, fn, 8, 8, Bit(bits, 7)
    SetFn m, fn, 7, 8, Bit(bits, 8)
    For i = 9 To 14
        SetFn m, fn, 14 - i, 8, Bit(bits, i)
    Next

    For i = 0 To 7                                    ' copy split between the other finders
        SetFn m, fn, size - 1 - i, 8, Bit(bits, i)
    Next
    For i = 8 To 14
        SetFn m, fn, 8, size - 15 + i, Bit(bits, i)
    Next
    SetFn m, fn, 8, size - 8, True                    ' always-dark module
End Sub

Private Sub DrawVersion(m() As Byte, fn() As Boolean, ByVal ver As Long)
    Dim size As Long, r As Long, bits As Long, i As Long, a As Long, b As Long

    If ver < 7 Then Exit Sub
    size = UBound(m, 1) + 1
    r = ver
    For i = 1 To 12
        r = (r * 2) Xor ((r \ 2048) * &H1F25)
    Next
    bits = (ver * 4096) Or r
    For i = 0 To 17
        a = size - 11 + (i Mod 3)
        b = i \ 3
        SetFn m, fn, a, b, Bit(bits, i)
        SetFn m, fn, b, a, Bit(bits, i)
    Next
End Sub

' Places the codewords in the zigzag order, skipping function modules.
Private Sub DrawCodewords(m() As Byte, fn() As Boolean, cw() As Byte)
    Dim size As Long, total As Long, i As Long, col As Long, vert As Long, j As Long
    Dim x As Long, y As Long, upward As Boolean

    size = UBound(m, 1) + 1
    total = (UBound(cw) + 1) * 8
    col = size - 1
    Do While col >= 1
        If col = 6 Then col = 5                       ' skip the vertical timing pattern
        upward = (((col + 1) And 2) = 0)
        For vert = 0 To size - 1
            If upward Then y = size - 1 - vert Else y = vert
            For j = 0 To 1
                x = col - j
                If Not fn(x, y) Then
                    If i < total Then
                        If (cw(i \ 8) And Pow2(7 - (i Mod 8))) <> 0 Then m(x, y) = 1
                        i = i + 1
                    End If
                End If
            Next
        Next
        col = col - 2
    Loop
End Sub

Private Sub ApplyMask(m() As Byte, fn() As Boolean, ByVal mask As Long)
    Dim size As Long, x As Long, y As Long, inv As Boolean
    size = UBound(m, 1) + 1
    For y = 0 To size - 1
        For x = 0 To size - 1
            If Not fn(x, y) Then
                Select Case mask
                    Case 0: inv = ((x + y) Mod 2 = 0)
                    Case 1: inv = (y Mod 2 = 0)
                    Case 2: inv = (x Mod 3 = 0)
                    Case 3: inv = ((x + y) Mod 3 = 0)
                    Case 4: inv = ((x \ 3 + y \ 2) Mod 2 = 0)
                    Case 5: inv = ((x * y) Mod 2 + (x * y) Mod 3 = 0)
                    Case 6: inv = (((x * y) Mod 2 + (x * y) Mod 3) Mod 2 = 0)
                    Case 7: inv = (((x + y) Mod 2 + (x * y) Mod 3) Mod 2 = 0)
                End Select
                If inv Then m(x, y) = 1 - m(x, y)
            End If
        Next
    Next
End Sub

' Mask penalty score (ISO/IEC 18004 section 7.8.3); lower is better.
Private Function Penalty(m() As Byte) As Long
    Dim size As Long, x As Long, y As Long, run As Long, p As Long
    Dim dark As Long, total As Long, k As Long

    size = UBound(m, 1) + 1

    ' Rule 1: five or more same-coloured modules in a row or column
    For y = 0 To size - 1
        run = 1
        For x = 1 To size - 1
            If m(x, y) = m(x - 1, y) Then
                run = run + 1
                If run = 5 Then p = p + 3
                If run > 5 Then p = p + 1
            Else
                run = 1
            End If
        Next
    Next
    For x = 0 To size - 1
        run = 1
        For y = 1 To size - 1
            If m(x, y) = m(x, y - 1) Then
                run = run + 1
                If run = 5 Then p = p + 3
                If run > 5 Then p = p + 1
            Else
                run = 1
            End If
        Next
    Next

    ' Rule 2: 2x2 blocks of the same colour
    For y = 0 To size - 2
        For x = 0 To size - 2
            If m(x, y) = m(x + 1, y) Then
                If m(x, y) = m(x, y + 1) Then
                    If m(x, y) = m(x + 1, y + 1) Then p = p + 3
                End If
            End If
        Next
    Next

    ' Rule 3: finder-like 1:1:3:1:1 patterns with 4 light modules on one side
    For y = 0 To size - 1
        For x = 0 To size - 7
            If FinderLike(m, size, x, y, 1, 0) Then p = p + 40
            If FinderLike(m, size, y, x, 0, 1) Then p = p + 40
        Next
    Next

    ' Rule 4: balance of dark and light modules
    For y = 0 To size - 1
        For x = 0 To size - 1
            dark = dark + m(x, y)
        Next
    Next
    total = size * size
    k = (Abs(dark * 20 - total * 10) + total - 1) \ total - 1
    Penalty = p + k * 10
End Function

Private Function FinderLike(m() As Byte, ByVal size As Long, ByVal x As Long, ByVal y As Long, _
                            ByVal dx As Long, ByVal dy As Long) As Boolean
    Dim k As Long, want As Byte
    For k = 0 To 6                                    ' dark-light-dark-dark-dark-light-dark
        If k = 1 Or k = 5 Then want = 0 Else want = 1
        If m(x + k * dx, y + k * dy) <> want Then Exit Function
    Next
    FinderLike = LightRun(m, size, x - 4 * dx, y - 4 * dy, dx, dy)
    If Not FinderLike Then FinderLike = LightRun(m, size, x + 7 * dx, y + 7 * dy, dx, dy)
End Function

' True if the 4 modules from (x, y) in direction (dx, dy) are light or outside.
Private Function LightRun(m() As Byte, ByVal size As Long, ByVal x As Long, ByVal y As Long, _
                          ByVal dx As Long, ByVal dy As Long) As Boolean
    Dim k As Long, xx As Long, yy As Long
    For k = 0 To 3
        xx = x + k * dx: yy = y + k * dy
        If xx >= 0 And xx < size And yy >= 0 And yy < size Then
            If m(xx, yy) = 1 Then Exit Function
        End If
    Next
    LightRun = True
End Function

'===============================================================================
' Small utilities
'===============================================================================

Private Function Pow2(ByVal i As Long) As Long
    Pow2 = CLng(2 ^ i)
End Function

Private Function Bit(ByVal v As Long, ByVal i As Long) As Boolean
    Bit = (((v \ Pow2(i)) And 1) = 1)
End Function
