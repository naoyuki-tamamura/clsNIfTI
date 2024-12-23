Imports System.IO
Public Class clsNIfTIImage
    Public Property BigEndianFlag As Boolean    'バイトオーダー（TrueでBigEndian）
    Private Property Dimension0 As Short
    Public Property MatrixX As Short             '横方向のマトリクスサイズ
    Public Property MatrixY As Short            '縦方向のマトリクスサイズ
    Public Property SliceCount As Short         'スライス枚数
    Private Property Dimension4 As Short
    Private Property Dimension5 As Short
    Private Property Dimension6 As Short
    Private Property Dimension7 As Short
    Public Property DataType As Short           'データタイプ（1:Binary 2:Unsigned char 4:signed short　8:signed long 16:float 64:double 512:unsigned short)
    Public Property BitsPerPixel As Short       'BitsPerPixel
    Private Property SizeDim0 As Single
    Public Property SizeX As Single             '横方向のピクセルサイズ
    Public Property SizeY As Single             '縦方向のピクセルサイズ
    Public Property SizeZ As Single             'スライス厚
    Private Property SizeDim4 As Single
    Private Property SizeDim5 As Single
    Private Property SizeDim6 As Single
    Private Property SizeDim7 As Single
    Public Property VoxelOffset As Single       'Imageのオフセット(通常は352）
    Public Property RescaleSlope As Single      'リスケールスロープ
    Public Property RescaleIntercept As Single  'リスケール切片
    Public Property qFormCode As Short          ' qform座標系コード
    Public Property sFormCode As Short          ' sform座標系コード
    Public Property qParamB As Single           ' 四元数パラメータB
    Public Property qParamC As Single           ' 四元数パラメータC
    Public Property qParamD As Single           ' 四元数パラメータD
    Public Property qOffsetX As Single          ' 四元数座標オフセットX
    Public Property qOffsetY As Single          ' 四元数座標オフセットY
    Public Property qOffsetZ As Single          ' 四元数座標オフセットZ
    Public Property SrowX As Single()           ' sform行列の第1行
    Public Property SrowY As Single()           ' sform行列の第2行
    Public Property SrowZ As Single()           ' sform行列の第3行
    Public Property Pixel As Double(,,)     ' 3次元の画素値配列（z,y,x)

    Private Const OFFSET_DIMENSION As Integer = 40
    Private Const OFFSET_MATRIX_X As Integer = 42
    Private Const OFFSET_MATRIX_Y As Integer = 44
    Private Const OFFSET_SLICE_COUNT As Integer = 46
    Private Const OFFSET_DATATYPE As Integer = 70
    Private Const OFFSET_BITPERPIXEL As Integer = 72
    Private Const OFFSET_PIXDIM As Integer = 76
    Private Const OFFSET_SIZE_X As Integer = 80
    Private Const OFFSET_SIZE_Y As Integer = 84
    Private Const OFFSET_SIZE_Z As Integer = 88
    Private Const OFFSET_VOXEL_OFFSET As Integer = 108
    Private Const OFFSET_RESCALE_SLOPE As Integer = 112
    Private Const OFFSET_RESCALE_INTERCEPT As Integer = 116
    Private Const OFFSET_QFORM_CODE As Integer = 252
    Private Const OFFSET_SFORM_CODE As Integer = 254
    Private Const OFFSET_QPARAM_B As Integer = 256
    Private Const OFFSET_QPARAM_C As Integer = 260
    Private Const OFFSET_QPARAM_D As Integer = 264
    Private Const OFFSET_QOFFSET_X As Integer = 268
    Private Const OFFSET_QOFFSET_Y As Integer = 272
    Private Const OFFSET_QOFFSET_Z As Integer = 276
    Private Const OFFSET_SROW_X As Integer = 280
    Private Const OFFSET_SROW_Y As Integer = 296
    Private Const OFFSET_SROW_Z As Integer = 312
    Private Const OFFSET_MAGIC_WORD As Integer = 344
    Private Const MAGIC_WORD As String = "n+1" & Chr(0)

    Sub New()
        BigEndianFlag = False
        Dimension0 = 3
        MatrixX = 0
        MatrixY = 0
        SliceCount = 0
        Dimension4 = 0
        Dimension5 = 0
        Dimension6 = 0
        Dimension7 = 0
        DataType = 0
        BitsPerPixel = 0
        SizeDim0 = 0
        SizeX = 0
        SizeY = 0
        SizeZ = 0
        SizeDim4 = 0
        SizeDim5 = 0
        SizeDim6 = 0
        SizeDim7 = 0
        VoxelOffset = 352
        RescaleSlope = 1
        RescaleIntercept = 0
        qFormCode = 0
        sFormCode = 0
        qParamB = 0
        qParamC = 0
        qParamD = 0
        qOffsetX = 0
        qOffsetY = 0
        qOffsetZ = 0
        SrowX = New Single(3) {}
        SrowY = New Single(3) {}
        SrowZ = New Single(3) {}
    End Sub

    Sub Read(ByVal FilePath As String)

        If Not File.Exists(FilePath) Then
            Throw New FileNotFoundException($"File not found: {FilePath}")
        End If

        Using stream As Stream = File.OpenRead(FilePath)
            ' streamから読み込むためのBinaryReaderを作成
            Using reader As New BinaryReader(stream)
                Dim HeaderBuff() As Byte = reader.ReadBytes(352)

                'Endian判定
                If ReadValue(Of Int32)(HeaderBuff, 0, 0) = 348 Then
                    BigEndianFlag = False
                ElseIf ReadValue(Of Int32)(HeaderBuff, 0, 1) = 348 Then
                    BigEndianFlag = True
                Else
                    Throw New InvalidDataException($"File is not NIfTI: {FilePath}")
                End If

                '次元数
                Dimension0 = ReadValue(Of Int16)(HeaderBuff, OFFSET_DIMENSION, BigEndianFlag)

                'Matrixサイズ
                MatrixX = ReadValue(Of Int16)(HeaderBuff, OFFSET_MATRIX_X, BigEndianFlag)
                MatrixY = ReadValue(Of Int16)(HeaderBuff, OFFSET_MATRIX_Y, BigEndianFlag)
                SliceCount = ReadValue(Of Int16)(HeaderBuff, OFFSET_SLICE_COUNT, BigEndianFlag)

                '未使用次元
                Dimension4 = ReadValue(Of Int16)(HeaderBuff, OFFSET_SLICE_COUNT + 2, BigEndianFlag)
                Dimension5 = ReadValue(Of Int16)(HeaderBuff, OFFSET_SLICE_COUNT + 4, BigEndianFlag)
                Dimension6 = ReadValue(Of Int16)(HeaderBuff, OFFSET_SLICE_COUNT + 6, BigEndianFlag)
                Dimension7 = ReadValue(Of Int16)(HeaderBuff, OFFSET_SLICE_COUNT + 8, BigEndianFlag)

                'Pixelバッファ
                Pixel = New Double(SliceCount - 1, MatrixY - 1, MatrixX - 1) {}

                'データタイプ
                DataType = ReadValue(Of Int16)(HeaderBuff, OFFSET_DATATYPE, BigEndianFlag)

                'BitsPerPixel
                BitsPerPixel = ReadValue(Of Int16)(HeaderBuff, OFFSET_BITPERPIXEL, BigEndianFlag)

                'ピクセルサイズ
                SizeDim0 = ReadValue(Of Single)(HeaderBuff, OFFSET_PIXDIM, BigEndianFlag)
                SizeX = ReadValue(Of Single)(HeaderBuff, OFFSET_SIZE_X, BigEndianFlag)
                SizeY = ReadValue(Of Single)(HeaderBuff, OFFSET_SIZE_Y, BigEndianFlag)
                SizeZ = ReadValue(Of Single)(HeaderBuff, OFFSET_SIZE_Z, BigEndianFlag)
                '未使用次元
                SizeDim4 = ReadValue(Of Single)(HeaderBuff, OFFSET_SIZE_Z + 4, BigEndianFlag)
                SizeDim5 = ReadValue(Of Single)(HeaderBuff, OFFSET_SIZE_Z + 8, BigEndianFlag)
                SizeDim6 = ReadValue(Of Single)(HeaderBuff, OFFSET_SIZE_Z + 12, BigEndianFlag)
                SizeDim7 = ReadValue(Of Single)(HeaderBuff, OFFSET_SIZE_Z + 16, BigEndianFlag)

                '画素データのオフセット
                VoxelOffset = ReadValue(Of Single)(HeaderBuff, OFFSET_VOXEL_OFFSET, BigEndianFlag)

                'スケーリングファクタ
                RescaleSlope = ReadValue(Of Single)(HeaderBuff, OFFSET_RESCALE_SLOPE, BigEndianFlag)
                RescaleIntercept = ReadValue(Of Single)(HeaderBuff, OFFSET_RESCALE_INTERCEPT, BigEndianFlag)

                'ORIENTATION AND LOCATION
                qFormCode = ReadValue(Of Short)(HeaderBuff, OFFSET_QFORM_CODE, BigEndianFlag)
                sFormCode = ReadValue(Of Short)(HeaderBuff, OFFSET_SFORM_CODE, BigEndianFlag)
                qParamB = ReadValue(Of Single)(HeaderBuff, OFFSET_QPARAM_B, BigEndianFlag)
                qParamC = ReadValue(Of Single)(HeaderBuff, OFFSET_QPARAM_C, BigEndianFlag)
                qParamD = ReadValue(Of Single)(HeaderBuff, OFFSET_QPARAM_D, BigEndianFlag)
                qOffsetX = ReadValue(Of Single)(HeaderBuff, OFFSET_QOFFSET_X, BigEndianFlag)
                qOffsetY = ReadValue(Of Single)(HeaderBuff, OFFSET_QOFFSET_Y, BigEndianFlag)
                qOffsetZ = ReadValue(Of Single)(HeaderBuff, OFFSET_QOFFSET_Z, BigEndianFlag)

                For col As Integer = 0 To 3
                    SrowX(col) = ReadValue(Of Single)(HeaderBuff, OFFSET_SROW_X + col * 4, BigEndianFlag)
                    SrowY(col) = ReadValue(Of Single)(HeaderBuff, OFFSET_SROW_Y + col * 4, BigEndianFlag)
                    SrowZ(col) = ReadValue(Of Single)(HeaderBuff, OFFSET_SROW_Z + col * 4, BigEndianFlag)
                Next

                '画素データ取り込み
                Dim BytesPerPixel As Integer = BitsPerPixel / 8
                Dim AllPixelBuff() As Byte = reader.ReadBytes(CLng(MatrixX) * CLng(MatrixY) * CLng(SliceCount) * BytesPerPixel)
                Dim Offset As Long = 0

                '画素値再現
                For Z As Integer = 0 To SliceCount - 1
                    For Y As Integer = 0 To MatrixY - 1
                        For X As Integer = 0 To MatrixX - 1
                            Select Case DataType
                                Case 1
                                    If ReadValue(Of Byte)(AllPixelBuff, Offset, BigEndianFlag) = 0 Then
                                        Pixel(Z, Y, X) = 0
                                    Else
                                        Pixel(Z, Y, X) = 1
                                    End If
                                Case 2
                                    Pixel(Z, Y, X) = CDbl(ReadValue(Of Byte)(AllPixelBuff, Offset, BigEndianFlag))
                                Case 4
                                    Pixel(Z, Y, X) = CDbl(ReadValue(Of Int16)(AllPixelBuff, Offset, BigEndianFlag))
                                Case 8
                                    Pixel(Z, Y, X) = CDbl(ReadValue(Of Int32)(AllPixelBuff, Offset, BigEndianFlag))
                                Case 16
                                    Pixel(Z, Y, X) = CDbl(ReadValue(Of Single)(AllPixelBuff, Offset, BigEndianFlag))
                                Case 64
                                    Pixel(Z, Y, X) = CDbl(ReadValue(Of Double)(AllPixelBuff, Offset, BigEndianFlag))
                                Case 512
                                    Pixel(Z, Y, X) = CDbl(ReadValue(Of UInt16)(AllPixelBuff, Offset, BigEndianFlag))
                                Case Else
                                    Throw New NotSupportedException("Type " & DataType.ToString() & " is not supported")
                            End Select

                            If DataType <> 1 Then
                                Pixel(Z, Y, X) = Pixel(Z, Y, X) * RescaleSlope + RescaleIntercept
                            End If

                            Offset += BytesPerPixel
                        Next
                    Next
                Next

            End Using
        End Using

    End Sub

    Sub Write(ByVal FilePath As String)

        Dim HeaderBuff(351) As Byte

        'Endian判定
        'Array.Copy(BitConverter.GetBytes(CLng(348)), 0, HeaderBuff, 0, 4)
        HeaderBuff(0) = &H5C
        HeaderBuff(1) = &H1
        HeaderBuff(2) = &H0
        HeaderBuff(3) = &H0

        '次元数（三次元固定）
        HeaderBuff(OFFSET_DIMENSION) = &H3
        HeaderBuff(OFFSET_DIMENSION + 1) = &H0

        'Matrixサイズ,スライス枚数
        Array.Copy(BitConverter.GetBytes(MatrixX), 0, HeaderBuff, OFFSET_MATRIX_X, 2)
        Array.Copy(BitConverter.GetBytes(MatrixY), 0, HeaderBuff, OFFSET_MATRIX_Y, 2)
        Array.Copy(BitConverter.GetBytes(SliceCount), 0, HeaderBuff, OFFSET_SLICE_COUNT, 2)
        Array.Copy(BitConverter.GetBytes(Dimension4), 0, HeaderBuff, OFFSET_SLICE_COUNT + 4, 2)
        Array.Copy(BitConverter.GetBytes(Dimension5), 0, HeaderBuff, OFFSET_SLICE_COUNT + 8, 2)
        Array.Copy(BitConverter.GetBytes(Dimension6), 0, HeaderBuff, OFFSET_SLICE_COUNT + 12, 2)
        Array.Copy(BitConverter.GetBytes(Dimension7), 0, HeaderBuff, OFFSET_SLICE_COUNT + 16, 2)

        'データタイプ（float型16固定）
        Array.Copy(BitConverter.GetBytes(CShort(16)), 0, HeaderBuff, OFFSET_DATATYPE, 2)
        'HeaderBuff(OFFSET_DATATYPE) = &H10

        'bit/pixel（float型=32固定）
        Array.Copy(BitConverter.GetBytes(CShort(32)), 0, HeaderBuff, OFFSET_BITPERPIXEL, 2)
        'HeaderBuff(OFFSET_BITPERPIXEL) = &H20

        'ピクセルサイズ
        Array.Copy(BitConverter.GetBytes(SizeDim0), 0, HeaderBuff, OFFSET_PIXDIM, 4)
        Array.Copy(BitConverter.GetBytes(SizeX), 0, HeaderBuff, OFFSET_SIZE_X, 4)
        Array.Copy(BitConverter.GetBytes(SizeY), 0, HeaderBuff, OFFSET_SIZE_Y, 4)
        Array.Copy(BitConverter.GetBytes(SizeZ), 0, HeaderBuff, OFFSET_SIZE_Z, 4)
        Array.Copy(BitConverter.GetBytes(SizeDim4), 0, HeaderBuff, OFFSET_SIZE_Z + 4, 4)
        Array.Copy(BitConverter.GetBytes(SizeDim5), 0, HeaderBuff, OFFSET_SIZE_Z + 8, 4)
        Array.Copy(BitConverter.GetBytes(SizeDim6), 0, HeaderBuff, OFFSET_SIZE_Z + 12, 4)
        Array.Copy(BitConverter.GetBytes(SizeDim7), 0, HeaderBuff, OFFSET_SIZE_Z + 16, 4)

        '画素データのオフセット
        Array.Copy(BitConverter.GetBytes(VoxelOffset), 0, HeaderBuff, OFFSET_VOXEL_OFFSET, 4)

        'スケーリングファクタ(スケーリングなし固定）
        Array.Copy(BitConverter.GetBytes(CSng(1)), 0, HeaderBuff, OFFSET_RESCALE_SLOPE, 4)
        Array.Copy(BitConverter.GetBytes(CSng(0)), 0, HeaderBuff, OFFSET_RESCALE_INTERCEPT, 4)

        'ORIENTATION AND LOCATION
        Array.Copy(BitConverter.GetBytes(qFormCode), 0, HeaderBuff, OFFSET_QFORM_CODE, 2)
        Array.Copy(BitConverter.GetBytes(sFormCode), 0, HeaderBuff, OFFSET_SFORM_CODE, 2)

        Array.Copy(BitConverter.GetBytes(qParamB), 0, HeaderBuff, OFFSET_QPARAM_B, 4)
        Array.Copy(BitConverter.GetBytes(qParamC), 0, HeaderBuff, OFFSET_QPARAM_C, 4)
        Array.Copy(BitConverter.GetBytes(qParamD), 0, HeaderBuff, OFFSET_QPARAM_D, 4)
        Array.Copy(BitConverter.GetBytes(qOffsetX), 0, HeaderBuff, OFFSET_QOFFSET_X, 4)
        Array.Copy(BitConverter.GetBytes(qOffsetY), 0, HeaderBuff, OFFSET_QOFFSET_Y, 4)
        Array.Copy(BitConverter.GetBytes(qOffsetZ), 0, HeaderBuff, OFFSET_QOFFSET_Z, 4)

        For col As Integer = 0 To 3
            Array.Copy(BitConverter.GetBytes(SrowX(col)), 0, HeaderBuff, OFFSET_SROW_X + col * 4, 4)
            Array.Copy(BitConverter.GetBytes(SrowY(col)), 0, HeaderBuff, OFFSET_SROW_Y + col * 4, 4)
            Array.Copy(BitConverter.GetBytes(SrowZ(col)), 0, HeaderBuff, OFFSET_SROW_Z + col * 4, 4)
        Next

        'Magic Word("n+1" 固定)
        Array.Copy(Text.Encoding.ASCII.GetBytes(MAGIC_WORD), 0, HeaderBuff, OFFSET_MAGIC_WORD, MAGIC_WORD.Length)


        '画素値をバッファリング
        Dim BytesPerPixel As Integer = 4
        Dim DestBuff(CLng(MatrixX) * CLng(MatrixY) * CLng(SliceCount) * BytesPerPixel - 1) As Byte
        For Z As Integer = 0 To SliceCount - 1
            For Y As Integer = 0 To MatrixY - 1
                For X As Integer = 0 To MatrixX - 1
                    Array.Copy(BitConverter.GetBytes(CSng(Pixel(Z, Y, X))), 0, DestBuff, ((MatrixX * MatrixY * Z) + (MatrixX * Y) + X) * BytesPerPixel, BytesPerPixel)
                Next
            Next
        Next

        'ファイル書き込み
        Using stream As Stream = File.OpenWrite(FilePath)
            ' streamに書き込むためのBinaryWriterを作成
            Using writer As New BinaryWriter(stream)

                'ヘッダ書き込み
                writer.Write(HeaderBuff)

                'ピクセル書き込み
                writer.Write(DestBuff)
            End Using
        End Using

    End Sub

    Private Function ReadValue(Of T)(Buffer() As Byte, Offset As Integer, isBigEndian As Boolean) As T

        ' サイズを型ごとに指定
        Dim BuffSize As Integer = Runtime.InteropServices.Marshal.SizeOf(GetType(T))

        ' 指定したサイズ分のバイトを取り出す
        Dim TempBuff(BuffSize - 1) As Byte
        Array.Copy(Buffer, Offset, TempBuff, 0, BuffSize)

        ' BigEndian対応
        If isBigEndian Then
            Array.Reverse(TempBuff)
        End If

        ' 型に応じたBitConverterのメソッドで変換
        Dim result As Object = Nothing
        Select Case GetType(T)
            Case GetType(Short)
                result = BitConverter.ToInt16(TempBuff, 0)
            Case GetType(Integer)
                result = BitConverter.ToInt32(TempBuff, 0)
            Case GetType(Long)
                result = BitConverter.ToInt64(TempBuff, 0)
            Case GetType(Single)
                result = BitConverter.ToSingle(TempBuff, 0)
            Case GetType(Double)
                result = BitConverter.ToDouble(TempBuff, 0)
            Case GetType(Byte)
                result = TempBuff(0)
            Case GetType(UShort)
                result = BitConverter.ToUInt16(TempBuff, 0)
            Case GetType(ULong)
                result = BitConverter.ToUInt64(TempBuff, 0)
        End Select

        Return CType(result, T)
    End Function

    Public Function GetPixels() As Double(,,)
        Return CType(Pixel.Clone(), Double(,,))
    End Function

    Public Sub SetPixels(Source As Double(,,))

        Dim sizeX As Integer = Source.GetLength(2) ' X軸のサイズ
        Dim sizeY As Integer = Source.GetLength(1) ' Y軸のサイズ
        Dim sizeZ As Integer = Source.GetLength(0) ' Z軸のサイズ

        ' サイズ検証
        If sizeX <> MatrixX OrElse sizeY <> MatrixY OrElse sizeZ <> SliceCount Then
            Throw New ArgumentException($"入力配列のサイズが不正です。" & vbCrLf &
                                    $"期待されるサイズ: ({MatrixX}, {MatrixY}, {SliceCount})" & vbCrLf &
                                    $"実際のサイズ: ({sizeX}, {sizeY}, {sizeZ})")
        End If

        Pixel = CType(Source.Clone(), Double(,,))
    End Sub

    Sub CloneHeader(Source As clsNIfTIImage)
        Me.BigEndianFlag = Source.BigEndianFlag
        Me.MatrixX = Source.MatrixX
        Me.MatrixY = Source.MatrixY
        Me.SliceCount = Source.SliceCount
        Me.DataType = Source.DataType
        Me.BitsPerPixel = Source.BitsPerPixel
        Me.SizeX = Source.SizeX
        Me.SizeY = Source.SizeY
        Me.SizeZ = Source.SizeZ
        Me.VoxelOffset = Source.VoxelOffset
        Me.RescaleSlope = Source.RescaleSlope
        Me.RescaleIntercept = Source.RescaleIntercept
        Me.qFormCode = Source.qFormCode
        Me.sFormCode = Source.sFormCode
        Me.qParamB = Source.qParamB
        Me.qParamC = Source.qParamC
        Me.qParamD = Source.qParamD
        Me.qOffsetX = Source.qOffsetX
        Me.qOffsetY = Source.qOffsetY
        Me.qOffsetZ = Source.qOffsetZ
        Me.SrowX = Source.SrowX
        Me.SrowY = Source.SrowY
        Me.SrowZ = Source.SrowZ
    End Sub

    Sub CreatePixelBuff(Source As clsNIfTIImage)
        Me.Pixel = New Double(Source.SliceCount - 1, Source.MatrixY - 1, Source.MatrixX - 1) {}
    End Sub
End Class

