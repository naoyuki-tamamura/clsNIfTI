Imports System.IO
Public Class clsNIfTIImage
    Public Property VoxelOffset As Int32         'Imageのオフセット(通常は352）
    Public Property BigEndianFlag As Boolean    'バイトオーダー（TrueでBigEndianFlag）
    Public Property MatrixX As Int32             '横方向のマトリクスサイズ
    Public Property MatrixY As Int32             '縦方向のマトリクスサイズ
    Public Property SliceCount As Int32          'スライス枚数
    Public Property SizeX As Single             '横方向のピクセルサイズ
    Public Property SizeY As Single             '縦方向のピクセルサイズ
    Public Property SizeZ As Single             'スライス厚
    Public Property DataType As Short           'データタイプ（1:Binary 2:Unsigned char 4:signed short　8:signed long 16:float 64:double 512:unsigned short)
    Public Property RescaleSlope As Single
    Public Property RescaleIntercept As Single

    Const OFFSET_MATRIX_X As Integer = 42
    Const OFFSET_MATRIX_Y As Integer = 44
    Const OFFSET_SLICE_COUNT As Integer = 46
    Const OFFSET_DATATYPE As Integer = 70
    Const OFFSET_SIZE_X As Integer = 80
    Const OFFSET_SIZE_Y As Integer = 84
    Const OFFSET_SIZE_Z As Integer = 88
    Const OFFSET_VOXEL_OFFSET As Integer = 108
    Const OFFSET_RESCALE_SLOPE As Integer = 112
    Const OFFSET_RESCALE_INTERCEPT As Integer = 116

    Sub New()
        VoxelOffset = 352
        BigEndianFlag = False
        MatrixX = 0
        MatrixY = 0
        SliceCount = 0
        SizeX = 0
        SizeY = 0
        SizeZ = 0
        RescaleSlope = 1
        RescaleIntercept = 0
    End Sub

    Sub ReadHeader(ByVal FilePath As String)

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

                'Matrixサイズ
                MatrixX = ReadValue(Of Int16)(HeaderBuff, OFFSET_MATRIX_X, BigEndianFlag)
                MatrixY = ReadValue(Of Int16)(HeaderBuff, OFFSET_MATRIX_Y, BigEndianFlag)
                SliceCount = ReadValue(Of Int16)(HeaderBuff, OFFSET_SLICE_COUNT, BigEndianFlag)

                'データタイプ
                DataType = ReadValue(Of Int16)(HeaderBuff, OFFSET_DATATYPE, BigEndianFlag)

                'ピクセルサイズ
                SizeX = ReadValue(Of Single)(HeaderBuff, OFFSET_SIZE_X, BigEndianFlag)
                SizeY = ReadValue(Of Single)(HeaderBuff, OFFSET_SIZE_Y, BigEndianFlag)
                SizeZ = ReadValue(Of Single)(HeaderBuff, OFFSET_SIZE_Z, BigEndianFlag)

                '画素データのオフセット
                VoxelOffset = ReadValue(Of Single)(HeaderBuff, OFFSET_VOXEL_OFFSET, BigEndianFlag)

                'スケーリングファクタ
                RescaleSlope = ReadValue(Of Single)(HeaderBuff, OFFSET_RESCALE_SLOPE, BigEndianFlag)
                RescaleIntercept = ReadValue(Of Single)(HeaderBuff, OFFSET_RESCALE_INTERCEPT, BigEndianFlag)

            End Using
        End Using

    End Sub

    Function ReadImage(ByVal HeaderInfo As clsNIfTIImage, ByVal FilePath As String) As Double(,,)

        If Not File.Exists(FilePath) Then
            Throw New FileNotFoundException($"File not found: {FilePath}")
        End If

        Dim ImageBuff(HeaderInfo.SliceCount - 1, HeaderInfo.MatrixY - 1, HeaderInfo.MatrixX - 1) As Double

        Using stream As Stream = File.OpenRead(FilePath)
            ' streamから読み込むためのBinaryReaderを作成
            Using reader As New BinaryReader(stream)
                '画素値取り込み

                Dim BytesPerPixel As Long

                Select Case HeaderInfo.DataType
                    Case 1      'Binary
                        BytesPerPixel = 1
                    Case 2      'Unsigned char
                        BytesPerPixel = 1
                    Case 4      'signed short
                        BytesPerPixel = 2
                    Case 8      'signed long
                        BytesPerPixel = 4
                    Case 16     'float
                        BytesPerPixel = 4
                    Case 64     'double
                        BytesPerPixel = 8
                    Case 512    'unsigned short
                        BytesPerPixel = 2
                    Case Else
                        Throw New NotSupportedException("Type " & HeaderInfo.DataType.ToString() & " is not supported")
                End Select

                'ヘッダ部分を読み飛ばす
                reader.ReadBytes(HeaderInfo.VoxelOffset)

                '画素データ取り込み
                Dim AllPixelBuff() As Byte = reader.ReadBytes(HeaderInfo.MatrixX * HeaderInfo.MatrixY * HeaderInfo.SliceCount * BytesPerPixel)
                Dim Offset As Long = 0

                '画素値再現
                For Z = 0 To HeaderInfo.SliceCount - 1
                    For Y = 0 To HeaderInfo.MatrixY - 1
                        For X = 0 To HeaderInfo.MatrixX - 1
                            Select Case HeaderInfo.DataType
                                Case 1
                                    If ReadValue(Of Byte)(AllPixelBuff, Offset, BigEndianFlag) = 0 Then
                                        ImageBuff(Z, Y, X) = 0
                                    Else
                                        ImageBuff(Z, Y, X) = 1
                                    End If
                                Case 2
                                    ImageBuff(Z, Y, X) = CDbl(ReadValue(Of Byte)(AllPixelBuff, Offset, BigEndianFlag)) * HeaderInfo.RescaleSlope + HeaderInfo.RescaleIntercept
                                Case 4
                                    ImageBuff(Z, Y, X) = CDbl(ReadValue(Of Int16)(AllPixelBuff, Offset, BigEndianFlag)) * HeaderInfo.RescaleSlope + HeaderInfo.RescaleIntercept
                                Case 8
                                    ImageBuff(Z, Y, X) = CDbl(ReadValue(Of Int32)(AllPixelBuff, Offset, BigEndianFlag)) * HeaderInfo.RescaleSlope + HeaderInfo.RescaleIntercept
                                Case 16
                                    ImageBuff(Z, Y, X) = CDbl(ReadValue(Of Single)(AllPixelBuff, Offset, BigEndianFlag)) * HeaderInfo.RescaleSlope + HeaderInfo.RescaleIntercept
                                Case 64
                                    ImageBuff(Z, Y, X) = CDbl(ReadValue(Of Double)(AllPixelBuff, Offset, BigEndianFlag)) * HeaderInfo.RescaleSlope + HeaderInfo.RescaleIntercept
                                Case 512
                                    ImageBuff(Z, Y, X) = CDbl(ReadValue(Of UInt16)(AllPixelBuff, Offset, BigEndianFlag)) * HeaderInfo.RescaleSlope + HeaderInfo.RescaleIntercept
                                Case Else
                                    Throw New NotSupportedException("Type " & HeaderInfo.DataType.ToString() & " is not supported")
                            End Select
                            Offset += BytesPerPixel
                        Next
                    Next
                Next
            End Using
        End Using

        ReadImage = ImageBuff
    End Function

    Sub WriteImage(ByVal HeaderInfo As clsNIfTIImage, ByVal ImageBuff As Double(,,), ByVal FilePath As String)

        Dim HeaderBuff(351) As Byte

        'Endian判定
        Array.Copy(BitConverter.GetBytes(CLng(348)), 0, HeaderBuff, 0, 4)

        'Matrixサイズ
        Array.Copy(BitConverter.GetBytes(CShort(HeaderInfo.MatrixX)), 0, HeaderBuff, OFFSET_MATRIX_X, 2)
        Array.Copy(BitConverter.GetBytes(CShort(HeaderInfo.MatrixY)), 0, HeaderBuff, OFFSET_MATRIX_Y, 2)

        'スライス枚数
        Array.Copy(BitConverter.GetBytes(CShort(HeaderInfo.SliceCount)), 0, HeaderBuff, OFFSET_SLICE_COUNT, 2)

        'データタイプ（float型固定）
        Array.Copy(BitConverter.GetBytes(CShort(16)), 0, HeaderBuff, OFFSET_DATATYPE, 2)

        'ピクセルサイズ
        Array.Copy(BitConverter.GetBytes(CSng(HeaderInfo.SizeX)), 0, HeaderBuff, OFFSET_SIZE_X, 4)
        Array.Copy(BitConverter.GetBytes(CSng(HeaderInfo.SizeY)), 0, HeaderBuff, OFFSET_SIZE_Y, 4)
        Array.Copy(BitConverter.GetBytes(CSng(HeaderInfo.SizeZ)), 0, HeaderBuff, OFFSET_SIZE_Z, 4)

        '画素データのオフセット
        Array.Copy(BitConverter.GetBytes(CLng(HeaderInfo.VoxelOffset)), 0, HeaderBuff, OFFSET_VOXEL_OFFSET, 4)

        'スケーリングファクタ
        Array.Copy(BitConverter.GetBytes(CSng(1)), 0, HeaderBuff, OFFSET_RESCALE_SLOPE, 4)
        Array.Copy(BitConverter.GetBytes(CSng(0)), 0, HeaderBuff, OFFSET_RESCALE_INTERCEPT, 4)

        Dim BytesPerPixel As Long = 4
        Dim DestBuff(HeaderInfo.MatrixX * HeaderInfo.MatrixY * HeaderInfo.SliceCount * BytesPerPixel) As Byte

        '画素値をバッファリング
        For Z = 0 To HeaderInfo.SliceCount - 1
            For Y = 0 To HeaderInfo.MatrixY - 1
                For X = 0 To HeaderInfo.MatrixX - 1
                    Array.Copy(BitConverter.GetBytes(CSng(ImageBuff(Z, Y, X))), 0, DestBuff, ((HeaderInfo.MatrixX * HeaderInfo.MatrixY * Z) + (HeaderInfo.MatrixX * Y) + X) * BytesPerPixel, BytesPerPixel)
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
        Dim BuffSize As Integer
        Select Case GetType(T)
            Case GetType(Short)
                BuffSize = 2
            Case GetType(Integer)
                BuffSize = 4
            Case GetType(Long)
                BuffSize = 8
            Case GetType(Single)
                BuffSize = 4
            Case GetType(Double)
                BuffSize = 8
            Case GetType(Byte)
                BuffSize = 1
            Case GetType(UShort)
                BuffSize = 2
            Case GetType(ULong)
                BuffSize = 8
            Case Else
                Throw New NotSupportedException("Type " & GetType(T).ToString() & " is not supported")
        End Select

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

End Class