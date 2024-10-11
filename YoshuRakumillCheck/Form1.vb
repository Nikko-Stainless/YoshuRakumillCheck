Imports System.IO
Imports Oracle.ManagedDataAccess.Client
Imports Nikko.Windows

Public Class Form1
    Dim folderPath As String = ""
    Dim connectionStringForYoshu As String = ""
    Dim connectionStringForNikko As String = ""
    Dim fileINI = ""

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'load connect Nikko
        Dim aaaa = New AAAA()
        connectionStringForNikko = aaaa.GetConnectionString
        'load connect Yoshu
        If Not LoadFileini() Then Me.Close()
        Label1.Text = ""

        SelectedIndexChanged()
    End Sub
    'load file ini 
    Private Function LoadFileini() As Boolean
        fileINI = Application.StartupPath + "\YoshuRakumillCheck_AppSettings.ini"

        If Not IO.File.Exists(fileINI) Then Return False

        Dim lines As String() = IO.File.ReadAllLines(fileINI)
        Dim section As String = ""

        For Each line As String In lines
            line = line.Trim()
            If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                section = line.Trim("["c, "]"c)
            ElseIf Not String.IsNullOrEmpty(line) Then
                Select Case section
                    Case "ConnectString"
                        connectionStringForYoshu = line.Trim()
                    Case "RakumillCheckFolder"
                        folderPath = line.Trim()
                End Select
            End If
        Next
        Return True
    End Function
    'SQL
    Private Function GetData(query As String, isConnectString As String) As DataTable
        Dim dataTable As New DataTable()

        Using connection As New OracleConnection(isConnectString)
            Using command As New OracleCommand(query, connection)
                Using adapter As New OracleDataAdapter(command)
                    connection.Open()
                    adapter.Fill(dataTable)
                End Using
            End Using
        End Using

        Return dataTable
    End Function
    Private Sub SelectedIndexChanged() Handles TabControl1.SelectedIndexChanged
        Dim selectedIndex As Integer = TabControl1.SelectedIndex
        Select Case selectedIndex
            Case 0
                selectDataTab1()
            Case 1
                selectDataTab2()
            Case 2
                selectDataTab3()
            Case 3
                checkFolder()
        End Select
    End Sub

#Region "tab_1"
    Dim sql_TAB1 = "
    SELECT a.URIKEIYAKUNO as 売契約No, 
           a.URIKEIYAKUGYOUNO as 行No,
           a.sunpou as 寸法,
           a.insuu as 員数,
           a.shukkaymd as 出荷日,
           a.W_NO , a.C_NO, 
           a.milsheetno as ミルシートNo,
           a.MAKERCD as メーカーCD,
           b.makerryakumei as メーカー略称,
           b.makermei as メーカー名
    FROM 
          RKM_shipping a, rkm_m_maker b 
   WHERE 
          a.makercd = b.makercd(+) AND a.URIKEIYAKUNO IN ({0}) 
   ORDER BY 
          1,2"
    Private Sub selectDataTab1()
        Dim pdfFiles As String() = Directory.GetFiles(folderPath, "*.pdf")
        Dim listKeys As String = ""

        For Each pdfFile As String In pdfFiles
            Dim fileName As String = Path.GetFileNameWithoutExtension(pdfFile)
            If fileName.Length >= 8 Then
                Dim first8Chars As String = fileName.Substring(0, 8)
                ' ファイル名の最初の8文字を取得
                If IsNumeric(first8Chars) Then
                    ' 最初の8文字が数字であるかを確認
                    listKeys = listKeys + "'" + first8Chars + "',"
                End If
            End If
        Next

        If listKeys.Length > 0 Then
            listKeys = listKeys.TrimEnd(","c)
        End If

        If Not String.IsNullOrEmpty(listKeys) Then
            Dim sqlQuery As String = String.Format(sql_TAB1, listKeys)
            Dim resultTable As DataTable = GetData(sqlQuery, connectionStringForYoshu)
            DataGridView1.DataSource = Nothing
            DataGridView1.DataSource = resultTable
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        selectDataTab1()
    End Sub
#End Region
#Region "tab_2"
    Dim sql_TAB2 = "
    SELECT
    distinct d.KAISHAKBN as 会社区分,
　　d.shimeidisp as 締切　,
　　decode(lengthb(a.tourokutime),7,0||substr( a.tourokutime,1,1) || ':' || substr(a.tourokutime, 2, 2),substr(a.tourokutime,1,2) || ':' || substr(a.tourokutime, 3, 2) ) as 時間,
　　a.USER_YOBI1 as orderNo,
　　a.sunpou as 寸法,
　　a.inzuu as 員数　,
　　a.MAKERCD as メーカーCD　,
　　b.MAKERRYAKUMEI as メーカー略称　,
　　a.c_no　,
　　a.MILSHEETNO as ミルシートNo　,
　　decode(c.MILSHEET_YOUHIKBN,1,'要','不要') as 区分,
　　c.SHUKKAYMD as 出荷日,
　　c.SHUKKAYMD2 as 出荷日２,
　　c.TOUROKUYMD as 出荷データ登録日,
　　c.KOUSHINKAISHAKBN as 更新区分　,
　　c.KOUSHINtantouno as 更新担当

FROM
  RKM_milsheet a,rkm_shipping c,rkm_m_shain d,
   rkm_m_maker b
where a.USER_YOBI1=c.URIKEIYAKUNO and c.KOUSHINTANTOUNO=d.tantouno and
 a.TOUROKUYMD=to_char(sysdate,'yyyymmdd') and a.MAKERCD=b.MAKERCD
order by
  1,2,3"

    Private Sub selectDataTab2()
        Dim resultTable As DataTable = GetData(sql_TAB2, connectionStringForYoshu)
        DataGridView2.DataSource = Nothing
        DataGridView2.DataSource = resultTable
        Label1.Text = DataGridView2.Rows.Count.ToString() + " 行"

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        selectDataTab2()
    End Sub
#End Region
#Region "tab_3"
    Dim sql_TAB3 = "
    SELECT distinct
      s.urikeiyakuno                                        as 豫洲受注NO
    , S.URIKEIYAKUGYOUNO                                    as 豫洲行NO
    , s.shukkaymd                                           as 豫洲出荷日
    , s.sunpou                                              as 豫洲寸法
    , S.INSUU                                               as 豫洲員数
    , A.URIKEIYAKUNO                                        as NS受注NO
    , A.URIKEIYAKUGYOUNO                                    as NS行NO
    , DECODE(
        A.MILSHEETFLG
        , 0
        , '不要'
        , 1
        , '要'
        , 2
        , '同時'
        , 4
        , '切証'
        , 5
        , '切同'
    )                                                       as MS区分
    , decode(c.milsheetno, null, '「未」紐付', '紐付済')    as 紐付
    , decode(a.milsheethakkouflg, 1, '未印刷', 0, '印刷済') as 印刷
    , C.W_NO
    , C.C_NO
FROM
    NS.URIKEIYAKU_M A
    , NS.tr_millprt C
    , NS.rkm_shipping@share4 S
WHERE
    s.MILSHEET_YOUHIKBN = '1'
    AND A.URIKEIYAKUNO = C.URIKEIYAKUNO
    AND A.URIKEIYAKUGYOUNO = C.URIKEIYAKUGYOUNO
    AND replace (a.kyakuchuuno, '-', '') = s.urikeiyakuno
    and s.milsheetno = '9999999'
    and s.USER_SHUKKASAKI like '日鋼ステンレス%'
    and NVL(S.SHUKKAYMD, S.SHUKKAYMD2) > '20211001'
    and a.gaikei = s.gaikei
    and a.nikuatu = s.nikuatu
    and a.NAGASA = s.nagasa
ORDER BY
    s.urikeiyakuno
    , S.URIKEIYAKUGYOUNO
    , S.SUNPOU"

    Private Sub selectDataTab3()
        Dim resultTable As DataTable = GetData(sql_TAB3, connectionStringForNikko)
        DataGridView3.DataSource = Nothing
        DataGridView3.DataSource = resultTable
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        selectDataTab3()
    End Sub
#End Region
#Region "tab_4"
    Public Class FolderInfo
        Public Property LinkFolder As String ' フォルダパス
        Public Property DisplayText As String ' 表示テキスト
        Public Property FileFormat As String ' ファイル形式

        ' コンストラクタ
        Public Sub New(link As String, display As String, format As String)
            LinkFolder = link
            DisplayText = display
            FileFormat = format
        End Sub
    End Class
    Private Sub checkFolder()
        Try
            lbBatch.Text = ""
            Dim folderInfos As New List(Of FolderInfo)()
            ' .iniファイルの読み込み
            If Not IO.File.Exists(fileINI) Then Return

            Dim lines As String() = File.ReadAllLines(fileINI)
            Dim isFolderCheckSection As Boolean = False
            Dim currentLink As String = ""
            Dim currentDisplayText As String = ""
            Dim currentFileFormat As String = ""

            For Each line As String In lines
                line = line.Trim()
                ' [FolderCheckForTab4] セクションの開始を確認
                If line.StartsWith("[FolderCheckForTab4]") Then
                    isFolderCheckSection = True
                    Continue For
                    ' セクションの終了を確認
                ElseIf line.StartsWith("[") AndAlso isFolderCheckSection Then
                    Exit For ' セクション終了後にループを抜ける
                End If

                ' セクション内の行を処理
                If isFolderCheckSection Then
                    If String.IsNullOrWhiteSpace(line) Then
                        Continue For ' 空白行をスキップ
                    End If

                    ' フォルダパスを保存
                    If String.IsNullOrEmpty(currentLink) Then
                        currentLink = line
                    ElseIf String.IsNullOrEmpty(currentDisplayText) Then
                        currentDisplayText = line ' 保存する表示テキスト
                    Else
                        currentFileFormat = line ' 保存するファイルフォーマット

                        ' FolderInfoオブジェクトをリストに追加
                        folderInfos.Add(New FolderInfo(currentLink, currentDisplayText, currentFileFormat))

                        ' リセット
                        currentLink = ""
                        currentDisplayText = ""
                        currentFileFormat = ""
                    End If
                End If
            Next

            ' フォルダのチェックとファイルのカウント
            For Each folderInfo In folderInfos
                ' フォルダが存在するか確認
                If Directory.Exists(folderInfo.LinkFolder) Then
                    Dim count As Integer = 0

                    ' ファイルのカウント
                    Dim files = Directory.GetFiles(folderInfo.LinkFolder, folderInfo.FileFormat)
                    Dim visibleFiles = files.Where(Function(f) (File.GetAttributes(f) And FileAttributes.Hidden) = 0).ToArray()
                    count += visibleFiles.Length

                    ' フォルダもカウントする場合 (FileFormat = "*.*")
                    If folderInfo.FileFormat = "*.*" Then
                        Dim directories = Directory.GetDirectories(folderInfo.LinkFolder)
                        Dim visibleDirectories = directories.Where(Function(d) (File.GetAttributes(d) And FileAttributes.Hidden) = 0).ToArray()
                        count += visibleDirectories.Length
                    End If

                    If count > 0 Then
                        ' 結果を表示
                        lbBatch.Text += $"{folderInfo.DisplayText}:" + vbNewLine + $" {count} ファイル/フォルダ" + vbNewLine + vbNewLine
                    End If
                Else
                    ' フォルダが見つからない場合のエラーメッセージ
                    lbBatch.Text += $"{folderInfo.DisplayText}: " + vbNewLine + $" フォルダが見つかりませんでした。" + vbNewLine + vbNewLine
                End If
            Next
        Catch ex As Exception
            ' エラー処理
            MessageBox.Show("エラー: " & ex.Message)
        End Try
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        checkFolder()
    End Sub

#End Region
End Class
