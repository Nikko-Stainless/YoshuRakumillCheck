Imports System.IO
Imports Oracle.ManagedDataAccess.Client
Public Class Form1
    Dim folderPath As String = ""
    Dim connectionString As String = ""

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Not LoadFileini() Then Me.Close()

        SelectedIndexChanged()
    End Sub
    'load file ini 
    Private Function LoadFileini() As Boolean
        Dim file = Application.StartupPath + "\YoshuRakumillCheck_AppSettings.ini"

        If Not IO.File.Exists(file) Then Return False

        Dim lines As String() = IO.File.ReadAllLines(file)
        Dim section As String = ""

        For Each line As String In lines
            line = line.Trim()
            If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                section = line.Trim("["c, "]"c)
            Else
                Select Case section
                    Case "ConnectString"
                        connectionString = line.Trim()
                    Case "RakumillCheckFolder"
                        folderPath = line.Trim()
                End Select
            End If
        Next
        Return True
    End Function
    'SQL
    Private Function GetData(query As String) As DataTable
        Dim dataTable As New DataTable()

        Using connection As New OracleConnection(connectionString)
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
            Dim resultTable As DataTable = GetData(sqlQuery)
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
        Dim resultTable As DataTable = GetData(sql_TAB2)
        DataGridView2.DataSource = Nothing
        DataGridView2.DataSource = resultTable
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        selectDataTab2()
    End Sub
#End Region




End Class
