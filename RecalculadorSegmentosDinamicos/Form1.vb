
Imports System.IO
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Data.Odbc
Public Class Form1


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim conexion As New SqlConnection(ModConexion.conEvoOtras)
        Dim sentencia As String
        Dim comando As SqlCommand

        Try
            conexion.Open()
            sentencia = "Delete From RecalculadorSegmentosEstaticos"
            comando = New SqlCommand(sentencia, conexion)
            MessageBox.Show("Se ha borrado la tabla")
            conexion.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        ''----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '' Cargamos el fichero 
        ''----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        OFD2.Title = "Abrir"
        OFD2.DefaultExt = "txt"
        OFD2.Filter = "csv files (*.csv)|*.csv"
        OFD2.ShowDialog()

        ' Llantas con slicks
        ' Dinamometrica
        ' Vaso tornillos seguridad
        ' Tuercas Standar
        ' Gato
        ' Guantes
        ' Zapatillas
        ' Casco
        ' Guantes de trabajo
        ' Cuñas Ruedas


        'declaramos la Tabla donde añadiremos los datos y la fila correspondiente
        Dim MiTabla As DataTable = New DataTable("MyTable")
        Dim MiFila As DataRow
        Dim Separador As Char = ";"
        'declaramos el resto de variables que nos harán falta
        Dim pos As Boolean = False
        Dim i As Integer
        Dim fieldValues As String()
        Dim sr2 As New System.IO.StreamReader(OFD2.FileName, System.Text.Encoding.Default)
        ' Dim miReader As IO.StreamReader
        Try
            'Abrimos el fichero y leemos la primera linea con el fin de determinar cuantos campos tenemos
            'miReader = File.OpenText(RutaCompletaArchivo)
            fieldValues = sr2.ReadLine().Split(Separador)
            'Creamos las columnas de la cabecera
            For i = 0 To fieldValues.Length() - 1
                MiTabla.Columns.Add(New DataColumn(fieldValues(i).ToString()))
            Next
            'Continuamos leyendo el resto de filas y añadiendolas a la tabla
            While sr2.Peek() <> -1
                fieldValues = sr2.ReadLine().Split(Separador)
                MiFila = MiTabla.NewRow
                For i = 0 To fieldValues.Length() - 1
                    MiFila.Item(i) = fieldValues(i).ToString
                Next
                MiTabla.Rows.Add(MiFila)
            End While
            'Cerramos el reader
            sr2.Close()
        Catch ex As Exception
            'Gestionamos las excepciones
            Dim mensaje As String
            mensaje = "alert ('Ha ocurrido el siguiente error al importar el archivo: " + ex.ToString + "');"
            MsgBox(mensaje)
        Finally
            'Si queremos ejecutar algo exista excepción o no
        End Try

        ''----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '' Volcamos el datatable del paso anterior a la tabla: RecalculadorSegmentosEstaticos
        ''----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Try
            conexion.Open()
            Dim copiar As New SqlBulkCopy(conexion)
            copiar.DestinationTableName = "RecalculadorSegmentosEstaticos"
            copiar.WriteToServer(MiTabla)
            conexion.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' En este paso 
        '   -   Lanzamos el procedimiento sp_carga para insertar los registros en Evolution
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Dim sentencia As String
        Dim comando As SqlCommand

        ModConexion.conexion = New SqlConnection()
        'Utiliza la conexion contra el servidor EVOLUTION
        ModConexion.conexion.ConnectionString = ModConexion.conEvoCaixa

        Try
            conexion.Open()
            sentencia = "SELECT DISTINCT IdCampanya,AtributoSegmento  FROM [OTRASCAMPANYAS].[dbo].[RecalculadorSegmentosEstaticos] "
            comando = New SqlCommand(sentencia, conexion)
            lector = comando.ExecuteReader
            While lector.Read()
                ModConexion.conexion = New SqlConnection()
                ModConexion.conexion.ConnectionString = ModConexion.conEvoDB
                Try
                    conexion.Open()
                    comando = New SqlCommand("sp_ActualizarContadoresSegmento", conexion)
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Parameters.Add(New SqlParameter("@idCampanya", lector("IdCampanya")))
                    comando.Parameters.Add(New SqlParameter("@sAtributo", lector("AtributoSegmento")))
                    comando.ExecuteNonQuery()
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            End While
            conexion.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub
End Class
