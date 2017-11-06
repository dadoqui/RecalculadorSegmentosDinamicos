Imports System.Data.SqlClient
Module ModConexion
    'Elementos de la conexion
    '***********************************************
    Public adaptador As SqlDataAdapter
    Public conjuntoDatos As DataSet
    Public comando As SqlCommandBuilder
    Public comando2 As SqlCommandBuilder
    Public conexion As SqlConnection
    Public conexion2 As SqlConnection
    Public lector As SqlDataReader
    Public procedimiento As SqlCommand

    'Conexiones: Evolution y Share para logins
    '***********************************************
    Public conEvoCaixa As String = "Initial Catalog=LACAIXA;Data Source=192.168.14.12;persist security info=True;User ID=sa;Password=8uh1cl1"
    Public conEvoDB As String = "Initial Catalog=EVOLUTIONDB;Data Source=192.168.14.12;persist security info=True;User ID=sa;Password=8uh1cl1"
    Public conEvoOtras As String = "Initial Catalog=OTRASCAMPANYAS;Data Source=192.168.14.12;persist security info=True;User ID=sa;Password=8uh1cl1"
    'Public conectSHARE As String = "Initial Catalog=PYRENALIA;Data Source=SHARE\SQLEXPRESS;integrated security=SSPI; persist security info=True;"
    Public conectShare As String = "Initial Catalog=PYRENALIA;Data Source=192.168.14.1\SQLEXPRESS;persist security info=True;User ID=choras;Password=pyre101"
End Module
