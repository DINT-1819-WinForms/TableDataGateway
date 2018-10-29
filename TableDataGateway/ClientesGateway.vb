Imports System.Data.SqlClient

''' <summary>
''' Clase que define un gateway para la tabla Clientes
''' </summary>

Public Class ClientesGateway
    Private conexion As SqlConnection
    Private comando As SqlCommand
    Private adaptador As SqlDataAdapter
    Private generador As SqlCommandBuilder


    ''' <summary>
    ''' Constructor: crea y configura los objetos de aceso a la base de datos
    ''' </summary>
    Public Sub New(ByRef cadenaConexion As String)
        conexion = New SqlConnection(cadenaConexion)
        comando = New SqlCommand()
        comando.Connection = conexion
        adaptador = New SqlDataAdapter("SELECT * FROM Clientes", conexion)
        generador = New SqlCommandBuilder(adaptador)
    End Sub

    ''' <summary>
    ''' Método para insertar un nuevo registro en la base de datos
    ''' </summary>
    ''' <param name="id">Identificador del cliente</param>
    ''' <param name="nombre">Nombre del cliente</param>
    ''' <param name="edad">Edad del cliente</param>
    ''' <returns>Número de filas afectadas por la consulta</returns>
    Public Function Insertar(id As Integer, nombre As String, edad As Integer) As Integer

        Dim filas As Integer
        'Creamos la sentencia SQL de inserción
        Dim consulta As String = "INSERT INTO Clientes(id,nombre,edad) VALUES (@id,@nombre,@edad)"
        comando.CommandText = consulta

        comando.Parameters.Add("@id", SqlDbType.Int)
        comando.Parameters.Add("@nombre", SqlDbType.NVarChar)
        comando.Parameters.Add("@edad", SqlDbType.Int)

        comando.Parameters("@id").Value = id
        comando.Parameters("@nombre").Value = nombre
        comando.Parameters("@edad").Value = edad

        'Ejecutamos la consulta
        Try
            conexion.Open()
            filas = comando.ExecuteNonQuery()
        Catch ex As Exception
            Throw New ClientesGatewayException("Error al ejecutar la inserción en base de datos", ex)
        Finally
            If (conexion IsNot Nothing) Then
                conexion.Close()
            End If
        End Try

        'Devolvemos el número de filas afectadas
        Return filas

    End Function

    ''' <summary>
    ''' Método para actualizar un registro en la base de datos
    ''' </summary>
    ''' <param name="id">Identificador del cliente</param>
    ''' <param name="nombre">Nombre del cliente</param>
    ''' <param name="edad">Edad del cliente</param>
    ''' <returns>Número de filas afectadas por la consulta</returns>
    Public Function Actualizar(id As Integer, nombre As String, edad As Integer) As Integer

        Dim filas As Integer
        'Creamos la sentencia SQL de inserción
        Dim consulta As String = "UPDATE Clientes SET nombre=@nombre, edad=@edad WHERE id=@id"

        comando.Parameters.Add("@id", SqlDbType.Int)
        comando.Parameters.Add("@nombre", SqlDbType.NVarChar)
        comando.Parameters.Add("@edad", SqlDbType.Int)

        comando.Parameters("@id").Value = id
        comando.Parameters("@nombre").Value = nombre
        comando.Parameters("@edad").Value = edad


        'Ejecutamos la consulta
        Try
            conexion.Open()
            comando.CommandText = consulta
            filas = comando.ExecuteNonQuery()
        Catch ex As Exception
            Throw New ClientesGatewayException("Error al ejecutar la actualización en base de datos", ex)
        Finally
            If (conexion IsNot Nothing) Then
                conexion.Close()
            End If
        End Try

        'Devolvemos el número de filas afectadas
        Return filas

    End Function

    ''' <summary>
    ''' Método para eliminar un registro de la base de datos
    ''' </summary>
    ''' <param name="id">Identificador del registro a eliminar</param>
    ''' <returns>Número de registros eliminados</returns>
    Public Function Eliminar(id As Integer) As Integer

        Dim filas As Integer
        'Creamos la sentencia SQL de inserción
        Dim consulta As String = "DELETE FROM Clientes WHERE id=@id"

        comando.Parameters.Add("@id", SqlDbType.Int)

        comando.Parameters("@id").Value = id

        'Ejecutamos la consulta
        Try
            conexion.Open()
            comando.CommandText = consulta
            filas = comando.ExecuteNonQuery()
        Catch ex As Exception
            Throw New ClientesGatewayException("Error al ejecutar el borrado en base de datos", ex)
        Finally
            If (conexion IsNot Nothing) Then
                conexion.Close()
            End If
        End Try

        'Devolvemos el número de filas afectadas
        Return filas

    End Function

    ''' <summary>
    ''' Método para seleccionar todos los registros de la tabla
    ''' </summary>
    ''' <returns>Un objeto DataTable con todos los registros</returns>
    Public Function SeleccionarTodos() As DataTable
        Dim resultado As New DataSet

        'Ejecutamos la consulta
        Try
            adaptador.Fill(resultado, "Clientes")
            Return resultado.Tables("Clientes")
        Catch ex As Exception
            Throw New ClientesGatewayException("Error al consultar la base de datos", ex)
        Finally
            If (conexion IsNot Nothing) Then
                conexion.Close()
            End If
        End Try

    End Function

    ''' <summary>
    ''' Método para seleccionar un registro concreto de la tabla por id
    ''' </summary>
    ''' <param name="id">Identificador del registro</param>
    ''' <returns>Un objeto DataTable con el registro seleccionado</returns>
    Public Function SeleccionarId(id As Integer) As DataTable
        'Creamos la sentencia SQL de selección
        Dim consulta As String = "SELECT * FROM Clientes WHERE id=@id"
        comando.CommandText = consulta

        comando.Parameters.Add("@id", SqlDbType.Int)

        comando.Parameters("@id").Value = id


        Dim resultado As New DataTable
        Dim lector As SqlDataReader

        'Ejecutamos la consulta
        Try
            conexion.Open()
            lector = comando.ExecuteReader()

            'Cargamos el DataTable
            resultado.Load(lector)
        Catch ex As Exception
            Throw New ClientesGatewayException("Error al consultar la base de datos", ex)
        Finally
            If (conexion IsNot Nothing) Then
                conexion.Close()
            End If
        End Try

        'Devolvemos el resultado
        Return resultado

    End Function

    ''' <summary>
    ''' Método para seleccionar un registro concreto de la tabla por nombre
    ''' </summary>
    ''' <param name="nombre">Nombre del cliente</param>
    ''' <returns>Un objeto DataTable con los registros seleccionado</returns>
    Public Function SeleccionarNombre(nombre As String) As DataTable
        'Creamos la sentencia SQL de selección
        Dim consulta As String = "SELECT * FROM Clientes WHERE nombre=@nombre"
        comando.CommandText = consulta

        comando.Parameters.Add("@nombre", SqlDbType.NVarChar)

        comando.Parameters("@nombre").Value = nombre


        Dim resultado As New DataTable
        Dim lector As SqlDataReader

        'Ejecutamos la consulta
        Try
            conexion.Open()
            lector = comando.ExecuteReader()

            'Cargamos el DataTable
            resultado.Load(lector)
        Catch ex As Exception
            Throw New ClientesGatewayException("Error al consultar la base de datos", ex)
        Finally
            If (conexion IsNot Nothing) Then
                conexion.Close()
            End If
        End Try

        'Devolvemos el resultado
        Return resultado

    End Function


    ''' <summary>
    ''' Método para actualizar la base de datos a partir de un DataTable
    ''' </summary>
    ''' <param name="tabla">DataTable origen de la actualización</param>
    ''' <returns>Número de filas afectadas</returns>
    Public Function ActualizarTabla(tabla As DataTable) As Integer
        'Ejecutamos la consulta
        Try
            Return adaptador.Update(tabla)
        Catch ex As Exception
            Throw New ClientesGatewayException("Error al actualizar la base de datos", ex)
        Finally
            If (conexion IsNot Nothing) Then
                conexion.Close()
            End If
        End Try

    End Function

End Class


