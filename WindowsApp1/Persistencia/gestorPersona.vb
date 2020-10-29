Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class gestorPersona

    ' Atributo de la clase gestorPersona ' 
    Private mlistaPersonas As Collection
    Private Agente As AgenteBD

    ' Constructor '
    Sub New()
        mlistaPersonas = New Collection
    End Sub

    ' Property listaPersonas '
    Public Property listaPersonas As Collection
        Get
            Return mlistaPersonas
        End Get
        Set(ByVal value As Collection)
            mlistaPersonas = value
        End Set
    End Property

    ' Lee la persona que tenga el DNI escogido entre sus atributos '
    Sub read(ByVal dni As String, ByRef per As Persona)
        Me.Agente = AgenteBD.getInstancia()
        Dim sSQL As String = "SELECT * FROM Personas WHERE DNI='" + dni + "'"
        Dim reader As OleDbDataReader
        reader = Me.Agente.read(sSQL)
        If reader.HasRows Then
            reader.Read()
            per.DNI = reader(0)
            per.Nombre = reader(1)
        End If
    End Sub

    ' Lee todas las personas de la base de datos '
       Sub readAll()
        Me.Agente = AgenteBD.getInstancia()
        Dim sSQL As String = "SELECT * FROM Personas ORDER BY DNI"
        Dim reader As OleDbDataReader
        reader = Me.Agente.read(sSQL)
        While reader.Read()
            Me.listaPersonas.Add(New Persona(reader.GetString(0), reader.GetString(1)))
        End While
    End Sub

    ' Inserta una persona en la base de datos '
    Sub insert(ByVal newPersona As Persona)
        Agente = AgenteBD.getInstancia()
        Dim sSQL As String = "INSERT INTO Personas VALUES ('" + newPersona.DNI + "', '" + newPersona.Nombre + "')"
        Agente.create(sSQL)
    End Sub

    ' Modifica la persona escogida por su DNI de la base de datos '
    Sub update(ByVal modPersona As Persona)
        Agente = AgenteBD.getInstancia()
        Dim sSQL As String = "UPDATE Personas SET Nombre='" + modPersona.Nombre + "'WHERE DNI='" + modPersona.DNI + "'"
        Agente.update(sSQL)
    End Sub

    ' Borra la persona escogida de la base de datos '
    Sub delete(ByVal noPersona As Persona)
        Agente = AgenteBD.getInstancia()
        Dim sSQL As String = "DELETE FROM Personas WHERE DNI='" + noPersona.DNI + "'"
        Agente.delete(sSQL)
    End Sub

    ' Borra todas las personas de la base de datos '
    Sub deleteAll()
        Agente = AgenteBD.getInstancia()
        Dim sSQL As String = "DELETE * FROM PERSONAS"
        Agente.delete(sSQL)
    End Sub
End Class
