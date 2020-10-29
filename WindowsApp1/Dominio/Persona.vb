Public Class Persona
    ' Atributos de la clase persona '
    Private mDNI As String
    Private mNombre As String
    Private Gestor As gestorPersona

    ' Constructor '
    Sub New(ByVal dni As String, ByVal nombre As String)
        Me.mDNI = dni
        Me.mNombre = nombre
        Gestor = New gestorPersona()
    End Sub

    Sub New()
        Gestor = New gestorPersona()
    End Sub

    ' Property DNI '
    Property DNI As String
        Get
            Return mDNI
        End Get
        Set(ByVal mDNI As String)
            Me.mDNI = mDNI
        End Set
    End Property

    ' Property Nombre '
    Property Nombre As String
        Get
            Return mNombre
        End Get
        Set(ByVal mNombre As String)
            Me.mNombre = mNombre
        End Set
    End Property

    ' Metodo leer Persona '
    Sub leerPersona(ByVal DNI As String)
        Gestor.read(DNI, Me)
    End Sub

    ' Funcion leer Personas '
    Function leertodos()
        Dim todos As Collection
        Me.Gestor.readAll()
        todos = Gestor.listaPersonas
        Return todos
    End Function

    ' Metodo leer Persona '
    Sub insertarPersona()
        Me.Gestor.insert(Me)
    End Sub

    ' Metodo modificar Persona '
    Sub modificarPersona()
        Me.Gestor.update(Me)
    End Sub

    ' Metodo eliminar Persona '
    Sub eliminarPersona()
        Me.Gestor.delete(Me)
    End Sub

    ' Metodo elimiar Personas '
    Sub eliminarTodo()
        Me.Gestor.deleteAll()
    End Sub
End Class
