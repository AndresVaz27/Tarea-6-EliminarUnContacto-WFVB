Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace Tarea_6_EliminarUnContacto_WFVB.Tarea_2___Segundo_Semestre
    Friend Class Contacto
        Inherits Persona

        Private telefono_ As String
        Private correo_ As String

        Public Property Correo As String
            Get
                Return correo_
            End Get
            Set(ByVal value As String)
                correo_ = value.Replace(" ", "").ToLower()
            End Set
        End Property

        Public Property Telefono As String
            Get
                Return telefono_
            End Get
            Set(ByVal value As String)
                telefono_ = value.Replace(" ", "").Replace("-", "")
            End Set
        End Property

        Public Sub New()
            MyBase.New()
            Telefono = String.Empty
        End Sub

        Public Sub New(ByVal nombre As String, ByVal fechaNacimiento As DateTime?, ByVal telefono As String)
            MyBase.New(nombre, fechaNacimiento)
            Me.Telefono = telefono
        End Sub

        Public Overrides Function ToString() As String
            Return MyBase.ToString() & " " & "---" & " TELEFONO: " & Telefono & " " & "---" & " CORREO: " & Correo.ToUpper()
        End Function
    End Class
End Namespace