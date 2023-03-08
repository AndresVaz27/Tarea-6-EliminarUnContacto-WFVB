Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace Tarea_6_EliminarUnContacto_WFVB.Tarea_2___Segundo_Semestre
    Friend Class Persona
        Protected nombre_ As String
        Protected fechaNacimiento_ As DateTime?


        Public Property Nombre As String
            Get
                Return nombre_
            End Get
            Set(ByVal value As String)
                nombre_ = value
            End Set
        End Property

        Public Property FechaNacimiento As DateTime?
            Get
                Return fechaNacimiento_
            End Get
            Set(ByVal value As DateTime?)
                fechaNacimiento_ = value
            End Set
        End Property

        Private edad_ As String

        Public Property Edad As String
            Set(ByVal value As String)
                edad_ = value
            End Set
            Get

                If FechaNacimiento.HasValue Then
                    Dim edad_ As Integer = DateTime.Now.Year - FechaNacimiento.Value.Year

                    If FechaNacimiento.Value.Month > DateTime.Now.Month OrElse (FechaNacimiento.Value.Month = DateTime.Now.Month AndAlso FechaNacimiento.Value.Day > DateTime.Now.Day) Then
                        edad_ -= 1
                    End If

                    Return edad_.ToString()
                Else
                    Return edad_
                End If
            End Get
        End Property

        Public Sub New()
            Nombre = String.Empty
            FechaNacimiento = Nothing
        End Sub

        Public Sub New(ByVal nombre As String, ByVal fechaNacimiento As DateTime?)
            Me.Nombre = nombre
            Me.FechaNacimiento = fechaNacimiento
        End Sub

        Public Overrides Function ToString() As String
            Return "NOMBRE: " & Nombre.ToUpper() & " " & "---" & " EDAD: " & Edad
        End Function
    End Class
End Namespace