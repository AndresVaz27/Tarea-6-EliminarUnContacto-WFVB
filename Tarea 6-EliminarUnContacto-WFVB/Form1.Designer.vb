Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Tarea_6_EliminarUnContacto_WFVB.Form1
Imports Tarea_6_EliminarUnContacto_WFVB.Tarea_2___Segundo_Semestre
Imports Tarea_6_EliminarUnContacto_WFVB.Tarea_6_EliminarUnContacto_WFVB.Form1
Imports Tarea_6_EliminarUnContacto_WFVB.Tarea_6_EliminarUnContacto_WFVB.Tarea_2___Segundo_Semestre

Namespace Tarea_4___Arreglos___Segundo_Semestre___WF_CS
    Partial Public Class Form1
        Inherits Form
        Private nombreArchivo As String
        Private cantidadContactos As Integer
        Private siguienteIndex As Integer = 0
        Private contactos As Contacto()
        Private nuevo As Contacto = New Contacto()

        Public Sub New()
            InitializeComponent()
            nombreArchivo = "contactos.txt"

            If Not File.Exists(nombreArchivo) Then
                File.Create(nombreArchivo).Dispose()
            Else

                Using sr As StreamReader = File.OpenText(nombreArchivo)
                    Dim linea As String
                    Dim encabezados As String = sr.ReadLine()
                    Dim encabezadosArray As String() = encabezados.Split(","c)
                    Dim columna As DataGridViewColumn = CSharpImpl.__Assign(columna, New DataGridViewTextBoxColumn())

                    For Each s As String In encabezadosArray
                        columna.HeaderText = s
                    Next

                    While (CSharpImpl.__Assign(linea, sr.ReadLine())) IsNot Nothing
                        Dim nuevo As Contacto = New Contacto()
                        Dim datos As String() = linea.Split(","c)

                        If datos.Length = encabezadosArray.Length Then
                            nuevo.Nombre = datos(0)
                            nuevo.Edad = datos(1)
                            nuevo.Telefono = datos(2)
                            nuevo.Correo = datos(3)
                            tabla.Rows.Add(nuevo.Nombre, nuevo.Edad, nuevo.Telefono, nuevo.Correo)
                        Else
                            MessageBox.Show("Elementos distintos.")
                        End If
                    End While

                    sr.Close()
                End Using
            End If
        End Sub

        Private Sub btnGuardar_Click(ByVal sender As Object, ByVal e As EventArgs)

        End Sub

        Private Sub btnEstablecer_Click(ByVal sender As Object, ByVal e As EventArgs)

        End Sub

        Private Sub tabla_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
            If e.KeyCode = Keys.Back OrElse e.KeyCode = Keys.Delete Then

                For Each rows As DataGridViewRow In tabla.SelectedRows
                    tabla.Rows.RemoveAt(rows.Index)
                Next
            End If
        End Sub

        Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs)

        End Sub

        Friend WithEvents btnEstablecer As Windows.Forms.Button
        Friend WithEvents btnGuardar As Windows.Forms.Button
        Friend WithEvents txtArregloLenght As Windows.Forms.TextBox
        Friend WithEvents lblArregloLenght As Label
        Friend WithEvents tabla As DataGridView

        Private Class CSharpImpl
            <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
            Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class

        Private Sub InitializeComponent()
            Me.btnEstablecer = New System.Windows.Forms.Button()
            Me.btnGuardar = New System.Windows.Forms.Button()
            Me.txtArregloLenght = New System.Windows.Forms.TextBox()
            Me.lblArregloLenght = New System.Windows.Forms.Label()
            Me.tabla = New System.Windows.Forms.DataGridView()
            Me.Nombre = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.Edad = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.Telefono = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.Correo = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.lblNombre = New System.Windows.Forms.Label()
            Me.lbl = New System.Windows.Forms.Label()
            Me.Label4 = New System.Windows.Forms.Label()
            Me.Label5 = New System.Windows.Forms.Label()
            Me.txtNombre = New System.Windows.Forms.TextBox()
            Me.txtFecha = New System.Windows.Forms.TextBox()
            Me.txtTelefono = New System.Windows.Forms.TextBox()
            Me.txtCorreo = New System.Windows.Forms.TextBox()
            CType(Me.tabla, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'btnEstablecer
            '
            Me.btnEstablecer.Location = New System.Drawing.Point(480, 30)
            Me.btnEstablecer.Name = "btnEstablecer"
            Me.btnEstablecer.Size = New System.Drawing.Size(101, 32)
            Me.btnEstablecer.TabIndex = 0
            Me.btnEstablecer.Text = "Establecer"
            Me.btnEstablecer.UseVisualStyleBackColor = True
            '
            'btnGuardar
            '
            Me.btnGuardar.Location = New System.Drawing.Point(480, 224)
            Me.btnGuardar.Name = "btnGuardar"
            Me.btnGuardar.Size = New System.Drawing.Size(101, 34)
            Me.btnGuardar.TabIndex = 1
            Me.btnGuardar.Text = "Guardar"
            Me.btnGuardar.UseVisualStyleBackColor = True
            '
            'txtArregloLenght
            '
            Me.txtArregloLenght.Location = New System.Drawing.Point(282, 31)
            Me.txtArregloLenght.Name = "txtArregloLenght"
            Me.txtArregloLenght.Size = New System.Drawing.Size(100, 22)
            Me.txtArregloLenght.TabIndex = 2
            '
            'lblArregloLenght
            '
            Me.lblArregloLenght.AutoSize = True
            Me.lblArregloLenght.Location = New System.Drawing.Point(61, 34)
            Me.lblArregloLenght.Name = "lblArregloLenght"
            Me.lblArregloLenght.Size = New System.Drawing.Size(215, 16)
            Me.lblArregloLenght.TabIndex = 3
            Me.lblArregloLenght.Text = "Número de Contactos por Agregar:"
            '
            'tabla
            '
            Me.tabla.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.tabla.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Nombre, Me.Edad, Me.Telefono, Me.Correo})
            Me.tabla.Location = New System.Drawing.Point(12, 331)
            Me.tabla.Name = "tabla"
            Me.tabla.RowHeadersWidth = 51
            Me.tabla.RowTemplate.Height = 24
            Me.tabla.Size = New System.Drawing.Size(569, 243)
            Me.tabla.TabIndex = 4
            '
            'Nombre
            '
            Me.Nombre.HeaderText = "Nombre"
            Me.Nombre.MinimumWidth = 6
            Me.Nombre.Name = "Nombre"
            Me.Nombre.Width = 125
            '
            'Edad
            '
            Me.Edad.HeaderText = "Edad"
            Me.Edad.MinimumWidth = 6
            Me.Edad.Name = "Edad"
            Me.Edad.Width = 125
            '
            'Telefono
            '
            Me.Telefono.HeaderText = "Telefono"
            Me.Telefono.MinimumWidth = 6
            Me.Telefono.Name = "Telefono"
            Me.Telefono.Width = 125
            '
            'Correo
            '
            Me.Correo.HeaderText = "Correo"
            Me.Correo.MinimumWidth = 6
            Me.Correo.Name = "Correo"
            Me.Correo.Width = 125
            '
            'lblNombre
            '
            Me.lblNombre.AutoSize = True
            Me.lblNombre.Location = New System.Drawing.Point(61, 88)
            Me.lblNombre.Name = "lblNombre"
            Me.lblNombre.Size = New System.Drawing.Size(59, 16)
            Me.lblNombre.TabIndex = 5
            Me.lblNombre.Text = "Nombre:"
            '
            'lbl
            '
            Me.lbl.AutoSize = True
            Me.lbl.Location = New System.Drawing.Point(61, 135)
            Me.lbl.Name = "lbl"
            Me.lbl.Size = New System.Drawing.Size(138, 16)
            Me.lbl.TabIndex = 6
            Me.lbl.Text = "Fecha de Nacimiento:"
            '
            'Label4
            '
            Me.Label4.AutoSize = True
            Me.Label4.Location = New System.Drawing.Point(61, 178)
            Me.Label4.Name = "Label4"
            Me.Label4.Size = New System.Drawing.Size(64, 16)
            Me.Label4.TabIndex = 7
            Me.Label4.Text = "Telefono:"
            '
            'Label5
            '
            Me.Label5.AutoSize = True
            Me.Label5.Location = New System.Drawing.Point(61, 224)
            Me.Label5.Name = "Label5"
            Me.Label5.Size = New System.Drawing.Size(51, 16)
            Me.Label5.TabIndex = 8
            Me.Label5.Text = "Correo:"
            '
            'txtNombre
            '
            Me.txtNombre.Location = New System.Drawing.Point(241, 88)
            Me.txtNombre.Name = "txtNombre"
            Me.txtNombre.Size = New System.Drawing.Size(100, 22)
            Me.txtNombre.TabIndex = 9
            '
            'txtFecha
            '
            Me.txtFecha.Location = New System.Drawing.Point(241, 135)
            Me.txtFecha.Name = "txtFecha"
            Me.txtFecha.Size = New System.Drawing.Size(100, 22)
            Me.txtFecha.TabIndex = 10
            '
            'txtTelefono
            '
            Me.txtTelefono.Location = New System.Drawing.Point(241, 175)
            Me.txtTelefono.Name = "txtTelefono"
            Me.txtTelefono.Size = New System.Drawing.Size(100, 22)
            Me.txtTelefono.TabIndex = 11
            '
            'txtCorreo
            '
            Me.txtCorreo.Location = New System.Drawing.Point(241, 224)
            Me.txtCorreo.Name = "txtCorreo"
            Me.txtCorreo.Size = New System.Drawing.Size(100, 22)
            Me.txtCorreo.TabIndex = 12
            '
            'Form1
            '
            Me.ClientSize = New System.Drawing.Size(593, 586)
            Me.Controls.Add(Me.txtCorreo)
            Me.Controls.Add(Me.txtTelefono)
            Me.Controls.Add(Me.txtFecha)
            Me.Controls.Add(Me.txtNombre)
            Me.Controls.Add(Me.Label5)
            Me.Controls.Add(Me.Label4)
            Me.Controls.Add(Me.lbl)
            Me.Controls.Add(Me.lblNombre)
            Me.Controls.Add(Me.tabla)
            Me.Controls.Add(Me.lblArregloLenght)
            Me.Controls.Add(Me.txtArregloLenght)
            Me.Controls.Add(Me.btnGuardar)
            Me.Controls.Add(Me.btnEstablecer)
            Me.Name = "Form1"
            CType(Me.tabla, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub

        Friend WithEvents lblNombre As Label
        Friend WithEvents lbl As Label
        Friend WithEvents Label4 As Label
        Friend WithEvents Label5 As Label
        Friend WithEvents txtNombre As Windows.Forms.TextBox
        Friend WithEvents txtFecha As Windows.Forms.TextBox
        Friend WithEvents txtTelefono As Windows.Forms.TextBox
        Friend WithEvents txtCorreo As Windows.Forms.TextBox
        Friend WithEvents Nombre As DataGridViewTextBoxColumn
        Friend WithEvents Edad As DataGridViewTextBoxColumn
        Friend WithEvents Telefono As DataGridViewTextBoxColumn
        Friend WithEvents Correo As DataGridViewTextBoxColumn

        Private Sub btnEstablecer_Click_1(sender As Object, e As EventArgs) Handles btnEstablecer.Click
            Try
                cantidadContactos = Integer.Parse(txtArregloLenght.Text)
            Catch ex As Exception
                MessageBox.Show("Únicamente números enteros para establecer la cantidad de contactos por agregar, por favor")
            End Try

            contactos = New Contacto(cantidadContactos - 1) {}
        End Sub

        Private Sub btnGuardar_Click_1(sender As Object, e As EventArgs) Handles btnGuardar.Click
            If contactos Is Nothing Then
                MessageBox.Show("Establezca el número de contactos por agregar.")
            ElseIf siguienteIndex >= contactos.Length Then
                MessageBox.Show("La cantidad de contactos por agregar se completó.")
            Else
                nuevo.Nombre = txtNombre.Text
                nuevo.Telefono = txtTelefono.Text
                Dim fechaNacimiento As DateTime

                If DateTime.TryParseExact(txtFecha.Text, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, fechaNacimiento) Then
                    nuevo.FechaNacimiento = fechaNacimiento
                Else
                    MessageBox.Show("La fecha de nacimiento debe tener el formato dd/MM/yyyy")
                    Return
                End If
                nuevo.Correo = txtCorreo.Text
                contactos(siguienteIndex) = nuevo
                tabla.Rows.Add(nuevo.Nombre, nuevo.Edad, nuevo.Telefono, nuevo.Correo)
                txtNombre.Clear()
                txtFecha.Clear()
                txtTelefono.Clear()
                txtCorreo.Clear()
                siguienteIndex += 1
            End If
        End Sub

        Private Sub tabla_KeyDown_1(sender As Object, e As KeyEventArgs) Handles tabla.KeyDown
            If e.KeyCode = Keys.Back OrElse e.KeyCode = Keys.Delete Then

                For Each rows As DataGridViewRow In tabla.SelectedRows
                    tabla.Rows.RemoveAt(rows.Index)
                Next
            End If
        End Sub

        Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        End Sub

        Private Sub Form1_FormClosing_1(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
            Using sw As StreamWriter = New StreamWriter(nombreArchivo)
                sw.WriteLine("Nombre,Edad,Telefono,Correo")
                For Each fila As DataGridViewRow In tabla.Rows
                    Dim columnas As String() = New String(tabla.Columns.Count - 1) {}

                    For i As Integer = 0 To tabla.Columns.Count - 1

                        If fila IsNot Nothing AndAlso fila.Cells(i).Value IsNot Nothing Then
                            columnas(i) = fila.Cells(i).Value.ToString()
                        Else
                            columnas(i) = ""
                        End If
                    Next

                    If columnas(tabla.Columns.Count - 1) = "" Then
                        Exit For
                    Else
                        sw.WriteLine(String.Join(",", columnas))
                    End If
                Next

                sw.Close()
            End Using
        End Sub
    End Class
End Namespace
