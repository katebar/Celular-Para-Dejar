Public Class Form1

    Dim modulo As Integer = 11
    Dim tipo As Integer = 2
    Dim seleccion As Integer

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        seleccion = 1
        cualva()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        seleccion = 2
        cualva()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        seleccion = 3
        cualva()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        seleccion = 4
        cualva()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        seleccion = 5
        cualva()
    End Sub

    Public Sub cualva()
        Select Case modulo
            Case 1
                modulo1(tipo, seleccion)
            Case 2
                modulo2(tipo, seleccion)
            Case 3
                modulo3(tipo, seleccion)
            Case 4
                modulo4(tipo, seleccion)
            Case 5
                modulo5(tipo, seleccion)
            Case 6
                modulo6(tipo, seleccion)
            Case 7
                modulo7(tipo, seleccion)
            Case 8
                modulo8(tipo, seleccion)
            Case 9
                modulo9(tipo, seleccion)
            Case 10
                modulo10(tipo, seleccion)
            Case 11
                modulo11(tipo, seleccion)
            Case 12
                modulo12(tipo, seleccion)
            Case 13
                modulo13(tipo, seleccion)
        End Select
    End Sub

    Public Sub modulo1(tip, sel)
        If tip = 1 Then
            TextBox1.Text = "Con referencia a su salud en general, como se ha sentido esta semana? Si piensa que su salud es excelente, oprima 1. Si piensa que su salud es buena, oprima 2. Si piensa que su salud es deficiente, oprima 3."
            TextBox1.Update()
            My.Computer.Audio.Play(My.Resources.q1, AudioPlayMode.WaitToComplete)
            'MsgBox("Como se siente usted el dia de hoy")
            modulo = 2
            'tipo = 2

            'ElseIf tip = 2 Then
            '    If sel = 1 Then
            '        My.Computer.Audio.Play(My.Resources.alegre, AudioPlayMode.WaitToComplete)
            '        MsgBox("Me alegra que se sienta mejor")
            '        tipo = 1
            '        modulo = 2
            '        cualva()
            '    ElseIf sel = 2 Then
            '        MsgBox("Me alegra que se sienta bien, sin embargo hay varias cosas que podría hacer para sentirse mejor")
            '        tipo = 1
            '        modulo = 2
            '        cualva()
            '    ElseIf sel = 3 Then
            '        MsgBox("Siento escuchar eso, seria bueno que contactara a su medico")
            '        tipo = 1
            '        modulo = 2
            '        cualva()
            '    End If
        End If
    End Sub

    Public Sub modulo2(tip, sel)
        If tip = 1 Then
            TextBox1.Text = "Con referencia al estado general de su salud  esta semana, cómo la compararía a la semana pasada? Si piensa que su salud es la misma, oprima 1. Si piensa que su salud es mejor, oprima 2. Si piensa que su salud es peor, oprima 3."
            TextBox1.Update()
            My.Computer.Audio.Play(My.Resources.q2, AudioPlayMode.WaitToComplete)
            'MsgBox("Como se siente usted en comparacion con la semana pasada2")
            tipo = 2

        ElseIf tip = 2 Then
            If sel = 1 Then
                TextBox1.Text = "Nos alegra que su estado de salud no haya empeorado. Continue la misma rutina."
                TextBox1.Update()
                My.Computer.Audio.Play(My.Resources.r1, AudioPlayMode.WaitToComplete)
                'MsgBox("Me alegra que se sienta mejor2")
                tipo = 1
                modulo = 3
                cualva()
            ElseIf sel = 2 Then
                TextBox1.Text = "Nos alegra que su estado de salud haya mejorado esta semana. Continue la misma rutina."
                TextBox1.Update()
                My.Computer.Audio.Play(My.Resources.r2, AudioPlayMode.WaitToComplete)
                'MsgBox("Me alegra que se sienta bien, sin embargo hay varias cosas que podría hacer para sentirse mejor2")
                tipo = 1
                modulo = 3
                cualva()
            ElseIf sel = 3 Then
                TextBox1.Text = "Sentimos mucho que usted se esté sintiendo peor esta semana. pero estamos aquí para apoyarle. Tenemos algunas preguntas más para usted."
                TextBox1.Update()
                My.Computer.Audio.Play(My.Resources.r3, AudioPlayMode.WaitToComplete)
                'MsgBox("Siento escuchar eso, seria bueno que contactara a su medico2")
                tipo = 1
                modulo = 4
                cualva()
            End If
        End If
    End Sub

    Public Sub modulo3(tip, sel)
        If tip = 1 Then
            TextBox1.Text = "Por esta semana, cuántos cigarrillos ha fumado cada día? Si este número es diferente cada día, piense en lo general. Si no ha fumado ningún cigarillo esta semana, oprima 1. Si ha fumado entre 1 a 3 cigarillos cada día de esta semana, oprima 2. Si ha fumado entre 4 a 10 cigarillos cada día de esta semana, oprima 3. Si ha fumado entre 11 a 20 cigarillos cada día de esta semana, oprima 4. Si ha fumado más de 20 cigarillos cada día de esta semana, oprima 5"
            TextBox1.Update()
            My.Computer.Audio.Play(My.Resources.q5, AudioPlayMode.WaitToComplete)
            tipo = 2

        ElseIf tip = 2 Then
            If sel = 1 Then
                TextBox1.Text = "Basado en su respuesta, nuestro sistema registra que ha fumado lo mismo que la semana pasada. Es bueno que no haya fumado más esta semana. Recuerde siempre la fecha que usted decidio para dejar de fumar y trate de disminuir el número de cigarillos cada semana. "
                TextBox1.Update()
                My.Computer.Audio.Play(My.Resources.r5, AudioPlayMode.WaitToComplete)
                tipo = 1
                modulo = 6
                cualva()
            ElseIf sel = 2 Then
                TextBox1.Text = "Basado en su respuesta, nuestro sistema registra que ha fumado lo mismo que la semana pasada. Es bueno que no haya fumado más esta semana. Recuerde siempre la fecha que usted decidio para dejar de fumar y trate de disminuir el número de cigarillos cada semana. "
                TextBox1.Update()
                My.Computer.Audio.Play(My.Resources.r5, AudioPlayMode.WaitToComplete)
                tipo = 1
                modulo = 6
                cualva()
            ElseIf sel = 3 Then
                TextBox1.Text = " Basado en su respuesta, nuestro sistema registra que ha fumado menos esta semana en comparación a la semana pasada. Esto es excelente, esperamos que pueda continuar disminuyendo el número de cigarillos cada semana antes de la fecha que usted decidio dejar de fumar."
                TextBox1.Update()
                My.Computer.Audio.Play(My.Resources.r6, AudioPlayMode.WaitToComplete)
                tipo = 1
                modulo = 6
                cualva()
            ElseIf sel = 4 Then
                TextBox1.Text = "Basado en su respuesta, nuestro sistema registra que ha fumado más cigarillos esta semana en comparación con la semana pasada. Sabemos que dejar de fumar es muy difícil, pero sería más fácil si disminuye el número de cigarillos poco a poco antes de la fecha que usted decidio dejar de fumar."
                TextBox1.Update()
                My.Computer.Audio.Play(My.Resources.r7, AudioPlayMode.WaitToComplete)
                tipo = 1
                modulo = 6
                cualva()
            ElseIf sel = 5 Then
                TextBox1.Text = "Basado en su respuesta, nuestro sistema registra que ha fumado más cigarillos esta semana en comparación con la semana pasada. Sabemos que dejar de fumar es muy difícil, pero sería más fácil si disminuye el número de cigarillos poco a poco antes de la fecha que usted decidio dejar de fumar."
                TextBox1.Update()
                My.Computer.Audio.Play(My.Resources.r7, AudioPlayMode.WaitToComplete)
                tipo = 1
                modulo = 6
                cualva()
            End If
        End If
    End Sub

    Public Sub modulo4(tip, sel)
        If tip = 1 Then
            TextBox1.Text = "Piensa que su salud es peor por causa de las dificultades presentadas al dejar de fumar o es  causado por algo diferente, como su salud física o mental? Oprima 1 si piensa que se siente peor por causa de dejar de fumar. Oprima 2 si piensa que se siente peor por causa de su salud física o mental. Oprima 3 si piensa que se siente peor por causa de dejar de fumar y por su salud física o mental."
            TextBox1.Update()
            My.Computer.Audio.Play(My.Resources.q3, AudioPlayMode.WaitToComplete)
            tipo = 2

        ElseIf tip = 2 Then
            If sel = 1 Then
                tipo = 1
                modulo = 3
                cualva()
            ElseIf sel = 2 Then
                tipo = 1
                modulo = 5
                cualva()
            ElseIf sel = 3 Then
                tipo = 1
                modulo = 5
                cualva()
            End If
        End If
    End Sub

    Public Sub modulo5(tip, sel)
        If tip = 1 Then
            TextBox1.Text = "En la semana pasada, ha contactado a su médico por su estado de salud? Si fue a su médico y recibio tratamiento específico, como medicamentos para una infección, oprima 1.
Si pidio una cita con su médico a causa de este problema, pero aun no ha asistido a esta cita, oprima 2.
Si decidio que no va a ir al médico porque cree que no necesite ayuda, oprima 3.
Si decidio que no va a ir al médico por otra razón, oprima 4."
            TextBox1.Update()
            My.Computer.Audio.Play(My.Resources.q4, AudioPlayMode.WaitToComplete)
            tipo = 2
        ElseIf tip = 2 Then
            If sel = 1 Then
                tipo = 1
                modulo = 3
                cualva()
            ElseIf sel = 2 Then
                tipo = 1
                modulo = 3
                cualva()
            ElseIf sel = 3 Then
                TextBox1.Text = "Es importante que se cuide. Si se siente peor, debe consultar a su médico."
                TextBox1.Update()
                My.Computer.Audio.Play(My.Resources.r4, AudioPlayMode.WaitToComplete)
                tipo = 1
                modulo = 3
                cualva()
            ElseIf sel = 4 Then
                TextBox1.Text = "Es importante que se cuide. Si se siente peor, debe consultar a su médico."
                TextBox1.Update()
                My.Computer.Audio.Play(My.Resources.r4, AudioPlayMode.WaitToComplete)
                tipo = 1
                modulo = 3
                cualva()
            End If
        End If
    End Sub

    Public Sub modulo6(tip, sel)
        If tip = 1 Then
            TextBox1.Text = "Durante la semana pasada, ha experimentado algunos síntomas de abstinencia de nicotina? Estos síntomas pueden incluir deseo de fumar, ansiedad, aumento de apetito, insomnio, irritabilidad, dificultad en concentrarse, disminución de la frecuencia cardiaca y depresión. Oprima 1 si ha experimentado algunos de estos síntomas. Oprima 2 si no los ha experimentado."
            TextBox1.Update()
            My.Computer.Audio.Play(My.Resources.q6, AudioPlayMode.WaitToComplete)
            tipo = 2
        ElseIf tip = 2 Then
            If sel = 1 Then
                TextBox1.Text = "Siga leyendo los mensajes de texto diarios porque tienen muchos consejos sobre como controlar estos síntomas."
                TextBox1.Update()
                My.Computer.Audio.Play(My.Resources.r8, AudioPlayMode.WaitToComplete)
                tipo = 1
                modulo = 7
                cualva()
            ElseIf sel = 2 Then
                TextBox1.Text = "Siga leyendo los mensajes de texto diarios porque tienen muchos consejos sobre como controlar estos síntomas."
                TextBox1.Update()
                My.Computer.Audio.Play(My.Resources.r8, AudioPlayMode.WaitToComplete)
                tipo = 1
                modulo = 7
                cualva()

            End If
        End If

    End Sub

    Public Sub modulo7(tip, sel)
        If tip = 1 Then
            TextBox1.Text = "Algunas personas han experimentado problemas al dejar de fumar, a causa que no están listos para dejar de fumar, hay otras personas cerca de ellos que fuman o no creen que los métodos para dejar de fumar funcionen para ellos mismos. Ha experimentado algunos de estos sintomas? Oprima 1 si su respuesta es si. oprima 2 si su respuesta es no."
            TextBox1.Update()
            My.Computer.Audio.Play(My.Resources.q7, AudioPlayMode.WaitToComplete)
            tipo = 2
        ElseIf tip = 2 Then
            If sel = 1 Then
                TextBox1.Text = "Sentimos que experimente estos sintomas, pero muchas personas enfrentan estas barreras. Si quiere escuchar otros testimonios al fin de la llamada, oprima 1. Si quiere que una enfermera le llame y le oriente sobre opciones terapeuticas para ayudarle a dejar de fumar, oprima 2. Si quiere ambos, oprima 3. Si no quiere ninguno, oprima 4 y siga a la proxima pregunta"
                TextBox1.Update()
                My.Computer.Audio.Play(My.Resources.r9, AudioPlayMode.WaitToComplete)
                tipo = 1
                modulo = 10
                cualva()
            ElseIf sel = 2 Then
                tipo = 1
                modulo = 8
                cualva()

            End If
        End If
    End Sub

    Public Sub modulo8(tip, sel)
        If tip = 1 Then
            TextBox1.Text = "Durante la semana pasada, se siente motivado a dejar de fumar o reducir el numero de cigarillos antes de la fecha que usted decidio dejar de fumar? Oprima 1 si su respuesta es si. Oprima 2 si su respuesta es no."
            TextBox1.Update()
            My.Computer.Audio.Play(My.Resources.q8, AudioPlayMode.WaitToComplete)
            tipo = 2
        ElseIf tip = 2 Then
            If sel = 1 Then
                TextBox1.Text = "Nos alegra que se sienta motivado. Es un largo camino a seguir pero puede lograrlo!"
                TextBox1.Update()
                My.Computer.Audio.Play(My.Resources.r12, AudioPlayMode.WaitToComplete)
                tipo = 1
                modulo = 9
                cualva()
            ElseIf sel = 2 Then
                TextBox1.Text = "Sentimos que no se sienta motivado. Trate de leer los mensajes de texto cada día y recuerde que no está solo en este largo camino."
                TextBox1.Update()
                My.Computer.Audio.Play(My.Resources.r13, AudioPlayMode.WaitToComplete)
                tipo = 1
                modulo = 9
                cualva()
            End If
        End If

    End Sub

    Public Sub modulo9(tip, sel)
        If tip = 1 Then
            TextBox1.Text = "Ha completado la llamada esta semana. Gracias otra vez por su participación en este programa. Esperamos que le ayude a estar tan saludable como sea posible. Le llamarémos la próxima semana a la hora programada."
            TextBox1.Update()
            My.Computer.Audio.Play(My.Resources.g11, AudioPlayMode.WaitToComplete)
        End If
    End Sub

    Public Sub modulo10(tip, sel)
        If tip = 1 Then
            'MsgBox("I'm here")
            tipo = 2
        ElseIf tip = 2 Then
            If sel = 1 Then
                tipo = 1
                modulo = 8
                cualva()
            ElseIf sel = 2 Then
                tipo = 1
                modulo = 8
                cualva()
            ElseIf sel = 3 Then
                tipo = 1
                modulo = 8
                cualva()
            ElseIf sel = 4 Then
                tipo = 1
                modulo = 8
                cualva()
            End If
        End If

    End Sub

    Public Sub modulo11(tip, sel)
        If tip = 2 Then
            If sel = 1 Then
                TextBox1.Text = "Si usted es el / la participante por favor oprima 1. Si usted no es el participante, pero puede hacer que el se acerque al teléfono por favor oprima 2. Si el participante no se puede acercar en este momento oprima 3. Si usted ha recibido este mensaje 
por error o si esta persona no se encuentra en este numero de teléfono, por favor oprima 4. "
                TextBox1.Update()
                My.Computer.Audio.Play(My.Resources.g3, AudioPlayMode.WaitToComplete)
                tipo = 1
                modulo = 12
                cualva()
            ElseIf sel = 2 Then
                TextBox1.Text = "Hola, este es el programa Llamada Saludable. Lamentamos no haber podido ponernos en contacto con usted. Trataremos de comunicarnos con usted nuevamente. Muchas gracias. Adiós"
                TextBox1.Update()
                My.Computer.Audio.Play(My.Resources.g2, AudioPlayMode.WaitToComplete)
                Me.Close()
            End If
        End If

    End Sub

    Public Sub modulo12(tip, sel)
        If tip = 1 Then
            tipo = 2
        ElseIf tip = 2 Then
            If sel = 1 Then
                TextBox1.Text = "Para verificar que hace parte del programa de llamada saludable, por favor digite la fecha de nacimiento oprimiendo las teclas del telefono. Por ejemplo, si nacio en 1952 por favor oprima el numero 1-9-5-2 con el teclado del telefono "
                TextBox1.Update()
                My.Computer.Audio.Play(My.Resources.g5, AudioPlayMode.WaitToComplete)
                tipo = 1
                modulo = 13
                cualva()
            ElseIf sel = 2 Then
                TextBox1.Text = "Gracias. El sistema esperará 5 minutos mientras trae el participante al teléfono. Cuando vuelva, oprima cualquier tecla para continuar."
                TextBox1.Update()
                My.Computer.Audio.Play(My.Resources.g4, AudioPlayMode.WaitToComplete)
                tipo = 1
                modulo = 13
                cualva()
            ElseIf sel = 3 Then
                TextBox1.Text = "Muchas gracias por su ayuda. Trataremos de ponernos en contacto en otro momento. Adiós."
                TextBox1.Update()
                My.Computer.Audio.Play(My.Resources.g6, AudioPlayMode.WaitToComplete)
                Me.Close()
            ElseIf sel = 4 Then
                TextBox1.Text = "Lamento haberle molestado. No le llamaré nuevamente."
                TextBox1.Update()
                My.Computer.Audio.Play(My.Resources.g7, AudioPlayMode.WaitToComplete)
                Me.Close()

            End If
        End If

    End Sub


    ''' ''''''''''''''''''''''


    Public Sub modulo13(tip, sel)
        If tip = 1 Then
            tipo = 2
        ElseIf tip = 2 Then
            If sel = 1 Then
                TextBox1.Text = "Gracias. Estamos llamando del programa Llamada Saludable. Hoy, le preguntaremos sobre su salud en general y el desarrollo presentado al dejar de fumar. Usted puede responder oprimiendo los números del teclado de su teléfono. Serán necesarios 5 a 10 minutos para responder a todas las preguntas. Yo le ire haciendo cada pregunta y luego le voy a ofrecer las opciones de posibles respuestas. Por favor elija la respuesta que usted considera que es la mejor para cada pregunta."
                TextBox1.Update()
                My.Computer.Audio.Play(My.Resources.g8, AudioPlayMode.WaitToComplete)
                TextBox1.Text = "Si necesita que repitamos una pregunta, oprima la tecla asterisco. Por favor, responda a todas las preguntas para que podamos obtener información completa sobre su salud. Si usted quiere omitir parte correspondiente a la información de retorno y desea pasar a la siguiente pregunta oprima la tecla numeral del teclado de su teléfono en cualquier momento."
                TextBox1.Update()
                My.Computer.Audio.Play(My.Resources.g9, AudioPlayMode.WaitToComplete)
                TextBox1.Text = "Ahora, vamos a empezar con la primera pregunta."
                TextBox1.Update()
                My.Computer.Audio.Play(My.Resources.g10, AudioPlayMode.WaitToComplete)
                tipo = 1
                modulo = 1
                cualva()
            End If
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        TextBox1.AppendText("Hola, este es el programa llamada saludable, oprima 1 para continuar")
        My.Computer.Audio.Play(My.Resources.g1, AudioPlayMode.WaitToComplete)
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Me.Close()
    End Sub
End Class
