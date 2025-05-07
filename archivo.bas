Sub AsignarRecursosMasivamente()
    Dim t As Task
    Dim recursosTrabajo As Variant
    Dim materiales As Variant
    Dim recurso As Resource
    Dim material As Resource
    Dim nombre As Variant

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            recursosTrabajo = Array()
            materiales = Array()

            Select Case Trim(t.Name)

            'example  add resources in the task 
            Case "name task"
                recursosTrabajo = Array("resources_1", "resources_2", "add more?")
                materiales = Array("resources_1", "resources_2", "add more?")
            
            Case "name task"
                recursosTrabajo = Array("resources_1", "resources_2", "add more?")
                materiales = Array("resources_1", "resources_2", "add more?")
            'end example

           End Select

            ' Asignar recursos de trabajo
            For Each nombre In recursosTrabajo
                On Error Resume Next
                Set recurso = ActiveProject.Resources(nombre)
                If Not recurso Is Nothing Then
                    t.Assignments.Add ResourceID:=recurso.ID
                End If
                On Error GoTo 0
            Next nombre

            ' Asignar materiales
            For Each nombre In materiales
                On Error Resume Next
                Set material = ActiveProject.Resources(nombre)
                If Not material Is Nothing Then
                    t.Assignments.Add ResourceID:=material.ID
                End If
                On Error GoTo 0
            Next nombre
        End If
    Next t

    MsgBox "¡Asignación completada! Verifica en la vista 'Diagrama de Gantt'.", vbInformation
End Sub
