'@2023 Newspoint Software
Imports System
Imports System.Text.RegularExpressions

Public Module MathFunc
    ''' <summary>
    ''' Classe per uso matematico
    ''' </summary>
    Public Class MathOperation
        ''' <summary>
        ''' Basical six operation (+ , - , * , / , % , √)
        ''' Sample :: PerformOperation(textbox1.text).tostring()
        ''' </summary>
        Public Function PerformOperation(ByVal input As String) As Double
            Dim result As Double = 0.0

            ' Usa espressioni regolari per estrarre i componenti (numeri e operatore)
            Dim pattern As String = "(\d+(?:\.\d+)?)\s*([-+*/√])\s*(\d+(?:\.\d+)?)\s*(?:([+-]?\d+(?:\.\d+)?)%)?"
            Dim match As Match = Regex.Match(input, pattern)

            If match.Success Then
                Dim num1 As Double = Double.Parse(match.Groups(1).Value)
                Dim operationType As Char = match.Groups(2).Value(0)
                Dim num2 As Double = Double.Parse(match.Groups(3).Value)
                Dim percentage As Double = If(match.Groups(4).Success, Double.Parse(match.Groups(4).Value), 0)

                ' Esegui l'operazione in base all'operatore
                Select Case operationType
                    Case "+"
                        If percentage <> 0 Then
                            result = num1 + (num2 * (1 + percentage / 100))
                        Else
                            result = num1 + num2
                        End If
                    Case "-"
                        If percentage <> 0 Then
                            result = num1 - (num2 * (1 - percentage / 100))
                        Else
                            result = num1 - num2
                        End If
                    Case "*"
                        If percentage <> 0 Then
                            result = num1 * (num2 * (1 + percentage / 100))
                        Else
                            result = num1 * num2
                        End If
                    Case "/"
                        If num2 <> 0 Then
                            If percentage <> 0 Then
                                result = num1 / (num2 * (1 - percentage / 100))
                            Else
                                result = num1 / num2
                            End If
                        Else
                            Return "Divisione per zero non consentita."
                            ' Puoi gestire l'errore in un modo appropriato qui e restituire un valore di errore se necessario.
                        End If
                    Case "√" ' Radice quadrata
                        If num1 >= 0 Then
                            result = Math.Sqrt(num1)
                        Else
                            Return "Impossibile calcolare la radice quadrata di un numero negativo."
                            ' Puoi gestire l'errore in un modo appropriato qui e restituire un valore di errore se necessario.
                        End If
                    Case Else
                        Return "Operatore non valido."
                        ' Puoi gestire l'errore in un modo appropriato qui e restituire un valore di errore se necessario.
                End Select
            Else
                Return "Input non valido. L'input deve essere nel formato 'numero operatore numero percentuale'."
                ' Puoi gestire l'errore in un modo appropriato qui e restituire un valore di errore se necessario.
            End If

            ' Restituisci il risultato
            Return result
        End Function
    End Class
End Module