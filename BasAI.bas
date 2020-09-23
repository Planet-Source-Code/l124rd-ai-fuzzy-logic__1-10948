Attribute VB_Name = "BasAI"
'Differnt Probobility/States for the enemy
Public Sleep As Integer     'Probability that the Enemy will sleep
Public Attack As Integer    'Probability that the Enemy will Attack
Public Gaurd As Integer     'Probability that the Enemy will Guard

'What it is doing currently
Public State As String       'What it is/will do(ing)

'Holds the Player's HitPoints
Public Hp As Integer         'Hps for the player

'Calculates the State for the Enemy
Public Function Calculate()
    'dims num for rnd
    Dim num As Integer
    'Calculates num for rnd
    num = Sleep + Attack + Gaurd + 1
    'Makes a int to hold the rnd in
    Dim prob As Integer
    'Holds the rnd in prob
    prob = CInt(Rnd(Time) * num)
    'Cint rounds whatever is inside the brackets
    'rnd(time) makes a random number off of the time
    'The rnd is ussually a single wich is why we use
    'Cint
    'the num efficiantly makes the number between 0
    'and (num) - 1 so we add one to num
    
    'The Following Functions take the random
    'Number and create a state based on those
    'having them grouped together makes it so
    'there is one decsision total.
    
    'Sees if its lower then sleep
    If prob < Sleep Then State = "Sleep"
    'Sees if its lower then Attack
    If prob < Attack Then State = "Attack"
    'Sees if its lower then gaurd
    If prob < Gaurd Then State = "Gaurd"
    If State = "" Then
        Calculate
    End If
    FrmFuzzy.Label4.Caption = State
End Function
