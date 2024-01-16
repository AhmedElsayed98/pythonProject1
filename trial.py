#X-O Game
from guizero import App, Box, Text, PushButton

def clear_board():
    new_board = [[None,None,None],[None,None,None],[None,None,None]]
    for x in range(3):
        for y in range(3):
            button = PushButton(
                board, text="", grid=[x, y], width=3, command=choose_square, args=[x,y]
            )
            new_board[x][y] = button
    return new_board
def choose_square(x,y):
    board_squares[x][y].text = turn
    board_squares[x][y].disable()
    toggle_player()

def toggle_player():
   global turn
   print(turn)
   if turn == "X":
       turn = "O"
   else:
       turn = "X"
   text1 = 'it is your turn ' + turn
   turn_msg.value = text1

turn = 'X'
app = App('XO Game')
board = Box(app, layout="grid")
app.title = "Game on"
text1 = 'it is your turn '+ turn
turn_msg = Text(app, text = text1)
for x in range(3):
    for y in range(3):
        button = PushButton(
            board, text="", grid=[x,y], width=3
        )
board_squares = clear_board()
print(board_squares)
message = Text(app, text="by Ahmed Zecka")
app.display()
