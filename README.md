# 2048 in Excel (VBA Edition)

This project is a recreation of the popular game **2048** — but built entirely using **Excel and VBA macros**! Move the tiles with your arrow keys, combine matching numbers, and try to reach the elusive 2048 tile.

---

## 🎮 Features

- Fully functional **4x4 2048 game board** inside Excel  
- **Arrow key controls** (Up, Down, Left, Right)  
- **Random tile spawning** (2s and 4s)  
- Automatic **score calculation**  
- **Game Over detection**  
- Easy-to-read and modify codebase  

---

## 🚀 Getting Started

### Prerequisites

- Microsoft Excel with **VBA support** (Office 365, Excel 2016+, or compatible)
- Macros must be enabled in Excel for the game to function

### Installation

1. **Clone this repository** or download the `.xlsm` file.
2. Open the file in Excel.
3. Make sure macros are enabled.
4. Press `Alt + F11` to open the VBA editor and inspect/edit the code if needed.

---

## 🕹️ How to Play

1. Run the macro `RunGame` to start the game:  
   Press `Alt + F8` in Excel and select `RunGame`.
2. Use your **arrow keys** to move the tiles:
    - Tiles slide in the direction pressed.
    - Matching tiles merge and double their value.
3. A new tile (2 or 4) will appear after every move.
4. The game ends when there are no possible moves left.

---

## 📁 Project Structure

- `InitializeGame` – Sets up the board and formatting  
- `AddRandomTile` – Spawns a new tile in an empty cell  
- `MoveUp`, `MoveDown`, `MoveLeft`, `MoveRight` – Handle movement and merging logic  
- `CalculateScore` – Sums up all tile values to calculate the current score  
- `IsGameOver` – Checks if the player has lost  
- `SetKeyBindings` – Binds arrow keys to the appropriate movement macros  
- `RunGame` – Main procedure to launch the game  

---

## 🛠️ Customization

You can easily tweak the game by:
- Changing the board size (requires code adjustments)
- Modifying the appearance (font size, color, etc.)
- Adjusting probabilities for spawning 2s vs 4s

---

## 📷 Screenshots
![image](https://github.com/user-attachments/assets/4d90ce0a-4588-4ec8-a634-33b6695b5699)
![image](https://github.com/user-attachments/assets/76655f04-0181-4bba-bb18-e1741e77bee1)
![image](https://github.com/user-attachments/assets/6b7f3d3a-558b-4639-b804-de46b619a0f9)
![image](https://github.com/user-attachments/assets/6654ec2f-82ae-4738-8c87-9fa7c0ba12d3)


---




