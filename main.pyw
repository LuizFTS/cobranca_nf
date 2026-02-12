import tkinter as tk
from src.presentation.main_window import StoreEmailConfigUI

def main():
    root = tk.Tk()
    app = StoreEmailConfigUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
