import Bombadil
import tkinter as tk
from tkinter import messagebox

Bombadil.ADMIN_MODE = False

if __name__ == "__main__":
    try:
        Bombadil.App().mainloop()
    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Startfehler", f"{type(e).__name__}: {e}")
        raise
