import tkinter as tk
from tkinter import messagebox

def show_message():
    messagebox.showinfo("Hello", "This is a macOS app built from Python!")

root = tk.Tk()
root.title("Mac App")
root.geometry("300x200")

btn = tk.Button(root, text="Click Me", command=show_message)
btn.pack(pady=50)

root.mainloop()
