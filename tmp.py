import tkinter as tk

class WrappingGridApp:
    def __init__(self, root):
        self.root = root
        root.geometry("600x300")
        root.title("Grille avec wrapping dynamique")

        self.frame = tk.Frame(root)
        self.frame.pack(fill="both", expand=True)

        self.widgets = []
        self.num_columns = 4  # Valeur par défaut

        for i in range(20):  # 20 widgets pour tester le wrapping
            btn = tk.Button(self.frame, text=f"Widget {i+1}", width=15)
            self.widgets.append(btn)

        self.arrange_widgets()

        self.root.bind("<Configure>", self.on_resize)

    def arrange_widgets(self):
        for widget in self.frame.winfo_children():
            widget.grid_forget()

        for index, widget in enumerate(self.widgets):
            row = index // self.num_columns
            col = index % self.num_columns
            widget.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")

        for col in range(self.num_columns):
            self.frame.columnconfigure(col, weight=1)

    def on_resize(self, event):
        width = self.frame.winfo_width()
        new_columns = max(1, width // 150)  # ajuster la largeur de seuil
        if new_columns != self.num_columns:
            self.num_columns = new_columns
            self.arrange_widgets()

if __name__ == "__main__":
    root = tk.Tk()
    app = WrappingGridApp(root)
    root.mainloop()

