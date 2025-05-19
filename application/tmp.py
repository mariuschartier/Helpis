    def create_result_box(self):
        # Cadre pour les résultats du test
        self.result_frame = tk.LabelFrame(self, text="4. Résultats du test", bg="#f4f4f4")
        self.result_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Barre de défilement verticale
        scroll_y = tk.Scrollbar(self.result_frame, orient="vertical")
        scroll_y.pack(side="right", fill="y")

        # Barre de défilement horizontale
        scroll_x = tk.Scrollbar(self.result_frame, orient="horizontal")
        scroll_x.pack(side="bottom", fill="x")

        # Zone de texte avec padding
            # Zone de texte avec padding
        self.result_text = tk.Text(
            self.result_frame, 
            height=10, 
            wrap="none",  # Pas de retour à la ligne automatique
            xscrollcommand=scroll_x.set, 
            yscrollcommand=scroll_y.set,
            padx=5, 
            pady=5
        )
        self.result_text.pack(fill="both", expand=True)

        # Méthode pour ajouter du texte sans le remplacer
 
        self.result_text.pack(fill="both", expand=True)

        # Configuration des barres de défilement
        scroll_y.config(command=self.result_text.yview)
        scroll_x.config(command=self.result_text.xview)








    def create_results_frame(self):
        self.results_frame = tk.LabelFrame(self.scrollable_frame, text="4. Résultats / Erreurs", bg="#f4f4f4")
        self.results_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.result_text = tk.Text(self.results_frame, height=10, wrap="none")
        self.result_text.pack(fill="both", expand=True, padx=10, pady=5)

        # Barres de défilement pour les résultats
        result_scroll_y = tk.Scrollbar(self.results_frame, command=self.result_text.yview)
        result_scroll_y.pack(side="right", fill="y")
        result_scroll_x = tk.Scrollbar(self.results_frame, orient="horizontal", command=self.result_text.xview)
        result_scroll_x.pack(side="bottom", fill="x")

        self.result_text.configure(yscrollcommand=result_scroll_y.set, xscrollcommand=result_scroll_x.set)
