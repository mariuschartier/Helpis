# import matplotlib.pyplot as plt

# # Exemple de liste de données
# data = [7, 15, 13, 9, 12, 8, 10, 14, 11, 16, 12, 9]

# # Création de la boîte à moustache
# plt.boxplot(data)

# # Ajout de titres et labels
# plt.title("Boîte à moustache")
# plt.ylabel("Valeurs")

# # Affichage du graphique
# plt.show()

import matplotlib.pyplot as plt

# Exemple de liste de données
categories = ['Catégorie 1', 'Catégorie 2', 'Catégorie 3', 'Catégorie 4']
valeur1 = [10, 15, 7, 12]
valeur2 = [8, 11, 9, 14]

# Création du bareplot
plt.bar(categories, valeur1, label='Série 1')
plt.bar(categories, valeur2, bottom=valeur1, label='Série 2')

# Ajout de titres et labels
plt.title("Barplot")
plt.xlabel("Catégories")
plt.ylabel("Valeurs")
plt.legend()

# Affichage du graphique
plt.show()