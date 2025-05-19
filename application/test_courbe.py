import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import scipy.stats as stats

# Chargement des données depuis un fichier Excel
# Supposons que vos données sont dans une colonne appelée 'Valeurs'
df = pd.read_excel('results/Nouveau support stats.xlsx', engine='openpyxl')
donnees = df["POIDS de l'œuf (g)"].dropna()

# Tracer l'histogramme
plt.hist(donnees, bins=30, density=True, alpha=0.6, color='g', label='Données')

# Calculer la densité de la courbe normale avec la moyenne et l'écart-type de vos données
mu, sigma = donnees.mean(), donnees.std()

# Créer une gamme de valeurs pour la courbe normale
x = np.linspace(donnees.min(), donnees.max(), 100)
normal_curve = stats.norm.pdf(x, mu, sigma)

# Tracer la courbe normale
plt.plot(x, normal_curve, color='r', linewidth=2, label='Courbe normale théorique')

plt.xlabel('Valeurs')
plt.ylabel('Densité')
plt.title('Histogramme et courbe normale')
plt.legend()
plt.show()

# Optionnel : Q-Q plot pour vérifier la normalité
import scipy.stats as stats

stats.probplot(donnees, dist="norm", plot=plt)
plt.title('Q-Q Plot')
plt.show()