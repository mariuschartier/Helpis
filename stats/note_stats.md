# Note stats

## periode etudiée

|Modalité|Salles|Sortie jardin|lumiere naturelle|
|:--|---|---|---|
|**Témoin**|1 & 2  |Semaine 11|oui|
|**Sortie anticipée**|3 & 4  |Semaine 7|oui|


## tests sur données

### Nouveau support stats_SALLE/histogramme

```
Histogramme mise en place > Poids(g) : stat=0.9720, p=0.0874 → ✅ Normale
```


### Nouveau support stats_SALLE/PeséeCorpo

JOUR1  / 2023-06-08 / SALLE : 1, 2, 3, 4 / 1 suppression
``` 
Poids : stat=0.9907, p=0.2320 → ✅ Normale
```

JOUR2  / 2023-06-15 / SALLE : 1, 2, 3, 4 / 1 suppression
``` 
Poids : stat=0.9920, p=0.3421 → ✅ Normale
```


JOUR3  / 2023-06-22 / SALLE : 1, 2, 3, 4 / 0 suppression
``` 
Poids : stat=0.9938, p=0.5746 → ✅ Normale
```

JOUR4  / 2023-07-20 / SALLE : 1, 2, 3, 4 / 1 suppression
``` 
Poids : stat=0.9890, p=0.1313 → ✅ Normale
```


JOUR5  / 2023-08-03 / SALLE : 1, 2, 3, 4 / 1 suppression
``` 
Poids : stat=0.9916, p=0.3069 → ✅ Normale
```

JOUR6  / 2023-09-07 / SALLE : 1, 2, 3, 4 / 0 suppression
``` 
Poids : stat=0.9846, p=0.0267 → ❌ Non normale
```

### Nouveau support stats_SALLE/PeséeOeufs