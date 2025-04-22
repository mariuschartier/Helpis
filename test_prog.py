from test_gen import Test_gen
from test_spe import Test_spe

from fichier import Fichier
from Feuille import Feuille




# =============================================================================
# Pour le facteur température, les valeurs des sondes 1 et 2 ne doivent pas présenter un écart de plus de 10 degrés avec les valeurs de la colonne consigne température.
# Pour l’hygrométrie tu peux conserver la plage 0-100%.
# Pour le poids, il doit être compris en 30 et 3500g.

# Pour la dépression, elle doit être comprise en 0 et 50 Pa (c’est ce qui nous  indique la « force » de la ventilation sur la plage de temps).
# Les paliers doivent être compris entre 0 et 9 ( cela nous indique la plage de ventilation que l’automate a sélectionné).
# Le CO2 doit être compris entre 0 et 3000ppm.
# Le réel femelle est le poids qui a été mesuré par l’automate, il doit être compris entre 0 et 3500g et présenter un écart inférieur à 20% avec la colonne Poids femelle (qui est la valeur théorique attendue).
# Le luxmètre doit avoir une valeur comprise entre 0 et 100 lux. (cela nous permet de contrôler la luminosité ambiante qui peut avoir un impact sur le comportement des animaux).
# L’inter crépuscule ne nous sert pas dans ce boitier, il s’agit d’une donnée constructeur que l’on ne peut pas enlever. Il n’y a donc pas besoin de la contrôler.
# L’occultant doit être compris entre 0 et 100% (ce sont les volets qui permettent l’entrée de lumière naturelle).
# =============================================================================



# =============================================================================
#                    TESTS GENERAUX
# =============================================================================

tg1 = Test_gen("test_%",["%"])
tg2 = Test_gen("test_°C",["°","°C"])

# =============================================================================
#                    TESTS Mesures.xlsx
# =============================================================================
f1 = Fichier("result/Mesures.xlsx")
f1_1 = Feuille(f1,"opti",2)
f1_1.clear_all_cell_colors()



# tg1.val_entre(f1_1,0, 100)

ts1 = Test_spe("test_Mesures", f1_1)

# ts1.val_entre(0,50,"Depression 1") 
# ts1.val_entre(30,3500,"Poids Femelles")
# ts1.val_entre(0,9,"Paliers") 
# ts1.val_entre(0,3000,"CO2")
# ts1.val_entre(0,3500,"Réel femelles")
# ts1.val_entre(0,100,"Luxmetre")

ts1.compare_col_fix(10, "Sonde 1", "Consigne Temp")
ts1.compare_col_fix(10, "Sonde 2", "Consigne Temp")

# ts1.compare_col_ratio(1.2, "Réel femelles", "Poids Femelles")


# mise à jour du fichier
f1_1.error_all_cell_colors()
# =============================================================================
#                    TESTS MinMax.xlsx
# =============================================================================
# f2 = Fichier("result/MinMax.xlsx")
# f2_1 = Feuille(f2,"opti",3)
# f2_1.clear_all_cell_colors()


# ts2 = Test_spe("test_Mesures", f2_1)


# tg1.val_entre(f2_1,0,100)

# ts2.val_entre(0,100,"Luxmetre")


#mise à jour du fichier
# f2_1.error_all_cell_colors()

 

