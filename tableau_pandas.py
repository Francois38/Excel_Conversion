#!/usr/bin/env python3
# -*- coding: utf8 -*-

import numpy as np 
import pandas as pd
import os
import time

####################### Gestion des chemins d'accès aux fichiers inputs et outputs ###########

## On asigne les chemins des répertoires d'entrée et de sortie à 2 variables
répertoire_entrée='/Users/Desktop/Tableaux_inputs/Puissances Excel'
répertoire_sortie='/Users/Desktop/Tableaux_outputs'

## On crée une liste de string avec les noms des fichiers présents dans le répertoire d'entrée
fichiers = os.listdir(répertoire_entrée)

## On initialise les variables (listes) des fichiers_excel des chemins d'entrée, de sortie
fichiers_excel_in= []
chemins_entrée = []
chemins_sortie = []

## On ne sélectionne que les fichiers Excels que l'on range dans fichiers_excel_in
for element in fichiers :
	if element.endswith('.xlsx'): fichiers_excel_in.append(element)

## On re crée les chemins complets d'accès aux fichiers, en entrée et en sortie
for element in fichiers_excel_in :
	 chemins_entrée.append(répertoire_entrée+"/"+element)
	 chemins_sortie.append(répertoire_sortie+"/"+element)

## On affiche les listes de chemins d'entrée et de chemin de sorties
print("Voici les chemins d'entrée: \n") 
print(chemins_entrée)
print("\nVoici les chemins de sortie: \n") 
print(chemins_sortie) 

##################### Création de la colonne des temps ##############################

## On crée la liste de 144 valeurs, entiers de 0 à 86400 secondes

time_Enedis = [i for i in range(0, 86400, 600)]

## On transforme les valeurs de cette listes en str au format HH:MM avec le module time

i=0
while i< len(time_Enedis):
	time_Enedis[i]=time.strftime("%HH%M", time.gmtime(time_Enedis[i]))
	i+=1

#################### Boucle de convertion des fichiers Excels du répertoire inputs ##############

i=0
while i < len(chemins_entrée):

	## On sélectionne l'élément i de la liste chemins_entrée
	a =chemins_entrée[i]
	puissance_KEP = []
	puissance_Enedis = []

	## Avec le module Pandas on lit le fichier excel d'entrée, uniquement la colonne C
	df_input = pd.read_excel(a,skiprows = 0,usecols="C")
	#print(df_input)

	## On crée une nouvelle liste en faisant la moyenne des valeurs du dataframe toutes les 60 valeurs 
	j=0
	while j < len(df_input):
		a= round(sum(df_input.iloc[j:j+59:1,0])/len(df_input.iloc[j:j+59:1,0]), 2)
		puissance_Enedis.append(a)
		j+=60

	## On transforme les listes time_Enedis et puissance_Enedis en un dataframe de sortie 
	df_output = pd.DataFrame({'Temps': time_Enedis, 'Puissance active reseau': puissance_Enedis})
	
	#df_output.set_index('Temps', inplace=True)
	#print(df_output)

	## On transforme ce dataframe en un fichier Excel avec ExcelWriter et notre chemin de sortie
	b = chemins_sortie[i]
	with pd.ExcelWriter(b, engine = 'xlsxwriter') as writer:
		df_output.to_excel(writer, index=False)
	i+=1

## A la fin de la boucle on affiche le nombre de fichiers convertis
print("C'est terminé,",len(chemins_sortie), " fichiers xlsx ont été convertis !")




#writer = pd.ExcelWriter(b, engine = 'xlsxwriter')

'''
#puissance_KEP = tableau_Entrée['Puissance active reseau'].tolist()
#print(puissance_KEP)	#8640 rows * 1 col

for i in df_input['Puissance active reseau']:
	print (df_input.iloc[i,0])

def sec_to_Heures(nb_sec):
	q,s=divmod=(nb_sec,60)
	h,m=divmod(q,60)
	return "%dH %d" %(h,m)

for i in time_Enedis:
	time_Enedis[i]= sec_to_Heures(time_Enedis[i])
	
print(time_Enedis)


#n = len(tableau_Sortie)
#m=len(tableau_Sortie[0])
#tableau_Sortie = [[tableau_Sortie[j][i] for j in range(n)] for i in range(m)]

j=0
while j < len(puissance_KEP):
	a= round(sum(puissance_KEP[j:j+59:1])/len(puissance_KEP[j:j+59:1]), 2)
	puissance_Enedis.append(a)
	j+=60

#print(puissance_Enedis)
print(len(puissance_Enedis)) #144 éléments

tuple_puissance = tuple(puissance_Enedis)
tuple_temps = tuple(time_Enedis)
print(type(tuple_puissance))

data = [tuple_temps, tuple_puissance]
print(type(data))

df = pd.DataFrame(data)
df.T
print(df)
#df.columns = []

time_Enedis = np.array([time_Enedis])
puissance_Enedis=np.array([puissance_Enedis])

tableau_Sortie = np.array([[time_Enedis],[puissance_Enedis]])
print(tableau_Sortie)

df = pd.DataFrame(tableau_Sortie)
print(df)

#tableau_Sortie = pd.DataFrame(tableau_Sortie, index = [], columns = ['Temps (min)','Puissance Active reseau (kW)'])
writer=df.to_excel('output1.xlsx')
#writer = pd.ExcelWriter("test1.xlsx", engine='xlsxwriter')
#df.to_excel(writer,sheet_name = 'tab1', index=False, columns = False)
#writer.save() 

'''
