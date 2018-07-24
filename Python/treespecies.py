"""
 *****************************************************************************
   FILE :           treespecies.py

   AUTHOR :         Jiin Jeong

   DATE :           June 4, 2018

   DESCRIPTION :	Simple program to find all tree species in the Excel file.

   OUTPUT :
    24              Acacia salicina
	4                Acacia saligna
	3        Chitalpa tashkentensis
	42          Corymbia citriodora
	58                Gingko biloba
	82        Jacaranda mimosifolia
	0     Lagerstroemia x 'Natchez'
	25        Lophostemon confertus
	2            Pistacia chinensis
	64                Quercus rubra
	26               Quercus rubra 

 *****************************************************************************
"""

import pandas as pd

# Skips first three rows. Works with .xlsx as well.
df = pd.read_excel("tree.xls", header=None, skiprows=3)

# Renames the column.
df.rename(columns={list(df)[1]: 'Species'}, inplace=True)

# Drops duplicate species and only keps the first instance.
df = df.drop_duplicates(subset=['Species'])

# Alphabetically sorts.
sorted_alph = df.sort_values(by=['Species'])
print(sorted_alph.iloc[:, 1])
