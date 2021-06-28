# -*- coding: utf-8 -*-
"""
Created on Fri Apr 23 13:49:59 2021

@author: mar_altermark
"""
import fonctions
import affichage


def main():
    MP = fonctions.lireMP()
    emissionsFixes = fonctions.lireEmissionsFixes()
    sacherie = fonctions.lireSacherie()
    produits = fonctions.lireProduits()
    produits = fonctions.lireCompos(produits)
    produits = fonctions.calculBC(produits, MP, sacherie, emissionsFixes) 
    produits_marque = fonctions.regrouper_par_marque(produits)
    
    nomFichier = affichage.enregistre_resultat(produits_marque)
    affichage.noterProduits(nomFichier)
    affichage.elements_manquants(nomFichier)
    return MP, sacherie, produits, emissionsFixes
MP, sacherie, produits, emissionsFixes = main()
