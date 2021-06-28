# -*- coding: utf-8 -*-
"""
Created on Fri Apr 23 14:39:11 2021

@author: mar_altermark
"""
import xlrd
import sys

SFMANQUANT = []
SACHERIEMANQUANTE = []
INTRANTMANQUANT = []
NOM_TABLEAU_MP = "tableauMP.xlsx"
NOM_TABLEAU_SACHERIE = "tableauSacherie.xlsx"
NOM_TABLEAU_COMPOS = "tableauCompos.xlsx"
NOM_TABLEAU_PSFPF = "tableauPFPSF.xlsx"
NOM_TABLEAU_CODEMP = "tableauEquivalencesMP.xlsx"
NOM_TABLEAU_FIXE = "tableauFixe.xlsx"

def lireMP():
    listeMP = [["Nom", "Identifiant", "Famille", "Tonnage 2020", "Fret routier moyen (kgCO2e/t)", "Fret naval moyen (kgCO2e/t)",
              "Masse volumique (t/m3)", "FE fabrication (kgCO2e/t)", "FE N2O (kgCO2e/t)", "FE tourbe (kgCO2e/t)", "Ref"]]
    try:
         document = xlrd.open_workbook(NOM_TABLEAU_MP)
    except FileNotFoundError:
         sys.exit("Le fichier des matières premières n'est pas trouvé à l'emplacement "+NOM_TABLEAU_MP)
    feuille_bdd = document.sheet_by_index(0)
    nbrows = feuille_bdd.nrows
    for ligne in range(1,nbrows):
        l = []
        for colonne in [0,1,2,7,10,11,12,13,17,18]:
            l.append(feuille_bdd.cell_value(ligne, colonne))
        listeMP.append(l)
    
    MP = [["Nom", "Code", "Réf"]]
    try:
         document = xlrd.open_workbook(NOM_TABLEAU_CODEMP)
    except FileNotFoundError:
         sys.exit("Le fichier des matières premières n'est pas trouvé à l'emplacement "+NOM_TABLEAU_MP)
    feuille_bdd = document.sheet_by_index(0)
    nbrows = feuille_bdd.nrows
    for ligne in range(1,nbrows):
        nom = feuille_bdd.cell_value(ligne, 0)
        code = feuille_bdd.cell_value(ligne, 3)
        ref = feuille_bdd.cell_value(ligne, 1)
        MP.append([nom, code, ref])
    
    for mp in listeMP[1:]:
        trouve = False
        nom = mp[0]
        code = mp[1]
        for mp2 in MP[1:]:
            if nom == mp2[0] or code == mp2[1]:
                trouve = True
                ref = mp2[2]
                mp.append(ref)
        if not trouve:
            mp.append(0)
            # print("Code MP introuvé pour :"+nom)
    return listeMP


def lireProduits():
    listeProduits= [["PF","PSF","Marque","Volume","Composition", "Adjuvants"]]
    try:
         document = xlrd.open_workbook(NOM_TABLEAU_PSFPF)
    except FileNotFoundError:
         sys.exit("Le fichier des matières premières n'est pas trouvé à l'emplacement "+NOM_TABLEAU_PSFPF)
    feuille_bdd = document.sheet_by_index(0)
    nbrows = feuille_bdd.nrows
    for ligne in range(1,nbrows):
        PF = feuille_bdd.cell_value(ligne, 4)
        PSF = feuille_bdd.cell_value(ligne, 7)
        marque = feuille_bdd.cell_value(ligne, 2)
        litrage = feuille_bdd.cell_value(ligne, 5)
        composition = [["Nom MP", "Code MP", "Pourcentage"]]
        adjuvants = [["Nom AD", "Code AD", "Quantité (kg/m3)"]]
        listeProduits.append([PF, PSF, marque, litrage, composition, adjuvants])
    return listeProduits

def lireCompos(listeProduits):
    #Lecture du fichier lien entre PSF et PF
    entete = []
    try:
         document = xlrd.open_workbook(NOM_TABLEAU_COMPOS)
    except FileNotFoundError:
         sys.exit("Le fichier des matières premières n'est pas trouvé à l'emplacement "+NOM_TABLEAU_COMPOS)
    feuille_bdd = document.sheet_by_index(0)
    nbrows = feuille_bdd.nrows
    nbcols = feuille_bdd.ncols
    for i in range(1,nbcols):
        entete.append(feuille_bdd.cell_value(1, i))
    
    n_total = entete.index('Total général')
    
    for ligne in range(2,nbrows):
        composition  = [["Ref MP", "Pourcentage","Routier", "Naval","Fabrication","EOL engrais", "EOL Tourbe"]]
        ammendements = [["Ref AD", "Qte (kg/m3)","Routier", "Naval","Fabrication","EOL engrais", "EOL Tourbe"]]
        refSF = feuille_bdd.cell_value(ligne, 0)
        for i in range(1,nbcols):
            if i-1!= n_total and feuille_bdd.cell_value(ligne, i)!="":
                qte = feuille_bdd.cell_value(ligne, i)
                mp = entete[i-1]
                if i-1<n_total:
                    #Il s'agit de matières premières
                    composition.append( [mp, qte]+5*[0])
                else:
                    #Il s'agit d'ammendements
                    ammendements.append([mp, qte]+5*[0])
        for PF in listeProduits:
            if PF[1][0] =="*":
                PF[1] = PF[1][1:]
            if PF[1]==refSF:
                PF[4]=composition
                PF[5]=ammendements
    
    res = []
    for PF in listeProduits[1:]:
        if len(PF[4])>1:
            res.append(PF)
        else:
            SFMANQUANT.append(PF[1])
            print("Composition introuvée pour "+str(PF[0])+" (mélange "+PF[1]+")")
    return res

def lireEmissionsFixes():
    try:
         document = xlrd.open_workbook(NOM_TABLEAU_FIXE)
    except FileNotFoundError:
         sys.exit("Le fichier des matières premières n'est pas trouvé à l'emplacement "+NOM_TABLEAU_COMPOS)
    feuille_bdd = document.sheet_by_index(0)
    nbrows = feuille_bdd.nrows
    ligne0 = 3
    resultat=[["Ventes pro","Ventes Jardineries","Interdepot","Electricité",
              "Fuel","Fret amont des emballages","Deplacements","Immobilisations"], 8*[0]]
    for i in range(ligne0,nbrows):
        intitule = feuille_bdd.cell_value(i, 1)
        try:
            idx = resultat[0].index(intitule)
            resultat[1][idx] = feuille_bdd.cell_value(i, 3)
        except ValueError:
            pass
    return resultat

def calculBC(produits, MP, sacherie,emissionsFixes) :
    produits.insert(0, ["Ref","SF", "Marque", "Litrage", "Composition", "Adjuvants"])
    produits[0].append("Total routier (kgCO2e)")
    produits[0].append("Total naval (kgCO2e)")
    produits[0].append("Total fabrication (kgCO2e)")
    produits[0].append("Total N2O (kgCO2e)")
    produits[0].append("Total tourbe décomposition (kgCO2e)")
    produits[0].append("Total sacherie (kgCO2e)")
    produits[0].append("Emissions fixes (kgCO2e)")
    produits[0].append("Bilan carbone par sac (kgCO2e)")
    produits[0].append("Bilan carbone par m3 (kgCO2e/m3)")
    produits[0].append("Part de la tourbe dans le BC (%)")
    produits[0].append("Part des engrais dans le BC (%)")
    for produit in produits[1:]:
        produit.append(11*[0])
        compteur_engrais= 0 #en kgCO2e
        compteur_tourbe = 0 #en kgCO2e
        volume      = produit[3] #en L
        composition = produit[4]
        adjuvants   = produit[5]
        for mp in composition[1:]:
            refMP = mp[0]
            pourcent = mp[1]
            trouvee = False
            for intrant in MP[1:]:
                if str(intrant[10]) in str(refMP) or str(refMP) in str(intrant[10]):
                    nomMP = intrant[0]
                    
                    trouvee = True
                    fret_routier= intrant[4] #en kgCO2e/t
                    fret_naval  = intrant[5] #en kgCO2e/t
                    masse_volum = intrant[6] #en t/m3
                    FE_fabricat = intrant[7] #en kgCO2e/t
                    FE_N2O      = intrant[8] #en kgCO2e/t
                    FE_tourbe   = intrant[9] #en kgCO2e/t
                    
                    
                    #tout arrive en kgco2e
                    mp[2] = fret_routier*pourcent*volume*masse_volum/100000 
                    mp[3] = fret_naval*pourcent*volume*masse_volum/100000 
                    mp[4] = FE_fabricat*pourcent*volume*masse_volum/100000 
                    mp[5] = FE_N2O*pourcent*volume*masse_volum/100000 
                    mp[6] = FE_tourbe*pourcent*volume*masse_volum/100000 
                    if 'tourbe' in nomMP.lower():
                        compteur_tourbe +=sum(mp[2:7])
                    break
            if not trouvee:
                print("Intrant introuvé ("+str(refMP)+") pour ref "+str(produit[1]))
                INTRANTMANQUANT.append(refMP)
        for mp in adjuvants[1:]:
            refMP = mp[0]
            qte = mp[1]
            trouvee = False
            for intrant in MP[1:]:
                if intrant[10] == refMP:
                    nomMP = intrant[0]
                    trouvee = True
                    fret_routier= intrant[4] #en kgCO2e/t
                    fret_naval  = intrant[5] #en kgCO2e/t
                    masse_volum = intrant[6] #en t/m3
                    FE_fabricat = intrant[7] #en kgCO2e/t
                    FE_N2O      = intrant[8] #en kgCO2e/t
                    FE_tourbe   = 0
                    
                    #tout arrive en kgco2e                   
                    mp[2] = fret_routier*qte*volume/1000000
                    mp[3] = fret_naval*qte*volume/1000000
                    mp[4] = FE_fabricat*qte*volume/1000000 
                    mp[5] = FE_N2O*qte*volume/1000000
                    mp[6] = FE_tourbe*qte*volume/1000000 
                    
                    compteur_engrais +=sum(mp[2:7])
                    
                    break
            if not trouvee:
                print("Intrant introuvé ("+str(refMP)+") pour ref "+str(produit[1]))
                INTRANTMANQUANT.append(refMP)
        sacherie_trouvee = False
        for i in range(len(sacherie)-1,0,-1):
            if sacherie[i][2] in produit[0]:
                sacherie_trouvee = True
                BC_sacherie = calc_BC_sacherie(sacherie[i])
                break
        if not sacherie_trouvee:
            if len(produit[0])>=5:
                recherche= "".join(produit[0][0:6])
                for i in range(len(sacherie)-1,0,-1):
                    if recherche in sacherie[i][2]:
                        sacherie_trouvee = True
                        BC_sacherie = calc_BC_sacherie(sacherie[i])
                        break
            if not sacherie_trouvee:
                print("Sacherie introuvée pour ref "+str(produit[0])+" ("+str(trouver_litrage(produit[0]))+"L)")
                SACHERIEMANQUANTE.append(produit[0])
                BC_sacherie = 0
        produit[6][0] = (sum([mp[2]for mp in composition[1:]])+sum([mp[2]for mp in adjuvants[1:]]))
        produit[6][1] = (sum([mp[3]for mp in composition[1:]])+sum([mp[3]for mp in adjuvants[1:]]))
        produit[6][2] = (sum([mp[4]for mp in composition[1:]])+sum([mp[4]for mp in adjuvants[1:]]))
        produit[6][3] = (sum([mp[5]for mp in composition[1:]])+sum([mp[5]for mp in adjuvants[1:]]))
        produit[6][4] = (sum([mp[6]for mp in composition[1:]])+sum([mp[6]for mp in adjuvants[1:]]))
        produit[6][5] = BC_sacherie/1000 #en kgCO2e
        produit[6][6] = sum(emissionsFixes[1])*volume/1000 #en kgCO2e
        #Bilan carbone par sac (gCO2e)
        produit[6][7] = sum(produit[6][0:7])
        #Bilan carbone par m3 (kgCO2e/m3)
        produit[6][8] = produit[6][7]*1000/volume
        if produit[6][7]!=0:
            produit[6][9] = (compteur_tourbe)/produit[6][7] #en kgCO2e
            produit[6][10] = (compteur_engrais)/produit[6][7] #en kgCO2e
        else:
            produit[6][9] = "nan"
            produit[6][10] = "nan"
    return produits

def calc_BC_sacherie(ligne_sacherie):
    FE_PEBD_v = 2090  #kgCO2e/t
    FE_PEBD_r = 202   #kgCO2e/t
    FE_papier_v = 100 #kgCO2e/t
    FE_papier_r = 100 #kgCO2e/t
    taux_recycle = ligne_sacherie[6]
    poids_PEBD = ligne_sacherie[7]
    poids_papier = ligne_sacherie[8]
    
    emissionsPEBD = poids_PEBD*(taux_recycle*FE_PEBD_r+(1-taux_recycle)*FE_PEBD_v)/1000
    emissionsPapier = poids_papier*(taux_recycle*FE_papier_r+(1-taux_recycle)*FE_papier_v)/1000
    
    return emissionsPEBD+emissionsPapier

def regrouper_par_marque(produits):
    marques_vues = []
    resultat = []
    for element in produits:
        marque = element[2]
        try:
            i = marques_vues.index(marque)
            resultat[i][1].append(element)
        except ValueError:
            marques_vues.append(marque)
            resultat.append([marque, [element]])
    return resultat

def lireSacherie():
    listeSacherie = [["Annee", "Gamme", "ReferencePSF","Reference Sacherie", "Designation",
                     "Materiau", "Taux recyclé", "Poids PEBD", "Poids Papier"]]
    try:
         document = xlrd.open_workbook(NOM_TABLEAU_SACHERIE)
    except FileNotFoundError:
         sys.exit("Le fichier des matières premières n'est pas trouvé à l'emplacement "+NOM_TABLEAU_SACHERIE)
    feuille_bdd = document.sheet_by_index(0)
    nbrows = feuille_bdd.nrows
    for ligne in range(10,nbrows):
        l = []
        for colonne in [1,2,3,4,5,10,11,12,13]:
            l.append(feuille_bdd.cell_value(ligne, colonne))
        if l[2] !="":
            listeSacherie.append(l)
    return listeSacherie
def trouver_litrage(reference):
    a = "".join([c for c in reference if c.isdigit()])
    if len(a)>2:
        a = int(a[0]+a[1])
    if a !="":
        return int(a)
    else:
        return 0
    
