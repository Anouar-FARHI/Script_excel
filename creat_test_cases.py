import openpyxl
from openpyxl.styles import Font

def create_test_cases():
    # Créer un nouveau classeur
    wb = openpyxl.Workbook()

    # Ajouter une feuille pour le test de paiement
    payment_sheet = wb.active
    payment_sheet.title = "Test de Paiement"

    # Ajouter des en-têtes
    headers = ["Étape", "Action", "Entrée", "Résultat Attendu", "Résultat Observé", "Statut"]
    payment_sheet.append(headers)

    # Rendre les en-têtes en gras
    for cell in payment_sheet["1:1"]:
        cell.font = Font(bold=True)

    # Ajouter des cas de test pour le paiement
    payment_tests = [
        ["1", "Naviguer vers le site", "URL du site", "Page d'accueil affichée", "", ""],
        ["2", "Rechercher un produit", "Nom du produit", "Page des résultats affichée", "", ""],
        ["3", "Sélectionner un produit", "Clique sur le produit", "Page du produit affichée", "", ""],
        ["4", "Ajouter au panier", "Clique sur 'Ajouter au panier'", "Produit ajouté au panier", "", ""],
        ["5", "Passer à la caisse", "Clique sur 'Passer à la caisse'", "Page de paiement affichée", "", ""],
        ["6", "Entrer les informations de paiement", "Détails de la carte", "Informations acceptées", "", ""],
        ["7", "Confirmer le paiement", "Clique sur 'Payer'", "Confirmation de la commande", "", ""],
    ]

    for test in payment_tests:
        payment_sheet.append(test)

    # Ajouter une feuille pour le test de suivi de commande
    order_tracking_sheet = wb.create_sheet(title="Suivi de Commande")

    # Ajouter des en-têtes
    order_tracking_sheet.append(headers)

    # Rendre les en-têtes en gras
    for cell in order_tracking_sheet["1:1"]:
        cell.font = Font(bold=True)

    # Ajouter des cas de test pour le suivi de commande
    tracking_tests = [
        ["1", "Naviguer vers le site", "URL du site", "Page d'accueil affichée", "", ""],
        ["2", "Se connecter", "Identifiants", "Compte connecté", "", ""],
        ["3", "Accéder au suivi de commande", "Clique sur 'Suivi de commande'", "Page de suivi affichée", "", ""],
        ["4", "Vérifier l'état de la commande", "Numéro de commande", "État de la commande affiché", "", ""],
    ]

    for test in tracking_tests:
        order_tracking_sheet.append(test)

    # Sauvegarder le fichier Excel
    wb.save("Cas_de_Test_Cdiscount.xlsx")
    print("Cas de test créés et sauvegardés dans 'Cas_de_Test_Cdiscount.xlsx'.")

# Appeler la fonction pour créer les cas de test
create_test_cases()
