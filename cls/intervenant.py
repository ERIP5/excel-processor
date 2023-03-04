class Intervenant:

    def __init__(self, nom:str, prenom:str, adresse:str, status:str, employeur:str):
        self.nom = nom
        self.prenom = prenom
        self.identifiant = Intervenant.getIdentifiant(self)
        self.status = status

        #Informations faculfatives
        self.adresse = adresse if adresse else ''
        self.employeur = employeur if employeur else ''

    @staticmethod
    def getIdentifiant(intervenant:Intervenant):
        Identifiant = f'{intervenant.nom[0]}{intervenant.prenom[0]}' #Concaténation de la première lettre du nom et prénom
        return Identifiant
