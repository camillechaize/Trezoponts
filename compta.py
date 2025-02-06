import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
import json
import os
import pdfplumber
from datetime import datetime
import shutil  # pour copier les fichiers
import pathlib  # pour gérer les noms de fichiers et de dossiers

FACTURES_ROOT_FOLDER = "factures"  # Dossier principal pour stocker les factures par compte
flag = True


# Classe représentant une opération
class Operation:
    def __init__(self, compte, moyen, nom, destinataire, montant, date, valeur=None, de=None, motif=None, ref=None, ref_2=None, ref_3=None,
                 pour=None, date_virement=None, remise=None, chez=None, lib=None, facture=None, repartition=None):
        self.compte = compte
        self.moyen = moyen
        self.nom = nom
        self.destinataire = destinataire
        self.montant = montant
        self.date = date  # Date de l'opération
        self.valeur = valeur  # Date de valeur
        self.de = de
        self.motif = motif
        self.ref = ref
        self.ref_2 = ref_2
        self.ref_3 = ref_3
        self.pour = pour
        self.date_virement = date_virement
        self.remise = remise
        self.chez = chez
        self.lib = lib
        self.facture = facture  # Chemin du fichier facture (s'il y en a un)
        self.repartition = repartition or []

    def to_dict(self):
        return {
            "compte": self.compte,
            "moyen": self.moyen,
            "nom": self.nom,
            "destinataire": self.destinataire,
            "montant": self.montant,
            "date": self.date.strftime("%d/%m/%Y"),
            "valeur": self.valeur.strftime("%d/%m/%Y") if self.valeur else None,
            "de": self.de,
            "motif": self.motif,
            "ref": self.ref,
            "ref_2": self.ref_2,
            "ref_3": self.ref_3,
            "pour": self.pour,
            "date_virement": self.date_virement,
            "remise": self.remise,
            "chez": self.chez,
            "lib": self.lib,
            "facture": self.facture,
            "repartition": self.repartition,
        }

    @classmethod
    def from_dict(cls, data):
        return cls(
            compte=data["compte"],
            moyen=data["moyen"],
            nom=data["nom"],
            destinataire=data["destinataire"],
            montant=data["montant"],
            date=datetime.strptime(data["date"], "%d/%m/%Y"),
            valeur=datetime.strptime(data["valeur"], "%d/%m/%Y") if data["valeur"] else None,
            de=data.get("de"),
            motif=data.get("motif"),
            ref=data.get("ref"),
            ref_2=data.get("ref_2"),
            ref_3=data.get("ref_3"),
            pour=data.get("pour"),
            date_virement=data.get("date_virement"),
            remise=data.get("remise"),
            chez=data.get("chez"),
            lib=data.get("lib"),
            facture=data.get("facture"),
            repartition=data.get("repartition", []),
        )

    def __repr__(self):
        return f"Operation({self.nom}, {self.date}, {self.montant})"


# Classe représentant une opération de cash
class CashOperation:
    def __init__(self, uni_id, nom, destinataire, montant, date, repartition=None):
        self.uni_id = uni_id
        self.nom = nom
        self.destinataire = destinataire
        self.montant = montant
        self.date = date  # Date de l'opération
        self.repartition = repartition or []

    def to_dict(self):
        return {
            "uni_id": self.uni_id,
            "nom": self.nom,
            "destinataire": self.destinataire,
            "montant": self.montant,
            "date": self.date.strftime("%d/%m/%Y"),
            "repartition": self.repartition,
        }

    @classmethod
    def from_dict(cls, data):
        return cls(
            uni_id=data["uni_id"],
            nom=data["nom"],
            destinataire=data["destinataire"],
            montant=data["montant"],
            date=datetime.strptime(data["date"], "%d/%m/%Y"),
            repartition=data.get("repartition", []),
        )

    def __repr__(self):
        return f"Operation({self.nom}, {self.date}, {self.montant})"


# Classe représentant un tiers
class Tiers:
    def __init__(self, nom_usage, noms_associes):
        self.nom_usage = nom_usage
        self.noms_associes = noms_associes  # Liste de noms associés

    def to_dict(self):
        return {
            "nom_usage": self.nom_usage,
            "noms_associes": self.noms_associes,
        }


class Event:
    def __init__(self, nom, couleur):
        self.nom = nom
        self.couleur = couleur

    def to_dict(self):
        return {
            "nom": self.nom,
            "couleur": self.couleur,
        }


class ComptaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Logiciel de Comptabilité")
        self.operations = []
        self.all_operations = []  # Liste pour stocker toutes les opérations des RDC (non filtrées)
        self.cash_operations = []  # Liste pour stocker toutes les opérations de Cash
        self.tiers = []
        self.events = []
        self.config = {"accounts": {}, "root_folder": None}
        self.page_num_operations = 0
        self.page_num_cash_operations = 0
        self.page_num_tiers = 0
        self.page_num_events = 0

        # Chargement des données à partir des fichiers JSON
        self.load_data()

        # Menu principal
        self.main_menu()

    def load_data(self):
        # Chargement des opérations
        try:
            with open("operations.json", "r") as f:
                operations_data = json.load(f)
                self.all_operations = [Operation.from_dict(op) for op in operations_data]
                self.operations = self.all_operations  # Par défaut, afficher toutes les opérations
        except FileNotFoundError:
            self.operations = []

        # Chargement des opérations de Cash
        try:
            with open("cash_operations.json", "r") as f:
                cash_operations_data = json.load(f)
                self.cash_operations = [CashOperation.from_dict(op) for op in cash_operations_data]
        except FileNotFoundError:
            self.cash_operations = []

        # Chargement des tiers
        try:
            with open("tiers.json", "r") as f:
                tiers_data = json.load(f)
                self.tiers = [Tiers(**tier) for tier in tiers_data]
        except FileNotFoundError:
            self.tiers = []

        # Chargement des événements
        try:
            with open("events.json", "r") as f:
                events_data = json.load(f)
                self.events = [Event(**event) for event in events_data]
        except FileNotFoundError:
            self.events = []

        # Chargement de la configuration (chemin des dossiers de relevés et relevés analysés)
        try:
            with open("config.json", "r") as f:
                self.config = json.load(f)
        except FileNotFoundError:
            self.config = {"accounts": {}, "root_folder": None}

    def save_data(self):
        # Sauvegarde des opérations
        with open("operations.json", "w") as f:
            json.dump([op.to_dict() for op in self.all_operations], f)

        # Sauvegarde des opérations de cash
        with open("cash_operations.json", "w") as f:
            json.dump([c_op.to_dict() for c_op in self.cash_operations], f)

        # Sauvegarde des tiers
        with open("tiers.json", "w") as f:
            json.dump([tier.to_dict() for tier in self.tiers], f)

        # Sauvegarde des événements
        with open("events.json", "w") as f:
            json.dump([event.to_dict() for event in self.events], f)

        # Sauvegarde de la configuration
        with open("config.json", "w") as f:
            json.dump(self.config, f)

    def main_menu(self):
        # Réinitialisation de la fenêtre
        for widget in self.root.winfo_children():
            widget.destroy()

        # Boutons principaux
        btn_operations = tk.Button(self.root, text="Consultation des opérations", command=self.open_operations)
        btn_operations.pack(pady=10)

        btn_tiers = tk.Button(self.root, text="Ajout de tiers", command=self.open_tiers)
        btn_tiers.pack(pady=10)

        btn_events = tk.Button(self.root, text="Événement", command=self.open_create_event_window)
        btn_events.pack(pady=10)

        btn_events = tk.Button(self.root, text="$", command=self.open_cash_operations_window)
        if flag: btn_events.pack(pady=10)

    def get_tiers_nom_usage(self, destinataire):
        """Renvoie le nom d'usage du tiers si le destinataire correspond à un tiers connu."""
        for tier in self.tiers:
            if destinataire in tier.noms_associes:
                return tier.nom_usage
        return destinataire  # Retourne le destinataire original si aucun tiers correspondant n'est trouvé.

    # region OPERATIONS
    def open_operations(self):
        if not self.config["root_folder"]:
            # Sélectionner le dossier racine au lancement si non sélectionné
            # noinspection PyTypedDict
            self.config["root_folder"] = filedialog.askdirectory(title="Sélectionner le dossier racine contenant les comptes")
            if not self.config["root_folder"]:
                messagebox.showerror("Erreur", "Aucun dossier racine sélectionné.")
                return
            self.save_data()

        # Fenêtre de consultation des opérations
        operations_window = tk.Toplevel(self.root)
        operations_window.title("Consultation des opérations")

        # Options de sélection de compte (dossiers dans le dossier racine)
        comptes_frame = tk.Frame(operations_window)
        comptes_frame.pack(side="right", padx=10, pady=10)

        tk.Label(comptes_frame, text="Sélection du compte :").pack()

        # Récupération des sous-dossiers du dossier racine pour les comptes
        # noinspection PyTypeChecker
        self.accounts = [d for d in os.listdir(self.config["root_folder"]) if os.path.isdir(os.path.join(self.config["root_folder"], d))]
        self.selected_account = tk.StringVar(value=self.accounts[0] if self.accounts else "")

        # Ajout des boutons radio pour les comptes
        for compte in self.accounts:
            tk.Radiobutton(comptes_frame, text=compte, variable=self.selected_account, value=compte,
                           command=self.update_operations_view).pack(anchor="w")

        # Bouton pour choisir le dossier des relevés
        btn_select_folder = tk.Button(comptes_frame, text="Analyser les comptes", command=self.select_releve_folder)
        btn_select_folder.pack(pady=5)

        # Bouton pour lier une facture à l'opération
        btn_add_invoice = tk.Button(comptes_frame, text="Ajouter une facture", command=self.attach_invoice)
        btn_add_invoice.pack(pady=5)

        # Bouton pour ouvrir la facture liée
        btn_open_invoice = tk.Button(comptes_frame, text="Voir la facture", command=self.open_invoice)
        btn_open_invoice.pack(pady=5)

        # Bouton pour ouvrir la répartition de l'opération
        btn_open_repartition = tk.Button(comptes_frame, text="Répartiton", command=self.open_repartition_window)
        btn_open_repartition.pack(pady=5)

        # Frame pour la liste des opérations
        operations_frame = tk.Frame(operations_window)
        operations_frame.pack(side="left", fill="both", expand=True)

        # Mise à jour des colonnes pour inclure la Date
        self.operations_tree = ttk.Treeview(operations_frame, columns=("ID", "Date", "MOY", "Nom", "Destinataire", "Montant", "Facture"),
                                            show="headings",
                                            height=15)
        self.operations_tree.pack(fill="both", expand=True)

        # Configuration des colonnes
        self.operations_tree.heading("ID", text="#")
        self.operations_tree.column("ID", anchor="e", width=5)
        self.operations_tree.heading("Date", text="Date")
        self.operations_tree.column("Date", anchor="center", width=70)
        self.operations_tree.heading("MOY", text="MOY")
        self.operations_tree.column("MOY", anchor="center", width=70)
        self.operations_tree.heading("Nom", text="Nom")
        self.operations_tree.heading("Destinataire", text="Destinataire")
        self.operations_tree.heading("Montant", text="Montant")
        self.operations_tree.heading("Facture", text="Facture")
        self.operations_tree.bind("<Double-1>", self.on_operation_double_click)

        pagination_frame = tk.Frame(operations_frame)
        pagination_frame.pack(pady=5)

        btn_prev = tk.Button(pagination_frame, text="Précédent", command=self.previous_page_operations)
        btn_prev.grid(row=0, column=0)
        self.month_page = tk.Label(pagination_frame, text=(self.page_num_operations + 1) % 12 + 1)
        self.month_page.grid(row=0, column=1)
        btn_next = tk.Button(pagination_frame, text="Suivant", command=self.next_page_operations)
        btn_next.grid(row=0, column=2)

        # Pagination des opérations
        self.load_operations_page()

    def update_operations_view(self):
        # Récupérer le compte sélectionné
        selected_account = self.selected_account.get()
        # Filtrer les opérations pour ce compte, mais sans modifier self.all_operations
        self.operations = [op for op in self.all_operations if op.compte == selected_account]  # Filtrage uniquement pour l'affichage
        # Réinitialiser l'affichage des opérations
        self.load_operations_page()

    def select_releve_folder(self):
        # Sélectionner le dossier principal contenant tous les sous-dossiers de comptes
        main_folder_path = self.config["root_folder"]
        if main_folder_path:
            # Parcourt les sous-dossiers du dossier principal et les lie à leurs comptes respectifs
            for account_name in os.listdir(main_folder_path):
                # noinspection PyTypeChecker
                account_path = os.path.join(main_folder_path, account_name)

                # Vérifier que chaque élément est bien un sous-dossier
                if os.path.isdir(account_path):
                    self.config["accounts"].setdefault(account_name, {"folder": account_path, "analyzed_files": []})

            # Sauvegarder la configuration mise à jour
            self.save_data()

            # Analyser les relevés pour chaque compte
            self.check_new_releves()

    def check_new_releves(self):
        new_operations_count = 0

        for account, account_info in self.config["accounts"].items():
            folder_path = account_info.get("folder")
            if folder_path:
                # Appel à la fonction d'analyse pour chaque sous-dossier correspondant à un compte
                new_operations = analyze_account_statements(account, folder_path)

                # Incrément du compteur d'opérations nouvellement ajoutées
                new_operations_count += len(new_operations)

        # Affichage du résultat à l'utilisateur
        if new_operations_count > 0:
            messagebox.showinfo("Nouveaux relevés détectés", f"{new_operations_count} opérations ajoutées depuis les nouveaux relevés.")
        else:
            messagebox.showinfo("Aucun nouveau relevé", "Aucun nouveau relevé à analyser dans les dossiers des comptes.")

        # Actualisation de l'affichage des opérations
        self.load_operations_page()

    def load_operations_page(self):
        self.operations_tree.delete(*self.operations_tree.get_children())
        self.month_page.config(text=(self.page_num_operations + 1) % 12 + 1)
        for i, op in enumerate(self.operations):
            # Afficher uniquement les opérations correspondant au mois de la page (la première page correspond à février)
            if op.date.month == (self.page_num_operations + 1) % 12 + 1:
                # Récupérer le nom d'usage du tiers si possible
                destinataire_affiche = self.get_tiers_nom_usage(op.destinataire)
                self.operations_tree.insert("", "end", values=(
                    i, op.date.strftime("%d/%m/%Y"), op.moyen, op.nom, destinataire_affiche, op.montant, op.facture),
                                            tags="rep" if len(op.repartition) > 0 else "")
        self.operations_tree.tag_configure("rep", background="salmon1")

    def previous_page_operations(self):
        if self.page_num_operations > 0:
            self.page_num_operations -= 1
            self.load_operations_page()

    def next_page_operations(self):
        if self.page_num_operations < 11:
            self.page_num_operations += 1
            self.load_operations_page()

    def on_operation_double_click(self, event):
        item_id = int(self.operations_tree.item(self.operations_tree.focus())['values'][0])
        operation = self.operations[item_id]
        details = (
            f"Date: {operation.date}\n"
            f"Valeur: {operation.valeur}\n"
            f"Nom: {operation.nom}\n"
            f"Destinataire: {operation.destinataire}\n"
            f"Montant: {operation.montant}\n"
            f"DE: {operation.de}\n"
            f"MOTIF: {operation.motif}\n"
            f"REF: {operation.ref}, {operation.ref_2}, {operation.ref_3}\n"
            f"POUR: {operation.pour}\n"
            f"DATE virement: {operation.date_virement}\n"
            f"REMISE: {operation.remise}\n"
            f"CHEZ: {operation.chez}\n"
            f"LIB: {operation.lib}\n"
            f"Facture: {operation.facture if operation.facture else 'Aucune'}"
        )
        messagebox.showinfo("Détails de l'opération", details)

    def attach_invoice(self):
        # Récupération de l'élément sélectionné
        item_id = self.operations_tree.focus()
        if item_id:
            operation_values = self.operations_tree.item(item_id, "values")
            operation_index = int(operation_values[0])
            operation = self.operations[operation_index]

            # Sélectionner le fichier de facture
            filepath = filedialog.askopenfilename(filetypes=[("All Files", "*.*")])
            if filepath:
                # Assurez-vous que le dossier principal pour les factures existe
                facture_root = pathlib.Path(FACTURES_ROOT_FOLDER)
                facture_root.mkdir(exist_ok=True)

                # Créez un sous-dossier pour le compte de l'opération
                compte_folder = facture_root / operation.compte
                compte_folder.mkdir(exist_ok=True)

                # Formatage du nom de fichier de la facture copiée
                date_str = operation.date.strftime("%m_%d_%Y")
                nom_operation = operation.nom.replace(" ", "_").replace("/", "_")  # Retire les espaces et / dans le nom
                montant_operation = str(operation.montant).replace(".", "")
                filename = f"{date_str}_{operation.compte}_{nom_operation}_{montant_operation}{pathlib.Path(filepath).suffix}"

                # Chemin cible dans le dossier des factures
                dest_path = compte_folder / filename
                shutil.copy(filepath, dest_path)

                # Mise à jour du chemin de la facture dans l'opération et sauvegarde
                operation.facture = str(dest_path)
                self.load_operations_page()  # Rafraîchir l'affichage
                self.save_data()

    def open_invoice(self):
        item_id = self.operations_tree.focus()
        operation_values = self.operations_tree.item(item_id, "values")
        if self.operations[int(operation_values[0])].facture:
            os.startfile(self.operations[int(operation_values[0])].facture)
        else:
            messagebox.showwarning("Facture manquante", "Aucune facture n'est liée à cette opération")

    # endregion

    # region CASH OPERATIONS
    def open_cash_operations_window(self):

        def add_cash_operation():
            motif = self.nom_var.get()

            try:
                montant = float(self.montant_var.get())
            except ValueError:
                messagebox.showerror("Erreur", "Veuillez entrer un montant valide.")
                return

            destinataire = self.selected_desti.get()
            if destinataire == "Autre":
                destinataire = self.destinataire_var.get()

            try:
                date = datetime.strptime(self.date_var.get(), '%d%m%Y')
            except ValueError:
                messagebox.showerror("Erreur", "Veuillez entrer un format de date valide (ddmmyyyy).")
                return

            self.cash_operations.append(CashOperation(int(datetime.now().timestamp()), motif, destinataire, montant, date))
            self.nom_var.delete(0, tk.END)
            self.montant_var.delete(0, tk.END)
            self.destinataire_var.delete(0, tk.END)
            self.selected_desti.set("Autre")
            self.save_data()
            self.load_cash_operations_page()

        def delete_cash_operation():
            # Récupère l'élément sélectionné
            selected_item = self.cash_operations_tree.selection()
            if not selected_item:
                messagebox.showwarning("Aucune sélection", "Veuillez sélectionner un élément à supprimer.")
                return

            # Récupère l'index de l'élément dans la liste et le supprime
            cash_operation_tag = self.cash_operations_tree.item(selected_item[0], "tags")[0]
            for c_op in self.cash_operations:
                if c_op.uni_id == int(cash_operation_tag):
                    self.cash_operations.remove(c_op)
                    break

            # Actualise la liste affichée
            self.save_data()
            self.load_cash_operations_page()

        # Fenêtre de consultation des opérations de cash
        cash_operations_window = tk.Toplevel(self.root)
        cash_operations_window.title("Coffre")

        # Panneau de droite pour créer/supprimer une opération et Répartir
        right_frame = tk.Frame(cash_operations_window)
        right_frame.pack(side="right", padx=10, pady=10)

        # region CREATION OPERATION
        # Panneau de création d'opération
        new_cash_operation_frame = tk.Frame(right_frame)
        new_cash_operation_frame.pack(side="top", padx=10, pady=10)

        # Motif
        tk.Label(new_cash_operation_frame, text="Motif").grid(row=0, column=0)
        self.nom_var = tk.Entry(new_cash_operation_frame)
        self.nom_var.grid(row=0, column=1)

        # Montant
        tk.Label(new_cash_operation_frame, text="Montant").grid(row=1, column=0)
        self.montant_var = tk.Entry(new_cash_operation_frame)
        self.montant_var.grid(row=1, column=1)

        # Destinataire
        tk.Label(new_cash_operation_frame, text="Destinataire").grid(row=2, column=0)
        self.selected_desti = tk.StringVar()
        self.selected_desti.set("Autre")
        # Menu déroulant pour choisir le destinataire
        if self.tiers:
            tk.OptionMenu(new_cash_operation_frame, self.selected_desti, *["Autre"] + [t.nom_usage for t in self.tiers]).grid(row=2, column=1)
        # Sinon le nom est rentré manuellement
        if self.selected_desti.get() == "Autre":
            self.destinataire_label = tk.Label(new_cash_operation_frame, text="nom si Autre")
            self.destinataire_label.grid(row=3, column=0)
            self.destinataire_var = tk.Entry(new_cash_operation_frame)
            self.destinataire_var.grid(row=3, column=1)

        # Date
        tk.Label(new_cash_operation_frame, text="Date (ddmmyyyy)").grid(row=4, column=0)
        self.date_var = tk.Entry(new_cash_operation_frame)
        self.date_var.grid(row=4, column=1)

        # Bouton d'ajout de l'opération
        btn_add_cash_operation = tk.Button(new_cash_operation_frame, text="Ajouter", command=add_cash_operation)
        btn_add_cash_operation.grid(row=5, columnspan=2, pady=5)
        # endregion

        # Bouton pour supprimer une opération
        btn_delete_cash_operation = tk.Button(right_frame, text="Supprimer", command=delete_cash_operation)
        btn_delete_cash_operation.pack(pady=5)
        # Bouton pour ouvrir la répartition de l'opération
        btn_open_repartition = tk.Button(right_frame, text="Répartition", command=lambda: self.open_repartition_window(cash=True))
        btn_open_repartition.pack(pady=5)

        # Frame pour la liste des opérations
        cash_operations_frame = tk.Frame(cash_operations_window)
        cash_operations_frame.pack(side="left", fill="both", expand=True)

        # Treeview pour présenter les opérations
        self.cash_operations_tree = ttk.Treeview(cash_operations_frame, columns=("ID", "Date", "Nom", "Destinataire", "Montant"),
                                                 show="headings",
                                                 height=15)
        self.cash_operations_tree.pack(fill="both", expand=True)

        # Configuration des colonnes
        self.cash_operations_tree.heading("ID", text="#")
        self.cash_operations_tree.column("ID", anchor="e", width=5)
        self.cash_operations_tree.heading("Date", text="Date")
        self.cash_operations_tree.column("Date", anchor="center", width=70)
        self.cash_operations_tree.heading("Nom", text="Nom")
        self.cash_operations_tree.heading("Destinataire", text="Destinataire")
        self.cash_operations_tree.heading("Montant", text="Montant")

        pagination_frame = tk.Frame(cash_operations_frame)
        pagination_frame.pack(pady=5)

        btn_prev = tk.Button(pagination_frame, text="Précédent", command=self.previous_page_cash_operations)
        btn_prev.grid(row=0, column=0)
        self.cash_page = tk.Label(pagination_frame, text=self.page_num_cash_operations + 1)
        self.cash_page.grid(row=0, column=1)
        btn_next = tk.Button(pagination_frame, text="Suivant", command=self.next_page_cash_operations)
        btn_next.grid(row=0, column=2)

        # Pagination des opérations
        self.load_cash_operations_page()

    def load_cash_operations_page(self):
        self.cash_operations_tree.delete(*self.cash_operations_tree.get_children())
        self.cash_page.config(text=self.page_num_cash_operations + 1)
        sorted_operations = sorted(self.cash_operations, key=lambda op: op.date)
        offset = self.page_num_cash_operations * 30
        for i, c_op in enumerate(sorted_operations[offset:offset + 30]):
            self.cash_operations_tree.insert("", "end", values=(
                i, c_op.date.strftime("%d/%m/%Y"), c_op.nom, c_op.destinataire, c_op.montant), tags=c_op.uni_id)

    def previous_page_cash_operations(self):
        if self.page_num_cash_operations > 0:
            self.page_num_cash_operations -= 1
            self.load_cash_operations_page()

    def next_page_cash_operations(self):
        if (self.page_num_cash_operations + 1) * 30 < len(self.cash_operations):
            self.page_num_cash_operations += 1
            self.load_cash_operations_page()

    # endregion

    # region REPARTITION
    def open_repartition_window(self, cash=False):
        item_id = self.operations_tree.focus() if not cash else self.cash_operations_tree.focus()
        if not item_id:
            messagebox.showwarning("Aucune opération", "Veuillez sélectionner une opération pour la répartition.")
            return
        operation_index = int(self.operations_tree.item(item_id, "values")[0]) if not cash else -1
        c_op_uni_id = int(self.cash_operations_tree.item(item_id, "tags")[0]) if cash else 0  # L'identifiant pour l'opération cash
        operation = self.operations[operation_index] if not cash else next(
            (cash_op for cash_op in self.cash_operations if cash_op.uni_id == c_op_uni_id), None)

        repartition_window = tk.Toplevel(self.root)
        repartition_window.title(f"Répartition pour l'opération : {operation.nom} - {operation.montant}€")

        def add_repartition():
            try:
                montant = float(self.repartition_montant.get())
                tier = self.selected_tier.get()
                event_name = self.selected_event.get()
                self.repartition_list.append((tier, montant, event_name))
                self.repartition_montant.delete(0, tk.END)
                update_list_repartition()
            except ValueError:
                messagebox.showerror("Erreur", "Veuillez entrer un montant valide.")

        def delete_repartition():
            # Récupère l'élément sélectionné dans l'affichage de la répartition
            selected_item = repartition_tree.selection()
            if not selected_item:
                messagebox.showwarning("Aucune sélection", "Veuillez sélectionner un élément à supprimer.")
                return

            # Récupère l'index de l'élément dans la liste et le supprime
            item_index = repartition_tree.index(selected_item[0])  # L'index de l'élément dans repartition_tree
            del self.repartition_list[item_index]  # Supprime le couple (tiers, montant) correspondant dans la liste

            # Actualise la liste affichée
            update_list_repartition()

        def complete_amount():
            self.repartition_montant.delete(0, tk.END)
            self.repartition_montant.insert(0, operation.montant - sum(repartition[1] for repartition in self.repartition_list))
            add_repartition()

        # Liste des tiers pour la répartition
        tiers_frame = tk.Frame(repartition_window)
        tiers_frame.pack(pady=5, padx=10)

        # Selection du tiers
        tk.Label(tiers_frame, text="Sélectionner un tiers").grid(row=0, column=0)
        self.selected_tier = tk.StringVar()
        self.selected_tier.set(self.tiers[0].nom_usage if self.tiers else "")
        if self.tiers:
            tk.OptionMenu(tiers_frame, self.selected_tier, *[t.nom_usage for t in self.tiers]).grid(row=0, column=1)

        # Selection de l'event
        self.selected_event = tk.StringVar()
        self.selected_event.set("Aucun")
        if self.events:
            tk.OptionMenu(tiers_frame, self.selected_event, *[e.nom for e in self.events]).grid(row=0, column=2)

        tk.Label(tiers_frame, text="Montant").grid(row=1, column=0)
        self.repartition_montant = tk.Entry(tiers_frame)
        self.repartition_montant.grid(row=1, column=1, sticky="w")
        tk.Button(tiers_frame, text="Compléter", command=complete_amount).grid(row=1, column=2, padx=10, pady=10, sticky="w")

        # Liste pour afficher la répartition actuelle
        self.repartition_list = operation.repartition
        repartition_tree = ttk.Treeview(repartition_window, columns=("Tiers", "Montant", "Événement"), show="headings")
        repartition_tree.heading("Tiers", text="Tiers")
        repartition_tree.heading("Montant", text="Montant")
        repartition_tree.heading("Événement", text="Événement")
        repartition_tree.pack(pady=5, padx=10, fill="x")

        tk.Button(tiers_frame, text="Ajouter à la répartition", command=add_repartition).grid(row=2, column=0, pady=10, sticky="w")
        tk.Button(tiers_frame, text="Supprimer", command=delete_repartition).grid(row=2, column=1, pady=10, sticky="e")
        tiers_frame.columnconfigure(0, weight=1)
        tiers_frame.columnconfigure(1, weight=1)

        def update_list_repartition():
            repartition_tree.delete(*repartition_tree.get_children())
            for i, rep in enumerate(self.repartition_list):
                repartition_tree.insert("", "end", values=(rep[0], rep[1], rep[2]))

        def save_repartition():
            operation.repartition = self.repartition_list
            self.save_data()
            repartition_window.destroy()
            if not cash: self.update_operations_view()

        tk.Button(repartition_window, text="Enregistrer la répartition", command=save_repartition).pack(pady=10)
        update_list_repartition()

    # endregion

    # region TIERS
    def open_tiers(self):
        # Fenêtre de gestion de tiers
        tiers_window = tk.Toplevel(self.root)
        tiers_window.title("Gestion de tiers")

        # Frame pour la liste des tiers
        tiers_frame = tk.Frame(tiers_window)
        tiers_frame.pack(fill="both", expand=True)

        self.tiers_tree = ttk.Treeview(tiers_frame, columns=("Nom d'usage", "Noms associés"), show="headings", height=15)
        self.tiers_tree.pack(fill="both", expand=True)

        # Configuration des colonnes
        self.tiers_tree.heading("Nom d'usage", text="Nom d'usage")
        self.tiers_tree.heading("Noms associés", text="Noms associés")

        # Pagination des tiers
        self.load_tiers_page()

        pagination_frame = tk.Frame(tiers_frame)
        pagination_frame.pack(pady=5)

        btn_prev_tiers = tk.Button(pagination_frame, text="Précédent", command=self.previous_page_tiers)
        btn_prev_tiers.grid(row=0, column=0)
        btn_next_tiers = tk.Button(pagination_frame, text="Suivant", command=self.next_page_tiers)
        btn_next_tiers.grid(row=0, column=1)

        # Section pour ajouter un nouveau tiers
        add_tiers_frame = tk.Frame(tiers_window)
        add_tiers_frame.pack(pady=10)

        tk.Label(add_tiers_frame, text="Nom d'usage").grid(row=0, column=0)
        self.nom_usage_var = tk.Entry(add_tiers_frame)
        self.nom_usage_var.grid(row=0, column=1)

        tk.Label(add_tiers_frame, text="Noms associés (séparés par des virgules)").grid(row=1, column=0)
        self.noms_associes_var = tk.Entry(add_tiers_frame)
        self.noms_associes_var.grid(row=1, column=1)

        btn_add_tiers = tk.Button(add_tiers_frame, text="Ajouter le tiers", command=self.add_tiers)
        btn_add_tiers.grid(row=2, columnspan=2, pady=5)

    def load_tiers_page(self):
        self.tiers_tree.delete(*self.tiers_tree.get_children())
        offset = self.page_num_tiers * 30
        for tiers in self.tiers[offset:offset + 30]:
            noms_associes_str = ", ".join(tiers.noms_associes)
            self.tiers_tree.insert("", "end", values=(tiers.nom_usage, noms_associes_str))

    def previous_page_tiers(self):
        if self.page_num_tiers > 0:
            self.page_num_tiers -= 1
            self.load_tiers_page()

    def next_page_tiers(self):
        if (self.page_num_tiers + 1) * 30 < len(self.tiers):
            self.page_num_tiers += 1
            self.load_tiers_page()

    def add_tiers(self):
        nom_usage = self.nom_usage_var.get()
        noms_associes = self.noms_associes_var.get().split(",")
        self.tiers.append(Tiers(nom_usage, noms_associes))
        self.load_tiers_page()
        self.save_data()
        self.nom_usage_var.delete(0, tk.END)
        self.noms_associes_var.delete(0, tk.END)

    # endregion

    # region EVENTS
    def open_create_event_window(self):
        event_window = tk.Toplevel(self.root)
        event_window.title("Gestion d'événements")

        # Frame principale pour la liste des événements
        events_frame = tk.Frame(event_window)
        events_frame.pack(fill="both", expand=True)

        self.events_tree = ttk.Treeview(events_frame, columns=("Nom", "Couleur"), show="headings", height=15)
        self.events_tree.pack(fill="both", expand=True)
        self.events_tree.bind("<Double-1>", self.on_event_double_click)

        # Configuration des colonnes
        self.events_tree.heading("Nom", text="Nom")
        self.events_tree.heading("Couleur", text="Couleur")

        # Pagination des événements
        self.load_events_page()

        # Frame pour la pagination
        pagination_frame = tk.Frame(event_window)  # Utilisation de `event_window` comme parent
        pagination_frame.pack(pady=5)

        btn_prev_events = tk.Button(pagination_frame, text="Précédent", command=self.previous_page_events)
        btn_prev_events.grid(row=0, column=0)
        btn_next_events = tk.Button(pagination_frame, text="Suivant", command=self.next_page_events)
        btn_next_events.grid(row=0, column=1)

        # Frame pour ajouter un nouvel événement
        add_event_frame = tk.Frame(event_window)
        add_event_frame.pack(pady=10)

        tk.Label(add_event_frame, text="Nom").grid(row=0, column=0)
        self.nom_var = tk.Entry(add_event_frame)
        self.nom_var.grid(row=0, column=1)

        tk.Label(add_event_frame, text="Couleur").grid(row=1, column=0)

        # Variable pour stocker la couleur sélectionnée
        self.event_color = tk.StringVar(value="#FFFFFF")  # Valeur par défaut

        # Fonction pour ouvrir le sélecteur de couleur
        def choose_color():
            color_code = colorchooser.askcolor(title="Choisir une couleur")[1]
            if color_code:
                # noinspection PyTypeChecker
                self.event_color.set(color_code)

        # Bouton pour ouvrir le sélecteur de couleur
        tk.Button(add_event_frame, text="Choisir couleur", command=choose_color).grid(row=1, column=1)

        # Bouton pour enregistrer l'événement
        tk.Button(add_event_frame, text="Enregistrer l'événement", command=self.add_event).grid(row=2, column=0, padx=5, pady=10)

        # Bouton pour supprimer l'événement
        tk.Button(add_event_frame, text="Supprimer l'événement", command=self.delete_event).grid(row=2, column=1, padx=5, pady=10)

        # Bouton pour supprimer l'événement
        tk.Button(add_event_frame, text="Supprimer l'événement", command=self.delete_event).grid(row=2, column=1, padx=5, pady=10)

        tk.Label(add_event_frame, text="Date début (ddmmyyyy)").grid(row=3, column=0)
        self.date_events_var = tk.Entry(add_event_frame)
        self.date_events_var.grid(row=3, column=1)

    def load_events_page(self):
        self.events_tree.delete(*self.events_tree.get_children())
        offset = self.page_num_events * 30
        for event in self.events[offset:offset + 30]:
            self.events_tree.insert("", "end", values=(event.nom, event.couleur), tags=event.couleur)
            self.events_tree.tag_configure(event.couleur, background=event.couleur)

    def previous_page_events(self):
        if self.page_num_events > 0:
            self.page_num_events -= 1
            self.load_events_page()

    def next_page_events(self):
        if (self.page_num_events + 1) * 30 < len(self.events):
            self.page_num_events += 1
            self.load_events_page()

    def add_event(self):
        event_name = self.nom_var.get()
        if not event_name:
            messagebox.showwarning("Nom manquant", "Veuillez entrer un nom pour l'événement.")
            return
        # Ajoute l'événement à la liste des événements
        self.events.append(Event(event_name, self.event_color.get()))
        self.nom_var.delete(0, tk.END)
        self.save_data()
        self.load_events_page()

    def delete_event(self):
        # Récupère l'élément sélectionné dans l'affichage des events
        selected_item = self.events_tree.selection()
        if not selected_item:
            messagebox.showwarning("Aucune sélection", "Veuillez sélectionner un élément à supprimer.")
            return

        # Récupère l'index de l'élément dans la liste et le supprime
        item_index = self.events_tree.index(selected_item[0])  # L'index de l'élément dans le tree
        del self.events[item_index]  # Supprime l'event correspondant dans la liste

        # Actualise la liste affichée
        self.save_data()
        self.load_events_page()

    def on_event_double_click(self, event):
        # Récupère l'événement sélectionné par un double-clic dans l'interface
        selected_item = self.events_tree.focus()
        if not selected_item:
            messagebox.showwarning("Aucun événement sélectionné", "Veuillez sélectionner un événement.")
            return

        # Récupère le nom de l'événement depuis l'interface et la date de début de comptage (utile pour les clubs)
        event_name = self.events_tree.item(selected_item, "values")[0]
        try:
            date_events = datetime.strptime(self.date_events_var.get(), '%d%m%Y')
        except ValueError:
            messagebox.showerror("Erreur", "Veuillez entrer un format de date valide (ddmmyyyy).")
            return

        # Initialiser le total général et un dictionnaire pour les détails par tiers
        total_recettes = 0.0
        total_charges = 0.0
        tiers_summary = {}

        # Parcourir toutes les opérations pour trouver celles liées à cet événement
        for operation in self.all_operations + self.cash_operations:
            if date_events < operation.date:
                for tier, montant, event in operation.repartition:
                    if event == event_name:  # Vérifier si la répartition est liée à l'événement sélectionné
                        # Initialiser les détails du tiers si ce n'est pas déjà fait
                        if tier not in tiers_summary:
                            tiers_summary[tier] = {"recettes": 0.0, "charges": 0.0, "total": 0.0}

                        # Mettre à jour les totaux en fonction du signe du montant
                        tiers_summary[tier]["total"] += montant
                        if montant > 0:
                            tiers_summary[tier]["recettes"] += montant
                            total_recettes += montant
                        else:
                            tiers_summary[tier]["charges"] += montant
                            total_charges += montant

        # Affichage des résultats
        details_window = tk.Toplevel(self.root)
        details_window.title(f"Détails de l'événement : {event_name}")
        details_window.geometry("800x400")

        # Treeview pour afficher les données
        tree = ttk.Treeview(details_window, columns=("Tiers", "Recettes", "Charges", "Total"), show="headings", height=15)
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        tree.heading("Tiers", text="Tiers")
        tree.heading("Recettes", text="Recettes (€)")
        tree.heading("Charges", text="Charges (€)")
        tree.heading("Total", text="Total (€)")

        # Détails des tiers
        for tier, details in tiers_summary.items():
            tree.insert("", "end", values=(
                tier,
                f"{details['recettes']:.2f}",
                f"{details['charges']:.2f}",
                f"{details['total']:.2f}"
            ))

        # Ajouter la ligne de total général en bas
        tree.insert("", "end", values=(
            "TOTAL",
            f"{total_recettes:.2f}",
            f"{total_charges:.2f}",
            f"{total_recettes + total_charges:.2f}"
        ), tags="Total")
        tree.tag_configure("Total", foreground="red", font=("Helvetica", 10, "bold"))

    # endregion


def extract_text_from_pdf(path):
    text = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            table = page.extract_table(table_settings={"vertical_strategy": "lines", "horizontal_strategy": "text"})
            if table is not None:  # Vérifie si une table a été trouvée
                text += table  # Ajoute la table extraite si elle existe
            else:
                print("Pas de tableau trouvé pour: ", path)
    return text


def analyze_account_statements(account, folder_path):
    account_info = app.config["accounts"].get(account, {})
    analyzed_files = set(account_info.get("analyzed_files", []))
    new_operations = []

    # On trie les relevés de compte présents dans le dossier par date afin d'ajouter les opérations dans le bon ordre
    sorted_pdfs_filenames = [pdf_name for pdf_name in os.listdir(folder_path)]
    sorted_pdfs_filenames.sort(key=lambda x: datetime.strptime(x.split('.')[0].split('_')[-1], "%d%m%Y"))

    for pdf_filename in sorted_pdfs_filenames:
        if not pdf_filename.endswith('.pdf') or pdf_filename in analyzed_files:
            continue

        path = os.path.join(folder_path, pdf_filename)
        text = extract_text_from_pdf(path)
        current_operation = None

        for l in text[3: len(text) - 1]:
            if l[0] != '' and l[0] is not None:
                if current_operation:
                    new_operations.append(current_operation)
                date = datetime.strptime(l[0], "%d/%m/%Y")
                valeur = datetime.strptime(l[1], "%d/%m/%Y")
                nom = l[2]
                moyen = "CARTE" if "CARTE" in nom else "VIR" if "VIR" in nom else "CHEQUE" if "CHEQUE" in nom else "_"
                debit = None if l[3] == '' else str_to_float(l[3])
                credit = None if l[4] == '' else str_to_float(l[4])
                montant = credit if credit is not None else -debit

                current_operation = Operation(compte=account, moyen=moyen,
                                              nom=nom, destinataire="", montant=montant, date=date, valeur=valeur
                                              )

            elif l[2] is not None:
                if l[2].startswith("DE:"):
                    current_operation.de = l[2].replace("DE:", "", 1).strip()
                    current_operation.destinataire = current_operation.de
                elif l[2].startswith("MOTIF:"):
                    current_operation.motif = l[2].replace("MOTIF:", "", 1).strip()
                elif l[2].startswith("REF:"):
                    if current_operation.ref is None:
                        current_operation.ref = l[2].replace("REF:", "", 1).strip()
                    elif current_operation.ref_2 is None:
                        current_operation.ref_2 = l[2].replace("REF:", "", 1).strip()
                    elif current_operation.ref_3 is None:
                        current_operation.ref_3 = l[2].replace("REF:", "", 1).strip()
                elif l[2].startswith("POUR:"):
                    current_operation.pour = l[2].replace("POUR:", "", 1).strip()
                    current_operation.destinataire = current_operation.pour
                elif l[2].startswith("DATE:"):
                    current_operation.date_virement = l[2].replace("DATE:", "", 1).strip()
                elif l[2].startswith("REMISE:"):
                    current_operation.remise = l[2].replace("REMISE:", "", 1).strip()
                elif l[2].startswith("CHEZ:"):
                    current_operation.chez = l[2].replace("CHEZ:", "", 1).strip()
                elif l[2].startswith("LIB:"):
                    current_operation.lib = l[2].replace("LIB:", "", 1).strip()

        if current_operation:
            new_operations.append(current_operation)

        analyzed_files.add(pdf_filename)

    app.all_operations.extend(new_operations)
    app.config["accounts"][account]["analyzed_files"] = list(analyzed_files)
    app.save_data()
    return new_operations


def str_to_float(text: str) -> float:
    return float(text.replace('*', '').replace('.', '').replace(',', '.'))


# Exécution de l'application
root = tk.Tk()
app = ComptaApp(root)
root.mainloop()
