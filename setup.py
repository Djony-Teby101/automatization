import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
from tkinter import ttk
import os
import sys
from datetime import datetime
import numpy as np

class MainMenuApp:
    def __init__(self):
        # Configuration de l'interface
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        # Création de la fenêtre principale
        self.root = ctk.CTk()
        self.root.title("Système de Traitement Excel")
        self.root.geometry("12000x800")
        
        
        
        # Set the window to full-screen mode
        self.root.attributes('-fullscreen', True)


        # Bind the Escape key to exit full-screen mode
        # This allows the user to easily close the full-screen window
        def exit_fullscreen(event=None):
            self.root.attributes('-fullscreen', False)
            # You might also want to destroy the window or exit the application
            # root.destroy() 
        self.root.bind('<Escape>', exit_fullscreen)
         # Create a CTkScrollableFrame
       
        
        # Gestion de la fermeture de la fenêtre
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        
        # Variables de données (mise à jours)
        self.original_df = None
        self.filtered_df = None
        self.processed_df = None
        self.final_df = None
        self.file_path = None
        
        # Variables d'état
        self.is_processed = False
        self.soworkflow_values = []
        self.selected_soworkflow = None
        self.is_filter_applied = False
        
        
        
        # Référence à l'application OC (pour fermeture propre)
        self.oc_app = None
        self.oc_drgt = None
        
        # Création de l'interface du menu principal
        self.create_main_menu()
        
    def create_main_menu(self):
        """Crée le menu principal avec les boutons"""
        # Frame principal
        main_frame = ctk.CTkFrame(self.root)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
    

        
        # Titre de l'application
        title_label = ctk.CTkLabel(main_frame, 
                                  text="Traitement des Instances FTTH.", 
                                  font=ctk.CTkFont(size=24, weight="bold"))
        title_label.pack(pady=40)
        
        # Sous-titre
        subtitle_label = ctk.CTkLabel(main_frame, 
                                     text="Sélectionnez le module de traitement", 
                                     font=ctk.CTkFont(size=14),
                                     text_color="gray")
        subtitle_label.pack(pady=10)
        
        # Frame pour les boutons
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(pady=30)
        
        # Bouton 1: Traitement des OC (notre interface existante)
        oc_button = ctk.CTkButton(button_frame, 
                                 text="Traitement Instances OC", 
                                 command=self.open_oc_processor,
                                  width=250,
                                  height=60,
                                 font=ctk.CTkFont(size=14),
                                 corner_radius=10,
                                 fg_color="#353836",
                                 hover_color="#59615B")
        oc_button.pack(pady=15)
        
        # Bouton 2: Rapport des Activités (placeholder)
        report_button = ctk.CTkButton(button_frame, 
                                     text="Traitement DRGTS", 
                                     command=self.open_reports,
                                     width=250,
                                     height=60,
                                     font=ctk.CTkFont(size=14),
                                     corner_radius=10,
                                     fg_color="#353836",
                                     hover_color="#59615B")
        report_button.pack(pady=15)
        
        # Bouton 3: Configuration (placeholder)
        config_button = ctk.CTkButton(button_frame, 
                                     text="Production", 
                                     command=self.open_configuration,
                                     state="disabled",
                                     width=250,
                                     height=60,
                                     font=ctk.CTkFont(size=16),
                                     corner_radius=10,
                                     fg_color="#353836",
                                     hover_color="#59615B",
                                     )
        config_button.pack(pady=15)
       
        
        # Bouton 4: Configuration (quit l'app)
        btn_quit=ctk.CTkButton(button_frame,
            text="Quitter",
            command=self.root.quit,
            height=40,
            font=ctk.CTkFont(size=14),
            fg_color="#353836",
            hover_color="#59615B"
        )
        btn_quit.pack(pady=16)
        
        # Footer avec informations
        footer_label = ctk.CTkLabel(main_frame, 
                                   text="© 2025 Système de Traitement Excel - Version 1.0", 
                                   font=ctk.CTkFont(size=10),
                                   text_color="darkgray")
        footer_label.pack(side="bottom", pady=10)
    
    def open_oc_processor(self):
        """Ouvre l'interface de traitement des OC"""
        self.root.withdraw()  # Cache le menu principal
        self.oc_app = OCProcessorApp(self)  # Passe la référence du menu principal
        self.oc_app.run()
    
    def open_reports(self):
        """Ouvre l'interface de traitement des Drgts"""
        self.root.withdraw()  # Cache le menu principal
        self.oc_drgt = DrgtProcessorApp(self)  # Passe la référence du menu principal
        self.oc_drgt.run()
    
    def open_configuration(self):
        """Ouvre l'interface de configuration (placeholder)"""
        messagebox.showinfo("Configuration", "Module de configuration en développement...")
    
    def show_main_menu(self):
        """Affiche à nouveau le menu principal"""
        self.oc_app = None  # Réinitialise la référence
        self.root.deiconify()  # Réaffiche la fenêtre principale
    
    def on_closing(self):
        """Gère la fermeture propre de l'application"""
        # Ferme d'abord l'application OC si elle est ouverte
        if self.oc_app and self.oc_app.root.winfo_exists():
            self.oc_app.root.destroy()
        
        # Ferme l'application principale
        self.root.destroy()
        sys.exit(0)  # Arrête proprement le programme
    
    def run(self):
        """Lance l'application"""
        self.root.mainloop()

class OCProcessorApp:
    def __init__(self, main_menu_app):
        # Référence au menu principal
        self.main_menu = main_menu_app
        
        # Configuration de l'interface
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        # Création de la fenêtre de traitement OC
        self.root = ctk.CTk()
        self.root.title("Traitement des OC - Filtre RatePlan")
        self.root.geometry("1200x800")
        
        
        # Gestion de la fermeture de la fenêtre
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Variables
        self.original_df = None
        self.processed_df = None
        self.filtered_df = None
        self.file_path = None
        self.is_processed = False
        
        # Variables d'état(Filtre mise à jours)
        self.is_processed = False
        self.soworkflow_values = []
        self.selected_soworkflow = None
        self.is_filter_applied = False
        
        # Création de l'interface
        self.create_widgets()
        
    def create_widgets(self):
        """Crée l'interface de traitement des OC"""
        # Frame principal
        main_frame = ctk.CTkFrame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Header avec bouton retour
        header_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        header_frame.pack(fill="x", padx=20, pady=10)
        
        # Bouton retour au menu principal
        back_button = ctk.CTkButton(header_frame, 
                                   text="← Retour au Menu", 
                                   command=self.return_to_main_menu,
                                   width=150,
                                   fg_color="gray",
                                   hover_color="darkgray")
        back_button.pack(side="left")
        
        # Titre
        title_label = ctk.CTkLabel(header_frame, 
                                  text="Traitement des Instances Oc", 
                                  font=ctk.CTkFont(size=16, weight="bold"))
        title_label.pack(side="left", padx=20)
        
        # Bouton fermer l'application
        close_button = ctk.CTkButton(header_frame, 
                                    text="❌ Fermer l'Application", 
                                    command=self.close_application,
                                    width=180,
                                    fg_color="#C41E3A",
                                    hover_color="#A61E3A")
        close_button.pack(side="right")
        
        # Frame pour les boutons d'action
        button_frame = ctk.CTkFrame(main_frame)
        button_frame.pack(fill="x", padx=20, pady=10)
        
        # Bouton pour téléverser le fichier
        upload_btn = ctk.CTkButton(button_frame, text="📁 Téléverser le Fichier Excel", 
                                  command=self.upload_file, width=200)
        upload_btn.pack(side="left", padx=5, pady=5)
        
        # Bouton pour appliquer le filtre RatePlan
        filter_btn = ctk.CTkButton(button_frame, text="🔍 Filtrer RatePlan", 
                                  command=self.apply_rateplan_filter,
                                  width=180,
                                  state="disabled")
        filter_btn.pack(side="left", padx=5, pady=5)
        self.filter_btn = filter_btn
        
        # Bouton pour appliquer le traitement
        process_btn = ctk.CTkButton(button_frame, text="⚙️ Appliquer le Traitement", 
                                   command=self.apply_processing, width=200, state="disabled")
        process_btn.pack(side="left", padx=5, pady=5)
        self.process_btn = process_btn
        
        # Bouton pour appliquer le Filtre individuel.
        step4_label = ctk.CTkLabel(button_frame, text="Recherche par:", font=ctk.CTkFont(size=12, weight="bold"))
        step4_label.pack(side="left", padx=(20, 5), pady=5)
        
        
        self.soworkflow_combobox = ctk.CTkComboBox(button_frame, 
                                                  values=["Tous les statuts"],
                                                  width=200,
                                                  state="disabled",
                                                  command=self.on_soworkflow_selected
                                                  )
        self.soworkflow_combobox.pack(side="left", padx=5, pady=5)
        self.soworkflow_combobox.set("Tous les statuts")
        
        self.soworkflow_filter_btn = ctk.CTkButton(button_frame, 
                                                  text="🔎 Appliquer Filtre", 
                                                  command=self.apply_soworkflow_filter,
                                                  width=150, 
                                                  state="disabled")
        self.soworkflow_filter_btn.pack(side="left", padx=5, pady=5)
        
        
        
        
        # Bouton pour créer le nouveau fichier
        save_btn = ctk.CTkButton(button_frame, text="💾 Créer Nouveau Fichier", 
                                command=self.save_file, width=200, state="disabled")
        save_btn.pack(side="left", padx=5, pady=5)
        self.save_btn = save_btn
        
        # Frame pour les onglets
        tabview = ctk.CTkTabview(main_frame)
        tabview.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Onglets
        tabview.add("Aperçu Original")
        tabview.add("Aperçu Filtre RatePlan")
        tabview.add("Aperçu Traité")
        tabview.add("Statistiques")
        
        # Treeview pour l'aperçu des données originales
        self.original_tree = ttk.Treeview(tabview.tab("Aperçu Original"))
        self.setup_treeview(self.original_tree, tabview.tab("Aperçu Original"))
        
        # Treeview pour l'aperçu des données filtrées
        self.filtered_tree = ttk.Treeview(tabview.tab("Aperçu Filtre RatePlan"))
        self.setup_treeview(self.filtered_tree, tabview.tab("Aperçu Filtre RatePlan"))
        
        # Treeview pour l'aperçu des données traitées
        self.processed_tree = ttk.Treeview(tabview.tab("Aperçu Traité"))
        self.setup_treeview(self.processed_tree, tabview.tab("Aperçu Traité"))
        
        # Frame pour les statistiques
        stats_frame = ctk.CTkFrame(tabview.tab("Statistiques"))
        stats_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.stats_text = ctk.CTkTextbox(stats_frame, height=200)
        self.stats_text.pack(fill="both", expand=True, padx=10, pady=10)
        self.stats_text.insert("1.0", "Aucune statistique disponible. Veuillez charger un fichier.")
        
        # Frame pour les informations
        info_frame = ctk.CTkFrame(main_frame)
        info_frame.pack(fill="x", padx=20, pady=10)
        
        # Informations sur le fichier
        self.info_label = ctk.CTkLabel(info_frame, text="📊 Aucun fichier chargé", 
                                      font=ctk.CTkFont(size=12))
        self.info_label.pack(pady=5)
        
        # Status du traitement
        self.status_label = ctk.CTkLabel(info_frame, text="Status: En attente de fichier", 
                                        text_color="gray", font=ctk.CTkFont(size=10))
        self.status_label.pack(pady=2)
    
    def setup_treeview(self, tree, parent):
        """Configure un treeview avec scrollbars"""
        v_scrollbar = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        h_scrollbar = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        tree.pack(side="left", fill="both", expand=True)
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")
    
    def return_to_main_menu(self):
        """Retourne au menu principal"""
        self.root.destroy()  # Ferme la fenêtre de traitement
        self.main_menu.show_main_menu()  # Réaffiche le menu principal
    
    def close_application(self):
        """Ferme complètement l'application"""
        # Demande confirmation à l'utilisateur
        if messagebox.askyesno("Fermeture", "Êtes-vous sûr de vouloir quitter l'application ?"):
            # Ferme cette fenêtre
            self.root.destroy()
            # Ferme aussi le menu principal s'il existe
            if hasattr(self.main_menu, 'root') and self.main_menu.root.winfo_exists():
                self.main_menu.root.destroy()
            sys.exit(0)  # Arrête proprement le programme
    
    def on_closing(self):
        """Gère la fermeture de la fenêtre OC"""
        # Retourne au menu principal au lieu de fermer l'application
        self.return_to_main_menu()
    
    def upload_file(self):
        """Téléverse un fichier Excel"""
        file_path = filedialog.askopenfilename(
            title="Sélectionner un fichier Excel",
            filetypes=[("Fichiers Excel", "*.xlsx *.xls"), ("Tous les fichiers", "*.*")]
        )
        
        if file_path:
            try:
                self.file_path = file_path
                self.original_df = pd.read_excel(file_path)
                self.filtered_df = None
                self.processed_df = None
                self.final_df = None
                self.is_processed = False
                
                self.selected_soworkflow = None
                self.is_filter_applied = False
                
                self.update_previews()
                self.filter_btn.configure(state="normal")
                self.process_btn.configure(state="disabled")
                self.save_btn.configure(state="disabled")
                
                file_info = f"📊 Fichier: {os.path.basename(file_path)} - {len(self.original_df)} lignes - {len(self.original_df.columns)} colonnes"
                self.info_label.configure(text=file_info)
                self.status_label.configure(text="Status: Fichier chargé - Prêt pour le filtre RatePlan", text_color="green")
                
                self.update_stats()
                
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de la lecture du fichier: {str(e)}")
    
    
    
    def on_soworkflow_selected(self, choice):
        """Callback quand une valeur COCodeDescription est sélectionnée"""
        self.selected_soworkflow = choice if choice != "Tous les statuts" else None
    
    def apply_soworkflow_filter(self):
        """Étape 4 FACULTATIVE: Applique le filtre COCodeDescription"""
        if self.processed_df is None:
            messagebox.showwarning("Avertissement", "Veuillez d'abord appliquer le traitement (Étape 3)")
            return
        
        # Si "Tous les statuts" est sélectionné, pas de filtre (toutes les données)
        if self.selected_soworkflow is None or self.selected_soworkflow == "Tous les statuts":
            # Pas de filtre appliqué, on utilise toutes les données traitées
            self.final_df = self.processed_df.copy()
            self.is_filter_applied = False
            
            messagebox.showinfo("Aucun filtre", 
                              "✅ Aucun filtre spécifique appliqué.\n"
                              "Toutes les données traitées seront sauvegardées.")
            
            self.status_label.configure(text="Status: Aucun filtre COCodeDescription appliqué - Prêt pour la sauvegarde", text_color="green")
            self.update_stats()
            return
        
        try:
            # Appliquer le filtre COCodeDescription sur les données traitées
            self.final_df = self.processed_df[self.processed_df['COCodeDescription'] == self.selected_soworkflow].copy()
            self.is_filter_applied = True
            
            if len(self.final_df) == 0:
                messagebox.showwarning("Avertissement", 
                                     f"Aucune donnée trouvée pour: {self.selected_soworkflow}")
                return
            
            messagebox.showinfo("Filtre appliqué", 
                              f"✅ Filtre facultatif appliqué avec succès!\n"
                              f"Valeur sélectionnée: {self.selected_soworkflow}\n"
                              f"Lignes filtrées: {len(self.final_df)}/{len(self.processed_df)}\n"
                              f"Prêt pour la sauvegarde (Étape 5)")
            
            self.status_label.configure(text=f"Status: Filtre COCodeDescription appliqué ({self.selected_soworkflow}) - Prêt pour la sauvegarde", text_color="green")
            self.update_stats()
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du filtrage COCodeDescription: {str(e)}")
    
    
    
    def apply_rateplan_filter(self):
        """Applique le filtre sur RatePlan"""
        if self.original_df is None:
            messagebox.showwarning("Avertissement", "Veuillez d'abord charger un fichier")
            return
        
        try:
            # Définir les valeurs RatePlan à filtrer
            rateplan_values = ["FTTH DUO MAITRISE", "FTTH DUO"]
            
            if 'RatePlan' in self.original_df.columns:
                # Appliquer le filtre
                self.filtered_df = self.original_df[self.original_df['RatePlan'].isin(rateplan_values)].copy()
                print(len(self.filtered_df))
                self.filtered_df=self.filtered_df.drop_duplicates("UnitID1")
                print(len(self.filtered_df))
                
                if len(self.filtered_df) == 0:
                    messagebox.showwarning("Avertissement", 
                                         f"Aucune donnée trouvée pour les RatePlan: {rateplan_values}")
                    return
                
                self.processed_df = self.filtered_df.copy()
                self.update_treeview(self.filtered_tree, self.filtered_df)
                self.process_btn.configure(state="normal")
                
                messagebox.showinfo("Filtre appliqué", 
                                  f"Filtre RatePlan appliqué avec succès!\n"
                                  f"Lignes filtrées: {len(self.filtered_df)}/{len(self.original_df)}\n"
                                  f"RatePlan: {rateplan_values}")
                
                self.status_label.configure(text="Status: Filtre RatePlan appliqué - Prêt pour le traitement SOWorkflowStage", text_color="blue")
                self.update_stats()
                
            else:
                messagebox.showwarning("Avertissement", "Colonne 'RatePlan' non trouvée dans le fichier")
                
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du filtrage: {str(e)}")
    
    def update_previews(self):
        """Met à jour les aperçus des données"""
        self.update_treeview(self.original_tree, self.original_df)
        self.update_treeview(self.filtered_tree, self.filtered_df)
        self.update_treeview(self.processed_tree, self.processed_df)
    
    def update_treeview(self, tree, df):
        """Met à jour un treeview spécifique avec des données"""
        # Clear existing data
        for item in tree.get_children():
            tree.delete(item)
        
        if df is None:
            return
        
        # Set columns
        columns = list(df.columns)
        tree["columns"] = columns
        tree["show"] = "headings"
        
        # Configure columns
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, minwidth=80)
        
        # Add sample data (first 30 rows)
        sample_df = df.head(30)
        for _, row in sample_df.iterrows():
            tree.insert("", "end", values=list(row))
    
    def apply_processing(self):
        """Applique le traitement SOWorkflowStage sur les données filtrées"""
        if self.filtered_df is None:
            messagebox.showwarning("Avertissement", "Veuillez d'abord appliquer le filtre RatePlan")
            return
        
        try:
            # Faire une copie pour le traitement
            self.processed_df = self.filtered_df.copy()
            
            # Définir les anciennes et nouvelles valeurs SOWorkflowStage[instance renvoyees en agences]
            old_instance_agence = ["Prise en compte de la demande client", "Déplacer - Prise en Compte de la demande"]
            new_instance_agence = "Instances OC Renvoyées au Agences"
            
            
            # Définir les anciennes et nouvelles valeurs SOWorkflowStage[instances en cours de qualification]
            old_instance_cours_qualifs=["Connection du NE au Transport","Installation Physique de la Ligne", "Equipe Commutation: Affectation du NE"]
            new_instance_cours_qualifs="Instances en cours de qualification"
            
             # Définir les anciennes et nouvelles valeurs SOWorkflowStage[Zones de desservie]
            old_instance_znd=["Zone non deservie","Faisabilite et Extension Reseau"]
            new_instance_znd="Zone non desservie"
            
            pco_sature="Zone saturée"
            
            
            # Convertir date pour traitement.
            self.processed_df["FirstActivatedDate"] = pd.to_datetime(self.processed_df["FirstActivatedDate"], format='%d/%m/%Y')
            # Recuperer la derniere date.
            Date_actuel=self.processed_df["FirstActivatedDate"][0]
            # Creer une Nouvelle colonnes et ajouter la derniere date.
            self.processed_df.insert(3, 'Date_Jours',Date_actuel )
            
            
            # Calculer la difference entre la date de souscription et la date actuelle.
            Delais_instance=self.processed_df["Date_Jours"]-self.processed_df["FirstActivatedDate"]
            Delais_instance=Delais_instance.dt.days
            self.processed_df.insert(4, 'Nombres_Jours',Delais_instance )

            # Creer les délais des instances oc.
            def get_status(age):
                if age <= 7:
                    return '≤ 7J'
                elif 8 <= age <=15:
                    return '8J ≤ X <15J'
                elif 15 < age < 31:
                    return '15J<X<1mois'
                elif 31 < age < 93:
                    return '1mois<X<3mois'

            
            delais_designation= self.processed_df['Nombres_Jours'].apply(get_status)
            
            

            #Ajouter la colonnes delais.
            self.processed_df.insert(5, "Délais",delais_designation)
            
            self.processed_df['year'] = self.processed_df['FirstActivatedDate'].dt.year
            
            #print(self.processed_df['year'])
           
            self.processed_df["Délais"]=self.processed_df["Délais"].mask(self.processed_df["year"]==2024, 2024)
            self.processed_df["Délais"]=self.processed_df["Délais"].mask(self.processed_df["year"]==2023, 2023)
            self.processed_df["Délais"]=self.processed_df["Délais"].mask(self.processed_df["year"]==2022, 2022)
            self.processed_df["Délais"]=self.processed_df["Délais"].mask(self.processed_df["Délais"].isna(), '>3mois')
            
            # Modifier les valeurs dans la colonne SOWorkflowStage
            if 'SOWorkflowStage' in self.processed_df.columns:
                # Compter les modifications
                avant_count = self.processed_df['SOWorkflowStage'].isin(old_instance_agence).sum()
               
                
                # Appliquer le remplacement[Instance renvoyees au agences au ag]
                self.processed_df['SOWorkflowStage'] = self.processed_df['SOWorkflowStage'].replace(old_instance_agence, new_instance_agence)
                
                self.processed_df['SOWorkflowStage'] = self.processed_df['SOWorkflowStage'].replace(old_instance_cours_qualifs, new_instance_cours_qualifs)
                self.processed_df['SOWorkflowStage'] = self.processed_df['SOWorkflowStage'].replace(old_instance_znd, new_instance_znd)
                self.processed_df['SOWorkflowStage'] = self.processed_df['SOWorkflowStage'].fillna("OC non escaladés aux CL")
                
                # Compter après modification
                apres_count_ag = (self.processed_df['SOWorkflowStage'] == new_instance_agence).sum()
                apres_count_qualif = (self.processed_df['SOWorkflowStage'] == new_instance_cours_qualifs).sum()
                apres_count_znd= (self.processed_df['SOWorkflowStage'] == new_instance_znd).sum()
                value_pco_sature= (self.processed_df['SOWorkflowStage'] == pco_sature).sum()
                
                
                self.is_processed = True
                self.update_treeview(self.processed_tree, self.processed_df)
                self.save_btn.configure(state="normal")
                
                # Charger les valeurs COCodeDescription pour le filtre facultatif
                self.load_soworkflow_values()
                
                messagebox.showinfo("Succès", 
                                  f"Traitement SOWorkflowStage appliqué avec succès!\n"
                                  f"Données filtrées (RatePlan): {len(self.filtered_df)} lignes\n"
                                  f"Modifications SOWorkflowStage: {avant_count} valeurs standardisées\n"
                                  f"Total '{new_instance_agence}': {apres_count_ag}\n"
                                  f"Total '{new_instance_cours_qualifs}': {apres_count_qualif}\n"
                                  f"Total '{new_instance_znd}': {apres_count_znd}\n"
                                  f"Total '{pco_sature}': {value_pco_sature}\n"
                                )
                
                self.status_label.configure(text="Status: Traitement appliqué - Prêt à sauvegarder", text_color="blue")
                self.update_stats()
                
            else:
                messagebox.showwarning("Avertissement", "Colonne 'SOWorkflowStage' non trouvée dans les données filtrées")
                
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du traitement: {str(e)}")
    
    
    def load_soworkflow_values(self):
        """Charge les valeurs uniques de COCodeDescription après traitement"""
        if self.processed_df is not None and 'COCodeDescription' in self.processed_df.columns:
            # Récupérer les valeurs uniques et trier
            self.soworkflow_values = sorted(self.processed_df['COCodeDescription'].dropna().unique().tolist())
            
            # Mettre à jour le menu déroulant
            values = ["Tous les statuts"] + self.soworkflow_values
            self.soworkflow_combobox.configure(values=values, state="normal")
            self.soworkflow_combobox.set("Tous les statuts")
            
            # Activer les boutons de sauvegarde et de filtre facultatif
            self.soworkflow_filter_btn.configure(state="normal")
            self.save_btn.configure(state="normal")
    
    def update_stats(self):
        """Met à jour les statistiques"""
        if self.original_df is None:
            return
        
        stats_text = "📈 STATISTIQUES COMPLÈTES\n\n"
        stats_text += f"• Total des lignes originales: {len(self.original_df)}\n"
        stats_text += f"• Total des colonnes: {len(self.original_df.columns)}\n"
        
        if self.filtered_df is not None:
            stats_text += f"• Lignes après filtre RatePlan: {len(self.filtered_df)}\n"
        
        if 'RatePlan' in self.original_df.columns:
            stats_text += f"\n📊 DISTRIBUTION RATEPLAN (original):\n"
            value_counts = self.original_df['RatePlan'].value_counts()
            for value, count in value_counts.items():
                stats_text += f"   - {value}: {count} lignes\n"
        
        # Ajouter les statistiques de COCodeDescription
        if self.filtered_df is not None and 'COCodeDescription' in self.filtered_df.columns:
            stats_text += f"\n📊 DISTRIBUTION COCodeDescription (APRÈS filtre RatePlan):\n"
            value_counts = self.filtered_df['COCodeDescription'].value_counts()
            for value, count in value_counts.items():
                stats_text += f"   - {value}: {count} lignes\n"
        
        if self.is_processed and 'COCodeDescription' in self.processed_df.columns:
            stats_text += f"\n📊 DISTRIBUTION COCodeDescription (APRÈS traitement):\n"
            value_counts = self.processed_df['COCodeDescription'].value_counts()
            for value, count in value_counts.items():
                stats_text += f"   - {value}: {count} lignes\n"
            
            # Afficher si un filtre est actif
            if self.is_filter_applied and self.selected_soworkflow:
                stats_text += f"\n🎯 FILTRE ACTUEL:\n"
                stats_text += f"   - COCodeDescription filtré: {self.selected_soworkflow}\n"
                if self.final_df is not None:
                    stats_text += f"   - Lignes après filtre: {len(self.final_df)}/{len(self.processed_df)}\n"
            else:
                stats_text += f"\n🎯 FILTRE ACTUEL: Aucun (toutes les valeurs)\n"
        
        self.stats_text.delete("1.0", "end")
        self.stats_text.insert("1.0", stats_text)
    
    def save_file(self):
        """Étape 5: Sauvegarde le fichier final (avec ou sans filtre)"""
        # Si final_df n'existe pas encore (première sauvegarde sans filtre)
        if self.final_df is None:
            if self.processed_df is None:
                messagebox.showwarning("Avertissement", "Aucune donnée à sauvegarder. Complétez d'abord l'Étape 3.")
                return
            # Par défaut, utiliser toutes les données traitées
            self.final_df = self.processed_df.copy()
            self.is_filter_applied = False
        
        file_path = filedialog.asksaveasfilename(
            title="Enregistrer le fichier final",
            defaultextension=".xlsx",
            filetypes=[("Fichiers Excel", "*.xlsx"), ("Tous les fichiers", "*.*")]
        )
        
        if file_path:
            try:
                # Sauvegarder avec le moteur openpyxl
                self.final_df.to_excel(file_path, index=False, engine='openpyxl')
                
                if self.is_filter_applied:
                    message_text = (f"🎉 Fichier final sauvegardé avec succès!\n\n"
                                  f"📁 Emplacement: {file_path}\n"
                                  f"📊 Lignes sauvegardées: {len(self.final_df)}\n"
                                  f"🔍 Filtre RatePlan: FTTH DUO MAITRISE, FTTH DUO\n"
                                  f"⚙️ Traitement SOWorkflowStage: Appliqué\n"
                                  f"🎯 Filtre COCodeDescription: {self.selected_soworkflow}\n\n"
                                  f"✅ Données filtrées sauvegardées!")
                else:
                    message_text = (f"🎉 Fichier final sauvegardé avec succès!\n\n"
                                  f"📁 Emplacement: {file_path}\n"
                                  f"📊 Lignes sauvegardées: {len(self.final_df)}\n"
                                  f"🔍 Filtre RatePlan: FTTH DUO MAITRISE, FTTH DUO\n"
                                  f"⚙️ Traitement SOWorkflowStage: Appliqué\n"
                                  f"🎯 Filtre COCodeDescription: Aucun (toutes les valeurs)\n\n"
                                  f"✅ Toutes les données traitées sauvegardées!")
                
                messagebox.showinfo("Succès", message_text)
                
                self.status_label.configure(text=f"Status: Fichier sauvegardé - {os.path.basename(file_path)}", text_color="green")
                
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de la sauvegarde: {str(e)}")
    



    def run(self):
        """Lance l'application"""
        self.root.mainloop()

class DrgtProcessorApp:
    def __init__(self, main_menu_app):
        # Référence au menu principal
        self.main_menu = main_menu_app
        
        # Configuration de l'interface
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        # Création de la fenêtre de traitement OC
        self.root = ctk.CTk()
        self.root.title("Traitement des Dérangements")
        self.root.geometry("1200x800")
        
        # Gestion de la fermeture de la fenêtre
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Variables
        self.original_df = None
        self.processed_df = None
        self.filtered_df = None
        self.final_df = None
        self.file_path = None
        self.is_processed = False
        
        # Variables d'état(mises à jours=> DRGTS.)
        self.is_processed = False
        self.followup_values = []
        self.selected_followup = None
        self.is_filter_applied = False

        # Création de l'interface
        self.create_widgets()
        
        # Design UI
    def create_widgets(self):
        """Crée l'interface de traitement des OC"""
        # Frame principal
        main_frame = ctk.CTkFrame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Header avec bouton retour
        header_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        header_frame.pack(fill="x", padx=20, pady=10)
        
        # Bouton retour au menu principal
        back_button = ctk.CTkButton(header_frame, 
                                text="← Retour au Menu", 
                                command=self.return_to_main_menu,
                                width=150,
                                fg_color="gray",
                                hover_color="darkgray")
        back_button.pack(side="left")
        
        # Titre
        title_label = ctk.CTkLabel(header_frame, 
                                text="Traitement des Drgts", 
                                font=ctk.CTkFont(size=16, weight="bold"))
        title_label.pack(side="left", padx=20)
        
        
        # Frame pour les boutons d'action
        button_frame = ctk.CTkFrame(main_frame)
        button_frame.pack(fill="x", padx=20, pady=10)
        
        
        # Bouton fermer l'application
        close_button = ctk.CTkButton(header_frame, 
                                    text="❌ Fermer l'Application", 
                                    command=self.close_application,
                                    width=180,
                                    fg_color="#C41E3A",
                                    hover_color="#A61E3A")
        close_button.pack(side="right")
        
        
        # Bouton pour téléverser le fichier
        upload_btn = ctk.CTkButton(button_frame, text="📁 Téléverser le Fichier Excel", 
                                command=self.upload_file, fg_color="darkgray",
                                hover_color="gray",
                                width=200)
        upload_btn.pack(side="left", padx=5, pady=5)
        
        # Bouton pour appliquer le filtre RatePlan
        filter_btn = ctk.CTkButton(button_frame, text="🔍 Filtrer la FTTH", 
                                command=self.apply_rateplan_filter_drgts,
                                width=180,fg_color="darkgray", hover_color="gray", state="disabled")
        filter_btn.pack(side="left", padx=5, pady=5)
        self.filter_btn = filter_btn
        
        # Bouton pour appliquer le traitement
        process_btn = ctk.CTkButton(button_frame, text="⚙️ Appliquer le Traitement", 
                                command=self.apply_processing_drgts,
                                width=200,fg_color="#B7B7B7", hover_color="gray", state="disabled")
        process_btn.pack(side="left", padx=5, pady=5)
        self.process_btn = process_btn
        
        # Menu déroulant pour FollowUp (mise à jours DRGTS)
        followup_label = ctk.CTkLabel(button_frame, text="Centre de suivi:", font=ctk.CTkFont(size=12))
        followup_label.pack(side="left", padx=(0, 5), pady=5)
        
        self.followup_combobox = ctk.CTkComboBox(button_frame, 
                                                  values=["Tous les centres"],
                                                  width=200,
                                                  state="disabled",
                                                  command=self.on_followup_selected)
        self.followup_combobox.pack(side="left", padx=5, pady=5)
        self.followup_combobox.set("Tous les centres")
        
        self.followup_filter_btn = ctk.CTkButton(button_frame, 
                                                  text="🔎 Appliquer Filtre", 
                                                  command=self.apply_followup_filter,
                                                  width=150,
                                                  fg_color="#B7B7B7",
                                                  hover_color="gray",
                                                  state="disabled")
        self.followup_filter_btn.pack(side="left", padx=5, pady=5)
        
        
        # Bouton pour créer le nouveau fichier
        save_btn = ctk.CTkButton(button_frame, text="💾 Créer Nouveau Fichier", 
                                command=self.save_file_drgt,
                                width=200,fg_color="#B7B7B7", hover_color="gray", state="disabled")
        save_btn.pack(side="left", padx=5, pady=5)
        self.save_btn = save_btn
        
        # Frame pour les onglets
        tabview = ctk.CTkTabview(main_frame)
        tabview.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Onglets
        tabview.add("Aperçu Original")
        tabview.add("Aperçu Filtre Status")
        tabview.add("Aperçu Traité")
        tabview.add("Statistiques")
        
        # Treeview pour l'aperçu des données originales
        self.original_tree = ttk.Treeview(tabview.tab("Aperçu Original"))
        self.setup_treeview(self.original_tree, tabview.tab("Aperçu Original"))
        
        # Treeview pour l'aperçu des données filtrées
        self.filtered_tree = ttk.Treeview(tabview.tab("Aperçu Filtre Status"))
        self.setup_treeview(self.filtered_tree, tabview.tab("Aperçu Filtre Status"))
        
        # Treeview pour l'aperçu des données traitées
        self.processed_tree = ttk.Treeview(tabview.tab("Aperçu Traité"))
        self.setup_treeview(self.processed_tree, tabview.tab("Aperçu Traité"))
        
        
        # Frame pour les statistiques
        stats_frame = ctk.CTkFrame(tabview.tab("Statistiques"))
        stats_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.stats_text = ctk.CTkTextbox(stats_frame, height=200)
        self.stats_text.pack(fill="both", expand=True, padx=10, pady=10)
        self.stats_text.insert("1.0", "Aucune statistique disponible. Veuillez charger un fichier.")
        
        # Frame pour les informations
        info_frame = ctk.CTkFrame(main_frame)
        info_frame.pack(fill="x", padx=20, pady=10)
        
        # Informations sur le fichier
        self.info_label = ctk.CTkLabel(info_frame, text="📊 Aucun fichier chargé", 
                                    font=ctk.CTkFont(size=12))
        self.info_label.pack(pady=5)
        
        # Status du traitement
        self.status_label = ctk.CTkLabel(info_frame, text="Status: En attente de fichier", 
                                        text_color="gray", font=ctk.CTkFont(size=10))
        self.status_label.pack(pady=2)
      
        #   
   
   
   
    def update_previews(self):
        """Met à jour les aperçus des données"""
        self.update_treeview(self.original_tree, self.original_df)
        self.update_treeview(self.filtered_tree, self.filtered_df)
        self.update_treeview(self.processed_tree, self.processed_df)
        
    
    def update_treeview(self, tree, df):
        """Met à jour un treeview spécifique avec des données"""
        # Clear existing data
        for item in tree.get_children():
            tree.delete(item)
        
        if df is None:
            return
        
        # Set columns
        columns = list(df.columns)
        tree["columns"] = columns
        tree["show"] = "headings"
        
        # Configure columns
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, minwidth=80)
        
        # Add sample data (first 30 rows)
        sample_df = df.head(30)
        for _, row in sample_df.iterrows():
            tree.insert("", "end", values=list(row))
    
    def on_followup_selected(self, choice):
        """Callback quand une valeur FollowUp est sélectionnée"""
        self.selected_followup = choice if choice != "Tous les centres" else None
    
    def apply_followup_filter(self):
        """Applique le filtre FollowUp (facultatif)"""
        if self.processed_df is None:
            messagebox.showwarning("Avertissement", "Veuillez d'abord appliquer le traitement (Étape 3)")
            return
        
        # Si "Tous les centres" est sélectionné, pas de filtre (toutes les données)
        if self.selected_followup is None or self.selected_followup == "Tous les centres":
            # Pas de filtre appliqué, on utilise toutes les données traitées
            self.final_df = self.processed_df.copy()
            self.is_filter_applied = False
            
            messagebox.showinfo("Aucun filtre", 
                              "✅ Aucun filtre spécifique appliqué.\n"
                              "Toutes les données traitées seront sauvegardées.")
            
            self.status_label.configure(text="Status: Aucun filtre FollowUp appliqué - Prêt pour la sauvegarde", text_color="green")
            self.update_stats()
            return
        
        try:
            # Appliquer le filtre FollowUp sur les données traitées
            self.final_df = self.processed_df[self.processed_df['FollowUp'] == self.selected_followup].copy()
            self.is_filter_applied = True
            
            if len(self.final_df) == 0:
                messagebox.showwarning("Avertissement", 
                                     f"Aucune donnée trouvée pour: {self.selected_followup}")
                return
            
            messagebox.showinfo("Filtre appliqué", 
                              f"✅ Filtre facultatif appliqué avec succès!\n"
                              f"Centre sélectionné: {self.selected_followup}\n"
                              f"Lignes filtrées: {len(self.final_df)}/{len(self.processed_df)}\n"
                              f"Prêt pour la sauvegarde")
            
            self.status_label.configure(text=f"Status: Filtre FollowUp appliqué ({self.selected_followup}) - Prêt pour la sauvegarde", text_color="green")
            self.update_stats()
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du filtrage FollowUp: {str(e)}")
            
    # DRGTS Function upload file
    def upload_file(self):
        """Téléverse un fichier Excel"""
        file_path = filedialog.askopenfilename(
            title="Sélectionner un fichier Excel",
            filetypes=[("Fichiers Excel", "*.xlsx *.xls"), ("Tous les fichiers", "*.*")]
        )
        
        if file_path:
            try:
                self.file_path = file_path
                self.original_df = pd.read_excel(file_path)
                self.filtered_df = None
                self.processed_df = None
                self.final_df = None
                self.is_processed = False
                self.selected_followup = None
                self.is_filter_applied = False
                
                self.update_previews()
                self.filter_btn.configure(state="normal")
                self.process_btn.configure(state="disabled")
                self.save_btn.configure(state="disabled")
                
                file_info = f"📊 Fichier: {os.path.basename(file_path)} - {len(self.original_df)} lignes - {len(self.original_df.columns)} colonnes"
                self.info_label.configure(text=file_info)
                self.status_label.configure(text="Status: Fichier chargé - Prêt pour le filtre RatePlan", text_color="green")
                
                self.update_stats()
                
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de la lecture du fichier: {str(e)}")
       
    # DRGTS Filtre de la FTTH, Statut, CallType et du Subject.
    
    def apply_rateplan_filter_drgts(self):
        """Applique le filtre sur Statut"""
        if self.original_df is None:
            messagebox.showwarning("Avertissement", "Veuillez d'abord charger un fichier")
            return
        
        try:
            # Définir les valeurs RatePlan à filtrer
            Status_values = ["Ouvert"]
            Calltype_values = ["FTTH"]
            if 'Statut' in self.original_df.columns:
                # Appliquer le filtre
                self.filtered_df = self.original_df[self.original_df['Statut'].isin(Status_values)].copy()
                self.filtered_df = self.filtered_df[self.filtered_df['CallType'].isin(Calltype_values)].copy()
                
                print(len(self.filtered_df))
                # Supprimer les Doublons
                self.filtered_df=self.filtered_df.drop_duplicates("ND")
                print(len(self.filtered_df))
                
                if len(self.filtered_df) == 0:
                    messagebox.showwarning("Avertissement", 
                                         f"Aucune donnée trouvée pour les CallType: {Calltype_values}")
                    return
                
                self.processed_df = self.filtered_df.copy()
                self.update_treeview(self.filtered_tree, self.filtered_df)
                self.process_btn.configure(state="normal")
                
                messagebox.showinfo("Filtre appliqué", 
                                  f"Filtre Statut appliqué avec succès!\n"
                                  f"Lignes filtrées: {len(self.filtered_df)}/{len(self.original_df)}\n"
                                  f": {Status_values}")
                
                self.status_label.configure(text="Status: Filtre RatePlan appliqué - Prêt pour le traitement SOWorkflowStage", text_color="blue")
                self.update_stats()
                
            else:
                messagebox.showwarning("Avertissement", "Colonne 'Status' non trouvée dans le fichier")
                
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du filtrage: {str(e)}")
    
    
    
    def apply_processing_drgts(self):
        """Applique le traitement sur les données filtrées"""
        if self.filtered_df is None:
            messagebox.showwarning("Avertissement", "Veuillez d'abord appliquer le filtre FTTH")
            return
        
        try:
            
            
            
            # Faire une copie pour le traitement
            self.processed_df = self.filtered_df.copy()
            
            
            # Supprimer les colonnes inutiles dans le dataFrame.
            list_colonnes=["Categorie","TypeFermeture","CallClass","Resolution","Cause"]
            self.processed_df=self.processed_df.drop(list_colonnes, axis=1)
            
            
            
            # Exclure les Centre de prod et les service Relation Clt.
            valeurs_a_exclure = ['Centre de Production', 'Service Relation Client','Centre Telephonique Makokou']
            
            self.processed_df = self.processed_df[~self.processed_df['FollowUp'].isin(valeurs_a_exclure)]
            self.processed_df=self.processed_df.dropna(subset=['FollowUp'])
            
            # Convertir date pour traitement.
            self.processed_df["DateCreated"] = pd.to_datetime(self.processed_df["DateCreated"], format='%d/%m/%Y')
            # Recuperer la derniere date.
            Date_actuel=self.processed_df["DateCreated"][0]
            # Creer une Nouvelle colonnes et ajouter la derniere date.
            self.processed_df['DateClosed']=Date_actuel
            
            # Calculer la difference entre la date de souscription et la date actuelle.
            Delais_instance=self.processed_df["DateClosed"]-self.processed_df["DateCreated"]
            Delais_instance=Delais_instance.dt.days
            self.processed_df['Nbr_Jour']=Delais_instance

            # Creer les délais des instances oc.
            def get_status(age):
                if age <= 1:
                    return '<=24H'
                elif age == 2:
                    return 'x=48H'
                elif 3 <= age <= 8:
                    return '3J ≤ X <=8J'
                elif age > 8:
                    return '3J ≤ X <=8J'

            
            delais_designation= self.processed_df['Nbr_Jour'].apply(get_status)
            
            #Ajouter la colonnes delais.
            self.processed_df.insert(8, "Délais",delais_designation)
            
            if 'FollowUp' in self.processed_df.columns:
               
                self.is_processed = True
                self.update_treeview(self.processed_tree, self.processed_df)
                self.save_btn.configure(state="normal")
                
                # Charger les valeurs FollowUp pour le filtre facultatif
                self.load_followup_values()
                
                self.status_label.configure(text="Status: Traitement appliqué - Prêt à sauvegarder", text_color="blue")
                self.update_stats()
                
            else:
                messagebox.showwarning("Avertissement", "Colonne 'FollowUp' non trouvée dans les données filtrées")
                
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du traitement: {str(e)}")
    
    def load_followup_values(self):
        """Charge les valeurs uniques de FollowUp après traitement"""
        if self.processed_df is not None and 'FollowUp' in self.processed_df.columns:
            # Récupérer les valeurs uniques et trier
            self.followup_values = sorted(self.processed_df['FollowUp'].dropna().unique().tolist())
            
            # Mettre à jour le menu déroulant
            values = ["Tous les centres"] + self.followup_values
            self.followup_combobox.configure(values=values, state="normal")
            self.followup_combobox.set("Tous les centres")
            
            # Activer les boutons de sauvegarde et de filtre facultatif
            self.followup_filter_btn.configure(state="normal")
            self.save_btn.configure(state="normal")
            
        
                
    def update_stats(self):
        """Met à jour les statistiques"""
        if self.original_df is None:
            return
        
        stats_text = "📈 STATISTIQUES COMPLÈTES\n\n"
        stats_text += f"• Total des lignes originales: {len(self.original_df)}\n"
        stats_text += f"• Total des colonnes: {len(self.original_df.columns)}\n"
        
        if self.filtered_df is not None:
            stats_text += f"• Lignes après filtre FTTH: {len(self.filtered_df)}\n"
        
        if 'FollowUp' in self.original_df.columns:
            stats_text += f"\n📊 DISTRIBUTION FollowUp (original):\n"
            value_counts = self.original_df['FollowUp'].value_counts().head(10)  # Limiter à 10 premiers
            for value, count in value_counts.items():
                stats_text += f"   - {value}: {count} lignes\n"
        
        if self.filtered_df is not None and 'FollowUp' in self.filtered_df.columns:
            stats_text += f"\n📊 DISTRIBUTION FollowUp (APRÈS filtre FTTH, AVANT traitement):\n"
            value_counts = self.filtered_df['FollowUp'].value_counts()
            for value, count in value_counts.items():
                stats_text += f"   - {value}: {count} lignes\n"
        
        if self.is_processed and 'FollowUp' in self.processed_df.columns:
            stats_text += f"\n📊 DISTRIBUTION FollowUp (APRÈS traitement):\n"
            value_counts = self.processed_df['FollowUp'].value_counts()
            for value, count in value_counts.items():
                stats_text += f"   - {value}: {count} lignes\n"
            
            # Afficher si un filtre est actif
            if self.is_filter_applied and self.selected_followup:
                stats_text += f"\n🎯 FILTRE ACTUEL:\n"
                stats_text += f"   - FollowUp filtré: {self.selected_followup}\n"
                if self.final_df is not None:
                    stats_text += f"   - Lignes après filtre: {len(self.final_df)}/{len(self.processed_df)}\n"
            else:
                stats_text += f"\n🎯 FILTRE ACTUEL: Aucun (tous les centres)\n"
        
        self.stats_text.delete("1.0", "end")
        self.stats_text.insert("1.0", stats_text)
        
    # Fonction de retour au menu principal.
    def on_closing(self):
        """Gère la fermeture de la fenêtre OC"""
        # Retourne au menu principal au lieu de fermer l'application
        self.return_to_main_menu()
        
    # Fonction de Fermeture de l'application.
    def return_to_main_menu(self):
        """Retourne au menu principal"""
        self.root.destroy()  # Ferme la fenêtre de traitement
        self.main_menu.show_main_menu()  # Réaffiche le menu principal
    
    def close_application(self):
        """Ferme complètement l'application"""
        # Demande confirmation à l'utilisateur
        if messagebox.askyesno("Fermeture", "Êtes-vous sûr de vouloir quitter l'application ?"):
            # Ferme cette fenêtre
            self.root.destroy()
            # Ferme aussi le menu principal s'il existe
            if hasattr(self.main_menu, 'root') and self.main_menu.root.winfo_exists():
                self.main_menu.root.destroy()
            sys.exit(0)  # Arrête proprement le programme
    
    # Fonction de View des donnees.
    def setup_treeview(self, tree, parent):
        """Configure un treeview avec scrollbars"""
        v_scrollbar = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        h_scrollbar = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        tree.pack(side="left", fill="both", expand=True)
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")
    
    # DRGTS Sauvegarder le Fichier apres Traitement.
    def save_file_drgt(self):
        """Sauvegarde le fichier final (avec ou sans filtre)"""
        # Si final_df n'existe pas encore (première sauvegarde sans filtre)
        if self.final_df is None:
            if self.processed_df is None:
                messagebox.showwarning("Avertissement", "Aucune donnée à sauvegarder. Complétez d'abord l'Étape 3.")
                return
            # Par défaut, utiliser toutes les données traitées
            self.final_df = self.processed_df.copy()
            self.is_filter_applied = False
        
        file_path = filedialog.asksaveasfilename(
            title="Enregistrer le fichier final",
            defaultextension=".xlsx",
            filetypes=[("Fichiers Excel", "*.xlsx"), ("Tous les fichiers", "*.*")]
        )
        
        if file_path:
            try:
                # Sauvegarder avec le moteur openpyxl
                self.final_df.to_excel(file_path, index=False, engine='openpyxl')
                
                if self.is_filter_applied:
                    message_text = (f"🎉 Fichier final sauvegardé avec succès!\n\n"
                                  f"📁 Emplacement: {file_path}\n"
                                  f"📊 Lignes sauvegardées: {len(self.final_df)}\n"
                                  f"🔍 Filtre FTTH: Ouvert\n"
                                  f"⚙️ Traitement des dérangements: Appliqué\n"
                                  f"🎯 Filtre FollowUp: {self.selected_followup}\n\n"
                                  f"✅ Données filtrées sauvegardées!")
                else:
                    message_text = (f"🎉 Fichier final sauvegardé avec succès!\n\n"
                                  f"📁 Emplacement: {file_path}\n"
                                  f"📊 Lignes sauvegardées: {len(self.final_df)}\n"
                                  f"🔍 Filtre FTTH: Ouvert\n"
                                  f"⚙️ Traitement des dérangements: Appliqué\n"
                                  f"🎯 Filtre FollowUp: Aucun (tous les centres)\n\n"
                                  f"✅ Toutes les données traitées sauvegardées!")
                
                messagebox.showinfo("Succès", message_text)
                
                self.status_label.configure(text=f"Status: Fichier sauvegardé - {os.path.basename(file_path)}", text_color="green")
                
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de la sauvegarde: {str(e)}")
    
    
    
    # Lancé l'onglet.
    def run(self):
        """Lance l'application"""
        self.root.mainloop()


    

    
    
    
# Lancement de l'application
if __name__ == "__main__":
    app = MainMenuApp()
    app.run()