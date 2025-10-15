using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using System.IO;
using System.Diagnostics;

namespace GESTION_MOYENNE_SEMESTRE_9_IC3__IT
{
    public partial class Form1 : Form
    {
        private TabControl? mainTabControl;
        private TabControl? specialitesTabControl;

        // Dictionnaire pour stocker toutes les matières
        private Dictionary<string, List<Matiere>> allMatieres = new Dictionary<string, List<Matiere>>();

        // Dictionnaires pour les TextBox
        private Dictionary<Matiere, TextBox> heuresTextBoxes = new Dictionary<Matiere, TextBox>();
        private Dictionary<Matiere, TextBox> moyennesTextBoxes = new Dictionary<Matiere, TextBox>();
        private Dictionary<string, Label> moyenneLabels = new Dictionary<string, Label>();

        private Label? lblMoyenneGenerale;

        public Form1()
        {
            // Forcer la culture invariante pour le formatage des nombres
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.InvariantCulture;
            System.Threading.Thread.CurrentThread.CurrentUICulture = System.Globalization.CultureInfo.InvariantCulture;

            InitialiserDonnees();
            InitializeComponent();
            CreerInterface();
        }

        private void InitialiserDonnees()
        {
            // UE de culture générale
            allMatieres["UE Culture Générale"] = new List<Matiere>
            {
                new Matiere { Nom = "Stage de production", Coefficient = 2 },
                new Matiere { Nom = "Séminaire des méthodes", Coefficient = 1 },
                new Matiere { Nom = "Choix économiques des projets", Coefficient = 1 },
                new Matiere { Nom = "E.P.S", Coefficient = 1 },
                new Matiere { Nom = "Anglais", Coefficient = 1 }
            };

            // Sciences de la construction
            allMatieres["Sciences de la construction"] = new List<Matiere>
            {
                new Matiere { Nom = "Procédés généraux de construction appliqués", Coefficient = 2 },
                new Matiere { Nom = "Technologie des engins de chantier", Coefficient = 1 },
                new Matiere { Nom = "Méthodologie de terrassement & démolition", Coefficient = 1 }
            };

            // Techniques d'entretien des infrastructures
            allMatieres["Techniques d'entretien des infrastructures"] = new List<Matiere>
            {
                new Matiere { Nom = "Technique d'entretien routier", Coefficient = 1 },
                new Matiere { Nom = "Pathologie des ouvrages", Coefficient = 1 },
                new Matiere { Nom = "Stratégie de suivi du réseau routier", Coefficient = 1 }
            };

            // Techniques et Méthodes de Contrôle et Gestion
            allMatieres["Techniques et Méthodes de Contrôle et Gestion"] = new List<Matiere>
            {
                new Matiere { Nom = "Contrôle gestion-Audit", Coefficient = 1 },
                new Matiere { Nom = "Gestion de projets", Coefficient = 1 }
            };

            // Gestion des infrastructures
            allMatieres["Gestion des infrastructures"] = new List<Matiere>
            {
                new Matiere { Nom = "Aéroport et transports aériens", Coefficient = 1 },
                new Matiere { Nom = "Economie des transports(acquisition)", Coefficient = 1 },
                new Matiere { Nom = "Chemin de fer", Coefficient = 1 },
                new Matiere { Nom = "Techniques d'exploitation des transports de marchandises", Coefficient = 1 },
                new Matiere { Nom = "Techniques d'exploitation des transports de voyageurs", Coefficient = 1 }
            };

            // Hydraulique appliquée
            allMatieres["Hydraulique appliquée"] = new List<Matiere>
            {
                new Matiere { Nom = "Hydraulique maritime", Coefficient = 1 },
                new Matiere { Nom = "Travaux maritime", Coefficient = 1 }
            };

            // Conception des infrastructures
            allMatieres["Conception des infrastructures"] = new List<Matiere>
            {
                new Matiere { Nom = "Conception d'ouvrages d'art", Coefficient = 1 },
                new Matiere { Nom = "Mécaniques des sols (amélioration des sols)", Coefficient = 1 },
                new Matiere { Nom = "Mécaniques des roches", Coefficient = 1 }
            };

            // Projet de fin d'études
            allMatieres["Projet de fin d'études"] = new List<Matiere>
            {
                new Matiere { Nom = "Projet de fin d'études (PFE)", Coefficient = 4 }
            };
        }

        private void CreerInterface()
        {
            this.Text = "GESTION MOYENNE SEMESTRE 9 IC3 IT/ESTP";
            this.Size = new Size(1400, 750);
            this.StartPosition = FormStartPosition.CenterScreen;

            // ========== PANEL DU HAUT (TITRE) ==========
            Panel topPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 80,
                BackColor = Color.FromArgb(70, 130, 180)
            };

            Label lblTitle = new Label
            {
                Text = "MES OBJECTIFS DU SEMESTRE 9 IC3 IT/ESTP",
                Font = new Font("Segoe UI", 18, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill
            };
            topPanel.Controls.Add(lblTitle);

            // ========== TABCONTROL PRINCIPAL ==========
            mainTabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };

            // Onglet 1: UE de culture générale
            TabPage tabCultureGenerale = new TabPage("UE de culture générale");
            CreerOngletMatiere(tabCultureGenerale, "UE Culture Générale");
            mainTabControl.TabPages.Add(tabCultureGenerale);

            // Onglet 2: UE de spécialités avec sous-onglets
            TabPage tabSpecialites = new TabPage("UE de spécialités");
            specialitesTabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };

            // Sous-onglets pour les spécialités
            string[] specialites = {
                "Sciences de la construction",
                "Techniques d'entretien des infrastructures",
                "Techniques et Méthodes de Contrôle et Gestion",
                "Gestion des infrastructures",
                "Hydraulique appliquée",
                "Conception des infrastructures",
                "Projet de fin d'études"
            };

            foreach (string specialite in specialites)
            {
                TabPage subTab = new TabPage(specialite);
                CreerOngletMatiere(subTab, specialite);
                specialitesTabControl.TabPages.Add(subTab);
            }

            tabSpecialites.Controls.Add(specialitesTabControl);
            mainTabControl.TabPages.Add(tabSpecialites);

            // ========== PANEL DU BAS (MOYENNE + BOUTONS) ==========
            Panel bottomPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 80,
                BackColor = Color.WhiteSmoke,
                BorderStyle = BorderStyle.FixedSingle
            };

            // Moyenne générale
            Label lblMoyenneText = new Label
            {
                Text = "MOYENNE GENERALE:",
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                Location = new Point(20, 25),
                AutoSize = true
            };
            bottomPanel.Controls.Add(lblMoyenneText);

            lblMoyenneGenerale = new Label
            {
                Text = "0.00 / 20",
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                ForeColor = Color.Red,
                Location = new Point(300, 22),
                AutoSize = true
            };
            bottomPanel.Controls.Add(lblMoyenneGenerale);

            // Boutons d'export
            Button btnExportPDF = new Button
            {
                Text = "Exporter PDF",
                Location = new Point(450, 20),
                Size = new Size(130, 40),
                Font = new Font("Segoe UI", 10)
            };
            btnExportPDF.Click += BtnExportPDF_Click;
            bottomPanel.Controls.Add(btnExportPDF);

            Button btnExportExcel = new Button
            {
                Text = "Exporter Excel",
                Location = new Point(590, 20),
                Size = new Size(130, 40),
                Font = new Font("Segoe UI", 10)
            };
            btnExportExcel.Click += BtnExportExcel_Click;
            bottomPanel.Controls.Add(btnExportExcel);

            Button btnExportTxt = new Button
            {
                Text = "Exporter TXT",
                Location = new Point(730, 20),
                Size = new Size(130, 40),
                Font = new Font("Segoe UI", 10)
            };
            btnExportTxt.Click += BtnExportTxt_Click;
            bottomPanel.Controls.Add(btnExportTxt);

            Button btnCalculer = new Button
            {
                Text = "Calculer",
                Location = new Point(870, 20),
                Size = new Size(130, 40),
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                BackColor = Color.FromArgb(70, 130, 180),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Standard
            };
            btnCalculer.Click += BtnCalculer_Click;
            bottomPanel.Controls.Add(btnCalculer);

            // ========== AJOUT DES CONTRÔLES ==========
            this.Controls.Add(mainTabControl);
            this.Controls.Add(topPanel);
            this.Controls.Add(bottomPanel);
        }

        private void CreerOngletMatiere(TabPage tabPage, string categorie)
        {
            Panel mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(20),
                BackColor = Color.White
            };

            int yPosition = 20;

            // En-têtes
            Panel headerPanel = new Panel
            {
                Location = new Point(20, yPosition),
                Size = new Size(1200, 35),
                BackColor = Color.LightGray
            };

            Label lblNom = new Label
            {
                Text = "MATIERE",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                Location = new Point(10, 8),
                Size = new Size(450, 25)
            };
            headerPanel.Controls.Add(lblNom);

            Label lblHeures = new Label
            {
                Text = "HEURES DEDIEES",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                Location = new Point(470, 8),
                Size = new Size(180, 25)
            };
            headerPanel.Controls.Add(lblHeures);

            Label lblCoef = new Label
            {
                Text = "COEF",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                Location = new Point(660, 8),
                Size = new Size(100, 25),
                TextAlign = ContentAlignment.MiddleCenter
            };
            headerPanel.Controls.Add(lblCoef);

            Label lblMoy = new Label
            {
                Text = "MOYENNE",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                Location = new Point(770, 8),
                Size = new Size(180, 25)
            };
            headerPanel.Controls.Add(lblMoy);

            mainPanel.Controls.Add(headerPanel);

            yPosition += 50;

            // Créer les champs pour chaque matière
            if (allMatieres.ContainsKey(categorie))
            {
                foreach (var matiere in allMatieres[categorie])
                {
                    Panel linePanel = new Panel
                    {
                        Location = new Point(20, yPosition),
                        Size = new Size(1200, 40),
                        BorderStyle = BorderStyle.FixedSingle
                    };

                    Label lblMatiereNom = new Label
                    {
                        Text = matiere.Nom,
                        Location = new Point(10, 10),
                        Size = new Size(450, 25),
                        Font = new Font("Segoe UI", 10)
                    };
                    linePanel.Controls.Add(lblMatiereNom);

                    TextBox txtHeures = new TextBox
                    {
                        Location = new Point(470, 8),
                        Size = new Size(170, 25),
                        Font = new Font("Segoe UI", 10)
                    };
                    txtHeures.TextChanged += (s, e) => TxtNumeric_TextChanged(s, e, matiere, true);
                    heuresTextBoxes[matiere] = txtHeures;
                    linePanel.Controls.Add(txtHeures);

                    Label lblCoefValue = new Label
                    {
                        Text = matiere.Coefficient.ToString(),
                        Location = new Point(660, 10),
                        Size = new Size(100, 25),
                        Font = new Font("Segoe UI", 10, FontStyle.Bold),
                        TextAlign = ContentAlignment.MiddleCenter
                    };
                    linePanel.Controls.Add(lblCoefValue);

                    TextBox txtMoyenne = new TextBox
                    {
                        Location = new Point(770, 8),
                        Size = new Size(170, 25),
                        Font = new Font("Segoe UI", 10)
                    };
                    txtMoyenne.TextChanged += (s, e) => TxtNumeric_TextChanged(s, e, matiere, false);
                    moyennesTextBoxes[matiere] = txtMoyenne;
                    linePanel.Controls.Add(txtMoyenne);

                    mainPanel.Controls.Add(linePanel);

                    yPosition += 50;
                }

                yPosition += 20;

                // Panel pour la moyenne de la catégorie
                Panel moyennePanel = new Panel
                {
                    Location = new Point(20, yPosition),
                    Size = new Size(1200, 50),
                    BackColor = Color.FromArgb(240, 240, 240),
                    BorderStyle = BorderStyle.FixedSingle
                };

                Label lblMoyenneCat = new Label
                {
                    Text = "MOYENNE " + categorie.ToUpper() + ":",
                    Font = new Font("Segoe UI", 12, FontStyle.Bold),
                    Location = new Point(10, 12),
                    Size = new Size(650, 30)
                };
                moyennePanel.Controls.Add(lblMoyenneCat);

                Label lblMoyenneValue = new Label
                {
                    Text = "0.00 / 20",
                    Font = new Font("Segoe UI", 13, FontStyle.Bold),
                    ForeColor = Color.Red,
                    Location = new Point(770, 12),
                    Size = new Size(170, 30)
                };
                moyenneLabels[categorie] = lblMoyenneValue;
                moyennePanel.Controls.Add(lblMoyenneValue);

                mainPanel.Controls.Add(moyennePanel);
            }

            tabPage.Controls.Add(mainPanel);
        }

        private void TxtNumeric_TextChanged(object? sender, EventArgs e, Matiere matiere, bool isHeures)
        {
            if (sender == null) return;

            TextBox txt = (TextBox)sender;

            if (decimal.TryParse(txt.Text.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal value))
            {
                // Validation pour les moyennes : doit être entre 0 et 20
                if (!isHeures && value > 20)
                {
                    txt.BackColor = Color.LightCoral;
                    MessageBox.Show("La moyenne ne peut pas dépasser 20 !", "Valeur invalide", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txt.Focus();
                    return;
                }

                // Validation pour les moyennes : pas de valeur négative
                if (!isHeures && value < 0)
                {
                    txt.BackColor = Color.LightCoral;
                    MessageBox.Show("La moyenne ne peut pas être négative !", "Valeur invalide", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txt.Focus();
                    return;
                }

                // Validation pour les heures : pas de valeur négative
                if (isHeures && value < 0)
                {
                    txt.BackColor = Color.LightCoral;
                    MessageBox.Show("Les heures dédiées ne peuvent pas être négatives !", "Valeur invalide", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txt.Focus();
                    return;
                }

                // Si tout est valide, enregistrer la valeur
                if (isHeures)
                {
                    matiere.HeuresDediees = value;
                }
                else
                {
                    matiere.Moyenne = value;
                }

                txt.BackColor = Color.White;
            }
            else if (!string.IsNullOrWhiteSpace(txt.Text))
            {
                txt.BackColor = Color.LightCoral;
            }
            else
            {
                txt.BackColor = Color.White;
            }
        }

        private void BtnCalculer_Click(object? sender, EventArgs e)
        {
            // Calculer les moyennes par catégorie
            foreach (var categorie in allMatieres.Keys)
            {
                decimal sommePonderee = 0;
                int sommeCoefficients = 0;

                foreach (var matiere in allMatieres[categorie])
                {
                    sommePonderee += matiere.MoyennePonderee;
                    sommeCoefficients += matiere.Coefficient;
                }

                decimal moyenneCategorie = sommeCoefficients > 0 ? sommePonderee / sommeCoefficients : 0;

                if (moyenneLabels.ContainsKey(categorie))
                {
                    moyenneLabels[categorie].Text = $"{moyenneCategorie:F2} / 20";

                    // Couleur selon la moyenne
                    if (moyenneCategorie >= 17)
                        moyenneLabels[categorie].ForeColor = Color.DarkGreen;
                    else if (moyenneCategorie >= 14)
                        moyenneLabels[categorie].ForeColor = Color.Green;
                    else if (moyenneCategorie >= 10)
                        moyenneLabels[categorie].ForeColor = Color.Orange;
                    else
                        moyenneLabels[categorie].ForeColor = Color.Red;
                }
            }

            // Calculer la moyenne générale
            decimal moyenneGenerale = CalculerMoyenneGenerale();

            if (lblMoyenneGenerale != null)
            {
                lblMoyenneGenerale.Text = $"{moyenneGenerale:F2} / 20";

                // Couleur selon la moyenne générale
                if (moyenneGenerale >= 17)
                    lblMoyenneGenerale.ForeColor = Color.DarkGreen;
                else if (moyenneGenerale >= 14)
                    lblMoyenneGenerale.ForeColor = Color.Green;
                else if (moyenneGenerale >= 10)
                    lblMoyenneGenerale.ForeColor = Color.Orange;
                else
                    lblMoyenneGenerale.ForeColor = Color.Red;
            }

            MessageBox.Show("Calcul effectué avec succès!", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnExportPDF_Click(object? sender, EventArgs e)
        {
            try
            {
                SaveFileDialog saveDialog = new SaveFileDialog
                {
                    Filter = "Fichier PDF|*.pdf",
                    Title = "Exporter en PDF",
                    FileName = GenererNomFichierUnique("Moyennes_Semestre9_IC3_IT", "pdf")
                };

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    // Vérifier si le fichier est déjà ouvert
                    if (FichierEstOuvert(saveDialog.FileName))
                    {
                        MessageBox.Show("Le fichier est déjà ouvert dans une autre application.\nVeuillez le fermer ou choisir un autre nom.",
                            "Fichier en cours d'utilisation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    using (PdfWriter writer = new PdfWriter(saveDialog.FileName))
                    using (PdfDocument pdf = new PdfDocument(writer))
                    using (Document document = new Document(pdf))
                    {
                        // Définir une police en gras
                        var boldFont = iText.Kernel.Font.PdfFontFactory.CreateFont(iText.IO.Font.Constants.StandardFonts.HELVETICA_BOLD);
                        var normalFont = iText.Kernel.Font.PdfFontFactory.CreateFont(iText.IO.Font.Constants.StandardFonts.HELVETICA);

                        // Titre principal
                        Table headerTable = new Table(4);
                        headerTable.SetWidth(iText.Layout.Properties.UnitValue.CreatePercentValue(100));

                        Cell headerCell = new Cell(1, 4)
                            .Add(new Paragraph("UE de culture générale")
                                .SetFont(boldFont)
                                .SetFontSize(12))
                            .SetBackgroundColor(iText.Kernel.Colors.ColorConstants.LIGHT_GRAY)
                            .SetTextAlignment(TextAlignment.LEFT);
                        headerTable.AddCell(headerCell);

                        // En-têtes des colonnes
                        headerTable.AddCell(new Cell().Add(new Paragraph("").SetFont(boldFont).SetFontSize(10)));
                        headerTable.AddCell(new Cell().Add(new Paragraph("Heure dédiées").SetFont(boldFont).SetFontSize(10)).SetBackgroundColor(iText.Kernel.Colors.ColorConstants.LIGHT_GRAY));
                        headerTable.AddCell(new Cell().Add(new Paragraph("Coef").SetFont(boldFont).SetFontSize(10)).SetBackgroundColor(iText.Kernel.Colors.ColorConstants.LIGHT_GRAY));
                        headerTable.AddCell(new Cell().Add(new Paragraph("Moyenne").SetFont(boldFont).SetFontSize(10)).SetBackgroundColor(iText.Kernel.Colors.ColorConstants.LIGHT_GRAY));

                        // Pour chaque catégorie
                        int categorieNumber = 1;
                        foreach (var categorie in allMatieres.Keys)
                        {
                            // Nom de la catégorie
                            Cell catCell = new Cell(1, 4)
                                .Add(new Paragraph(categorie).SetFont(boldFont).SetFontSize(10))
                                .SetBackgroundColor(new iText.Kernel.Colors.DeviceRgb(173, 216, 230));
                            headerTable.AddCell(catCell);

                            // Matières
                            foreach (var matiere in allMatieres[categorie])
                            {
                                headerTable.AddCell(new Cell().Add(new Paragraph(matiere.Nom).SetFont(normalFont).SetFontSize(9)));
                                headerTable.AddCell(new Cell().Add(new Paragraph(matiere.HeuresDediees.ToString("F0")).SetFont(normalFont).SetFontSize(9)));
                                headerTable.AddCell(new Cell().Add(new Paragraph(matiere.Coefficient.ToString()).SetFont(normalFont).SetFontSize(9)));
                                headerTable.AddCell(new Cell().Add(new Paragraph(matiere.Moyenne.ToString("F0")).SetFont(normalFont).SetFontSize(9)));
                            }

                            // Ligne TOTAL
                            int sommeCoef = allMatieres[categorie].Sum(m => m.Coefficient);
                            Cell totalCell = new Cell(1, 3)
                                .Add(new Paragraph($"TOTAL {categorieNumber}").SetFont(boldFont).SetFontSize(9))
                                .SetBackgroundColor(iText.Kernel.Colors.ColorConstants.LIGHT_GRAY);
                            headerTable.AddCell(totalCell);
                            headerTable.AddCell(new Cell().Add(new Paragraph(sommeCoef.ToString()).SetFont(boldFont).SetFontSize(9)).SetBackgroundColor(iText.Kernel.Colors.ColorConstants.LIGHT_GRAY));

                            // Ligne Moyenne
                            decimal sommePonderee = allMatieres[categorie].Sum(m => m.MoyennePonderee);
                            decimal moy = sommeCoef > 0 ? sommePonderee / sommeCoef : 0;
                            Cell moyCell = new Cell(1, 3)
                                .Add(new Paragraph($"Moy. Générale {categorieNumber}").SetFont(boldFont).SetFontSize(9))
                                .SetBackgroundColor(new iText.Kernel.Colors.DeviceRgb(173, 216, 230));
                            headerTable.AddCell(moyCell);
                            headerTable.AddCell(new Cell()
                                .Add(new Paragraph(moy.ToString("F2")).SetFont(boldFont).SetFontSize(9).SetFontColor(iText.Kernel.Colors.ColorConstants.BLUE))
                                .SetBackgroundColor(new iText.Kernel.Colors.DeviceRgb(173, 216, 230)));

                            categorieNumber++;
                        }

                        // Ligne TOTAL final
                        Cell totalFinalCell = new Cell(1, 4)
                            .Add(new Paragraph("TOTAL").SetFont(boldFont).SetFontSize(12))
                            .SetBackgroundColor(iText.Kernel.Colors.ColorConstants.LIGHT_GRAY)
                            .SetTextAlignment(TextAlignment.CENTER);
                        headerTable.AddCell(totalFinalCell);

                        // MOYENNE GENERALE
                        decimal moyGen = CalculerMoyenneGenerale();
                        Cell moyGenCell = new Cell(1, 3)
                            .Add(new Paragraph("MOYENNE GENERALE DU SEMESTRE 9").SetFont(boldFont).SetFontSize(12))
                            .SetBackgroundColor(iText.Kernel.Colors.ColorConstants.YELLOW);
                        headerTable.AddCell(moyGenCell);
                        headerTable.AddCell(new Cell()
                            .Add(new Paragraph(moyGen.ToString("F2")).SetFont(boldFont).SetFontSize(12).SetFontColor(iText.Kernel.Colors.ColorConstants.RED))
                            .SetBackgroundColor(iText.Kernel.Colors.ColorConstants.YELLOW));

                        document.Add(headerTable);
                    }

                    MessageBox.Show("Export PDF réussi!", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // Ouvrir le fichier PDF automatiquement
                    OuvrirFichier(saveDialog.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de l'export PDF: {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnExportExcel_Click(object? sender, EventArgs e)
        {
            try
            {
                SaveFileDialog saveDialog = new SaveFileDialog
                {
                    Filter = "Fichier Excel|*.xlsx",
                    Title = "Exporter en Excel",
                    FileName = GenererNomFichierUnique("Moyennes_Semestre9_IC3_IT", "xlsx")
                };

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    // Vérifier si le fichier est déjà ouvert
                    if (FichierEstOuvert(saveDialog.FileName))
                    {
                        MessageBox.Show("Le fichier est déjà ouvert dans une autre application.\nVeuillez le fermer ou choisir un autre nom.",
                            "Fichier en cours d'utilisation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("MOYENNE GENERALE DU SEMESTRE 9");

                        int row = 1;

                        // Titre principal
                        worksheet.Cell(row, 1).Value = "UE de culture générale";
                        worksheet.Range(row, 1, row, 4).Merge();
                        worksheet.Range(row, 1, row, 4).Style.Font.Bold = true;
                        worksheet.Range(row, 1, row, 4).Style.Font.FontSize = 12;
                        worksheet.Range(row, 1, row, 4).Style.Fill.BackgroundColor = XLColor.LightGray;
                        row++;

                        // En-têtes
                        worksheet.Cell(row, 2).Value = "Heure dédiées";
                        worksheet.Cell(row, 3).Value = "Coef";
                        worksheet.Cell(row, 4).Value = "Moyenne";
                        worksheet.Range(row, 1, row, 4).Style.Font.Bold = true;
                        worksheet.Range(row, 1, row, 4).Style.Fill.BackgroundColor = XLColor.LightGray;
                        row++;

                        // Pour chaque catégorie
                        int categorieIndex = 0;
                        foreach (var categorie in allMatieres.Keys)
                        {
                            // Nom de la catégorie
                            worksheet.Cell(row, 1).Value = categorie;
                            worksheet.Range(row, 1, row, 4).Merge();
                            worksheet.Range(row, 1, row, 4).Style.Font.Bold = true;
                            worksheet.Range(row, 1, row, 4).Style.Fill.BackgroundColor = XLColor.LightBlue;
                            row++;

                            // Matières
                            foreach (var matiere in allMatieres[categorie])
                            {
                                worksheet.Cell(row, 1).Value = matiere.Nom;
                                worksheet.Cell(row, 2).Value = matiere.HeuresDediees;
                                worksheet.Cell(row, 3).Value = matiere.Coefficient;
                                worksheet.Cell(row, 4).Value = matiere.Moyenne;
                                worksheet.Range(row, 1, row, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                row++;
                            }

                            // Ligne TOTAL
                            decimal sommePonderee = allMatieres[categorie].Sum(m => m.MoyennePonderee);
                            int sommeCoef = allMatieres[categorie].Sum(m => m.Coefficient);
                            decimal moy = sommeCoef > 0 ? sommePonderee / sommeCoef : 0;

                            worksheet.Cell(row, 1).Value = $"TOTAL {categorieIndex + 1}";
                            worksheet.Range(row, 1, row, 3).Merge();
                            worksheet.Cell(row, 4).Value = sommeCoef;
                            worksheet.Range(row, 1, row, 4).Style.Font.Bold = true;
                            worksheet.Range(row, 1, row, 4).Style.Fill.BackgroundColor = XLColor.LightGray;
                            row++;

                            // Ligne Moyenne
                            worksheet.Cell(row, 1).Value = $"Moy. Générale {categorieIndex + 1}";
                            worksheet.Range(row, 1, row, 3).Merge();
                            worksheet.Cell(row, 4).Value = moy;
                            worksheet.Cell(row, 4).Style.NumberFormat.Format = "0.00";
                            worksheet.Range(row, 1, row, 4).Style.Font.Bold = true;
                            worksheet.Range(row, 1, row, 4).Style.Fill.BackgroundColor = XLColor.LightBlue;
                            worksheet.Cell(row, 4).Style.Font.FontColor = XLColor.Blue;
                            row++;

                            categorieIndex++;
                        }

                        // Ligne TOTAL final
                        worksheet.Cell(row, 1).Value = "TOTAL";
                        worksheet.Range(row, 1, row, 4).Merge();
                        worksheet.Range(row, 1, row, 4).Style.Font.Bold = true;
                        worksheet.Range(row, 1, row, 4).Style.Font.FontSize = 14;
                        worksheet.Range(row, 1, row, 4).Style.Fill.BackgroundColor = XLColor.LightGray;
                        worksheet.Range(row, 1, row, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        row++;

                        // MOYENNE GENERALE DU SEMESTRE 9
                        decimal moyGen = CalculerMoyenneGenerale();
                        worksheet.Cell(row, 1).Value = "MOYENNE GENERALE DU SEMESTRE 9";
                        worksheet.Range(row, 1, row, 3).Merge();
                        worksheet.Cell(row, 4).Value = moyGen;
                        worksheet.Cell(row, 4).Style.NumberFormat.Format = "0.00";
                        worksheet.Range(row, 1, row, 4).Style.Font.Bold = true;
                        worksheet.Range(row, 1, row, 4).Style.Font.FontSize = 14;
                        worksheet.Range(row, 1, row, 4).Style.Fill.BackgroundColor = XLColor.Yellow;
                        worksheet.Cell(row, 4).Style.Font.FontColor = XLColor.Red;

                        // Bordures autour de tout le tableau
                        worksheet.Range(1, 1, row, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

                        // Ajuster les colonnes
                        worksheet.Column(1).Width = 50;
                        worksheet.Column(2).Width = 15;
                        worksheet.Column(3).Width = 10;
                        worksheet.Column(4).Width = 15;

                        workbook.SaveAs(saveDialog.FileName);
                    }

                    MessageBox.Show("Export Excel réussi!", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // Ouvrir le fichier Excel automatiquement
                    OuvrirFichier(saveDialog.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de l'export Excel: {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnExportTxt_Click(object? sender, EventArgs e)
        {
            try
            {
                SaveFileDialog saveDialog = new SaveFileDialog
                {
                    Filter = "Fichier Texte|*.txt",
                    Title = "Exporter en TXT",
                    FileName = GenererNomFichierUnique("Moyennes_Semestre9_IC3_IT", "txt")
                };

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    // Vérifier si le fichier est déjà ouvert
                    if (FichierEstOuvert(saveDialog.FileName))
                    {
                        MessageBox.Show("Le fichier est déjà ouvert dans une autre application.\nVeuillez le fermer ou choisir un autre nom.",
                            "Fichier en cours d'utilisation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    using (StreamWriter writer = new StreamWriter(saveDialog.FileName))
                    {
                        writer.WriteLine("==========================================================");
                        writer.WriteLine("     MES OBJECTIFS DU SEMESTRE 9 IC3 IT/ESTP");
                        writer.WriteLine("==========================================================");
                        writer.WriteLine();

                        foreach (var categorie in allMatieres.Keys)
                        {
                            writer.WriteLine($"--- {categorie} ---");
                            writer.WriteLine();
                            writer.WriteLine($"{"Matière",-50} {"Heures",10} {"Coef",8} {"Moyenne",10}");
                            writer.WriteLine(new string('-', 80));

                            foreach (var matiere in allMatieres[categorie])
                            {
                                writer.WriteLine($"{matiere.Nom,-50} {matiere.HeuresDediees,10:F2} {matiere.Coefficient,8} {matiere.Moyenne,10:F2}");
                            }

                            decimal sommePonderee = allMatieres[categorie].Sum(m => m.MoyennePonderee);
                            int sommeCoef = allMatieres[categorie].Sum(m => m.Coefficient);
                            decimal moy = sommeCoef > 0 ? sommePonderee / sommeCoef : 0;

                            writer.WriteLine(new string('-', 80));
                            writer.WriteLine($"Moyenne {categorie}: {moy:F2} / 20");
                            writer.WriteLine();
                            writer.WriteLine();
                        }

                        writer.WriteLine("==========================================================");
                        decimal moyGen = CalculerMoyenneGenerale();
                        writer.WriteLine($"MOYENNE GENERALE: {moyGen:F2} / 20");
                        writer.WriteLine("==========================================================");
                    }

                    MessageBox.Show("Export TXT réussi!", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // Ouvrir le fichier TXT automatiquement
                    OuvrirFichier(saveDialog.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de l'export TXT: {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private decimal CalculerMoyenneGenerale()
        {
            decimal sommeTotalePonderee = 0;
            int sommeTotaleCoefficients = 0;

            foreach (var matieres in allMatieres.Values)
            {
                foreach (var matiere in matieres)
                {
                    sommeTotalePonderee += matiere.MoyennePonderee;
                    sommeTotaleCoefficients += matiere.Coefficient;
                }
            }

            return sommeTotaleCoefficients > 0 ? sommeTotalePonderee / sommeTotaleCoefficients : 0;
        }

        /// <summary>
        /// Ouvre un fichier avec l'application par défaut du système
        /// </summary>
        private void OuvrirFichier(string cheminFichier)
        {
            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = cheminFichier,
                    UseShellExecute = true
                };
                Process.Start(startInfo);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Impossible d'ouvrir le fichier automatiquement: {ex.Message}\n\nVous pouvez l'ouvrir manuellement à l'emplacement : {cheminFichier}",
                    "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// Vérifie si un fichier est déjà ouvert par une autre application
        /// </summary>
        private bool FichierEstOuvert(string cheminFichier)
        {
            if (!File.Exists(cheminFichier))
                return false;

            try
            {
                using (FileStream stream = File.Open(cheminFichier, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    stream.Close();
                }
                return false;
            }
            catch (IOException)
            {
                return true;
            }
        }

        /// <summary>
        /// Génère un nom de fichier unique en ajoutant un numéro si le fichier existe déjà
        /// </summary>
        private string GenererNomFichierUnique(string nomBase, string extension)
        {
            string cheminDocuments = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string nomFichier = $"{nomBase}.{extension}";
            string cheminComplet = Path.Combine(cheminDocuments, nomFichier);

            if (!File.Exists(cheminComplet) || !FichierEstOuvert(cheminComplet))
                return nomFichier;

            int compteur = 1;
            do
            {
                nomFichier = $"{nomBase}_{compteur}.{extension}";
                cheminComplet = Path.Combine(cheminDocuments, nomFichier);
                compteur++;
            }
            while (File.Exists(cheminComplet) && FichierEstOuvert(cheminComplet));

            return nomFichier;
        }
    }
}