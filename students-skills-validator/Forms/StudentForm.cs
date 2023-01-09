using Microsoft.VisualBasic.Logging;
using students_skills_validator.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Button = System.Windows.Forms.Button;
using Label = System.Windows.Forms.Label;
using ProgressBar = System.Windows.Forms.ProgressBar;
using TextBox = System.Windows.Forms.TextBox;

namespace students_skills_validator
{
    public class StudentForm : Form
    {
        public FormSkillsValidator formSkillsValidator;

        private OpenFileDialog _openFileDialog;

        private static string _applicationDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        private static string _applicationPath = Path.Combine(_applicationDataPath, "SkillsValidator.NET");
        private static string _applicationRefPath = Path.Combine(_applicationPath, "Referentials");

        private readonly XmlWriterSettings _writerSettings;
        private XmlSerializer _xmlSerializer;
        private List<StudentFile> _studentFiles;

        private Label lbName;
        private Label lbReferential;
        private TextBox txtbFirstName;
        private Button btnCreateStudent;
        private TextBox txtbLastName;
        private Label lbFirstName;
        private Button btnLoadRef;
        private Label lblLoadRef;

        public StudentForm()
        {
            InitializeComponent();

            _writerSettings = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = ("\t"),
                OmitXmlDeclaration = true
            };
            _studentFiles = new List<StudentFile>();
            _xmlSerializer = new XmlSerializer(typeof(List<StudentFile>));

            if (!Directory.Exists(_applicationPath))
            {
                Directory.CreateDirectory(_applicationPath);
            }

            _openFileDialog = new OpenFileDialog();
        }

        public void InitializeComponent()
        {
            this.lbName = new System.Windows.Forms.Label();
            this.lbReferential = new System.Windows.Forms.Label();
            this.txtbFirstName = new System.Windows.Forms.TextBox();
            this.btnCreateStudent = new System.Windows.Forms.Button();
            this.txtbLastName = new System.Windows.Forms.TextBox();
            this.lbFirstName = new System.Windows.Forms.Label();
            this.btnLoadRef = new System.Windows.Forms.Button();
            this.lblLoadRef = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lbName
            // 
            this.lbName.AutoSize = true;
            this.lbName.Location = new System.Drawing.Point(30, 99);
            this.lbName.Name = "lbName";
            this.lbName.Size = new System.Drawing.Size(34, 15);
            this.lbName.TabIndex = 0;
            this.lbName.Text = "Nom";
            // 
            // lbReferential
            // 
            this.lbReferential.AutoSize = true;
            this.lbReferential.Location = new System.Drawing.Point(30, 167);
            this.lbReferential.Name = "lbReferential";
            this.lbReferential.Size = new System.Drawing.Size(63, 15);
            this.lbReferential.TabIndex = 1;
            this.lbReferential.Text = "Referentiel";
            // 
            // txtbFirstName
            // 
            this.txtbFirstName.Location = new System.Drawing.Point(122, 27);
            this.txtbFirstName.Name = "txtbFirstName";
            this.txtbFirstName.Size = new System.Drawing.Size(167, 23);
            this.txtbFirstName.TabIndex = 2;
            // 
            // btnCreateStudent
            // 
            this.btnCreateStudent.BackColor = System.Drawing.SystemColors.ScrollBar;
            this.btnCreateStudent.Location = new System.Drawing.Point(95, 238);
            this.btnCreateStudent.Name = "btnCreateStudent";
            this.btnCreateStudent.Size = new System.Drawing.Size(153, 23);
            this.btnCreateStudent.TabIndex = 4;
            this.btnCreateStudent.Text = "Créer";
            this.btnCreateStudent.UseVisualStyleBackColor = false;
            this.btnCreateStudent.Click += new System.EventHandler(this.btnCreateStudent_ClickAsync);
            // 
            // txtbLastName
            // 
            this.txtbLastName.Location = new System.Drawing.Point(122, 91);
            this.txtbLastName.Name = "txtbLastName";
            this.txtbLastName.Size = new System.Drawing.Size(167, 23);
            this.txtbLastName.TabIndex = 5;
            // 
            // lbFirstName
            // 
            this.lbFirstName.AutoSize = true;
            this.lbFirstName.Location = new System.Drawing.Point(30, 35);
            this.lbFirstName.Name = "lbFirstName";
            this.lbFirstName.Size = new System.Drawing.Size(49, 15);
            this.lbFirstName.TabIndex = 6;
            this.lbFirstName.Text = "Prénom";
            // 
            // btnLoadRef
            // 
            this.btnLoadRef.Location = new System.Drawing.Point(194, 163);
            this.btnLoadRef.Name = "btnLoadRef";
            this.btnLoadRef.Size = new System.Drawing.Size(95, 23);
            this.btnLoadRef.TabIndex = 7;
            this.btnLoadRef.Text = "Charger";
            this.btnLoadRef.UseVisualStyleBackColor = true;
            this.btnLoadRef.Click += new System.EventHandler(this.btnLoadRef_Click);
            // 
            // lblLoadRef
            // 
            this.lblLoadRef.AutoSize = true;
            this.lblLoadRef.Location = new System.Drawing.Point(122, 167);
            this.lblLoadRef.Name = "lblLoadRef";
            this.lblLoadRef.Size = new System.Drawing.Size(42, 15);
            this.lblLoadRef.TabIndex = 8;
            this.lblLoadRef.Text = "Aucun";
            // 
            // StudentForm
            // 
            this.ClientSize = new System.Drawing.Size(396, 284);
            this.Controls.Add(this.lblLoadRef);
            this.Controls.Add(this.btnLoadRef);
            this.Controls.Add(this.lbFirstName);
            this.Controls.Add(this.txtbLastName);
            this.Controls.Add(this.btnCreateStudent);
            this.Controls.Add(this.txtbFirstName);
            this.Controls.Add(this.lbReferential);
            this.Controls.Add(this.lbName);
            this.MinimumSize = new System.Drawing.Size(360, 230);
            this.Name = "StudentForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Création de l\'élève";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void btnCreateStudent_ClickAsync(object sender, EventArgs e)
        {
            var alreadyExists = false;

            var newFileName = txtbFirstName.Text.ToLower() + "-" + txtbLastName.Text.ToLower() + "-" + lblLoadRef.Text.ToLower() + ".xml";
            var filePath = Path.Combine(_applicationPath, newFileName);

            // Récupère les fichiers contenus dans un dossier et les intègres dans un Array<string>.
            string[] filePaths = Directory.GetFiles(_applicationPath);
            foreach (var file in filePaths)
            {
                if (file == filePath)
                {
                    MessageBox.Show(
                        "Un fichier portant ce nom existe déjà dans SkillsValidator.NET",
                        "ERREUR",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                    alreadyExists = true;
                    break;
                }
            }

            if (!alreadyExists)
            {
                StudentFile studentFile = new StudentFile();
                studentFile.FileName = txtbFirstName.Text + '-' + txtbLastName.Text;
                studentFile.FirstName = txtbFirstName.Text.ToUpper();
                studentFile.LastName = txtbLastName.Text.ToUpper();
                studentFile.RefName = lblLoadRef.Text.ToUpper();

                _studentFiles.Add(studentFile);

                // Formate le XML à la création du fichier grâce à SerializerNamespace et XmlWriter.
                var emptyNamespace = new XmlSerializerNamespaces(new[]
                {
                    XmlQualifiedName.Empty
                });
                using (XmlWriter writer = XmlWriter.Create(filePath, _writerSettings))
                {
                    _xmlSerializer.Serialize(writer, _studentFiles, emptyNamespace);
                }
                Hide();

                // Copie le contenu du référentiel dans le fichier de l'élève.
                try
                {
                    XDocument xmlStudentDoc = XDocument.Load(filePath);
                    XElement xmlStudentRoot = xmlStudentDoc.Root.Element("StudentFile");

                    if (xmlStudentRoot == null && Path.GetExtension(filePath) != ".xml")
                    {
                        MessageBox.Show(
                            "Ce fichier ne semble pas correspondre au format attendu.",
                            "ERREUR",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );
                    }

                    // Ajout d'une date au bon format.
                    xmlStudentRoot.Element("RefName").AddAfterSelf(
                        new XElement("updatedAt", DateTime.Now.ToString("MMM ddd d HH:mm yyyy"))
                    );

                    // Création des Logs dans le nouveau fichier élève.
                    xmlStudentRoot.AddAfterSelf(
                        new XComment("This XML Logs contain change of this student for SkillsValidator.NET application"),
                        new XComment("XML logs generated by Software Installer"),
                        new XElement("Logs",
                                new XElement("Log",
                                new XAttribute("created-at", DateTime.Now),
                                new XAttribute("message", "New Student file : " + studentFile.FirstName + " " + studentFile.LastName + " is successfull create.")
                            )
                        )
                    );

                    var targetRefFileName = lblLoadRef.Text.ToLower() + ".xml";
                    string refPath = Path.Combine(_applicationRefPath, targetRefFileName);
                    if (xmlStudentRoot != null && refPath != null)
                    {
                        using (var xmlReader = new StreamReader(refPath))
                        {
                            XDocument refDoc = XDocument.Load(xmlReader);
                            xmlStudentRoot.Add(refDoc.Root);

                            xmlStudentDoc.Save(filePath);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        "Problème à l'ouverture du fichier : " + ex.Message,
                        "ERREUR",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
            }
        }

        private void btnLoadRef_Click(object sender, EventArgs e)
        {
            if (_openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Clone le fichier de référentiel XML dans le dossier cible de l'application.
                string targetRefName = Path.GetFileNameWithoutExtension(_openFileDialog.FileName);
                string targetRefFileName = targetRefName + ".xml";
                string refPath = Path.Combine(_applicationRefPath, targetRefFileName);

                lblLoadRef.Text = Path.GetFileNameWithoutExtension(_openFileDialog.FileName);

                if (!File.Exists(refPath))
                {
                    File.Copy(_openFileDialog.FileName, refPath);
                }
            }
        }
    }
}
