using students_skills_validator.Models;
using System.Globalization;
using System.Reflection;
using System.Security.Principal;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Xml.XPath;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using ProgressBar = System.Windows.Forms.ProgressBar;

namespace students_skills_validator
{
    public partial class FormSkillsValidator : Form
    {
        private OpenFileDialog _openFileDialog;
        private SaveFileDialog _saveFileDialog;

        public StudentFile? currentStudentFile;
        public RichTextBox? currentRichTextBox;

        List<Label> _Labels = new List<Label>();
        List<ProgressBar> _progressBars = new List<ProgressBar>();

        XmlSerializer xmlSerializer;
        List<StudentFile> studentFiles;
        private readonly XmlWriterSettings _writerSettings;

        //private static string _applicationDataPath = "C:\\Users\\goupi\\Desktop\\";
        private static string _applicationDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        private static string _applicationFolderPath = Path.Combine(_applicationDataPath, "SkillsValidator.NET");
        private static string _applicationRefPath = Path.Combine(_applicationFolderPath, "Referentials");

        public FormSkillsValidator()
        {
            InitializeComponent();

            studentFiles = new List<StudentFile>();
            _writerSettings = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = ("\t"),
                OmitXmlDeclaration = true
            };
            xmlSerializer = new XmlSerializer(typeof(List<StudentFile>));

            // Création des dossiers
            if (!Directory.Exists(_applicationFolderPath))
            {
                Directory.CreateDirectory(_applicationFolderPath);
            }
            if (!Directory.Exists(_applicationRefPath))
            {
                Directory.CreateDirectory(_applicationRefPath);
            }

            _openFileDialog = new OpenFileDialog();
            _saveFileDialog = new SaveFileDialog();
        }

        private void nouveauToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var _studentForm = new StudentForm();
            _studentForm.Show();
        }

        public void ouvrirToolStripMenuItem_ClickAsync(object sender, EventArgs e)
        {
            // Ouverture d'un fichier élève.
            _openFileDialog.InitialDirectory = _applicationFolderPath;
            if (_openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    XDocument xmlStudentDoc = XDocument.Load(_openFileDialog.FileName);
                    XElement xmlStudentRoot = xmlStudentDoc.Root.Element("StudentFile");

                    if (xmlStudentRoot == null && Path.GetExtension(_openFileDialog.FileName) != ".xml")
                    {
                        MessageBox.Show(
                            "Ce fichier ne semble pas correspondre au format attendu.",
                            "ERREUR",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );
                    }

                    string firstName = xmlStudentRoot.Element("FirstName").Value;
                    string lastName = xmlStudentRoot.Element("LastName").Value;
                    string referential = xmlStudentRoot.Element("RefName").Value;
                    string updatedAt = xmlStudentRoot.Element("updatedAt").Value;

                    // Netoie la ListView.
                    studentListView.Clear();

                    // Défini la ListView dans l'interface.
                    ColumnHeader header = new ColumnHeader();
                    header.Text = "Student";
                    header.TextAlign = HorizontalAlignment.Left;
                    studentListView.Columns.Add("Student");

                    studentListView.View = View.Details;
                    studentListView.GridLines = true;

                    studentListView.Items.Add(new ListViewItem(new string[]
                    {
                        firstName
                    }));
                    studentListView.Items.Add(new ListViewItem(new string[]
                    {
                        lastName
                    }));
                    studentListView.Items.Add(new ListViewItem(new string[]
                    {
                        referential
                    }));
                    studentListView.Items.Add(new ListViewItem(new string[]
                    {
                        updatedAt
                    }));

                    studentListView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
                    studentListView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
                    studentListView.Columns[0].Width = studentListView.Width - 4 - SystemInformation.VerticalScrollBarWidth;

                    // Renomme le nom du validateur et longlet du menu strip.
                    this.Text = "Validateur de compétences : " + referential.ToUpper();
                    toolActualLoaded.Text = referential.ToUpper();
                    toolActualLoaded.Checked = true;

                    // Si compétences définis, on initialise le TreeView et on affiche les compétences.
                    if (xmlStudentRoot.Element("Referential") != null)
                    {
                        skillsTreeView.Nodes.Clear();
                        skillsTreeView.CheckBoxes = true;
                        TreeNode node;
                        int index = 0;

                        IEnumerable<XElement> blocks = from rskl in xmlStudentRoot.Descendants("Referential").Elements("Skills") select rskl;
                        foreach (XElement block in blocks)
                        {
                            // Obtenir la règle des 70% de validation.
                            int ruleSeventyPercent = (int)Math.Ceiling(block.Descendants("Skill").Count() * 0.7);

                            // Nombre de compétences majeurs non-validées.
                            int numbersOfInvalidMajSkills = block.Descendants("Skill").Where(s => s.Attribute("type").Value == "maj" && s.Attribute("validate").Value == "0").Count();

                            // Ajout du titre du Block
                            skillsTreeView.Nodes.Add(block.Attribute("title").Value);

                            // Ajout des jauges.
                            var label = new Label();
                            label.Text = "Block - " + block.Attribute("title").Value.Substring(0, 1);

                            var progressBar = new ProgressBar();
                            progressBar.Visible = true;
                            progressBar.Minimum = 0;
                            progressBar.Maximum = block.Descendants("Skill").Count();
                            progressBar.Value = 0;
                            progressBar.Step = 1;

                            _Labels.Add(label);
                            _progressBars.Add(progressBar);

                            FlowLayoutPanel.Controls.Add(_Labels[index]);
                            FlowLayoutPanel.Controls.Add(_progressBars[index]);

                            foreach (XElement skill in block.Descendants("Skill"))
                            {
                                node = skillsTreeView.Nodes[index].Nodes.Add(skill.Value);

                                // Compétences majeurs en gras.
                                if (skill.Attribute("type").Value == "maj")
                                {
                                    node.NodeFont = new Font(skillsTreeView.Font, FontStyle.Bold);
                                }

                                // Coche la checkbox si l'attribut est à 1.
                                if (skill.Attribute("validate").Value == "1")
                                {
                                    node.Checked = true;
                                    _progressBars[index].PerformStep();
                                }
                                else
                                {
                                    node.Checked = false;
                                }
                            }

                            // Applique règle visuel des 70%.
                            if (_progressBars[index].Value >= ruleSeventyPercent && numbersOfInvalidMajSkills == 0)
                            {
                                _progressBars[index].Value = _progressBars[index].Maximum;
                            }
                            index++;
                        }
                    }

                    // Sauvegarde de l'horodatage.
                    xmlStudentRoot.Element("updatedAt").Remove();
                    xmlStudentRoot.Element("RefName").AddAfterSelf(
                        new XElement("updatedAt", DateTime.Now.ToString("MMM ddd d HH:mm yyyy"))
                    );
                    xmlStudentDoc.Save(_openFileDialog.FileName);
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

        private void tsmiQuit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void editReferential_Click(object sender, EventArgs e)
        {
            var isEditable = true;

            // Chargement d'un nouveau référentiel.
            _openFileDialog.InitialDirectory = _applicationRefPath;
            if (_openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Clone le fichier de référentiel XML dans le dossier cible de l'application.
                string targetRefName = Path.GetFileNameWithoutExtension(_openFileDialog.FileName);
                string refPath = Path.Combine(_applicationRefPath, targetRefName + ".xml");

                if (!File.Exists(refPath))
                {
                    File.Copy(_openFileDialog.FileName, refPath);
                }

                if (toolActualLoaded.Text != "")
                {
                    string actualStudentFileName = studentListView.Items[0].Text.ToLower() + "-" + studentListView.Items[1].Text.ToLower() + "-" + studentListView.Items[2].Text.ToLower() + ".xml";
                    string actualStudentFileNamePath = Path.Combine(_applicationFolderPath, actualStudentFileName);


                    // Vérification qu'une compétence est validée.
                    try
                    {
                        TreeNodeCollection nodes = skillsTreeView.Nodes;
                        foreach (TreeNode node in nodes)
                        {
                            foreach (TreeNode subNode in node.Nodes)
                            {
                                if (subNode.Checked)
                                {
                                    isEditable = false;
                                    throw new Exception("Une ou plusieurs compétences sont validé et empêchent le changement de référentiels pour cet élève.");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            ex.Message,
                            "ERREUR",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );
                    }

                    // Si la modification est possible.
                    if (isEditable)
                    {
                        string newStudentFileName = studentListView.Items[0].Text.ToLower() + "-" + studentListView.Items[1].Text.ToLower() + "-" + targetRefName.ToLower() + ".xml";
                        string newStudentFileNamePath = Path.Combine(_applicationFolderPath, newStudentFileName);

                        this.Text = "Validateur de compétences : " + targetRefName.ToUpper();
                        toolActualLoaded.Text = targetRefName.ToUpper();
                        toolActualLoaded.Checked = true;

                        // Copie le contenu du nouveau référentiel dans le fichier de l'élève si aucune compétences validé.
                        try
                        {
                            XDocument xmlStudentDoc = XDocument.Load(actualStudentFileNamePath);
                            XElement xmlStudentRoot = xmlStudentDoc.Root.Element("StudentFile");

                            using (var xmlReader = new StreamReader(refPath))
                            {
                                XDocument refDoc = XDocument.Load(xmlReader);
                                xmlStudentRoot.Element("Referential").Remove();
                                xmlStudentRoot.Add(refDoc.Root);

                                xmlStudentDoc.Save(actualStudentFileNamePath);
                                File.Move(actualStudentFileNamePath, newStudentFileNamePath);
                            }

                            studentListView.Items[2].Text = targetRefName.ToUpper();
                            skillsTreeView.Nodes.Clear();

                            // Reconstruction de l'interface avec le nouveau referentiel.
                            if (xmlStudentRoot.Element("Referential") != null)
                            {
                                skillsTreeView.CheckBoxes = true;
                                TreeNode node;
                                int index = 0;

                                IEnumerable<XElement> blocks = from rskl in xmlStudentRoot.Descendants("Referential").Elements("Skills") select rskl;
                                foreach (XElement block in blocks)
                                {
                                    // Obtenir la règle des 70% de validation.
                                    int ruleSeventyPercent = (int)Math.Ceiling(block.Descendants("Skill").Count() * 0.7);

                                    // Nombre de compétences majeurs non-validées.
                                    int numbersOfInvalidMajSkills = block.Descendants("Skill").Where(s => s.Attribute("type").Value == "maj" && s.Attribute("validate").Value == "0").Count();

                                    // Ajout du titre du Block
                                    skillsTreeView.Nodes.Add(block.Attribute("title").Value);

                                    // Ajout des jauges.
                                    var label = new Label();
                                    label.Text = "Block - " + block.Attribute("title").Value.Substring(0, 1);

                                    var progressBar = new ProgressBar();
                                    progressBar.Visible = true;
                                    progressBar.Minimum = 0;
                                    progressBar.Maximum = block.Descendants("Skill").Count();
                                    progressBar.Value = 0;
                                    progressBar.Step = 1;

                                    _Labels.Add(label);
                                    _progressBars.Add(progressBar);

                                    FlowLayoutPanel.Controls.Add(_Labels[index]);
                                    FlowLayoutPanel.Controls.Add(_progressBars[index]);

                                    foreach (XElement skill in block.Descendants("Skill"))
                                    {
                                        node = skillsTreeView.Nodes[index].Nodes.Add(skill.Value);

                                        // Compétences majeurs en gras.
                                        if (skill.Attribute("type").Value == "maj")
                                        {
                                            node.NodeFont = new Font(skillsTreeView.Font, FontStyle.Bold);
                                        }

                                        // Coche la checkbox si l'attribut est à 1.
                                        if (skill.Attribute("validate").Value == "1")
                                        {
                                            node.Checked = true;
                                            _progressBars[index].PerformStep();
                                        }
                                        else
                                        {
                                            node.Checked = false;
                                        }
                                    }

                                    // Applique règle visuel des 70%.
                                    if (_progressBars[index].Value >= ruleSeventyPercent && numbersOfInvalidMajSkills == 0)
                                    {
                                        _progressBars[index].Value = _progressBars[index].Maximum;
                                    }
                                    index++;
                                }
                            }

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(
                                ex.Message,
                                "ERREUR",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error
                            );
                        }
                    }
                }
                else
                {
                    MessageBox.Show(
                        "Vous devez d'abord ouvrir le fichier d'un élève avant de changer le référentiel pour celui-ci.",
                        "ERREUR",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
        }

        private void skillsTreeView_AfterCheck(object sender, TreeViewEventArgs e)
        {
            string actualStudentFileName = studentListView.Items[0].Text.ToLower() + "-" + studentListView.Items[1].Text.ToLower() + "-" + studentListView.Items[2].Text.ToLower() + ".xml";
            string actualStudentFileNamePath = Path.Combine(_applicationFolderPath, actualStudentFileName);

            try
            {
                //using (var xmlReader = new StreamReader(actualStudentFileNamePath))
                XDocument xmlActualStudentDoc = XDocument.Load(actualStudentFileNamePath);
                XElement xmlActualStudentRoot = xmlActualStudentDoc.Root.Element("StudentFile");
                int index = 0;

                // Coche l'intégralité des noeuds enfants.
                foreach (TreeNode tn in e.Node.Nodes)
                {
                    tn.Checked = e.Node.Checked;
                }

                IEnumerable<XElement> blocks = from rskl in xmlActualStudentRoot.Descendants("Referential").Elements("Skills") select rskl;
                foreach (XElement block in blocks)
                {
                    // Obtenir la règle des 70% de validation.
                    int ruleSeventyPercent = (int)Math.Ceiling(block.Descendants("Skill").Count() * 0.7);

                    // Nombre de compétences validées dans le fichier.
                    int numbersOfValidSkills = block.Descendants("Skill").Where(s => s.Attribute("validate").Value == "1").Count();

                    // Nombre de compétences invalidées dans le fichier.
                    int numbersOfInvalidSkills = block.Descendants("Skill").Where(s => s.Attribute("validate").Value == "0").Count();

                    // Nombre de compétences majeurs non-validées.
                    int numbersOfInvalidMajSkills = block.Descendants("Skill").Where(s => s.Attribute("type").Value == "maj" && s.Attribute("validate").Value == "0").Count();


                    foreach (XElement skill in block.Descendants("Skill"))
                    {
                        if (skill.Value == e.Node.Text)
                        {
                            if (e.Node.Checked)
                            {
                                skill.SetAttributeValue("validate", "1");
                                _progressBars[index].Increment(1);

                                xmlActualStudentDoc.Root.Element("Logs").Add(
                                    new XElement("Log",
                                        new XAttribute("created-at", DateTime.Now),
                                        new XAttribute("message", "The Skill : " + e.Node.Text + " has been validate.")
                                    )
                                );

                                // Applique la règle des 70% en incrémentation.
                                if (_progressBars[index].Value >= ruleSeventyPercent && numbersOfInvalidMajSkills == 0)
                                {
                                    _progressBars[index].Value = _progressBars[index].Maximum;
                                }

                                // Log valide block.
                                if (_progressBars[index].Value == _progressBars[index].Maximum)
                                {
                                    xmlActualStudentDoc.Root.Element("Logs").Add(
                                        new XElement("Log",
                                            new XAttribute("created-at", DateTime.Now),
                                            new XAttribute("message", "The Block : " + block.Attribute("title").Value + " is completly validate.")
                                        )
                                    );
                                }
                            }
                            else
                            {
                                skill.SetAttributeValue("validate", "0");

                                // Applique la règle des 70% en décrémentation.
                                if (numbersOfValidSkills == ruleSeventyPercent)
                                {
                                    _progressBars[index].Value = numbersOfValidSkills - 1;
                                    xmlActualStudentDoc.Root.Element("Logs").Add(
                                        new XElement("Log",
                                            new XAttribute("created-at", DateTime.Now),
                                            new XAttribute("message", "The Block : " + block.Attribute("title").Value + " is completly invalidate.")
                                        )
                                    );
                                }
                                else
                                {
                                    _progressBars[index].Increment(-1);
                                }
                                xmlActualStudentDoc.Root.Element("Logs").Add(
                                    new XElement("Log",
                                        new XAttribute("created-at", DateTime.Now),
                                        new XAttribute("message", "The Skill : " + e.Node.Text + " has been invalidate.")
                                    )
                                );
                            }
                            xmlActualStudentDoc.Save(actualStudentFileNamePath);
                            break;
                        }
                    }
                    index++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "ERREUR",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }
    }
}