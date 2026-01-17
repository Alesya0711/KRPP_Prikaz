using System;
using System.IO;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace КРПП_тест
{
    public partial class Form1 : Form
    {
        private Label lblNumber, lblCity, lblDay, lblResponsible, lblDeputy, lblPerson1, lblPerson2, lblPerson3;
        private TextBox txtNumber, txtCity, txtDay, txtResponsible, txtDeputy, txtPerson1, txtPerson2, txtPerson3;
        private Button btnSave;

        public Form1()
        {
            InitializeComponent();
            SetupUI();
        }

        private void SetupUI()
        {
            int y = 20;
            int spacing = 40;

            lblNumber = new Label { Text = "№ приказа:", Location = new System.Drawing.Point(20, y), AutoSize = true };
            txtNumber = new TextBox { Location = new System.Drawing.Point(120, y), Width = 100 };
            y += spacing;

            lblCity = new Label { Text = "Город:", Location = new System.Drawing.Point(20, y), AutoSize = true };
            txtCity = new TextBox { Location = new System.Drawing.Point(120, y), Width = 350 };
            y += spacing;

            lblDay = new Label { Text = "Дата:", Location = new System.Drawing.Point(20, y), AutoSize = true };
            txtDay = new TextBox { Location = new System.Drawing.Point(120, y), Width = 100 };
            y += spacing;

            lblResponsible = new Label { Text = "Ответственный:", Location = new System.Drawing.Point(20, y), AutoSize = true };
            txtResponsible = new TextBox { Location = new System.Drawing.Point(120, y), Width = 350 };
            y += spacing;

            lblDeputy = new Label { Text = "Заместитель:", Location = new System.Drawing.Point(20, y), AutoSize = true };
            txtDeputy = new TextBox { Location = new System.Drawing.Point(120, y), Width = 350 };
            y += spacing;

            lblPerson1 = new Label { Text = "Ознакомлен 1:", Location = new System.Drawing.Point(20, y), AutoSize = true };
            txtPerson1 = new TextBox { Location = new System.Drawing.Point(120, y), Width = 350 };
            y += spacing;

            lblPerson2 = new Label { Text = "Ознакомлен 2:", Location = new System.Drawing.Point(20, y), AutoSize = true };
            txtPerson2 = new TextBox { Location = new System.Drawing.Point(120, y), Width = 350 };
            y += spacing;

            lblPerson3 = new Label { Text = "Ознакомлен 3:", Location = new System.Drawing.Point(20, y), AutoSize = true };
            txtPerson3 = new TextBox { Location = new System.Drawing.Point(120, y), Width = 350 };
            y += spacing;

            btnSave = new Button
            {
                Text = "Создать приказ",
                Location = new System.Drawing.Point(170, y + 20),
                Width = 150
            };

            btnSave.Click += btnSave_Click;

            this.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                lblNumber, txtNumber,
                lblCity, txtCity,
                lblResponsible, txtResponsible,
                lblDeputy, txtDeputy,
                lblPerson1, txtPerson1,
                lblPerson2, txtPerson2,
                lblPerson3, txtPerson3,
                btnSave, lblDay,txtDay
            });

            this.Text = "Генератор приказа";
            this.Size = new System.Drawing.Size(520, 480);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
        }
