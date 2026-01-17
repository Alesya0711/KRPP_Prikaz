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

        private void btnSave_Click(object sender, EventArgs e)
        {
            string number = txtNumber.Text.Trim();
            string city = txtCity.Text.Trim();
            string day = txtDay.Text.Trim();
            string responsible = txtResponsible.Text.Trim();
            string deputy = txtDeputy.Text.Trim();
            string person1 = txtPerson1.Text.Trim();
            string person2 = txtPerson2.Text.Trim();
            string person3 = txtPerson3.Text.Trim();

            if (string.IsNullOrEmpty(number) || string.IsNullOrEmpty(city) ||
                string.IsNullOrEmpty(day) || string.IsNullOrEmpty(responsible))
            {
                MessageBox.Show("Пожалуйста, заполните все обязательные поля.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                string fileName = $"Приказ_№{number}_{DateTime.Now:yyyyMMdd_HHmmss}.docx";
                string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName);

                CreateOrderDocument(
                    filePath, number, city, day,
                    responsible, deputy,
                    person1, person2, person3
                );

                MessageBox.Show($"Приказ успешно создан:\n{filePath}", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при создании документа:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CreateOrderDocument(
            string filePath, string number, string city, string day,
            string responsible, string deputy,
            string person1, string person2, string person3)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                body.AppendChild(new Paragraph());

                var p1 = new Paragraph(
                    new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                    new Run(new Text($"Приказ № {number}"))
                );
                body.AppendChild(p1);

                var p2 = new Paragraph(
                    new ParagraphProperties(new Justification() { Val = JustificationValues.Right }),
                    new Run(new Text($"Г. {city}"))
                );
                body.AppendChild(p2);

                var p3 = new Paragraph(
                    new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                    new Run(new Text("О назначении ответственного за архив"))
                );
                body.AppendChild(p3);

                var p4 = new Paragraph(
                    new ParagraphProperties(new Justification() { Val = JustificationValues.Right }),
                    new Run(new Text($"{day}г."))
                );
                body.AppendChild(p4);

                body.AppendChild(new Paragraph());

                AddParagraph(body, "В целях обеспечения качественного документооборота на предприятии, приказываю:");
                body.AppendChild(new Paragraph());

                AddParagraph(body, $"1. С {day} назначить ответственным за архив: {responsible}.");
                AddParagraph(body, $"2. Назначить заместителем ответственного за архив: {deputy}.");
                AddParagraph(body, $"3. {deputy} выполнять обязанности ответственного в отсутствие {responsible} на рабочем месте.");
                AddParagraph(body, $"4. Ознакомить {responsible} и {deputy} с должностной инструкцией № {number} от {day}.");
                AddParagraph(body, $"5. Контроль за исполнением возлагаю на {person1}.");

                body.AppendChild(new Paragraph());

                AddParagraph(body, $"Руководитель  {person1} (__________)");


                AddParagraph(body, "С приказом ознакомлены:");
                //body.AppendChild(new Paragraph());
                AddParagraph(body, $"________________   {person1}   (___________)");
                AddParagraph(body, $"________________   {person2}   (___________)");
                AddParagraph(body, $"________________   {person3}   (___________)");
            }
        }

        private void AddParagraph(Body body, string text)
        {
            body.AppendChild(new Paragraph(
                new Run(new Text(text))
            ));
        }
    } 
}
