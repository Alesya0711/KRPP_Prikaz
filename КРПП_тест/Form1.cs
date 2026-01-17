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

