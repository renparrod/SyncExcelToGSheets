using System;
using System.Windows.Forms;

namespace SyncExcelToGSheets
{
    public partial class ConfigSync : Form
    {
        //public string GoogleSheetId { get; set; }
        //public string GoogleSheetName { get; set; }
        //public string GoogleSheetRange { get; set; }
        //public string LocalSheetName { get; set; }
        //public string LocalSheetRange { get; set; }
        //public bool EnableSaveSheets { get; set; }

        public ConfigSync()
        {
            InitializeComponent();
        }

        private void ConfigGoogleSheet_Load(object sender, EventArgs e)
        {
            //var sheet = GoogleSheetSettings.Default.ExcelSheetName;
            //DropExcelSheetName.Items.AddRange(Globals.ThisAddIn.GetAllSheets().ToArray());
            //TextGoogleSheetId.Text = GoogleSheetSettings.Default.GoogleSheetId;
            //DropExcelSheetName.SelectedIndex = DropExcelSheetName.FindStringExact(sheet);
            //TextGoogleSheetName.Text = GoogleSheetSettings.Default.GoogleSheetName;

            DropExcelSheetName.Items.AddRange(Globals.ThisAddIn.GetAllSheets().ToArray());
            TextGoogleSheetId.Text = ConfigValues.GoogleSheetId;
            DropExcelSheetName.SelectedIndex = DropExcelSheetName.FindStringExact(ConfigValues.ExcelSheetName);
            TextGoogleSheetName.Text = ConfigValues.GoogleSheetName;
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            if (DropExcelSheetName.SelectedItem == null ||
                string.IsNullOrEmpty(DropExcelSheetName.SelectedItem.ToString()))
            {
                MessageBox.Show("Seleccione la hoja para copiar los datos");
                DropExcelSheetName.Focus();
                return;
            }

            if (string.IsNullOrEmpty(TextGoogleSheetId.Text))
            {
                MessageBox.Show("Ingrese el id de Google Sheet");
                TextGoogleSheetId.Focus();
                return;
            }

            if (string.IsNullOrEmpty(TextGoogleSheetName.Text))
            {
                MessageBox.Show("Ingrese el nombre de la hoja en Google sheet");
                TextGoogleSheetName.Focus();
                return;
            }

            ConfigValues.ExcelSheetName = DropExcelSheetName.SelectedItem.ToString();
            ConfigValues.GoogleSheetId = TextGoogleSheetId.Text;
            ConfigValues.GoogleSheetName = TextGoogleSheetName.Text;
            ConfigValues.ExcelSheetRange = "A1";
            ConfigValues.GoogleSheetRange = "A1";

            Close();
        }
    }
}
