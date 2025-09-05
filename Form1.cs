using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Text.Json;

namespace ExcelToDB
{
    
   

    public partial class Form1 : Form
    {
       
        private List<ConnectionInfo> _savedConnections = new List<ConnectionInfo>();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cmbSavedConnections.SelectedIndex = -1;
            txtConnectionName.Text = "";
            txtServer.Text = "";
            txtUser.Text = "";
            txtPassword.Text = "";


            LoadSavedConnections();
        }

        private void LoadSavedConnections()
        {
           
            cmbSavedConnections.SelectedIndexChanged -= cmbSavedConnections_SelectedIndexChanged;

            // Ayarlardan JSON verisini oku
            string connectionsJson = Properties.Settings.Default.SavedConnections;

            if (!string.IsNullOrEmpty(connectionsJson))
            {
                try
                {
                    _savedConnections = JsonSerializer.Deserialize<List<ConnectionInfo>>(connectionsJson);
                }
                catch (JsonException)
                {
                    _savedConnections = new List<ConnectionInfo>();
                }
            }
            else
            {
                _savedConnections = new List<ConnectionInfo>();
            }

            
            cmbSavedConnections.DataSource = null;
            cmbSavedConnections.DataSource = _savedConnections;
            cmbSavedConnections.DisplayMember = "Name";
            cmbSavedConnections.SelectedIndex = -1; 

           
            cmbSavedConnections.SelectedIndexChanged += cmbSavedConnections_SelectedIndexChanged;
        }

        private void SaveConnections()
        {
            
            string connectionsJson = JsonSerializer.Serialize(_savedConnections);
           
            Properties.Settings.Default.SavedConnections = connectionsJson;
            Properties.Settings.Default.Save(); 
        }

       
        private void cmbSavedConnections_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbSavedConnections.SelectedItem is ConnectionInfo selectedConnection)
            {
              
                txtConnectionName.Text = selectedConnection.Name;
                txtServer.Text = selectedConnection.Server;
                txtUser.Text = selectedConnection.User;
                txtPassword.Text = selectedConnection.Password;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            if (chkSaveDetails.Checked)
            {
                if (string.IsNullOrWhiteSpace(txtConnectionName.Text))
                {
                    MessageBox.Show("Lütfen kaydetmek için bir bağlantı adı girin.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                
                var existingConnection = _savedConnections.Find(c => c.Name.Equals(txtConnectionName.Text, StringComparison.OrdinalIgnoreCase));

                if (existingConnection != null)
                {
                  
                    existingConnection.Server = txtServer.Text;
                    existingConnection.User = txtUser.Text;
                    existingConnection.Password = txtPassword.Text;
                }
                else
                {
                    
                    _savedConnections.Add(new ConnectionInfo
                    {
                        Name = txtConnectionName.Text,
                        Server = txtServer.Text,
                        User = txtUser.Text,
                        Password = txtPassword.Text
                    });
                }

               
                SaveConnections();
                LoadSavedConnections();
                MessageBox.Show("Bağlantı bilgileri kaydedildi!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }


          
            string server = txtServer.Text;
            string user = txtUser.Text;
            string password = txtPassword.Text;
            string connStr = $"Server={server};User Id={user};Password={password};TrustServerCertificate=True;";

            try
            {
                using (SqlConnection connection = new SqlConnection(connStr))
                {
                    connection.Open();
                    MessageBox.Show("Bağlantı başarılı!");

                    Form2 tablesForm = new Form2(connStr);
                    tablesForm.Show();
                    this.Hide();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Bağlantı hatası: " + ex.Message);
            }
        }

       
        private void btnDeleteConnection_Click(object sender, EventArgs e)
        {
            if (cmbSavedConnections.SelectedItem is ConnectionInfo selectedConnection)
            {
                var result = MessageBox.Show($"'{selectedConnection.Name}' adlı profili silmek istediğinize emin misiniz?", "Silme Onayı", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    _savedConnections.Remove(selectedConnection);
                    SaveConnections();
                    LoadSavedConnections();
                    
                    txtConnectionName.Text = "";
                    txtServer.Text = "";
                    txtUser.Text = "";
                    txtPassword.Text = "";
                    MessageBox.Show("Profil silindi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                MessageBox.Show("Lütfen silmek için bir profil seçin.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        public class ConnectionInfo
        {
            public string Name { get; set; }
            public string Server { get; set; }
            public string User { get; set; }
            public string Password { get; set; }
        }
    }
  
}