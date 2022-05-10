using ModelodeAprov.Controller;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ModelodeAprov
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }

        private void Login_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            LoginUser.user= textBox1.Text;
            LoginUser.password = textBox2.Text;

            
            ConectaSAP.ConectaSap(LoginUser.user, LoginUser.password);

            if (ConectaSAP.oCompany.Connected)
            {
                
                Form1 form1 = new Form1();
                this.Hide();
                form1.Show();
             
            }
            else
            {
                string erroconect = ConectaSAP.oCompany.GetLastErrorDescription();
                int errorcode = ConectaSAP.oCompany.GetLastErrorCode();

                if ( errorcode == -132)
                {
                    MessageBox.Show("Verificar usuário ou senha incorretos!!" + erroconect, "Erro de Login" ,MessageBoxButtons.OKCancel,MessageBoxIcon.Exclamation);

                }

            }
         


        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
