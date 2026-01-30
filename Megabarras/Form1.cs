using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Net.Mail;
using System.Data.SqlClient;

namespace Megabarras
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int li_filas = 0 ;
            string usuario, clave;
            SqlConnection con = null;
            SqlCommand cmd = null;

            try
            {
                usuario = textBox1.Text.Trim();
                clave = textBox2.Text.Trim();
                string ls_query = " SELECT tbl_clave,tbl_Apellidos ,tbl_Nombre,  tbl_fecha_creacion FROM Megabarras.dbo.Tbl_usuarios  where tbl_usuario=@us and tbl_clave=@cl ";
                con = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["barras"].ConnectionString);
                con.Open();
                cmd = new SqlCommand(ls_query, con);
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.AddWithValue("@us", usuario);
                cmd.Parameters.AddWithValue("@cl", clave);
                li_filas = Convert.ToInt32(cmd.ExecuteScalar());

                if (li_filas > 0)
                {
                    cmd.Parameters.Clear();
                    DialogResult = DialogResult.OK;

                }
                else
                {
                    MessageBox.Show("Error de Inicio de Sesion", "Inicio De Sesion");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Relacionado con la base datos :"+' '+ex.ToString(),"Error al iniciar sesion");
            }
            finally
            {
                if (con != null)
                    con.Dispose();

                if (cmd != null)
                    cmd.Dispose();
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            //
            if (e.CloseReason == CloseReason.UserClosing)
            {
                Application.Exit();
            }      
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            Application.Exit();
        }
        
    }
}
