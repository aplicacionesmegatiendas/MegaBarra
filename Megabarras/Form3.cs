using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using BarcodeLib;
using System.Drawing.Imaging;
using System.Net.Mail;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Net;

namespace Megabarras
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();

            //{
            string FileName = folderBrowserDialog1.SelectedPath;
            tbruta.Text = FileName;
            tab2btgenerar.Enabled = true;
        }

        private void tab2btgenerar_Click(object sender, EventArgs e)
        {
            generobarra(tbbarra.Text, tbruta.Text + "\\" + tbcedula.Text + ".JPG");
            tab2btenviar.Enabled = true;

        }

        /* private void generobarra2(string codigo, string ruta)
         {
             BarcodeLib.Barcode b = new BarcodeLib.Barcode();
             Image img = b.Encode(BarcodeLib.TYPE.CODE128, codigo, Color.Black, Color.White, 200, 50);
             img.Save(ruta, ImageFormat.Jpeg);
             DialogResult resultado = MessageBox.Show("Codigo de barras generado, desea enviarlo al correo electronico", "Generando Barras", MessageBoxButtons.YesNo);
             if (resultado == DialogResult.Yes)
             {
                 string archivo = tbruta2.Text + "\\" + tbcedula2.Text + ".JPG";
                 if (tbcorreo2.Text == "")
                 {
                     MessageBox.Show("La direccion de correo no puede estar vacia, ingrese al menos un destinatario", "Envio de Correos");
                 }
                 else
                 {
                     EnviarCorreo(tbcorreo2.Text, "Codigo de barra del usuario" + " " + tbcedula2.Text,
                     "Este un mensaje del generador de automatico de codigos de barras, si usted no es el destinatario por favor eliminelo y haga caso omiso de éste ", archivo);
                     MessageBox.Show("Correo enviado correctamente", "Envio de Correos");
                 }
             }

         }*/

        private void generobarra(string codigo, string ruta)
        {
            try
            {
                BarcodeLib.Barcode b = new BarcodeLib.Barcode();
                Image img = b.Encode(BarcodeLib.TYPE.CODE128, codigo, Color.Black, Color.White, 200, 50);
                img.Save(ruta, ImageFormat.Jpeg);
                DialogResult resultado = MessageBox.Show("Codigo de barras generado, desea enviarlo al correo electronico", "Generando Barras", MessageBoxButtons.YesNo);
                if (resultado == DialogResult.Yes)
                {
                    string archivo = tbruta.Text + "\\" + tbcedula.Text + ".JPG";
                    if (tbcorreo.Text == "")
                    {
                        MessageBox.Show("La direccion de correo no puede estar vacia, ingrese al menos un destinatario", "Envio de Correos");
                    }
                    else
                    {
                        EnviarCorreo(tbcorreo.Text, "Codigo de barra del usuario" + " " + tbcedula.Text,
                        "Este un mensaje del generador de automatico de codigos de barras, si usted no es el destinatario por favor eliminelo y haga caso omiso de éste ", archivo);
                        MessageBox.Show("Correo enviado correctamente", "Envio de Correos");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void validoc()
        {
            string cedula = tbcedula.Text;
            string ls_query2 = "select  t9766_pdv_enrolamiento.f9766_id_cod_barras as barra " +
                "from t9766_pdv_enrolamiento inner join t200_mm_terceros on f9766_rowid_tercero = f200_rowid " +
                "where f200_id = @cedula1 and f200_id_cia = '1'";
            SqlConnection con2 = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["unoee"].ConnectionString);
            con2.Open();
            SqlCommand cmd2 = new SqlCommand(ls_query2, con2);
            cmd2.CommandType = CommandType.Text;
            //Realizar update en base de datos principal
            Application.DoEvents();
            cmd2.Parameters.AddWithValue("@cedula1", cedula);
            SqlDataReader reader = cmd2.ExecuteReader();
            if (reader.Read())
            {
                tbbarra.Text = reader["barra"] as string;
                reader.Close();
                cmd2.Dispose();
                btruta.Enabled = true;

            }
            else
            {
                MessageBox.Show("Error el tercero no esta enrolado en la base de datos UNOEE", "Consultando Enrolamiento De Tercero en UNOEE");
                tbcedula.Focus();
                tbbarra.Text = "";
                tab2btgenerar.Enabled = false;
                tab2btenviar.Enabled = false;
            }


        }

        private void Form3_Load(object sender, EventArgs e)
        {
            this.Top = (this.Parent.ClientSize.Height - this.Height) / 2;
            this.Left = (this.Parent.ClientSize.Width - this.Width) / 2;
        }

        private void tab2btenviar_Click(object sender, EventArgs e)
        {
            try
            {
                string archivo = tbruta.Text + "\\" + tbcedula.Text + ".JPG";
                if (tbcorreo.Text == "")
                {
                    MessageBox.Show("La direccion de correo no puede estar vacia, ingrese al menos un destinatario", "Envio de Correos");
                }
                else
                {
                    EnviarCorreo(tbcorreo.Text, "Codigo de barra del usuario" + " " + tbcedula.Text,
                    "Este un mensaje del generador de automatico de codigos de barras, si usted no es el destinatario por favor eliminelo y haga caso omiso de éste ", archivo);
                    MessageBox.Show("Correo enviado correctamente", "Envio de Correos");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void EnviarCorreo(string To, string Subject, string Body, string adjunto)//Funcion para Envio de correos
        {
            System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();
            mail.From = new MailAddress("contacto@megatiendas.com.co"/*"cambiobarras@megatiendas.co"*/);
            mail.To.Add(To);
            mail.Subject = Subject;
            mail.Body = Body;
            mail.Attachments.Add(new Attachment(adjunto));
            SmtpClient smtp = new SmtpClient();
            smtp.Host = "smtp.gmail.com"; //"mail.megatiendas.co";
            smtp.Port = 587;//2025; //465; //25
            smtp.EnableSsl = true;//false;
            smtp.UseDefaultCredentials = false;
            smtp.Credentials = new System.Net.NetworkCredential("contacto@megatiendas.com.co","Invercomer1."/*"cambiobarras@megatiendas.co", "RS4-R3@CT-1nv3rc0m3r"*/);

            try
            {
                ServicePointManager.ServerCertificateValidationCallback +=
                       delegate (
                       Object sender1,
                       X509Certificate certificate,
                       X509Chain chain,
                       SslPolicyErrors sslPolicyErrors)
                       {
                           return true;
                       };

                smtp.Send(mail);
            }
            catch (SmtpException ex)
            {

                if (ex.InnerException != null)
                {
                    throw new Exception("No se ha podido enviar el email: " + ex.InnerException.Message);
                }
                else
                {
                    throw new Exception("No se ha podido enviar el email: " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                if (ex.InnerException != null)
                {
                    MessageBox.Show(ex.InnerException.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            finally
            {
                smtp.Dispose();
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {


        }
        /* private void Validarcedula(string cedula)
         {
             cedula = tbcedula2.Text;
             string ls_query2 = "select  t9766_pdv_enrolamiento.f9766_id_cod_barras as barra " +
                 "from t9766_pdv_enrolamiento inner join t200_mm_terceros on f9766_rowid_tercero = f200_rowid " +
                 "where f200_id = @cedula1 and f200_id_cia = '1'";
             SqlConnection con1 = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["unoee"].ConnectionString);
             con1.Open();
             SqlCommand cmd2 = new SqlCommand(ls_query2, con1);
             cmd2.CommandType = CommandType.Text;
             //Realizar update en base de datos principal
             Application.DoEvents();
             cmd2.Parameters.AddWithValue("@cedula1", cedula);
             SqlDataReader reader = cmd2.ExecuteReader();
             if (reader.Read())
             {
                 con1.Close();
                 bt2ruta.Enabled = true;
                 btgenerar2.Enabled = true;

             }
             else
             {
                 MessageBox.Show("Error el tercero no esta enrolado en la base de datos UNOEE", "Consultando Enrolamiento De Tercero en UNOEE");
                 tbcedula2.Focus();
                 tbbarra2.Text = "";
                 btgenerar2.Enabled = false;
                 btactualizar.Enabled = false;
                 bt_enviar.Enabled = false;
             }
         }*/

        private void button1_Click_2(object sender, EventArgs e)
        {

        }

        private void tbcedula2_Leave(object sender, EventArgs e)
        {
            /* if (tbcedula2.Text == "")
             {

             }
             else
             {
                 Validarcedula(tbcedula2.Text);

             }*/


        }

        private void tbcedula2_TextChanged(object sender, EventArgs e)
        {

        }

        private void bt2ruta_Click(object sender, EventArgs e)
        {
            /*if (tbcedula2.Text == "")
            {

            }
            else
            {
                folderBrowserDialog1.ShowDialog();
                string FileName = folderBrowserDialog1.SelectedPath;
                tbruta2.Text = FileName;
                btgenerar2.Enabled = true;

            }*/
        }

        private void btgenerar2_Click(object sender, EventArgs e)
        {
            /*
            if (tbruta2.Text == "")
            {
                MessageBox.Show("Debe seleccionar una ruta destino", "Generando Barra");
            }
            else { 
                long lo_barras, lo_barras2;
                string barra;
                Random rnd1 = new Random();
                Random rnd2 = new Random();
                lo_barras = rnd1.Next(111111111, 999999999);
                lo_barras2 = rnd1.Next(111, 999);
                barra = string.Concat(lo_barras.ToString(), lo_barras2.ToString());
                tbbarra2.Text = barra.Trim();
                generobarra2(tbbarra2.Text, tbruta2.Text + "\\" + tbcedula2.Text + ".JPG");
                btactualizar.Enabled = true;
                bt_enviar.Enabled = true;
                
            }
            */
        }

        private void button5_Click(object sender, EventArgs e)
        {
            /*int li_filas = 0;
            string identificacion, barra;
             identificacion  = tbcedula2.Text;
            barra = tbbarra2.Text;
            string ls_query = "update t9766_pdv_enrolamiento set f9766_id_cod_barras = @barra " +
                "from t9766_pdv_enrolamiento inner join t200_mm_terceros on f9766_rowid_tercero = f200_rowid  " +
                "where f200_id = @cedula and f200_id_cia = '1'";
            SqlConnection con = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["unoee"].ConnectionString);
            con.Open();
            SqlCommand cmd = new SqlCommand(ls_query, con);
            cmd.CommandType = CommandType.Text;
                           //Realizar update en base de datos principal
            Application.DoEvents();
               
            cmd.Parameters.AddWithValue("@cedula", identificacion);
            cmd.Parameters.AddWithValue("@barra", barra);
            li_filas = cmd.ExecuteNonQuery();
                    if (li_filas > 0)
                    {
                        MessageBox.Show("Codigo De barras Actualizado correctamente","Actualizando Unoee");
                        cmd.Parameters.Clear();
                    }      */
        }

        private void tbcedula_Leave(object sender, EventArgs e)
        {
            if (tbcedula.Text == "")
            {

            }
            else
            {
                validoc();
            }
        }

        private void tbcedula_TextChanged(object sender, EventArgs e)
        {

        }

        private void bt_enviar_Click(object sender, EventArgs e)
        {
            /*string archivo = tbruta2.Text + "\\" + tbcedula2.Text + ".JPG";
            if (tbcorreo2.Text == "")
            {
                MessageBox.Show("La direccion de correo no puede estar vacia, ingrese al menos un destinatario", "Envio de Correos");
            }
            else
            {
                EnviarCorreo(tbcorreo2.Text, "Codigo de barra del usuario" + " " + tbcedula2.Text,
                "Este un mensaje del generador de automatico de codigos de barras, si usted no es el destinatario por favor eliminelo y haga caso omiso de éste ", archivo);
                MessageBox.Show("Correo enviado correctamente", "Envio de Correos");
            }*/
        }
    }
}
