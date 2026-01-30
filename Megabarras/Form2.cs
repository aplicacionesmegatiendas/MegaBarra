using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Imaging;
using System.Data.OleDb;
using System.Drawing.Text;
using System.Net.Mail;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Net;
using System.Data.SqlClient;
using BarcodeLib;


namespace Megabarras
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)//Genera codigo 12 digitos aleatoreamente y los convierte en codigos de barra
        {
            int li_tam;
            long lo_barras, lo_barras2;
            string dato_barra, ls_barra; ;
            li_tam = dataGridView1.RowCount;
            Random rnd1 = new Random();
            Random rnd2 = new Random();
            for (int i = 0; i < li_tam; i++)
            {

                lo_barras = rnd1.Next(11111111/*1*/, 99999999/*9*/);
                lo_barras2 = rnd1.Next(111, 999);
                ls_barra = string.Concat(lo_barras.ToString(), lo_barras2.ToString());
                dataGridView1.Rows[i].Cells[1].Value = ls_barra;
                //dataGridView1.Columns[5].Visible = false;
                dato_barra = dataGridView1.Rows[i].Cells[0].Value.ToString();
                string nombre = dataGridView1.Rows[i].Cells[2].Value.ToString();
                if (dataGridView1.Rows[i].DefaultCellStyle.ForeColor == Color.Red)
                {

                }
                else
                {
                    generobarra(dato_barra, textBox1.Text + "\\" + nombre + ".JPG");//genero barra y envio a archivo las imagenes
                }
            }
            bt_correos.Enabled = true;
            bt_actualizar.Enabled = true;

        }

        private void generobarra(string codigo, string ruta)
        {
            BarcodeLib.Barcode b = new BarcodeLib.Barcode();
            Image img = b.Encode(BarcodeLib.TYPE.CODE128, codigo, Color.Black, Color.White, 200, 50);
            img.Save(ruta, ImageFormat.Jpeg);
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
            smtp.Host = "smtp.gmail.com";//"mail.megatiendas.co";
            smtp.Port = 587;//2025; //465; //25
            smtp.EnableSsl = true;//false;
            smtp.UseDefaultCredentials = false;
            smtp.Credentials = new System.Net.NetworkCredential("contacto@megatiendas.com.co","Invercomer1."/*"cambiobarras@megatiendas.co", "RS4-R3@CT-1nv3rc0m3r"*/);
            smtp.EnableSsl = true;

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
                if (ex.InnerException!=null)
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
                    throw new Exception("No se ha podido enviar el email: " + ex.InnerException.Message);
                }
                else
                {
                    throw new Exception("No se ha podido enviar el email: " + ex.Message);
                }
            }
            finally
            {
                smtp.Dispose();
            }

        }


        private void LLenarGrid(string archivo, string hoja)//Funcion para importar datos desde hoja1 de excel
        {
            //declaramos las variables         
            OleDbConnection conexion = null;
            DataSet dataSet = null;
            OleDbDataAdapter dataAdapter = null;
            string consultaHojaExcel = "Select * from [" + hoja + "$]";

            //esta cadena es para archivos excel 2007 y 2010
            string cadenaConexionArchivoExcel = "provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + archivo + "';Extended Properties=Excel 12.0;";

            //para archivos de 97-2003 usar la siguiente cadena
            //string cadenaConexionArchivoExcel = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + archivo + "';Extended Properties=Excel 8.0;";

            //Validamos que el usuario ingrese el nombre de la hoja del archivo de excel a leer
            if (string.IsNullOrEmpty(hoja))
            {
                MessageBox.Show("No hay una hoja para leer");
            }
            else
            {
                try
                {
                    //Si el usuario escribio el nombre de la hoja se procedera con la busqueda
                    conexion = new OleDbConnection(cadenaConexionArchivoExcel);//creamos la conexion con la hoja de excel
                    conexion.Open(); //abrimos la conexion
                    dataAdapter = new OleDbDataAdapter(consultaHojaExcel, conexion); //traemos los datos de la hoja y las guardamos en un dataSdapter
                    dataSet = new DataSet(); // creamos la instancia del objeto DataSet
                    dataAdapter.Fill(dataSet, hoja);//llenamos el dataset
                    dataGridView1.DataSource = dataSet.Tables[0]; //le asignamos al DataGridView el contenido del dataSet
                    conexion.Close();//cerramos la conexion
                    dataGridView1.AllowUserToAddRows = false;       //eliminamos la ultima fila del datagridview que se autoagrega
                }
                catch (Exception ex)
                {
                    //en caso de haber una excepcion que nos mande un mensaje de error
                    MessageBox.Show("Error, Verificar el archivo o el nombre de la hoja", ex.Message);
                }
            }
        }


        private void button2_Click(object sender, EventArgs e)//Llamar folder dialogo para establecer ruta destino de imagen
        {
            folderBrowserDialog1.ShowDialog();

            string FileName = folderBrowserDialog1.SelectedPath;
            textBox1.Text = FileName;
            bt_generar.Enabled = true;
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            this.Top = (this.Parent.ClientSize.Height - this.Height) / 2;
            this.Left = (this.Parent.ClientSize.Width - this.Width) / 2;
        }



        private void button3_Click(object sender, EventArgs e)// Anre el cuadro de dialogo para importar fichero de excel
        {
            //creamos un objeto OpenDialog que es un cuadro de dialogo para buscar archivos
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Archivos de Excel (*.xls;*.xlsx)|*.xls;*.xlsx"; //le indicamos el tipo de filtro en este caso que busque
                                                                             //solo los archivos excel

            dialog.Title = "Seleccione el archivo de Excel";//le damos un titulo a la ventana

            dialog.FileName = string.Empty;//inicializamos con vacio el nombre del archivo

            //si al seleccionar el archivo damos Ok
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //el nombre del archivo sera asignado al textbox
                txtArchivo.Text = dialog.FileName;
                //hoja = "HOJA1"; //la variable hoja tendra el valor del textbox donde colocamos el nombre de la hoja
                LLenarGrid(txtArchivo.Text, "hoja1"); //se manda a llamar al metodo
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill; //se ajustan las
                bt_validar.Enabled = true;
                bt_generar.Enabled = false;
                bt_exportar.Enabled = false;
                bt_destino.Enabled = false;
                bt_correos.Enabled = false;
                bt_actualizar.Enabled = false;

                //columnas al ancho del DataGridview para que no quede espacio en blanco (opcional)
            }
        }

        private void button4_Click(object sender, EventArgs e)//Envia el correo electronico a cada destino en la hoja de excel
        {
            try
            {
                int li_tam = dataGridView1.RowCount;
                progressBar1.Maximum = li_tam - 1;
                label1.Visible = true;
                progressBar1.Visible = true;
                for (int i = 0; i < li_tam; i++)
                {
                    Application.DoEvents();
                    string nombre = dataGridView1.Rows[i].Cells[2].Value.ToString();
                    string archivo = textBox1.Text + "\\" + nombre + ".JPG";
                    string correo = dataGridView1.Rows[i].Cells[6].Value.ToString();
                    EnviarCorreo(correo, "Codigo de barra del usuario" + " " + nombre, "Este un mensaje del generador de automatico de codigos de barras, si usted no es el destinatario por favor eliminelo y haga caso omiso de éste ", archivo);
                    label1.Text = "Enviando " + " " + i.ToString() + " " + "de" + progressBar1.Maximum.ToString();
                    progressBar1.Value = i;

                }

                MessageBox.Show("Se enviaron" + " " + progressBar1.Maximum.ToString() + "correctamente", "Envio via correo");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)//Convierte la columna 5 a campo password
        {
            if (e.ColumnIndex == 5 && e.Value != null)
            {
                e.Value = new String('*', e.Value.ToString().Length);
            }
        }

        private void button5_Click(object sender, EventArgs e)// Actualiza base de datos Unoee
        {
            int li_filas = 0, li_tam, li_acum = 0, li_acum4 = 0;
            li_tam = dataGridView1.RowCount;
            progressBar1.Maximum = li_tam - 1;
            label1.Visible = true;
            progressBar1.Visible = true;
            string ls_query = @"if 
                                (
	                                select 
		                                COUNT(*) 
                                    from 
		                                t9766_pdv_enrolamiento 
		                                inner join t200_mm_terceros on f9766_rowid_tercero = f200_rowid 
	                                where 
		                                f200_id = @cedula
		                                and f200_id_cia = '1'
                                ) = 1 
                                begin 
	                                update 
		                                t9766_pdv_enrolamiento 
	                                set 
		                                f9766_id_cod_barras = @barra 
                                    from 
		                                t9766_pdv_enrolamiento 
		                                inner join t200_mm_terceros on f9766_rowid_tercero = f200_rowid  
                                    where 
		                                f200_id = @cedula
		                                and f200_id_cia = '1'
                                end";
            SqlConnection con = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["unoee"].ConnectionString);
            con.Open();
            SqlCommand cmd = new SqlCommand(ls_query, con);
            cmd.CommandType = CommandType.Text;
            for (int i = 0; i < li_tam; i++)
            {
                li_acum4++;
                if (dataGridView1.Rows[i].DefaultCellStyle.ForeColor == Color.Red)
                {

                }
                else
                {                    //Realizar update en base de datos principal
                    Application.DoEvents();
                    string identificacion = dataGridView1.Rows[i].Cells[0].Value.ToString();
                    string barra = dataGridView1.Rows[i].Cells[1].Value.ToString();
                    cmd.Parameters.AddWithValue("@cedula", identificacion);
                    cmd.Parameters.AddWithValue("@barra", barra);
                    li_filas = cmd.ExecuteNonQuery();
                    if (li_filas > 0)
                    {
                        li_acum++;
                    }

                    cmd.Parameters.Clear();
                    label1.Text = "Actualizando barra" + " " + li_acum.ToString() + " " + "de" + li_acum4.ToString();
                    progressBar1.Value = i;

                }
            }
            MessageBox.Show("Se actualizaron correctamente" + " " + li_acum.ToString() + " " + "codigos de barra", "Actualizando UNOEE");
            con.Close();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {

            var grid = sender as DataGridView;
            var rowIdx = (e.RowIndex + 1).ToString();

            var centerFormat = new StringFormat()
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };

            var headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height);
            e.Graphics.DrawString(rowIdx, this.Font, SystemBrushes.ControlText, headerBounds, centerFormat);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            int li_filas2 = 0, li_tam2, li_acum2 = 0, i = 0, li_acum3 = 0;
            li_tam2 = dataGridView1.RowCount;
            progressBar1.Maximum = li_tam2 - 1;
            label1.Visible = true;
            progressBar1.Visible = true;
            string ls_query2 = "declare @id varchar(40); set @id = @cedula1  select COUNT(*) from t9766_pdv_enrolamiento inner join t200_mm_terceros on f9766_rowid_tercero = f200_rowid where f200_id = @id and f200_id_cia = '1'";
            SqlConnection con2 = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["unoee"].ConnectionString);
            con2.Open();
            SqlCommand cmd2 = new SqlCommand(ls_query2, con2);
            cmd2.CommandType = CommandType.Text;
            for (i = 0; i < li_tam2; i++)
            {
                li_acum3++;                //Realizar update en base de datos principal
                Application.DoEvents();
                string cedula = dataGridView1.Rows[i].Cells[0].Value.ToString();
                cmd2.Parameters.AddWithValue("@cedula1", cedula);
                li_filas2 = Convert.ToInt32(cmd2.ExecuteScalar());
                if (li_filas2 > 0)
                {
                    li_acum2++;
                }
                else
                {
                    dataGridView1.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
                }

                cmd2.Parameters.Clear();
                label1.Text = "Se validaron " + " " + li_acum2.ToString() + " " + "identificaciones correctas de" + " " + li_acum3;
                progressBar1.Value = i;


            }
            if (li_acum2 == i)
            {
                bt_destino.Enabled = true;
                progressBar1.Visible = false;
                progressBar1.Value = 0;
                label1.Text = "";
            }
            else
            {
                MessageBox.Show("Algunos de las identificaciones no tienen enrrolamiento, edite el archivo o cree el enrolamiento en el maestro de usuarios para poder generar la barra", "Validando listado", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            con2.Close();
        }
    }
}
