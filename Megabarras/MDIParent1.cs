using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Megabarras
{
    public partial class MDIParent1 : Form
    {
       // private int childFormNumber = 0;
      //  private int i = 0;


        public MDIParent1()
        {
            InitializeComponent();
            
        }

        private void ShowNewForm(object sender, EventArgs e)
        {
            Form2 childForm = new Form2();
            childForm.MdiParent = this;
            childForm.Text = "Generacion de Barra ";
            childForm.Show();
        }

        private void ExitToolsStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        
        private void CascadeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.Cascade);
        }

        private void TileVerticalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.TileVertical);
        }

        private void TileHorizontalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.TileHorizontal);
        }

        private void ArrangeIconsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.ArrangeIcons);
        }

        private void CloseAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (Form childForm in MdiChildren)
            {
                childForm.Close();
            }
        }

        private void MDIParent1_Load(object sender, EventArgs e)
        {
            DialogResult result;
            Form1 w_acceso = new Form1();
            w_acceso.MdiParent = this.ParentForm;
            result = w_acceso.ShowDialog();
            
            if (result== DialogResult.OK)
            {
                //this.Close();
                w_acceso.Close();
            }
        }

        private void nuevaBarraIndividualToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 childForm = new Form3();
            childForm.MdiParent = this;
            childForm.Text = "Generar Barra individual";
            childForm.Show();
           
        }

        private void barrasIndividualActualizaUnoEEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form4 childForm = new Form4();
            childForm.MdiParent = this;
            childForm.Text = "Generar Barra individual Y Actualizar UnoEE";
            childForm.Show();
        }
    }
}
