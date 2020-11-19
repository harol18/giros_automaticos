using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Usuarios_planta.Formularios;
using FontAwesome.Sharp; // libreria para utilizar los iconbutton y los iconos


namespace Usuarios_planta
{
    public partial class desembolso : Form
    {
        
        private IconButton currentBtn;
        private Panel leftBorderBtn;

        public desembolso()
        {
            InitializeComponent();
            leftBorderBtn = new Panel();
            leftBorderBtn.Size = new Size(7, 60);
            panelSideMenu.Controls.Add(leftBorderBtn);
            

            this.Text = string.Empty;
            this.ControlBox = false; //quitar caja de control
            this.DoubleBuffered = true; // activar el buffer para reducri el parpadeo en los graficos del formulario
            this.MaximizedBounds = Screen.FromHandle(this.Handle).WorkingArea; // establecer limites para dejar el formulario como el area del escritorio
        }

        private struct RGBColors
        {
            public static Color color1 = Color.FromArgb(251, 187, 33);
            public static Color color2 = Color.FromArgb(52, 179, 29);
            public static Color color3 = Color.FromArgb(253, 138, 114);
            public static Color color4 = Color.FromArgb(37, 146, 110);
            public static Color color5 = Color.FromArgb(56, 171, 179);
            

        }

        private void ActivateButton(object senderBtn, Color color)
        {
            if (senderBtn != null)
            {
                DisableButton();
                //Button
                currentBtn = (IconButton)senderBtn;
                currentBtn.BackColor = Color.FromArgb(37, 36, 81);
                currentBtn.ForeColor = Color.FromArgb(215, 219, 222);
                currentBtn.TextAlign = ContentAlignment.MiddleCenter;
                currentBtn.IconColor = color;
                currentBtn.TextImageRelation = TextImageRelation.TextBeforeImage;
                currentBtn.ImageAlign = ContentAlignment.MiddleRight;
                //Left border button
                leftBorderBtn.BackColor = color;
                leftBorderBtn.Location = new Point(0, currentBtn.Location.Y);
                leftBorderBtn.Visible = true;
                leftBorderBtn.BringToFront();
            }
        }

        private void DisableButton()
        {
            if (currentBtn != null)
            {
                currentBtn.BackColor = Color.FromArgb(0, 66, 84);
                currentBtn.ForeColor = Color.Gainsboro;
                currentBtn.TextAlign = ContentAlignment.MiddleLeft;
                currentBtn.IconColor = Color.Gainsboro;
                currentBtn.TextImageRelation = TextImageRelation.ImageBeforeText;
                currentBtn.ImageAlign = ContentAlignment.MiddleLeft;
            }
        }

        private void BtnSalir_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

       
        private void BtnCheques_MouseHover(object sender, EventArgs e)
        {
            ActivateButton(sender, RGBColors.color1);
            labelInicio.Text = BtnAsignacion.Text;
            icon2.IconChar = BtnAsignacion.IconChar;
            icon2.IconColor = BtnAsignacion.IconColor;
            
        }
        private void BtnSalir_MouseHover(object sender, EventArgs e)
        {
            ActivateButton(sender, RGBColors.color5);
        }

        private void AbrirFormHijo(object formhijo)
        {
            if (this.panelContenedor.Controls.Count > 0)
                this.panelContenedor.Controls.RemoveAt(0);
            Form fh = formhijo as Form;
            fh.TopLevel = false; // decimos que es un formulario secundario
            fh.Dock = DockStyle.Fill;
            this.panelContenedor.Controls.Add(fh);
            this.panelContenedor.Tag = fh;
            fh.Show();
        }     

        private void Btngiros_Click(object sender, EventArgs e)
        {            
            Form formulario = new FormGiros();
            formulario.Show();
        }

        private void Btngiros_MouseHover(object sender, EventArgs e)
        {
            ActivateButton(sender, RGBColors.color3);
            labelInicio.Text = Btngiros.Text;
            icon2.IconChar = Btngiros.IconChar;
            icon2.IconColor = Btngiros.IconColor;
        }

        private void Btnlogin_Click(object sender, EventArgs e)
        {
            AbrirFormHijo(new Informes());
        }

        private void Btncheques_MouseHover_1(object sender, EventArgs e)
        {
            ActivateButton(sender, RGBColors.color4);
            labelInicio.Text = Btncheques.Text;
            icon2.IconChar = Btncheques.IconChar;
            icon2.IconColor = Btncheques.IconColor;
        }
    }
}
