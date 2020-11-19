using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.Data.SqlClient;
using System.Configuration;


namespace Usuarios_planta.Formularios
{
    public partial class Informes : Form
    {
        #region DeclaracionVariables
        private System.Windows.Forms.Timer timer;
        #endregion

        public Informes()
        {
            InitializeComponent();
        }

        private void InicioServicio(object sender, EventArgs e)
        {
            try
            {
                timer = new System.Windows.Forms.Timer();
                timer.Interval = Convert.ToInt32(ConfigurationManager.AppSettings["IntervaloEjecucion"]);
                timer.Enabled = true;
                this.timer.Tick += new EventHandler(EventoTemporizador);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void EventoTemporizador(object sender, EventArgs e)
        {
            try
            {
                //Declaracion de variable para conectar la base de datos
                string cadenaConexion = ConfigurationManager.ConnectionStrings["ConexionDB"].ToString();
                MySqlConnection conexion = new MySqlConnection(cadenaConexion);
                MySqlCommand comando = new MySqlCommand("Windows_Service", conexion); //se pasa el nombre del procedimiento almacena y la conexion
                comando.CommandType = CommandType.StoredProcedure;
                conexion.Open();
                comando.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void DetenerServicio(object sender, EventArgs e)
        {
            timer.Enabled = false;
            timer.Stop();
        }
    }
}
