using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Spreadsheet;
using MySql.Data.MySqlClient;
using SpreadsheetLight;
using Outlook = Microsoft.Office.Interop.Outlook;




namespace Usuarios_planta.Formularios
{
    public partial class FormGiros : Form
    {
        MySqlConnection con = new MySqlConnection("server=;Uid=;password=;database=dblibranza;port=3306;persistsecurityinfo=True;");
        Comandos cmds = new Comandos();
        
        public FormGiros()
        {
            InitializeComponent();
        }       

        DateTime hoy = DateTime.Today;       
           
        private void TxtCod_oficina_TextChanged(object sender, EventArgs e)
        {
            MySqlCommand comando = new MySqlCommand("SELECT * FROM tf_oficinas WHERE codigo_oficina = @codigo ", con);
            comando.Parameters.AddWithValue("@codigo", Txtcod_oficina.Text);
            con.Open();
            MySqlDataReader registro = comando.ExecuteReader();
            if (registro.Read())
            {
                datos_correo.correo_gerente= registro["correo_gerente"].ToString();
                datos_correo.correo_subgerente = registro["correo_subgerente"].ToString();                
            }
            con.Close();
        }

        private void TxtCoordinador_TextChanged(object sender, EventArgs e)
        {
            MySqlCommand comando = new MySqlCommand("SELECT * FROM tf_coordinador WHERE nombre_coordinador = @coordinador ", con);
            comando.Parameters.AddWithValue("@coordinador", Txtcoordinador.Text);
            con.Open();
            MySqlDataReader registro = comando.ExecuteReader();
            if (registro.Read())
            {
                datos_correo.correo_coordinador = registro["correo_coordinador"].ToString();
                datos_correo.correo_apoyo = registro["correo_apoyo"].ToString();
            }
            con.Close();
        }

        public static string GetHtml(DataGridView grid)
        {
            try
            {
                string messageBody = "<font>Señores: </font><br><br><br>Oficina  " + datos_correo.oficina+ "<br><br><br>Buen Día,<br><br>Por motivo del desembolso de la compra de cartera del cliente en referencia, se generó para su oficina el(los) Giro(s) de Cheque(s) de acuerdo con la información adjunta,para su respectiva impresión, custodia y contacto a cliente para su entrega. <br><br>" +
                    "La operatoria que se debe realizar:<br>Operatoria 2 / Operatoria activos / Prestamos / Formalización / Imprimir Cheques - Desembolso Crédito<br><br><br>";
                if (grid.RowCount == 0) return messageBody;
                string htmlTableStart = "<table style=\"border-collapse:collapse; text-align:center;\" >";
                string htmlTableEnd = "</table>";
                string htmlHeaderRowStart = "<tr style=\"background-color:#004254; color:#FFFFFF;\">";
                string htmlHeaderRowEnd = "</tr>";
                string htmlTrStart = "<tr style=\"color:#000000;\">";
                string htmlTrEnd = "</tr>";
                string htmlTdStart = "<td style=\" border-color:#000000; border-style:solid; border-width:thin; padding: 5px;\">";
                string htmlTdEnd = "</td>";
                string htmlTdparrafo = "<font><br><br><br>Por favor proceder con el giro de cheque de forma inmediata, dado que en caso contrario la partida quedará pendiente en la cuenta 259595201 de su centro de costos y la cual estará siendo monitoreada por CONTROL CONTABLE.<br><br>" +
                    "<br>Así mismo, una vez realizada la impresión del cheque se solicita realizar el endoso de cada una de las obligaciones correspondientes según la información suministrada; igualmente de requerirse esta información se podrá validar en Bonita.<br><br>" +
                    "SI PRESENTA ALGÚN INCONVENIENTE EN LA IMPRESIÓN,POR FAVOR DEVOLVER EL CORREO CON COPIA A TODOS LOS BUZONES ADJUNTANDO PANTALLAS PASO A PASO DE COMO SE ESTA INGRESANDO TODA LA INFORMACIÓN. Vale aclarar que se debe ingresar el valor informado en el correo y NO el valor de la partida Contable<br><br><br>" +
                    "BBVA - INDRA.<br>Centro de formalización.<br>Calle 75a # 27a - 28.<br>cheques.libranza@bbva.com.co</font>";
                messageBody += htmlTableStart;
                messageBody += htmlHeaderRowStart;
                messageBody += htmlTdStart + "Radicado" + htmlTdEnd;
                messageBody += htmlTdStart + "Codigo" + htmlTdEnd;
                messageBody += htmlTdStart + "Fecha" + htmlTdEnd;
                messageBody += htmlTdStart + "Oficina" + htmlTdEnd;
                messageBody += htmlTdStart + "Ciudad" + htmlTdEnd;
                messageBody += htmlTdStart + "Cedula" + htmlTdEnd;
                messageBody += htmlTdStart + "Nombre" + htmlTdEnd;
                messageBody += htmlTdStart + "Nit" + htmlTdEnd;
                messageBody += htmlTdStart + "Entidad" + htmlTdEnd;
                messageBody += htmlTdStart + "Valor" + htmlTdEnd;
                messageBody += htmlTdStart + "Obligacion" + htmlTdEnd;
                messageBody += htmlTdStart + "Scoring" + htmlTdEnd;
                messageBody += htmlTdStart + "Gestor" + htmlTdEnd;
                messageBody += htmlTdStart + "Coordinador" + htmlTdEnd;
                messageBody += htmlTdStart + "Cuenta" + htmlTdEnd;
                messageBody += htmlTdStart + "Ref" + htmlTdEnd;
                messageBody += htmlHeaderRowEnd;
                
                //Loop all the rows from grid vew and added to html td  
                for (int i = 0; i <= grid.RowCount - 1; i++)
                {
                    messageBody = messageBody + htmlTrStart;
                    messageBody = messageBody + htmlTdStart + grid.Rows[i].Cells[0].Value; //Radicado
                    messageBody = messageBody + htmlTdStart + grid.Rows[i].Cells[1].Value; //Codigo
                    messageBody = messageBody + htmlTdStart + grid.Rows[i].Cells[2].Value; //Fecha  
                    messageBody = messageBody + htmlTdStart + grid.Rows[i].Cells[3].Value; //Oficina
                    messageBody = messageBody + htmlTdStart + grid.Rows[i].Cells[4].Value; //Ciudad 
                    messageBody = messageBody + htmlTdStart + grid.Rows[i].Cells[5].Value; //Cedula 
                    messageBody = messageBody + htmlTdStart + grid.Rows[i].Cells[6].Value; //Nombre 
                    messageBody = messageBody + htmlTdStart + grid.Rows[i].Cells[7].Value; //Nit
                    messageBody = messageBody + htmlTdStart + grid.Rows[i].Cells[8].Value; //Entidad
                    messageBody = messageBody + htmlTdStart + grid.Rows[i].Cells[9].Value; //Valor
                    messageBody = messageBody + htmlTdStart + grid.Rows[i].Cells[10].Value; //Obligacion 
                    messageBody = messageBody + htmlTdStart + grid.Rows[i].Cells[11].Value; //Scoring 
                    messageBody = messageBody + htmlTdStart + grid.Rows[i].Cells[12].Value; //Gestor
                    messageBody = messageBody + htmlTdStart + grid.Rows[i].Cells[12].Value; //Coordinador 
                    messageBody = messageBody + htmlTdStart + grid.Rows[i].Cells[14].Value; //Cuenta
                    messageBody = messageBody + htmlTdStart + grid.Rows[i].Cells[15].Value; //Ref
                    messageBody = messageBody + htmlTrEnd;
                    

                }
                messageBody = messageBody + htmlTableEnd;
                messageBody = messageBody + htmlTdparrafo;
                return messageBody; // devuelve la tabla HTML como cadena de esta función  
            }
            catch (Exception)
            {
                return null;
            }
        }

        private void BtnBuscar_Click(object sender, EventArgs e)
        {
            cmds.Buscar_giro(TxtRadicado, Txtcedula, Txtnombre, Txtcuenta, Txtscoring, TxtCedula_Gestor, Txtnom_gestor,
           Txtcoordinador, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtobligacion1, TxtNom_entidad1, TxtNit1, TxtValor1,
           Txtobligacion2, TxtNom_entidad2, TxtNit2, TxtValor2, Txtobligacion3, TxtNom_entidad3, TxtNit3, TxtValor3,
           Txtobligacion4, TxtNom_entidad4, TxtNit4, TxtValor4, Txtobligacion5, TxtNom_entidad5, TxtNit5, TxtValor5,
           Txtobligacion6, TxtNom_entidad6, TxtNit6, TxtValor6, Txtobligacion7, TxtNom_entidad7, TxtNit7, TxtValor7,
           Txtobligacion8, TxtNom_entidad8, TxtNit8, TxtValor8);
        }

        private void Txtnombre_TextChanged(object sender, EventArgs e)
        {
            TxtAsunto.Text = "GIRO CHEQUE CPK " + Txtnombre.Text + " CC " + Txtcedula.Text;
        }

        private void btnAñadir_Carteras_Click(object sender, EventArgs e)
        {
           
        }

        private void TxtCedula_Gestor_TextChanged(object sender, EventArgs e)
        {
            MySqlCommand comando = new MySqlCommand("SELECT * FROM gestores WHERE Cedula_Gestor = @Cedula_Gestor ", con);
            comando.Parameters.AddWithValue("@Cedula_Gestor", TxtCedula_Gestor.Text);
            con.Open();
            MySqlDataReader registro = comando.ExecuteReader();
            if (registro.Read())
            {
                datos_correo.correo_gestor = registro["Correo_Gestor"].ToString();                
            }
            con.Close();
        }

        private void btnEnviar_Correo_Click(object sender, EventArgs e)
        {
            string correo_oficina = datos_correo.correo_gerente + " ; " + datos_correo.correo_subgerente + " ; ";
            string correo_F_comercial = datos_correo.correo_coordinador + " ; " + datos_correo.correo_apoyo + " ; " + datos_correo.correo_gestor + " ; ";
            string destinatarios = correo_oficina + correo_F_comercial;
            string correo_copia = datos_correo.copia_correo + TxtCopia_Correo.Text;
            string htmlString = GetHtml(dataGridView1);            

            try
            {
                Outlook.Application outlookApp = new Outlook.Application();
                Outlook._MailItem oMailItem = (Outlook._MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Inspector oInspector = oMailItem.GetInspector;


                oMailItem.Subject = TxtAsunto.Text;
                oMailItem.To = destinatarios;
                oMailItem.CC = correo_copia;
                oMailItem.HTMLBody = htmlString;                
                oMailItem.Attachments.Add(@"D:\Guia_Rapida.pdf");
                //oMailItem.BCC = "hsmartinez@indracompany.com";//Copia oculta
                oMailItem.Importance = Outlook.OlImportance.olImportanceNormal;//Asignar Importancia del correo
                oMailItem.Display(true);
                //oMailItem.Send();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            if (TxtNom_entidad1.Text != "" && TxtNom_entidad2.Text == "" && TxtNom_entidad3.Text == "" && TxtNom_entidad4.Text == "" && TxtNom_entidad5.Text == "" && TxtNom_entidad6.Text == "" && TxtNom_entidad7.Text == "" && TxtNom_entidad8.Text == "")
            {
                cmds.Insertar_cartera1(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                      Txtobligacion1, TxtNit1, TxtNom_entidad1, TxtValor1, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                MessageBox.Show("Información Registrada");
                btnNuevo.PerformClick(); // dar click automaticamente en el boton para limpiar
            }
            else if (TxtNom_entidad1.Text != "" && TxtNom_entidad2.Text != "" && TxtNom_entidad3.Text == "" && TxtNom_entidad4.Text == "" && TxtNom_entidad5.Text == "" && TxtNom_entidad6.Text == "" && TxtNom_entidad7.Text == "" && TxtNom_entidad8.Text == "")
            {
                cmds.Insertar_cartera1(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                      Txtobligacion1, TxtNit1, TxtNom_entidad1, TxtValor1, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera2(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion2, TxtNit2, TxtNom_entidad2, TxtValor2, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                MessageBox.Show("Información Registrada");
                btnNuevo.PerformClick(); // dar click automaticamente en el boton para limpiar
            }
            else if (TxtNom_entidad1.Text != "" && TxtNom_entidad2.Text != "" && TxtNom_entidad3.Text != "" && TxtNom_entidad4.Text == "" && TxtNom_entidad5.Text == "" && TxtNom_entidad6.Text == "" && TxtNom_entidad7.Text == "" && TxtNom_entidad8.Text == "")
            {
                cmds.Insertar_cartera1(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                      Txtobligacion1, TxtNit1, TxtNom_entidad1, TxtValor1, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera2(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion2, TxtNit2, TxtNom_entidad2, TxtValor2, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera3(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion3, TxtNit3, TxtNom_entidad3, TxtValor3, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                MessageBox.Show("Información Registrada");
                btnNuevo.PerformClick(); // dar click automaticamente en el boton para limpiar
            }
            else if (TxtNom_entidad1.Text != "" && TxtNom_entidad2.Text != "" && TxtNom_entidad3.Text != "" && TxtNom_entidad4.Text != "" && TxtNom_entidad5.Text == "" && TxtNom_entidad6.Text == "" && TxtNom_entidad7.Text == "" && TxtNom_entidad8.Text == "")
            {
                cmds.Insertar_cartera1(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                      Txtobligacion1, TxtNit1, TxtNom_entidad1, TxtValor1, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera2(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion2, TxtNit2, TxtNom_entidad2, TxtValor2, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera3(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion3, TxtNit3, TxtNom_entidad3, TxtValor3, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera4(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion4, TxtNit4, TxtNom_entidad4, TxtValor4, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                MessageBox.Show("Información Registrada");
                btnNuevo.PerformClick(); // dar click automaticamente en el boton para limpiar
            }
            else if (TxtNom_entidad1.Text != "" && TxtNom_entidad2.Text != "" && TxtNom_entidad3.Text != "" && TxtNom_entidad4.Text != "" && TxtNom_entidad5.Text != "" && TxtNom_entidad6.Text == "" && TxtNom_entidad7.Text == "" && TxtNom_entidad8.Text == "")
            {
                cmds.Insertar_cartera1(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                      Txtobligacion1, TxtNit1, TxtNom_entidad1, TxtValor1, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera2(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion2, TxtNit2, TxtNom_entidad2, TxtValor2, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera3(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion3, TxtNit3, TxtNom_entidad3, TxtValor3, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera4(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion4, TxtNit4, TxtNom_entidad4, TxtValor4, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera5(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion5, TxtNit5, TxtNom_entidad5, TxtValor5, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                MessageBox.Show("Información Registrada");
                btnNuevo.PerformClick(); // dar click automaticamente en el boton para limpiar
            }
            else if (TxtNom_entidad1.Text != "" && TxtNom_entidad2.Text != "" && TxtNom_entidad3.Text != "" && TxtNom_entidad4.Text != "" && TxtNom_entidad5.Text != "" && TxtNom_entidad6.Text != "" && TxtNom_entidad7.Text == "" && TxtNom_entidad8.Text == "")
            {
                cmds.Insertar_cartera1(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                      Txtobligacion1, TxtNit1, TxtNom_entidad1, TxtValor1, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera2(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion2, TxtNit2, TxtNom_entidad2, TxtValor2, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera3(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion3, TxtNit3, TxtNom_entidad3, TxtValor3, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera4(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion4, TxtNit4, TxtNom_entidad4, TxtValor4, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera5(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion5, TxtNit5, TxtNom_entidad5, TxtValor5, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera6(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion6, TxtNit6, TxtNom_entidad6, TxtValor6, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                MessageBox.Show("Información Registrada");
                btnNuevo.PerformClick(); // dar click automaticamente en el boton para limpiar
            }
            else if (TxtNom_entidad1.Text != "" && TxtNom_entidad2.Text != "" && TxtNom_entidad3.Text != "" && TxtNom_entidad4.Text != "" && TxtNom_entidad5.Text != "" && TxtNom_entidad6.Text != "" && TxtNom_entidad7.Text != "" && TxtNom_entidad8.Text == "")
            {
                cmds.Insertar_cartera1(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                    Txtobligacion1, TxtNit1, TxtNom_entidad1, TxtValor1, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera2(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion2, TxtNit2, TxtNom_entidad2, TxtValor2, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera3(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion3, TxtNit3, TxtNom_entidad3, TxtValor3, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera4(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion4, TxtNit4, TxtNom_entidad4, TxtValor4, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera5(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion5, TxtNit5, TxtNom_entidad5, TxtValor5, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera6(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion6, TxtNit6, TxtNom_entidad6, TxtValor6, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera7(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion7, TxtNit7, TxtNom_entidad7, TxtValor7, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                MessageBox.Show("Información Registrada");
                btnNuevo.PerformClick(); // dar click automaticamente en el boton para limpiar
            }
            else if (TxtNom_entidad1.Text != "" && TxtNom_entidad2.Text != "" && TxtNom_entidad3.Text != "" && TxtNom_entidad4.Text != "" && TxtNom_entidad5.Text != "" && TxtNom_entidad6.Text != "" && TxtNom_entidad7.Text != "" && TxtNom_entidad8.Text != "")
            {
                cmds.Insertar_cartera1(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                    Txtobligacion1, TxtNit1, TxtNom_entidad1, TxtValor1, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera2(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion2, TxtNit2, TxtNom_entidad2, TxtValor2, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera3(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion3, TxtNit3, TxtNom_entidad3, TxtValor3, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera4(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion4, TxtNit4, TxtNom_entidad4, TxtValor4, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera5(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion5, TxtNit5, TxtNom_entidad5, TxtValor5, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera6(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion6, TxtNit6, TxtNom_entidad6, TxtValor6, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera7(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion7, TxtNit7, TxtNom_entidad7, TxtValor7, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                cmds.Insertar_cartera8(TxtRadicado, Txtcedula, Txtnombre, Txtcod_oficina, Txtnom_oficina, Txtciudad, Txtscoring, Txtcuenta,
                                       Txtobligacion8, TxtNit8, TxtNom_entidad8, TxtValor8, TxtCedula_Gestor, Txtnom_gestor, Txtcoordinador);
                MessageBox.Show("Información Registrada");
                btnNuevo.PerformClick(); // dar click automaticamente en el boton para limpiar
            }
            else
            {
                MessageBox.Show("No hay carteras para almacenar", "Información", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }           
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            MessageBox.Show(datos_correo.correo_gerente + datos_correo.correo_subgerente);            
        }

        private void btnNuevo_Click(object sender, EventArgs e)
        {
            this.Close();
            Form formulario = new FormGiros();
            formulario.Show();
        }

        private void FormGiros_Load(object sender, EventArgs e)
        {
            datos_correo.copia_correo = "luis.zarate@bbva.com ; CUENTAS-PAGARCF@BBVA.COM.CO ; DESGLOSESCF@bbva.com.co ; brianduvan.garzon@bbva.com ; controldecambiosfabrica.co@bbva.com";
        }

        private void btnPunto_Control_Click(object sender, EventArgs e)
        {
            cmds.Base_punto(dtpFecha_Punto, dgvDatos_Punto);
            lbltotal.Text = dgvDatos_Punto.Rows.Count.ToString();
        }

        private void btnDescargar_Excel_Click(object sender, EventArgs e)
        {            
            SLDocument sl = new SLDocument();
            SLStyle style = new SLStyle();
            style.Font.Bold = true;
            style.Font.FontSize = 11;
            style.Font.FontName = "Calibri";
            style.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.Lavender, System.Drawing.Color.LightGray);
            style.Alignment.Horizontal = HorizontalAlignmentValues.Center;

            int i = 1;
            foreach (DataGridViewColumn columna in dgvDatos_Punto.Columns)
            {
                sl.SetCellValue(1, i, columna.HeaderText.ToString());
                sl.SetCellStyle(1, i, style);
                i++;
            }

            int j = 2;
            foreach (DataGridViewRow row in dgvDatos_Punto.Rows)
            {
                sl.SetCellValue(j, 1, row.Cells[0].Value.ToString());
                sl.SetCellValue(j, 2, row.Cells[1].Value.ToString());
                sl.SetCellValue(j, 3, row.Cells[2].Value.ToString());
                sl.SetCellValue(j, 4, row.Cells[3].Value.ToString());
                sl.SetCellValue(j, 5, row.Cells[4].Value.ToString());
                sl.SetCellValue(j, 6, row.Cells[5].Value.ToString());
                sl.SetCellValue(j, 7, row.Cells[6].Value.ToString());
                sl.SetCellValue(j, 8, row.Cells[7].Value.ToString());
                sl.SetCellValue(j, 9, row.Cells[8].Value.ToString());
                sl.SetCellValue(j, 10, row.Cells[9].Value.ToString());
                sl.SetCellValue(j, 11, row.Cells[10].Value.ToString());
                sl.SetCellValue(j, 12, row.Cells[11].Value.ToString());
                sl.SetCellValue(j, 13, row.Cells[12].Value.ToString());
                sl.SetCellValue(j, 14, row.Cells[13].Value.ToString());
                sl.SetCellValue(j, 15, row.Cells[14].Value.ToString());
                sl.SetCellValue(j, 16, row.Cells[15].Value.ToString());
                j++;
            }            
            sl.SaveAs(@"D:\punto_giros.xlsx");
            MessageBox.Show("Ok archivo creado", "Información", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
        }

        private void btnAgregar_Carteras_Click(object sender, EventArgs e)
        {
            if (TxtNom_entidad1.Text != "" && TxtNom_entidad2.Text == "" && TxtNom_entidad3.Text == "" && TxtNom_entidad4.Text == "" && TxtNom_entidad5.Text == "" && TxtNom_entidad6.Text == "" && TxtNom_entidad7.Text == "" && TxtNom_entidad8.Text == "")
            {
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit1.Text,
                                   TxtNom_entidad1.Text, TxtValor1.Text, Txtobligacion1.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
            }
            else if (TxtNom_entidad1.Text != "" && TxtNom_entidad2.Text != "" && TxtNom_entidad3.Text == "" && TxtNom_entidad4.Text == "" && TxtNom_entidad5.Text == "" && TxtNom_entidad6.Text == "" && TxtNom_entidad7.Text == "" && TxtNom_entidad8.Text == "")
            {
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit1.Text,
                                   TxtNom_entidad1.Text, TxtValor1.Text, Txtobligacion1.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit2.Text,
                                   TxtNom_entidad2.Text, TxtValor2.Text, Txtobligacion2.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
            }
            else if (TxtNom_entidad1.Text != "" && TxtNom_entidad2.Text != "" && TxtNom_entidad3.Text != "" && TxtNom_entidad4.Text == "" && TxtNom_entidad5.Text == "" && TxtNom_entidad6.Text == "" && TxtNom_entidad7.Text == "" && TxtNom_entidad8.Text == "")
            {
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit1.Text,
                                   TxtNom_entidad1.Text, TxtValor1.Text, Txtobligacion1.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit2.Text,
                                   TxtNom_entidad2.Text, TxtValor2.Text, Txtobligacion2.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit3.Text,
                                  TxtNom_entidad3.Text, TxtValor3.Text, Txtobligacion3.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
            }
            else if (TxtNom_entidad1.Text != "" && TxtNom_entidad2.Text != "" && TxtNom_entidad3.Text != "" && TxtNom_entidad4.Text != "" && TxtNom_entidad5.Text == "" && TxtNom_entidad6.Text == "" && TxtNom_entidad7.Text == "" && TxtNom_entidad8.Text == "")
            {
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit1.Text,
                                      TxtNom_entidad1.Text, TxtValor1.Text, Txtobligacion1.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit2.Text,
                                   TxtNom_entidad2.Text, TxtValor2.Text, Txtobligacion2.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit3.Text,
                                  TxtNom_entidad3.Text, TxtValor3.Text, Txtobligacion3.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit4.Text,
                                  TxtNom_entidad4.Text, TxtValor4.Text, Txtobligacion4.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
            }
            else if (TxtNom_entidad1.Text != "" && TxtNom_entidad2.Text != "" && TxtNom_entidad3.Text != "" && TxtNom_entidad4.Text != "" && TxtNom_entidad5.Text != "" && TxtNom_entidad6.Text == "" && TxtNom_entidad7.Text == "" && TxtNom_entidad8.Text == "")
            {
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit1.Text,
                                      TxtNom_entidad1.Text, TxtValor1.Text, Txtobligacion1.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit2.Text,
                                   TxtNom_entidad2.Text, TxtValor2.Text, Txtobligacion2.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit3.Text,
                                  TxtNom_entidad3.Text, TxtValor3.Text, Txtobligacion3.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit4.Text,
                                  TxtNom_entidad4.Text, TxtValor4.Text, Txtobligacion4.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit5.Text,
                                  TxtNom_entidad5.Text, TxtValor5.Text, Txtobligacion5.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
            }
            else if (TxtNom_entidad1.Text != "" && TxtNom_entidad2.Text != "" && TxtNom_entidad3.Text != "" && TxtNom_entidad4.Text != "" && TxtNom_entidad5.Text != "" && TxtNom_entidad6.Text != "" && TxtNom_entidad7.Text == "" && TxtNom_entidad8.Text == "")
            {
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit1.Text,
                                     TxtNom_entidad1.Text, TxtValor1.Text, Txtobligacion1.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit2.Text,
                                   TxtNom_entidad2.Text, TxtValor2.Text, Txtobligacion2.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit3.Text,
                                  TxtNom_entidad3.Text, TxtValor3.Text, Txtobligacion3.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit4.Text,
                                  TxtNom_entidad4.Text, TxtValor4.Text, Txtobligacion4.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit5.Text,
                                  TxtNom_entidad5.Text, TxtValor5.Text, Txtobligacion5.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit6.Text,
                                  TxtNom_entidad6.Text, TxtValor6.Text, Txtobligacion6.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
            }
            else if (TxtNom_entidad1.Text != "" && TxtNom_entidad2.Text != "" && TxtNom_entidad3.Text != "" && TxtNom_entidad4.Text != "" && TxtNom_entidad5.Text != "" && TxtNom_entidad6.Text != "" && TxtNom_entidad7.Text != "" && TxtNom_entidad8.Text == "")
            {
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit1.Text,
                                     TxtNom_entidad1.Text, TxtValor1.Text, Txtobligacion1.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit2.Text,
                                   TxtNom_entidad2.Text, TxtValor2.Text, Txtobligacion2.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit3.Text,
                                  TxtNom_entidad3.Text, TxtValor3.Text, Txtobligacion3.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit4.Text,
                                  TxtNom_entidad4.Text, TxtValor4.Text, Txtobligacion4.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit5.Text,
                                  TxtNom_entidad5.Text, TxtValor5.Text, Txtobligacion5.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit6.Text,
                                  TxtNom_entidad6.Text, TxtValor6.Text, Txtobligacion6.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit7.Text,
                                  TxtNom_entidad7.Text, TxtValor7.Text, Txtobligacion7.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
            }
            else if (TxtNom_entidad1.Text != "" && TxtNom_entidad2.Text != "" && TxtNom_entidad3.Text != "" && TxtNom_entidad4.Text != "" && TxtNom_entidad5.Text != "" && TxtNom_entidad6.Text != "" && TxtNom_entidad7.Text != "" && TxtNom_entidad8.Text != "")
            {
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit1.Text,
                                     TxtNom_entidad1.Text, TxtValor1.Text, Txtobligacion1.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit2.Text,
                                   TxtNom_entidad2.Text, TxtValor2.Text, Txtobligacion2.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit3.Text,
                                  TxtNom_entidad3.Text, TxtValor3.Text, Txtobligacion3.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit4.Text,
                                  TxtNom_entidad4.Text, TxtValor4.Text, Txtobligacion4.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit5.Text,
                                  TxtNom_entidad5.Text, TxtValor5.Text, Txtobligacion5.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit6.Text,
                                  TxtNom_entidad6.Text, TxtValor6.Text, Txtobligacion6.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit7.Text,
                                  TxtNom_entidad7.Text, TxtValor7.Text, Txtobligacion7.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
                dataGridView1.Rows.Add(TxtRadicado.Text, Txtcod_oficina.Text, hoy.ToShortDateString(), Txtnom_oficina.Text, Txtciudad.Text, Txtcedula.Text, Txtnombre.Text, TxtNit8.Text,
                                  TxtNom_entidad8.Text, TxtValor8.Text, Txtobligacion8.Text, Txtscoring.Text, Txtnom_gestor.Text, Txtcoordinador.Text, Txtcuenta.Text, "LIBRANZA");
            }
            else
            {
                MessageBox.Show("No hay carteras para remitir", "Información", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
