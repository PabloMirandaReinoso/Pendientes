using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;

using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Windows;
using System.Threading;

using System.Runtime.InteropServices;
using IoOutLook = Microsoft.Office.Interop.Outlook;
using System.Reflection;

using System.Data.OracleClient;
using System.Diagnostics;

using Ionic.Zip;
using System.Net.Http;
using System.Globalization;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;


namespace Pendientes
{
    public partial class Form1 : Form
    {


        private System.Windows.Forms.ContextMenu contextMenu1;
        private System.Windows.Forms.MenuItem menuItem1;

        private void button1_Click(object sender, EventArgs e)
        {
            //FileWatcher(txtRuta.Text);
            //dsConfiguracion(dgCarpetas, "Carpetas", "", txtRuta.Text);
            //Persistencia(dgCarpetas, "Carpetas");
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Merger();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            BaseModificada(txtServer.Text);
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            Persistencia(dgBitacora, "Bitacora");
            dgPedidos.DataSource = EjecucionTotalizar("");
        }

        private void btnBase_Click(object sender, EventArgs e)
        {
            //Aca deberian quedar en la carpeta de paso, pero asociados a un harvest
            BasePackage(txtServer.Text, txtPackage.Text,"C:\\Pendientes\\Liberaciones\\");
            BasePackage("QISCTOS", txtPackage.Text, "C:\\Pendientes\\Liberaciones\\");
            dsBase(dgObjetos, "Packages", cbBaseHarvest.Text, Fecha(), txtServer.Text + '.' + txtPackage.Text);
            Persistencia(dgObjetos, "Packages");
            Proyectos(cbDesarrolloProyecto);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Persistencia(dgBitacora, "Bitacora");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (txtGlosa.Text.Trim()!="")
            {
                CeldaSeleccionada(dgPedidos, "Solicitud", txtGlosa.Text.Trim());
            }

            Persistencia(dgPedidos, "Planificacion");
            CargarHarvest();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Persistencia(dgProyectos, "Proyectos");
            Proyectos(cbDesarrolloProyecto);
            Proyectos(cbAnalisisProyecto);            
        }

        private void btnLiberacionesGuardar_Click(object sender, EventArgs e)
        {
            Persistencia(dgLiberaciones, "Liberaciones");
        }

        private void button6_Click_3(object sender, EventArgs e)
        {
            //ExcelLoad("C:\\Harvest.xls");
        }

        private void btnCorreoExportar_Click(object sender, EventArgs e)
        {
            //Correo(cboCorreoHarvest.Text, txtCorreoCarpeta.Text);
        }

        private void button6_Click_2(object sender, EventArgs e)
        {
            InformeMensual(Mes());
        }

        private void btnCorreoCarpeta_Click(object sender, EventArgs e)
        {
            fbdSoporte.ShowDialog();
            txtDesarrolloRuta.Text = fbdSoporte.SelectedPath;
        }

        private void btnParametrosGuardar_Click(object sender, EventArgs e)
        {
            Persistencia(dgCarpetas, "Carpetas");
            Persistencia(dgArchivosExcluidos, "ArchivosExcluidos");
            Persistencia(dgCorreosFiltro, "CorreosFiltro");
            Persistencia(dgBatLiberacion, "ArchivosBat");
            Persistencia(dgEquipo, "Equipo");
            Persistencia(dgParametrosPEE, "Estimaciones");
            Persistencia(dgParametrosEtapas, "Etapas");

            Tecnologia(dgParametrosPEE, cbAnalisisTecno, "Tipo");
        }

        private void cbPlanificacionEstado_SelectedIndexChanged(object sender, EventArgs e)
        {
            String mess;
            String mesc;
            String avan;
            //int fila;
            int filas;
            bool bVisible = false;
            DataGridViewRow dr;

            dgPedidos.CurrentCell = null;
            mess = (cbPlanificacionEstado.SelectedIndex).ToString();
            filas = dgPedidos.Rows.Count;
            avan = "";

            for (int fila = 0; fila < filas; fila++)
            {
                bVisible = true;
                dr = dgPedidos.Rows[fila];
                if ((mess == "0") || (dr.Cells["Estado"].Value == null))
                {
                    bVisible = true;
                }
                else
                {
                    avan = dr.Cells["Estado"].Value.ToString();
                    if (avan == "Cerrada")
                    {
                        bVisible = false;
                    }
                }
                dr.Visible = bVisible;

            }
        }

        private void btnDesarrolloDetectar_Click(object sender, EventArgs e)
        {
            String sIzquierda;
            String sDerecha;

            string[] aIzquierda ;
            string[] aDerecha;

            String sArchivoI;
            String sArchivoD;

            String sTipo;

            dgApp.DataSource = null;

            sIzquierda = Configuracion(cbDesarrolloProyecto.Text, "Ruta LOC");
            sDerecha = Configuracion(cbDesarrolloProyecto.Text, "Ruta DES");
            Merger(sIzquierda, sDerecha);            

            aIzquierda = Archivos(sIzquierda);
            aDerecha = Archivos(sDerecha);

            //De izquierda a Derecha
            foreach (var izq in aIzquierda)
            {

                sTipo = ArchivoComparar(izq, izq.Replace(sIzquierda,sDerecha));
                if (sTipo != "Identico" || (opcFiltroTodos.Checked))
                {                            
                 if (sTipo == "NED")
                 {
                   sTipo = "No Existe en la Nueva Version";
                 }                            
                 }

                dsComponentes(dgApp, "Aplicativos", cbDesarrolloHarvest.Text, sTipo, izq, Fecha());
            }

            foreach (var der in aDerecha)
            {

                sTipo = ArchivoComparar(der, der.Replace(sDerecha, sIzquierda));
                if (sTipo != "Identico" || (opcFiltroTodos.Checked))
                {
                    if (sTipo == "NEI")
                    {
                        sTipo = "No Existe en la Version Anterior";
                        dsComponentes(dgApp, "Aplicativos", cbDesarrolloHarvest.Text, sTipo, der, Fecha());
                    }
                }

                
            }
                //sArchivoI = izq.Replace(sIzquierda, "");
                //sTipo = "";

                //foreach (var der in aDerecha)
                //{
                //    sArchivoD = der.Replace(sDerecha, "");
                //    if (sArchivoI == sArchivoD)
                //    {
                //        sTipo=ArchivoComparar(izq, der);
                //        if (sTipo != "Identico" || (opcFiltroTodos.Checked))
                //        {
                //            if (sTipo=="NEI")
                //            {
                //                sTipo = "No Existe en Nueva Versión";
                //            }
                //            if (sTipo == "NED")
                //            {
                //                sTipo = "No Existe en Versión Vigente";
                //            }
                //            dsComponentes(dgApp, "Aplicativos", cbDesarrolloHarvest.Text, sTipo, izq, Fecha());
                //        }
                //        break;
                //    }                    
                //}  

                //if (sTipo=="")
                //{
                //    sTipo = ArchivoComparar(izq, "");
                //    if (sTipo != "Identico" || (opcFiltroTodos.Checked))
                //    {
                //        dsComponentes(dgApp, "Aplicativos", cbDesarrolloHarvest.Text, sTipo, izq, Fecha());
                //}
                //}
            
            
            //foreach (var der in aDerecha)
            //{
            //    sArchivoD = der.Replace(sDerecha, "");
            //    sTipo = "";

            //    foreach (var izq in aIzquierda)
            //    {
            //        sArchivoI = izq.Replace(sIzquierda, "");
            //        if (sArchivoD == sArchivoI)
            //        {
            //            if (sTipo != "Identico" || (opcFiltroTodos.Checked))
            //            {
            //                sTipo = ArchivoComparar(izq, der);
            //            }
            //            break;
            //        }
            //    }

            //    if (sTipo == "")
            //    {
            //        sTipo = ArchivoComparar(der, "");
            //        if (sTipo != "Identico" || (opcFiltroTodos.Checked))
            //        {
            //            dsComponentes(dgApp, "Aplicativos", cbDesarrolloHarvest.Text, sTipo, der, Fecha());
            //        }
            //    }
            //}
            

        }

        private void btnDesarrolloGenerar_Click(object sender, EventArgs e)
        {
            string sPass = "explotacion";
            string sTipo;
            string sHarvest = cbDesarrolloHarvest.Text;
            string sArchivo;                         
            string sProyecto="";
                       

            sProyecto=cbDesarrolloProyecto.Text;
            sTipo = Celda(dgProyectos, "Nombre", sProyecto, "Tipo");

          

            if (sTipo == "ASPX")
            {
                sTipo = Compactar(dgApp, txtDesarrolloRuta.Text, "objeto", "explotacion");
            }
            else
            {
                sTipo = Compactar(dgApp, txtDesarrolloRuta.Text, "objeto", "");
                sTipo = ArchivoBAT(sProyecto, txtDesarrolloRuta.Text);
            }

            sArchivo = ExcelDesa(sHarvest, cbAmbienteAplicativo.Text,sProyecto, txtDesarrolloRuta.Text, sTipo);
            dsLiberaciones(dgLiberaciones, "Liberaciones", sHarvest,cbAmbienteAplicativo.Text, Fecha(), sArchivo);
            Persistencia(dgLiberaciones, "Liberaciones");            
            LiberarArchivos(sHarvest);                        
            Respaldar(sProyecto, txtDesarrolloRuta.Text);
        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            fbdSoporte.SelectedPath = txtDesarrolloRuta.Text;
            fbdSoporte.ShowDialog();
            txtDesarrolloRuta.Text = fbdSoporte.SelectedPath;
        }

        private void btnDesarrolloAgregar_Click(object sender, EventArgs e)
        {
            string sRuta = "";
            if (txtDesarrolloRuta.Text == "")
            {
                fbdSoporte.ShowDialog();
                sRuta = fbdSoporte.SelectedPath;
            }
            else
            {
                sRuta = txtDesarrolloRuta.Text.Trim();
            }

            ofdSoporte.Multiselect = true;
            ofdSoporte.InitialDirectory = sRuta;
            ofdSoporte.ShowDialog();
            AgregarArchivos(dgApp, "" , "Aplicativos");
        }

        private void btnDesarrolloGuardar_Click(object sender, EventArgs e)
        {
            Persistencia(dgApp, "Aplicativos");
        }
        private void btnSoporteSqlPlus_Click(object sender, EventArgs e)
        {
            SqlPlus(CeldaSeleccionada(dgSoporte, "Objeto"));
        }

        private void btnSoporteVer_Click(object sender, EventArgs e)
        {
            Ver(CeldaSeleccionada(dgSoporte, "Objeto"));
        }

        private void cboSoporteHarvest_SelectedIndexChanged(object sender, EventArgs e)
        {

            {
                String mess;
                String mesc;
                //int fila;
                int filas;
                DataGridViewRow dr;

                dgSoporte.CurrentCell = null;
                mess = (cboSoporteHarvest.Text).ToString();
                filas = dgSoporte.Rows.Count;

                for (int fila = 0; fila < filas; fila++)
                {
                    dr = dgSoporte.Rows[fila];
                    if (dr.Cells["Id Harvest"].Value != null)
                    {
                        mesc = dr.Cells["Id Harvest"].Value.ToString();
                        if (mess == mesc)
                        {
                            dr.Visible = true;
                        }
                        else
                        {
                            dr.Visible = false;
                        }
                    }

                }
            }
        }
        private void button3_Click_1(object sender, EventArgs e)
        {
            Correo(cboCorreoHarvest.Text, "");
        }

        private void button7_Click(object sender, EventArgs e)
        {

            Char delimiter = '.';
            String sFecha;
            String sArchivo;
            String sLeft;
            String sRigth;
            String sServer;
            String sPackage;
            String sRuta;

            sRuta = "C:\\Pendientes\\";
            sArchivo = CeldaSeleccionada(dgObjetos, "Objeto");
            sServer = sArchivo.Split(delimiter)[0];
            sPackage = sArchivo.Split(delimiter)[1];
            
            sFecha = CeldaSeleccionada(dgObjetos, "Fecha");
            sArchivo = CeldaSeleccionada(dgObjetos, "Objeto");
            sLeft = sRuta + sArchivo + "." + FechaArchivo(sFecha) + ".sql";

            SqlPlus(sServer,sLeft );
        }

        private void btnDecompilar_Click(object sender, EventArgs e)
        {
            AnalizarEnsamblados();

        }

        private void btnApp_Click(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            FiltrarBitacora(Mes());
        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            fbdSoporte.SelectedPath = txtSoporteRuta.Text;
            fbdSoporte.ShowDialog();
            txtSoporteRuta.Text = fbdSoporte.SelectedPath;

            ofdSoporte.Multiselect = true;
            ofdSoporte.InitialDirectory = fbdSoporte.SelectedPath;
            ofdSoporte.ShowDialog();
            AgregarArchivos(dgSoporte,cboSoporteHarvest.Text, "Soporte");
        }

        private void btnSoporte_Click(object sender, EventArgs e)
        {
            Persistencia(dgSoporte, "Soporte");
        }

        private void btnSoporteGenerar_Click(object sender, EventArgs e)
        {
            string sAdjunto;
            string sLiberacion = "";
            string sReversa = "";
            string sRuta = "";
            string sArchivo = "";
            string sHarvest = cboSoporteHarvest.Text;

            string sAdjLibe;
            string sAdjReve;

            string sValidar = "";
            string[] sArchivos;
            string item;
            DialogResult drs;

            sArchivos = Archivos(txtSoporteRuta.Text);
            foreach (DataGridViewRow fila in dgSoporte.Rows)
            {
                if (fila.Visible == true & fila.Cells["Objeto"].Value != null)
                {
                    item = fila.Cells["Objeto"].Value.ToString();
                    if (item.ToUpper().IndexOf(".SQL") > -1)
                    {
                        sValidar = ArchivoValidar(item);
                        if (sValidar != "")
                        {
                            drs = DialogResult.Yes;
                            //drs=MessageBox.Show(sValidar + " ¿Desea Continuar?", item, MessageBoxButtons.YesNo );
                            
                            if (drs!=DialogResult.Yes)
                            { 
                                Ver(item);
                                break;
                            }
                            else
                            {
                                    sValidar="";
                            }
                            
                        }
                    }    
                }
            }            

            if (sValidar=="")
            {
                sLiberacion = Listado(dgSoporte, "Liberacion", "Objeto");
                sReversa = Listado(dgSoporte, "Reversa", "Objeto");
                sRuta = txtSoporteRuta.Text;

                sAdjLibe = CompactarSoporte(dgSoporte,"Liberacion",  sRuta, "objeto", "");
                sAdjReve = CompactarSoporte(dgSoporte, "Reversa", sRuta, "objeto", "");
                sArchivo = ExcelSopo(sHarvest, sRuta, sLiberacion, sAdjLibe, sReversa, sAdjReve);

                dsLiberaciones(dgLiberaciones, "Liberaciones", sHarvest,cbAmbienteSoporte.Text, Fecha(), sArchivo);
                Persistencia(dgLiberaciones, "Liberaciones");
            }

        }
        private void notifyIcon1_DoubleClick(object Sender, EventArgs e)
        {
            // Show the form when the user double clicks on the notify icon.

            // Set the WindowState to normal if the form is minimized.
            if (this.WindowState == FormWindowState.Minimized)
                this.WindowState = FormWindowState.Normal;

            // Activate the form.
            this.Activate();
        }

        private void menuItem1_Click(object Sender, EventArgs e)
        {
            // Close the form, which closes the application.
            this.Close();
        }


        public Form1()
        {
            InitializeComponent();
            this.Show();
            FormSet();

            SetFontAndColors(dgArchivos);
            SetFontAndColors(dgPedidos);
        }

        private void FormSet()
        {

            //Posicion("FIRMA");
            //TodosLosObjetos();
            getImage();
            //ArchivoValidar("C:\\Soporte\\FURE\\CIERREMES201810\\OP10036-01.sql");

            timBackUp.Interval = 6000 * 10;
            
            
            Transacciones(dgArchivos, "Archivos", "ARC");
            Transacciones(dgBitacora, "Bitacora", "REG");
            Transacciones(dgPedidos, "Planificacion", "HAR");
            Transacciones(dgPlanificacion, "Ejecucion", "EJE");

            Transacciones(dgCarpetas, "Carpetas", "CFG");
            Transacciones(dgArchivosExcluidos, "ArchivosExcluidos", "CFG");
            Transacciones(dgCorreosFiltro, "CorreosFiltro", "CFG");

            Transacciones(dgObjetos, "Packages", "ARC");
            Transacciones(dgApp, "Aplicativos", "ARC");
            Transacciones(dgProyectos, "Proyectos", "PRO");
            Transacciones(dgSoporte, "Soporte", "SOP");
            Transacciones(dgLiberaciones, "Liberaciones", "LIB");
            Transacciones(dgBatLiberacion, "ArchivosBat", "CFG");
            Transacciones(dgEquipo, "Equipo", "USR");
            Transacciones(dgParametrosPEE, "Estimaciones", "PEE");
            Transacciones(dgRequerimientos, "Requerimientos", "REQ");
            Transacciones(dgAnalisisDocs, "Documentos", "SOP");
            Transacciones(dgParametrosEtapas, "Etapas", "CFG");

            CargarHarvest();
            Proyectos(cbDesarrolloProyecto);
            Proyectos(cbAnalisisProyecto);            
            
            dgPedidos.DataSource = EjecucionTotalizar("");
            ArchivoMonitor(dgCarpetas);
            TaskBar();
            FiltrarBitacora((DateTime.Now.Month - 1).ToString());
            dgProyectos.Columns["Nombre"].Frozen = true;
            dgBitacora.Columns["Fecha"].Frozen = true;
            dgPlanificacion.Columns["Id Harvest"].Frozen = true;


            Formato(dgArchivos);
            Formato(dgBitacora);
            Formato(dgPedidos);

            Formato(dgCarpetas);
            Formato(dgArchivosExcluidos);
            Formato(dgCorreosFiltro);

            Formato(dgObjetos);
            Formato(dgApp);
            Formato(dgProyectos);
            Formato(dgSoporte);
            Formato(dgLiberaciones);
            Formato(dgBatLiberacion);
            Formato(dgEquipo);
            Formato(dgParametrosPEE);
            Formato(dgRequerimientos);
            Formato(dgAnalisisDocs);
            Formato(dgParametrosEtapas);
            Formato(dgPlanificacion);
            NoSort(dgPlanificacion);
            NoSort(dgBatLiberacion);
            txtInicioPlan.Text = Fecha();

            Tecnologia(dgParametrosPEE, cbAnalisisTecno, "Tipo");
    
        }

        private void Icono(String sTipo)
        {
            notifyIcon1.Icon = new Icon(@"C:\Pendientes\TOOLS\lighton.ico");
        }
        // Define the event handlers.
        private void OnChanged(object source, FileSystemEventArgs e)
        {
            ArchivoCambiado(e.FullPath);

        }

        private void OnRenamed(object source, RenamedEventArgs e)
        {
            ArchivoCambiado(e.FullPath);
        }

        private void FileWatcher(String sRuta)
        {
            try
            {
                // Create a new FileSystemWatcher and set its properties.
                FileSystemWatcher watcher = new FileSystemWatcher();
                watcher.Path = sRuta;

                // Watch both files and subdirectories.
                watcher.IncludeSubdirectories = true;

                // Watch for all changes specified in the NotifyFilters
                //enumeration.
                watcher.NotifyFilter = NotifyFilters.Attributes |
                NotifyFilters.CreationTime |
                NotifyFilters.DirectoryName |
                NotifyFilters.FileName |
                NotifyFilters.LastAccess |
                NotifyFilters.LastWrite |
                NotifyFilters.Security |
                NotifyFilters.Size;

                // Watch all files.
                watcher.Filter = "*.*";

                // Add event handlers.
                watcher.Changed += new FileSystemEventHandler(OnChanged);
                watcher.Created += new FileSystemEventHandler(OnChanged);
                watcher.Deleted += new FileSystemEventHandler(OnChanged);
                watcher.Renamed += new RenamedEventHandler(OnRenamed);

                //Start monitoring.
                watcher.EnableRaisingEvents = true;

            }
            catch (IOException e)
            {
                Console.WriteLine("A Exception Occurred :" + e);
            }

            catch (Exception oe)
            {
                Console.WriteLine("An Exception Occurred :" + oe);
            }
        }

        private DataSet dsPlanificacion(DataGridView dg, String sTabla, String sHarvest, String Solicitud, String JP)
        {
            DataSet ds = new DataSet();
            DataTable dt;
            DataRow nr;


            if (dg.DataSource == null)
            {
                dt = new DataTable(sTabla);

                dt.Columns.Add("Fecha");
                dt.Columns.Add("Id Harvest");
                dt.Columns.Add("Solicitante");
                dt.Columns.Add("Solicitud");
                dt.Columns.Add("Horas Estimadas");
                dt.Columns.Add("Horas Consumidas");
                dt.Columns.Add("Clasificacion");
                dt.Columns.Add("Estado");
            }
            else
            {

                dt = (DataTable)dg.DataSource;
                ds.Tables.Add(dt.Copy());
            }

            nr = dt.NewRow();
            nr["Id Harvest"] = sHarvest;
            nr["Estado"] = "Abierto";
            nr["Fecha"] = Fecha();
            nr["Solicitud"] = Solicitud;
            nr["Solicitante"] = JP;
            dt.Rows.Add(nr);
            dt.AcceptChanges();
            ds.AcceptChanges();

            dg.Invoke((Action)(() => dg.DataSource = dt));
            return ds.Copy();
        }

        private DataSet dsEjecucion(DataGridView dg, String sTabla, String sHarvest, String Proyecto, String Componente , string Descripcion, String Fecha, String Horas)
        {
            DataSet ds = new DataSet();
            DataTable dt;
            DataRow nr;


            if (dg.DataSource == null)
            {
                dt = new DataTable(sTabla);                                
                dt.Columns.Add("Id Harvest");
                dt.Columns.Add("Proyecto");
                dt.Columns.Add("Componente");
                dt.Columns.Add("Id Componente");
                dt.Columns.Add("Descripcion");
                dt.Columns.Add("Iniciar");
                dt.Columns.Add("Finalizar");
                dt.Columns.Add("Horas");
                dt.Columns.Add("% Avance");                
            }
            else
            {

                dt = (DataTable)dg.DataSource;
                ds.Tables.Add(dt.Copy());
            }

            nr = dt.NewRow();
            nr["Id Harvest"] = sHarvest;
            nr["Proyecto"] = Proyecto;
            nr["Id Componente"] = Componente;
            nr["Descripcion"] = Descripcion;
            nr["Iniciar"] = Fecha;
            nr["Finalizar"] = "";
            nr["% Avance"] = "";
            nr["Horas"] = Horas;           
            
            dt.Rows.Add(nr);
            dt.AcceptChanges();
            ds.AcceptChanges();

            dg.Invoke((Action)(() => dg.DataSource = dt));
            return ds.Copy();
        }

        private DataSet dsBitacora(DataGridView dg, String sTabla, String sHarvest, String sSolicitante, String sSolicitud, String sFecha)
        {
            DataSet ds = new DataSet();
            DataTable dt;
            DataRow nr;


            if (dg.DataSource == null)
            {
                dt = new DataTable(sTabla);
                dt.Columns.Add("Fecha");
                dt.Columns.Add("Id Harvest");
                dt.Columns.Add("Solicitante");
                dt.Columns.Add("Solicitud");
                dt.Columns.Add("Actividades");
                dt.Columns.Add("Hora Inicio");
                dt.Columns.Add("Hora Cierre");
                dt.Columns.Add("Horas Consumidas");
                dt.Columns.Add("Tipo");
                dt.Columns.Add("Observaciones");
                dt.Columns.Add("Objeto");                
                dt.Columns.Add("Etapa");
                ds.Tables.Add(dt);
            }
            else
            {

                dt = (DataTable)dg.DataSource;
                ds.Tables.Add(dt.Copy());
            }

            nr = dt.NewRow();
            nr["Id Harvest"] = sHarvest;
            nr["Solicitud"] = sSolicitud;
            nr["Fecha"] = sFecha;
            nr["Solicitante"] = sSolicitante;
            
            dt.Rows.Add(nr);
            dt.AcceptChanges();
            ds.AcceptChanges();

            dg.Invoke((Action)(() => dg.DataSource = dt));
            return ds.Copy();
        }

        private DataSet dsArchivos(DataGridView dg, String sTabla, String sFecha, String sObjeto)
        {
            DataSet ds = new DataSet();
            DataTable dt;
            DataRow nr;
            String sTemp;

            try
            {
                sFecha = sFecha.Substring(0, 16);
                if (dg.DataSource == null)
                {
                    dt = new DataTable(sTabla);
                    dt.Columns.Add("Fecha");
                    dt.Columns.Add("Objeto");
                    ds.Tables.Add(dt);
                }
                else
                {

                    dt = (DataTable)dg.DataSource;
                    ds.Tables.Add(dt.Copy());
                }

                sTemp = sObjeto.Replace(Char.ConvertFromUtf32(10), " ");
                sTemp = sTemp.Replace(Char.ConvertFromUtf32(13), " ");
                sTemp = sTemp.Replace(Char.ConvertFromUtf32(39), "´");

                if (dt.Select("Fecha='" + sFecha + "' AND Objeto='" + sTemp + "'").Length == 0)
                {
                    nr = dt.NewRow();
                    nr["Fecha"] = sFecha;
                    nr["Objeto"] = sObjeto;
                    dt.Rows.Add(nr);
                    dt.AcceptChanges();
                    ds.AcceptChanges();
                }

                dg.Invoke((Action)(() => dg.DataSource = dt));
                return ds.Copy();
            }
            catch (Exception)
            {
                return new DataSet();
            }
            
        }

        private DataSet dsBase(DataGridView dg, String sTabla, String sHarvest, String sFecha, String sObjeto)
        {
            DataSet ds = new DataSet();
            DataTable dt;
            DataRow nr;
            String sTemp;

            sFecha = sFecha.Substring(0, 16);
            if (dg.DataSource == null)
            {
                dt = new DataTable(sTabla);
                //sHarvest
                dt.Columns.Add("Id Harvest");
                dt.Columns.Add("Fecha");
                dt.Columns.Add("Objeto");
                ds.Tables.Add(dt);
            }
            else
            {

                dt = (DataTable)dg.DataSource;
                ds.Tables.Add(dt.Copy());
            }

            nr = dt.NewRow();
            nr["Fecha"] = sFecha;
            nr["Id Harvest"] = sHarvest;
            nr["Objeto"] = sObjeto;
            dt.Rows.Add(nr);
            dt.AcceptChanges();
            ds.AcceptChanges();
            
            dg.Invoke((Action)(() => dg.DataSource = dt));
            return ds.Copy();
        }
        private DataSet dsConfiguracion(DataGridView dg, String sTabla, String sFecha, String sObjeto)
        {
            DataSet ds = new DataSet();
            DataTable dt;
            DataRow nr;

            if (dg.DataSource == null)
            {
                dt = new DataTable(sTabla);
                dt.Columns.Add("Tipo");
                dt.Columns.Add("Objeto");
                ds.Tables.Add(dt);
            }
            else
            {

                dt = (DataTable)dg.DataSource;
                ds.Tables.Add(dt.Copy());
            }

            nr = dt.NewRow();
            nr["Tipo"] = sFecha;
            nr["Objeto"] = sObjeto;
            dt.Rows.Add(nr);
            dt.AcceptChanges();
            ds.AcceptChanges();

            dg.Invoke((Action)(() => dg.DataSource = dt));
            return ds.Copy();
        }
        private DataSet dsProyectos(DataGridView dg, String sTabla)
        {
            DataSet ds = new DataSet();
            DataTable dt;
            DataRow nr;

            if (dg.DataSource == null)
            {
                dt = new DataTable(sTabla);
                //dt.Columns.Add("Fecha");                                
                dt.Columns.Add("Nombre");
                dt.Columns.Add("Tipo");
                dt.Columns.Add("Ruta DNS");
                dt.Columns.Add("Ruta DES");
                dt.Columns.Add("Ruta TES");
                dt.Columns.Add("Ruta NET");
                dt.Columns.Add("Puerto");
                dt.Columns.Add("Jefe Proyecto");
                dt.Columns.Add("Desarrollador");
                ds.Tables.Add(dt);
            }
            else
            {

                dt = (DataTable)dg.DataSource;
                ds.Tables.Add(dt.Copy());
            }

            nr = dt.NewRow();
            //nr["Fecha"] = sFecha;            
            dt.Rows.Add(nr);
            dt.AcceptChanges();
            ds.AcceptChanges();

            dg.Invoke((Action)(() => dg.DataSource = dt));
            return ds.Copy();
        }
        private DataSet dsComponentes(DataGridView dg, String sTabla, String sHarvest, String sTipo, String sObjeto,String sFecha)
        {
            DataSet ds = new DataSet();
            DataTable dt;
            DataRow nr;

            if (dg.DataSource == null)
            {
                dt = new DataTable(sTabla);
                dt.Columns.Add("Id Harvest");
                dt.Columns.Add("Tipo");
                dt.Columns.Add("Objeto");
                dt.Columns.Add("Fecha");
                ds.Tables.Add(dt);
            }
            else
            {

                dt = (DataTable)dg.DataSource;
                ds.Tables.Add(dt.Copy());
            }

            nr = dt.NewRow();
            nr["Id Harvest"] = sHarvest;
            nr["Tipo"] = sTipo;
            nr["Objeto"] = sObjeto;
            nr["Fecha"] = sFecha;            
            dt.Rows.Add(nr);
            dt.AcceptChanges();
            ds.AcceptChanges();

            dg.Invoke((Action)(() => dg.DataSource = dt));
            return ds.Copy();
        }

        private DataSet dsLiberaciones(DataGridView dg, String sTabla, String sHarvest, string sAmbiente,  String sFecha, String sObjeto)
        {
            DataSet ds = new DataSet();
            DataTable dt;
            DataRow nr;

            if (dg.DataSource == null)
            {
                dt = new DataTable(sTabla);
                dt.Columns.Add("Id Harvest");
                dt.Columns.Add("Ambiente");
                dt.Columns.Add("Archivo");
                dt.Columns.Add("Fecha");                
                ds.Tables.Add(dt);
            }
            else
            {

                dt = (DataTable)dg.DataSource;
                ds.Tables.Add(dt.Copy());
            }

            nr = dt.NewRow();
            nr["Id Harvest"] = sHarvest;
            nr["Ambiente"] = sAmbiente;
            nr["Archivo"] = sObjeto;
            nr["Fecha"] = sFecha;
            dt.Rows.Add(nr);
            dt.AcceptChanges();
            ds.AcceptChanges();

            dg.Invoke((Action)(() => dg.DataSource = dt));
            return ds.Copy();
        }

        private DataSet dsUsuarios(DataGridView dg, String sTabla, String sNombre, String sCorreo, String sFono)
        {
            DataSet ds = new DataSet();
            DataTable dt;
            DataRow nr;

            if (dg.DataSource == null)
            {
                dt = new DataTable(sTabla);
                dt.Columns.Add("Nombre");
                dt.Columns.Add("Correo");
                dt.Columns.Add("Fono");
                ds.Tables.Add(dt);
            }
            else
            {

                dt = (DataTable)dg.DataSource;
                ds.Tables.Add(dt.Copy());
            }

            nr = dt.NewRow();
            nr["Nombre"] = sNombre;
            nr["Correo"] = sCorreo;
            nr["Fono"] = sFono;
            dt.Rows.Add(nr);
            dt.AcceptChanges();
            ds.AcceptChanges();

            dg.Invoke((Action)(() => dg.DataSource = dt));
            return ds.Copy();
        }
        private DataSet dsRequerimiento(DataGridView dg, String sTabla, string sHarvest, string sProyecto, String sCodigo, String Descripcion, String sCantidad)
        {
            DataSet ds = new DataSet();
            DataTable dt;
            DataRow nr;

            if (dg.DataSource == null)
            {
                dt = new DataTable(sTabla);
                dt.Columns.Add("Id Harvest");
                dt.Columns.Add("Proyecto");
                dt.Columns.Add("Componente");
                dt.Columns.Add("Tipo Componente");
                dt.Columns.Add("Id Componente");
                dt.Columns.Add("Titulo");
                dt.Columns.Add("Descripcion");
                dt.Columns.Add("% Reutilizacion");
                dt.Columns.Add("Horas");
                ds.Tables.Add(dt);
            }
            else
            {

                dt = (DataTable)dg.DataSource;
                ds.Tables.Add(dt.Copy());
            }

            nr = dt.NewRow();
            nr["Id Harvest"] = sHarvest;
            nr["Proyecto"] = sProyecto;
            nr["Tipo Componente"] = sCodigo;
            nr["Titulo"] = Descripcion;
            nr["Horas"] = sCantidad;
            dt.Rows.Add(nr);
            dt.AcceptChanges();
            ds.AcceptChanges();

            dg.Invoke((Action)(() => dg.DataSource = dt));
            return ds.Copy();
        }

        private DataSet dsPEE(DataGridView dg, String sTabla)
        {
            DataSet ds = new DataSet();
            DataTable dt;
            DataRow nr;

            if (dg.DataSource == null)
            {
                dt = new DataTable(sTabla);
                //dt.Columns.Add("Fecha");                                
                dt.Columns.Add("Codigo");
                dt.Columns.Add("Tipo");
                dt.Columns.Add("Descripcion");
                dt.Columns.Add("Optimista");
                dt.Columns.Add("Pesimista");
                dt.Columns.Add("Sugerido");
                ds.Tables.Add(dt);
            }
            else
            {

                dt = (DataTable)dg.DataSource;
                ds.Tables.Add(dt.Copy());
            }

            nr = dt.NewRow();
            //nr["Fecha"] = sFecha;            
            dt.Rows.Add(nr);
            dt.AcceptChanges();
            ds.AcceptChanges();

            dg.Invoke((Action)(() => dg.DataSource = dt));
            return ds.Copy();
        }


        private void Persistencia(DataGridView dg, String sArchivo)
        {
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();

            try
            {
                ds = new DataSet();
                dt = (DataTable)dg.DataSource;
                dt.TableName = sArchivo;
                ds.Tables.Add(dt.Copy());
                ds.WriteXml("C:\\Pendientes\\Data\\" + sArchivo + ".Xml", XmlWriteMode.IgnoreSchema);
                //ds.WriteXmlSchema(".\\" + sArchivo + ".Xsd");
            }
            catch (Exception)
            {

                return;
            }

        }
        private void Transacciones(DataGridView dg, String sArchivo, string sDefault)
        {
            DataSet ds = new DataSet();
            string sRuta = "C:\\Pendientes\\Data\\";

            try
            {

                if (File.Exists(sRuta + sArchivo + ".xsd"))
                {
                    ds.ReadXmlSchema(sRuta + sArchivo + ".xsd");
                    ds.ReadXml(sRuta + sArchivo + ".xml", XmlReadMode.ReadSchema);                
                }
                else
                {
                    ds.ReadXml(sRuta + sArchivo + ".xml", XmlReadMode.InferSchema);
                }
                
                
                dg.DataSource = ds.Tables[sArchivo];
                dg.Refresh();
            }
            catch (Exception ex)
            {
                if (sDefault == "REG")
                    dsBitacora(dg, sArchivo, "", "","", Fecha());

                if (sDefault == "CFG")
                    dsConfiguracion(dg, sArchivo, Fecha(), "");

                if (sDefault == "ARC")
                    dsArchivos(dg, sArchivo, Fecha(), "");

                if (sDefault == "PRO")
                    dsProyectos(dg, sArchivo);

                if (sDefault == "SOP")
                    dsComponentes(dg, sArchivo, "","", "","");

                if (sDefault == "LIB")
                    dsLiberaciones(dg, sArchivo, "","", "", Fecha());

                if (sDefault == "USR")
                    dsUsuarios(dg, sArchivo, "", "", "");

                if (sDefault == "REQ")
                    dsRequerimiento(dg, sArchivo, "", "", "", "","");

                if (sDefault == "HAR")
                    dsPlanificacion(dg, sArchivo, "", "", "");

                if (sDefault == "BAS")
                    dsBase(dg, sArchivo, "", "", "");

                if (sDefault == "EJE")
                    dsEjecucion(dg, sArchivo, "", "", "","","","");

                if (sDefault == "PEE")
                    dsPEE(dg, sArchivo);                


            }
        }
        private void Correo(string sHarvest, string sRuta)
        {

            DataSet ds = new DataSet();
            Char delimiter = ';';

            Boolean esImputable;

            IoOutLook.Application oApp;
            IoOutLook._NameSpace oNS;
            IoOutLook.MAPIFolder oFolder;
            IoOutLook._Explorer oExp;

            oApp = new IoOutLook.Application();
            oNS = (IoOutLook._NameSpace)oApp.GetNamespace("MAPI");
            oFolder = oNS.GetDefaultFolder(IoOutLook.OlDefaultFolders.olFolderInbox);
            oExp = oFolder.GetExplorer(false);
            oNS.Logon(Missing.Value, "intel.2018", false, true);

            string sFechaItem;
            string sFechaCons;

            var ci = new CultureInfo("es-CL");
                //DateTime dt = DateTime.ParseExact(yourDateInputString, yourFormatString, ci);
                DateTime dParm, dMail;
            
            
            DataTable dt = new DataTable();
            dt.Columns.Add("Fecha");
            dt.Columns.Add("Objeto");
            dt.Columns.Add("Adjuntos");

            Int32 iItem, iItems;
            String sTexto, sAdjun;

            iItems = oFolder.Items.Count;
            IoOutLook.Items items = oFolder.Items;
            sFechaCons = txtFechaEmail.Text;

            for (iItem = 1; iItem <= iItems; iItem++)
            {

                sFechaItem = items[iItem].CreationTime.ToString("dd/MM/yyyy");

                dParm = Convert.ToDateTime(sFechaCons, ci);
                dMail = Convert.ToDateTime(sFechaItem, ci);

                if (dMail > dParm)
                {

                    sTexto = OutLookItem(items[iItem]);

                    if ((sTexto.IndexOf(sHarvest) > -1) && (sTexto != ""))
                    {
                        sAdjun = OutLookAdju(items[iItem], sRuta);
                        dt.Rows.Add(dt.NewRow());
                        dt.Rows[(dt.Rows.Count - 1)]["Fecha"] = items[iItem].CreationTime.ToString();
                        dt.Rows[(dt.Rows.Count - 1)]["Objeto"] = sTexto.Split(delimiter) [0];
                        dt.Rows[(dt.Rows.Count - 1)]["Adjuntos"] = sAdjun;

                    }
                }
            }

            dt.AcceptChanges();

            oExp = null;
            oFolder = null;
            oNS = null;
            oApp = null;

            if (dt.Rows.Count > 0)
            {
                dgCorreos.Invoke((Action)(() => dgCorreos.DataSource = dt));
                Persistencia(dgCorreos, "Correos");
            }
        }
        private void TaskBar()
        {
            this.components = new System.ComponentModel.Container();
            this.contextMenu1 = new System.Windows.Forms.ContextMenu();
            this.menuItem1 = new System.Windows.Forms.MenuItem();

            // Initialize contextMenu1
            this.contextMenu1.MenuItems.AddRange(
                        new System.Windows.Forms.MenuItem[] {this.menuItem1
        });

            // Initialize menuItem1
            this.menuItem1.Index = 0;
            this.menuItem1.Text = "E&xit";
            this.menuItem1.Click += new System.EventHandler(this.menuItem1_Click);

            // Set up how the form should be displayed.
            this.ClientSize = new System.Drawing.Size(292, 266);
            this.Text = "Notify Icon Example";

            // Create the NotifyIcon.
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);

            // The Icon property sets the icon that will appear
            // in the systray for this application.
            notifyIcon1.Icon = new Icon(@"C:\Pendientes\TOOLS\appicon.ico");


            // The ContextMenu property sets the menu that will
            // appear when the systray icon is right clicked.
            notifyIcon1.ContextMenu = this.contextMenu1;

            // The Text property sets the text that will be displayed,
            // in a tooltip, when the mouse hovers over the systray icon.
            notifyIcon1.Text = "Form1 (NotifyIcon example)";
            notifyIcon1.Visible = true;

            // Handle the DoubleClick event to activate the form.
            notifyIcon1.DoubleClick += new System.EventHandler(this.notifyIcon1_DoubleClick);

        }

        private Boolean MailLaboral(String sTexto)
        {
            Boolean bLaborar = false;

            try
            {
                if (sTexto.IndexOf("-INC-") > -1)
                { bLaborar = true; }

                if (sTexto.IndexOf("-REQ-") > -1)
                { bLaborar = true; }

            }
            catch (Exception ex)
            {

                bLaborar = false;
            }

            return bLaborar;
        }

        private String OutLookItem(object selObject )
        {
            string sTexto;
            Boolean EsImputable;
            
            sTexto = "";
            if (selObject is IoOutLook.MailItem)
            {
                IoOutLook.MailItem mailItem =
                    (selObject as IoOutLook.MailItem);

                if ((mailItem.Subject != null) && (mailItem.Body !=null))
                {
                    sTexto = mailItem.Subject.ToString().Replace(Char.ConvertFromUtf32(59), " ") + Char.ConvertFromUtf32(59) + mailItem.Body;                
                }
                else
                {
                sTexto = mailItem.Subject  ;
                }

            }
            else if (selObject is IoOutLook.ContactItem)
            {
                IoOutLook.ContactItem contactItem =
                    (selObject as IoOutLook.ContactItem);

                sTexto = contactItem.Subject + " ";
            }
            else if (selObject is IoOutLook.AppointmentItem)
            {
                IoOutLook.AppointmentItem apptItem =
                    (selObject as IoOutLook.AppointmentItem);

                sTexto = apptItem.Subject + " ";
            }
            else if (selObject is IoOutLook.TaskItem)
            {
                IoOutLook.TaskItem taskItem =
                    (selObject as IoOutLook.TaskItem);

                sTexto = taskItem.Body + " ";
            }
            else if (selObject is IoOutLook.MeetingItem)
            {
                IoOutLook.MeetingItem meetingItem =
                    (selObject as IoOutLook.MeetingItem);

                sTexto = meetingItem.Subject + " " + meetingItem.Body;
            }
            else
            {
                sTexto = "Que paso???";
            }

            if (sTexto == null)
            {
                sTexto = "";
            }
            return sTexto;
        }
        private String OutLookAdju(object selObject, string sRuta)
        {
            string attachmentList = string.Empty;

            if (selObject is IoOutLook.MailItem)
            {
                Microsoft.Office.Interop.Outlook.Attachments attachments = null;
                Microsoft.Office.Interop.Outlook.MailItem mail = null;

                mail = (Microsoft.Office.Interop.Outlook.MailItem)selObject;
                attachments = mail.Attachments;
                for (int i = 1; i <= attachments.Count; i++)
                {
                    Microsoft.Office.Interop.Outlook.Attachment attachment = attachments[i];
                    attachmentList += " " + attachment.DisplayName + Environment.NewLine;

                    if (sRuta != "")
                    {
                        Directory.CreateDirectory(sRuta + "\\" + cboCorreoHarvest.Text);
                        attachment.SaveAsFile(sRuta + "\\" + cboCorreoHarvest.Text + "\\" + attachment.DisplayName);
                    }
                    Marshal.ReleaseComObject(attachment);
                }
            }
            return attachmentList;
        }
        private String OutLookSend()
        {
            return "";
        }
        private Boolean Filtro(DataGridView dgObjeto, string sObjeto)
        {
            Boolean bExiste = false;
            DataTable dt = ((DataTable)dgObjeto.DataSource);
            String sFiltro;

            foreach (DataRow fila in dt.Rows)
            {
                sFiltro = fila.ItemArray[0].ToString();
                if (sObjeto.IndexOf(sFiltro) > -1)
                {
                    bExiste = true;
                    break;
                }

            }


            return bExiste;
        }
        private void SetFontAndColors(DataGridView dataGridView1)
        {

            return;
            dataGridView1.DefaultCellStyle.Font = new System.Drawing.Font("Tahoma", 11);

            dataGridView1.Dock = DockStyle.Fill;
            dataGridView1.BackgroundColor = Color.LightSteelBlue;
            dataGridView1.BorderStyle = BorderStyle.Fixed3D;

            // Set property values appropriate for read-only display and 
            // limited interactivity. 
            dataGridView1.AllowUserToAddRows = true;
            dataGridView1.AllowUserToDeleteRows = true;
            dataGridView1.AllowUserToOrderColumns = true;
            dataGridView1.ReadOnly = false;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.MultiSelect = true;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            dataGridView1.AllowUserToResizeColumns = true;
            dataGridView1.ColumnHeadersHeightSizeMode =
                DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
            dataGridView1.AllowUserToResizeRows = true;
            dataGridView1.RowHeadersWidthSizeMode =
                DataGridViewRowHeadersWidthSizeMode.EnableResizing;

            // Set the selection background color for all the cells.
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.White;
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;

            // Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default
            // value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
            dataGridView1.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty;

            // Set the background color for all rows and for alternating rows. 
            // The value for alternating rows overrides the value for all rows. 
            dataGridView1.RowsDefaultCellStyle.BackColor = Color.LightGray;
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.LightSteelBlue;

            // Set the row and column header styles.
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Black;

            // Set the Format property on the "Last Prepared" column to cause
            // the DateTime to be formatted as "Month, Year".
            //dataGridView1.Columns["Last Prepared"].DefaultCellStyle.Format = "y";

            // Specify a larger font for the "Ratings" column. 
            // using (Font font = new Font(
            //     dataGridView1.DefaultCellStyle.Font.FontFamily, 25, FontStyle.Bold))
            // {
            //     dataGridView1.Columns["Rating"].DefaultCellStyle.Font = font;
            // }

            // Attach a handler to the CellFormatting event.
            //dataGridView1.CellFormatting += new
            //    DataGridViewCellFormattingEventHandler(dataGridView1_CellFormatting);
        }
        private void ArchivoCambiado(String sArchivo)
        {
            Icono("REV");
            if (!(Filtro(dgArchivosExcluidos, sArchivo)))
            {
                dsArchivos(dgArchivos, "Archivos", Fecha(), sArchivo);
                Persistencia(dgArchivos, "Archivos");
            }
        }
        private Boolean ArchivoMonitor(DataGridView dgObjeto)
        {
            Boolean bExiste = false;
            DataTable dt = ((DataTable)dgObjeto.DataSource);
            String sFiltro;

            foreach (DataRow fila in dt.Rows)
            {
                sFiltro = fila.ItemArray[0].ToString();
                FileWatcher(sFiltro);

            }
            return bExiste;
        }

        private string BaseObjeto(string sServer, string sObject, string sRuta)
        {
            //string sServer = txtServer.Text;
            string sConnection = "Data Source=" + sServer + ";User Id=afil;Password=sistafil;";
            string sSource = "";
            string sNombre = "";


            string sPCKDEF = "";
            string sPCKBOD = "";
            string sTipo="";

            string sFecArchivo = FechaArchivo(Fecha());
            using (OracleConnection connection = new OracleConnection(sConnection))
            {                
                OracleDataReader odrSource;
                OracleCommand command = new OracleCommand("SELECT  DISTINCT  DECODE(OBJECT_TYPE,'PACKAGE BODY','PACKAGE','JOB','PROCOBJ','TYPE BODY','TYPE', OBJECT_TYPE) TYPE FROM  DBA_OBJECTS where owner ='AFIL' AND OBJECT_NAME='" + sObject + "'");
                command.Connection = connection;
                try
                {
                    connection.Open();
                    odrSource = command.ExecuteReader();

                    while (odrSource.Read())
                    {                        
                       sTipo =  odrSource.GetOracleString(0).Value ;
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }              

                command = new OracleCommand("SELECT dbms_metadata.get_ddl('" + sTipo + "','" + sObject + "') FROM dual");
                command.Connection = connection;
                    try
                    {
                        //connection.Open();
                        odrSource = command.ExecuteReader();

                        while (odrSource.Read())
                        {
                            sPCKBOD = sPCKBOD + odrSource.GetOracleLob(0).Value;
                        }

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                    }
                }
                

                sPCKBOD = sPCKBOD + Environment.NewLine + "/" + Environment.NewLine;
                sSource = "SET DEFINE OFF;" + Environment.NewLine;
                sSource = sSource + "ALTER SESSION SET CURRENT_SCHEMA = AFIL;" + Environment.NewLine + sPCKDEF + sPCKBOD;
                //sSource = sSource + "pause Pulse Enter para salir" + Environment.NewLine + "QUIT" + Environment.NewLine;
                sNombre = sServer + '.' + sTipo + "." + sObject + "." +  sFecArchivo + ".sql";
                File.WriteAllText(sNombre, sSource, Encoding.GetEncoding(1252));            

                return sNombre;
        }

        private string BaseObjeto(string sServer, string sObject, string sRuta,string sUsr, string sPwd)
        {
            //string sServer = txtServer.Text;
            string sConnection = "Data Source=" + sServer + ";User Id=" + sUsr + ";Password=" + sPwd +" ;";
            string sSource = "";
            string sNombre = "";


            string sPCKDEF = "";
            string sPCKBOD = "";
            string sTipo = "";

            string sFecArchivo = FechaArchivo(Fecha());
            using (OracleConnection connection = new OracleConnection(sConnection))
            {
                OracleDataReader odrSource;
                OracleCommand command = new OracleCommand("SELECT  DISTINCT  DECODE(OBJECT_TYPE,'PACKAGE BODY','PACKAGE','JOB','PROCOBJ','TYPE BODY','TYPE', OBJECT_TYPE) TYPE FROM  DBA_OBJECTS where owner ='AFIL' AND OBJECT_NAME='" + sObject + "'");
                command.Connection = connection;
                try
                {
                    connection.Open();
                    odrSource = command.ExecuteReader();

                    while (odrSource.Read())
                    {
                        sTipo = odrSource.GetOracleString(0).Value;
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                command = new OracleCommand("SELECT dbms_metadata.get_ddl('" + sTipo + "','" + sObject + "') FROM dual");
                command.Connection = connection;
                try
                {
                    //connection.Open();
                    odrSource = command.ExecuteReader();

                    while (odrSource.Read())
                    {
                        sPCKBOD = sPCKBOD + odrSource.GetOracleLob(0).Value;
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                finally
                {
                    connection.Close();
                }
            }


            sPCKBOD = sPCKBOD + Environment.NewLine + "/" + Environment.NewLine;
            sSource = "SET DEFINE OFF;" + Environment.NewLine;
            sSource = sSource + "ALTER SESSION SET CURRENT_SCHEMA = AFIL;" + Environment.NewLine + sPCKDEF + sPCKBOD;
            //sSource = sSource + "pause Pulse Enter para salir" + Environment.NewLine + "QUIT" + Environment.NewLine;
            sNombre = sServer + '.' + sTipo + "." + sObject + "." + sFecArchivo + ".sql";
            File.WriteAllText(sNombre, sSource, Encoding.GetEncoding(1252));

            return sNombre;
        }

        private string BasePackage(string sServer, string sPackage, string sRuta)
        {
            //string sServer = txtServer.Text;
            string sConnection = "Data Source=" + sServer + ";User Id=afil;Password=sistafil;";
            string sSource = "";
            string sNombre = "";


            string sPCKDEF = "";
            string sPCKBOD = "";

            using (OracleConnection connection = new OracleConnection(sConnection))
            {
                string sFecArchivo = FechaArchivo(Fecha());
                OracleDataReader odrSource;

                OracleCommand command = new OracleCommand("select text  from user_source where upper(name)='" + sPackage.ToUpper() + "' and type='PACKAGE'  order by type, line");
                command.Connection = connection;
                try
                {
                    connection.Open();
                    odrSource = command.ExecuteReader();

                    while (odrSource.Read())
                    {
                        if (sPCKDEF == "")
                        {
                            sPCKDEF = "CREATE OR REPLACE " + odrSource.GetOracleString(0).Value + Environment.NewLine;
                        }
                        else
                        {
                            sPCKDEF = sPCKDEF + odrSource.GetOracleString(0).Value;
                        }
                        
                    }
                    
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                if (sPCKDEF != "")
                {
                    sPCKDEF = sPCKDEF + Environment.NewLine + "/" + Environment.NewLine;

                    command = new OracleCommand("select text  from user_source where upper(name)='" + sPackage.ToUpper() + "' and type='PACKAGE BODY'  order by type, line");
                    command.Connection = connection;
                    try
                    {
                        //connection.Open();
                        odrSource = command.ExecuteReader();

                        while (odrSource.Read())
                        {
                            //if ((odrSource.GetOracleString(0).Value.IndexOf("PACKAGE") > -1) || (odrSource.GetOracleString(0).Value.IndexOf("PACKAGE BODY") > -1))
                            if (sPCKBOD == "")
                            {
                                sPCKBOD = "CREATE OR REPLACE " + odrSource.GetOracleString(0).Value + Environment.NewLine;
                            }
                            else
                            {
                                sPCKBOD = sPCKBOD + odrSource.GetOracleString(0).Value;
                            }

                        }

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                    }
                }
                else
                {                    
                    //TRIGGER
                    command = new OracleCommand("select text  from user_source where upper(name)='" + sPackage.ToUpper() + "' and type='TRIGGER'  order by type, line");
                    command.Connection = connection;
                    try
                    {
                        //connection.Open();
                        odrSource = command.ExecuteReader();

                        while (odrSource.Read())
                        {
                            //if ((odrSource.GetOracleString(0).Value.IndexOf("PACKAGE") > -1) || (odrSource.GetOracleString(0).Value.IndexOf("PACKAGE BODY") > -1))
                            if (sPCKBOD == "")
                            {
                                sPCKBOD = "CREATE OR REPLACE " + odrSource.GetOracleString(0).Value + Environment.NewLine;
                            }
                            else
                            {
                                sPCKBOD = sPCKBOD + odrSource.GetOracleString(0).Value;
                            }

                        }

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                    }

                }                


                    sPCKBOD = sPCKBOD + Environment.NewLine + "/" + Environment.NewLine;
                    sSource = "SET DEFINE OFF;" + Environment.NewLine;
                    sSource = sSource  +  "ALTER SESSION SET CURRENT_SCHEMA = AFIL;" + Environment.NewLine + sPCKDEF + sPCKBOD ;
                    //sSource = sSource + "pause Pulse Enter para salir" + Environment.NewLine + "QUIT" + Environment.NewLine;
                    sNombre = sServer + '.' + sPackage + "." + sFecArchivo + ".sql";
                    File.WriteAllText(sRuta + "\\" + sNombre, sSource, Encoding.GetEncoding(1252));
            }

            return sNombre;
        }

        private string Fecha()
        {

            return DateTime.Now.Day.ToString("00") + "-" + DateTime.Now.Month.ToString("00") + "-" + DateTime.Now.Year.ToString("0000") + " " + DateTime.Now.Hour.ToString("00") + ":" + DateTime.Now.Minute.ToString("00");
        }

        private string FechaAMD()
        {

            return DateTime.Now.Year.ToString("0000") + "-" + DateTime.Now.Month.ToString("00") + "-" + DateTime.Now.Day.ToString("00") + "-" + DateTime.Now.Month.ToString("00") + "T" + DateTime.Now.Hour.ToString("00") + ":" + DateTime.Now.Minute.ToString("00") + ":" + DateTime.Now.Second.ToString("00");
        }

        private string Fecha(string Formato)
        {

            return DateTime.Now.ToString(Formato) ;
        }
        private string Duracion(String Hrs)
        {

            Char delimiter = '.';            
            //PT8H0M0S
            string[] Hora;
            string Mins ="0M";
            Hora=Hrs.Split(delimiter);
            if (Hora.Length>1)
            {
                Mins = (60 * (Convert.ToSingle(Hora[1]) / 10)).ToString() + "M";
            }
            return "PT" + Hora[0] + "H" + Mins + "0S" ;
        }

        private string FechaArchivo()
        {

            return DateTime.Now.Year.ToString("0000") + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00") + DateTime.Now.Hour.ToString("00") + ":" + DateTime.Now.Minute.ToString("00");
        }

        private string FechaArchivo(string sFecha)
        {
            //
            //01/01/2018 01:01
            return sFecha.Substring(6, 4) + sFecha.Substring(3, 2) + sFecha.Substring(0, 2) + sFecha.Substring(11, 2) + sFecha.Substring(14, 2);
        }

        private string FechaCorta(string sFecha)
        {
            //
            //01/01/2018
            return sFecha.Substring(0, 10) ;
        }
        private string CeldaSeleccionada(DataGridView dg, String sNombre)
        {
            String sDato = "";

            if (dg.SelectedCells.Count > 0)
            {
                int selectedrowindex = dg.SelectedCells[0].RowIndex;
                DataGridViewRow selectedRow = dg.Rows[selectedrowindex];
                sDato = Convert.ToString(selectedRow.Cells[sNombre].Value);
            }

            return sDato;
        }

        private string CeldaSeleccionada(DataGridView dg, String sNombre, String sValor)
        {
            String sDato = "";

            if (dg.SelectedCells.Count > 0)
            {
                int selectedrowindex = dg.SelectedCells[0].RowIndex;
                DataGridViewRow selectedRow = dg.Rows[selectedrowindex];
                selectedRow.Cells[sNombre].Value=sValor;
            }

            return sDato;
        }
        private string Celda(DataGridView dg, String Columna, String Fila, String sDato)
        {
            String sValor = "";
            int iFilas = 0;
            DataTable dt = new DataTable();

            dt = (DataTable)dg.DataSource;

            foreach (DataColumn colum in dt.Columns)
            {
                if (colum.ColumnName.ToUpper() == Columna.ToUpper())
                {
                    foreach (DataRow item in dt.Rows)
                    {
                        if (item[Columna].ToString() == Fila)
                        {
                            sValor = item[sDato].ToString();
                            break;
                        }
                    }
                    break;
                }
            }

            return sValor;
        }
        private string BaseModificada(string sServidor)
        {
            //string sServer = txtServer.Text;
            string sConnection = "Data Source=" + sServidor +";User Id=afil;Password=sistafil;";
            string sSource = "";
            string sNombre = "";
            string sFecha;
            string sObjeto;

            using (OracleConnection connection = new OracleConnection(sConnection))
            {
                string sFecArchivo = FechaArchivo(Fecha());
                string sSql = "";
                string sFiltro = "";

                sFiltro = txtBaseFiltro.Text.ToUpper().Trim();
                //sSql = "SELECT  object_type  , OBJECT_NAME, last_ddl_time FROM    DBA_OBJECTS where owner ='AFIL' and (object_type like '%PACKAGE%' or object_type in ('TABLE','INDEX','JOB','FUNCTION')) and last_ddl_time > to_date('" + txtFecha.Text + "','dd/mm/yyyy hh24:mi')  and last_ddl_time <>  created ";
                sSql = "SELECT  distinct  OBJECT_NAME , max(last_ddl_time) FROM    DBA_OBJECTS where owner ='AFIL' and last_ddl_time > to_date('" + txtFecha.Text + "','dd/mm/yyyy hh24:mi') ";
                if (sFiltro!="")
                {
                    sSql = sSql + " and upper(object_name) like '%" + sFiltro + "%'";
                }
                sSql = sSql + " group by OBJECT_NAME order by object_name ";
                OracleDataReader odrSource;
                OracleCommand command = new OracleCommand(sSql);
                command.Connection = connection;
                try
                {
                    connection.Open();
                    odrSource = command.ExecuteReader();
                    dgCambios.DataSource = null;
                    while (odrSource.Read())
                    {
                        sObjeto = odrSource.GetOracleString(0).Value;// +'.' + odrSource.GetOracleString(1).Value;
                        sFecha = odrSource.GetOracleDateTime(1).ToString();

                        dsArchivos(dgCambios, "Cambios", sFecha, sObjeto);
                    }


                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            return sNombre;
        }

        
        private string TodosLosObjetos()
        {
            //string sServer = txtServer.Text;
            string sConnection = "Data Source=TISCTOS;User Id=afil;Password=sistafil;";
            string sSource = "";
            string sNombre = "";
            string sFecha;
            string sObjeto;

            Char delimiter = '.';
            //PT8H0M0S
            string[] aCre = new string[67] ;
            string Mins = "0M";
            
            aCre[0] = "ADM_CAEC|ADM_CAEC";
            aCre[1] = "AFIL|SISTAFIL";
            aCre[2] = "BENEFICIOS|BENEFICIOS";
            aCre[3] = "BONELEC|bonelec";
            aCre[4] = "CANAL_V_CON|CANAL_V_CON";
            aCre[5] = "CANAL_V_DAT|CANAL_V_DAT";
            aCre[6] = "CANAL_V_SOF|CANAL_V_SOF";
            aCre[7] = "CANALES_VIRTUALES|CANALES_VIRTUALES+2";
            aCre[8] = "CASMED|CASMED";
            aCre[9] = "COLECTIVOS|LECCO1234$";
            aCre[10] = "COMPAG|COMPAG02";
            aCre[11] = "CONMED|CONMED";
            aCre[12] = "CONTENT|CONTENT";
            aCre[13] = "CONVENIOS|CONVENIOS";
            aCre[14] = "COTIZ|COTIZ";
            aCre[15] = "CREDITO|CREDITO";
            aCre[16] = "CTAIND|CTAIND";
            aCre[17] = "CUELIMED_SOF|CUELIMED_SOF";
            aCre[18] = "CUENTASMEDICAS|cuentasmedicast";
            aCre[19] = "CUENTAUNICA|Cuniq_x13";
            aCre[20] = "DATA_CAJTES|CAJATESO";
            aCre[21] = "DATA_CORREOSCHILE|CORREOSCHILE";
            aCre[22] = "ENCUESTA|ENCUESTA";
            aCre[23] = "EXPLOTACION|Explota+22";
            aCre[24] = "FARMAEXC_SOF|FARMAEXC_SOF";
            aCre[25] = "FRAUDE|FRAUDE";
            aCre[26] = "FUGA|FUGA";
            aCre[27] = "GESARA|GESARA";
            aCre[28] = "GESSOL|GESSOL";
            aCre[29] = "GLOBAL|GLOGLO";
            aCre[30] = "INGRESOS|INGRESOS+2";
            aCre[31] = "ISAPRE|ISAPRE+2";
            aCre[32] = "LICMED|LICMED";
            aCre[33] = "LIQUIDADOR|LIQUIDADOR";
            aCre[34] = "MALENTERADAS|MAEN04";
            aCre[35] = "PAGOS_PREST|PAGOS_PREST";
            aCre[36] = "PEPA|PEPA04";
            aCre[37] = "PLANBC|PLANBC";
            aCre[38] = "PLANES|PLAN01";
            aCre[39] = "PLATAFORMA|PLAT2003";
            aCre[40] = "PORTAL_EMPRESAS|EmPr2009_";
            aCre[41] = "PORTAL_PLAT_SOF|PORTAL_PLAT_SOF";
            aCre[42] = "PREMIOS|BANCONSALUD";
            aCre[43] = "PREVENTIVO|AUGE";
            aCre[44] = "PRODCOM|PRODCOM";
            aCre[45] = "PRODUCTOS_CONVENIOS|PRODUCTOS+2";
            aCre[46] = "RECLAMOS|RECLAMOS";
            aCre[47] = "REEMBOLSOS|ReembolsosT";
            aCre[48] = "REGCTA|REGCTA";
            aCre[49] = "RESCATE|RESC01";
            aCre[50] = "SALCOBRAND|SALCOBRAND";
            aCre[51] = "SALUDADMINISTRADA|SALUDADMINISTRADA";
            aCre[52] = "SEG_SOLTAR|SEG_SOLTAR";
            aCre[53] = "SEGUIMIENTO|SEGUIMIENTO";
            aCre[54] = "SEGURIDAD|SEGURIDAD";
            aCre[55] = "SEGUROC|SEGUROC";
            aCre[56] = "SFECOB|SFECOB";
            aCre[57] = "SIL|SILTECNO";
            aCre[58] = "SIL_PAGOPRESTADORES|SIL_PAGOPRESTADORES+2";
            aCre[59] = "SOC|SOC";
            aCre[60] = "SUPERINTENDENCIA|SUPERA+0";
            aCre[61] = "TARICOL|TARICOL";
            aCre[62] = "USERCOMPAG|USERCOMPAG";
            aCre[63] = "USERSFECOB|USERSFECOB";
            aCre[64] = "VFUNES|SISTVFU";
            aCre[65] = "VIP|VIP$ABC1";
            aCre[66] = "WEB|TECNOGEST";


            for (int i = 0; i < 67; i++)
			{
			 			
            sConnection = "Data Source=TISCTOS;User Id=" + aCre[i].Split(delimiter)[0] + ";Password=" + aCre[i].Split(delimiter)[1] + ";";
            using (OracleConnection connection = new OracleConnection(sConnection))
            {
                string sFecArchivo = FechaArchivo(Fecha());
                string sSql = "";
                string sFiltro = "";

                sFiltro = txtBaseFiltro.Text.ToUpper().Trim();
                //sSql = "SELECT  object_type  , OBJECT_NAME, last_ddl_time FROM    DBA_OBJECTS where owner ='AFIL' and (object_type like '%PACKAGE%' or object_type in ('TABLE','INDEX','JOB','FUNCTION')) and last_ddl_time > to_date('" + txtFecha.Text + "','dd/mm/yyyy hh24:mi')  and last_ddl_time <>  created ";
                sSql = "SELECT  distinct  OBJECT_NAME , max(last_ddl_time) FROM    DBA_OBJECTS where owner ='" +  aCre[i].Split(delimiter)[0] + "'";                
                sSql = sSql + " group by OBJECT_NAME order by object_name ";
                OracleDataReader odrSource;
                OracleCommand command = new OracleCommand(sSql);
                command.Connection = connection;
                try
                {
                    connection.Open();
                    odrSource = command.ExecuteReader();
                    dgCambios.DataSource = null;
                    while (odrSource.Read())
                    {
                        sObjeto = odrSource.GetOracleString(0).Value;
                        BaseObjeto("TISCTOS", sObjeto, "@c:\todos", aCre[i].Split(delimiter)[0], aCre[i].Split(delimiter)[1]);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
          }

            return sNombre;
        }

        private void Merger()
        {


            String sLeft;
            String sRigth;
            String sArchivo;
            String sServer;
            String sPackage;
            String sMerger;
            Char delimiter = '.';


            //Este se usa pero hay que usar el campo txtServer
            sArchivo = CeldaSeleccionada(dgObjetos, "Objeto");
            sServer = sArchivo.Split(delimiter)[0];
            sPackage = sArchivo.Split(delimiter)[1];
            sMerger = "C:\\Pendientes\\Liberaciones\\";

            sLeft = BasePackage(sServer, sPackage, sMerger);
            sRigth = BasePackage(txtServer.Text, sPackage, sMerger);

            Process compiler = new Process();
            compiler.StartInfo.FileName = "C:\\Program Files (x86)\\WinMerge\\WinMergeU.exe";
            compiler.StartInfo.Arguments =sMerger + sLeft + " " + sMerger +  sRigth;
            compiler.StartInfo.UseShellExecute = false;
            compiler.StartInfo.RedirectStandardOutput = true;
            compiler.Start();
            compiler.WaitForExit(5000);

        }
        private void Merger(String sLeft, String sRigth)
        {

            Process compiler = new Process();
            compiler.StartInfo.FileName = "C:\\Program Files (x86)\\WinMerge\\WinMergeU.exe";
            compiler.StartInfo.Arguments = sLeft + " " + sRigth;
            compiler.StartInfo.UseShellExecute = false;
            compiler.StartInfo.RedirectStandardOutput = true;
            compiler.Start();
            compiler.WaitForExit(5000);
        }
        private void SqlPlus()
        {


            String sLeft;
            String sRigth;
            String sArchivo;
            String sServer;
            String sPackage;
            Char delimiter = '.';
            String sRuta;

            var location = new Uri(Assembly.GetEntryAssembly().GetName().CodeBase);
            sRuta = new FileInfo(location.AbsolutePath).Directory.ToString() + "\\";


            sArchivo = CeldaSeleccionada(dgObjetos, "Objeto");
            sServer = sArchivo.Split(delimiter)[0];
            sPackage = sArchivo.Split(delimiter)[1];

            sRigth = BasePackage(sServer, sPackage, "C:\\Pendientes\\Liberaciones\\");

            Process compiler = new Process();
            compiler.StartInfo.FileName = "C:\\ODAC1025x64\\BIN\\sqlplusw.exe";
            compiler.StartInfo.Arguments = "afil/sistafil@qisctos @" + sRuta + sRigth;
            compiler.StartInfo.UseShellExecute = false;
            compiler.StartInfo.RedirectStandardOutput = true;
            compiler.Start();
            compiler.WaitForExit(5000);

        }

        private void SqlPlus(string sArchivo)
        {

            String sLeft;
            String sRigth;
            String sServer;
            String sPackage;
            Char delimiter = '.';
            String sRuta;

            var location = new Uri(Assembly.GetEntryAssembly().GetName().CodeBase);
            sRuta = new FileInfo(location.AbsolutePath).Directory.ToString() + "\\";

            Process compiler = new Process();
            compiler.StartInfo.FileName = "C:\\ODAC1025x64\\BIN\\sqlplusw.exe";
            compiler.StartInfo.Arguments = "afil/sistafil@qisctos @" + sArchivo;
            compiler.StartInfo.UseShellExecute = false;
            compiler.StartInfo.RedirectStandardOutput = true;
            compiler.Start();
            compiler.WaitForExit(5000);

        }

        private void SqlPlus(string sServidor, string sArchivo)
        {

            String sLeft;
            String sRigth;
            String sServer;
            String sPackage;
            Char delimiter = '.';
            String sRuta;

            var location = new Uri(Assembly.GetEntryAssembly().GetName().CodeBase);
            sRuta = new FileInfo(location.AbsolutePath).Directory.ToString() + "\\";

            Process compiler = new Process();
            compiler.StartInfo.FileName = "C:\\ODAC1025x64\\BIN\\sqlplusw.exe";
            compiler.StartInfo.Arguments = "afil/sistafil@" + sServidor +" @" + sArchivo;
            compiler.StartInfo.UseShellExecute = false;
            compiler.StartInfo.RedirectStandardOutput = true;
            compiler.Start();
            compiler.WaitForExit(5000);

        }

        private void Ver(string sArchivo)
        {

            String sLeft;
            String sRigth;
            String sServer;
            String sPackage;
            Char delimiter = '.';
            String sRuta;

            var location = new Uri(Assembly.GetEntryAssembly().GetName().CodeBase);
            sRuta = new FileInfo(location.AbsolutePath).Directory.ToString() + "\\";

            Process compiler = new Process();
            compiler.StartInfo.FileName = "C:\\Program Files\\Notepad++\\notepad++.exe";
            compiler.StartInfo.Arguments = sArchivo;
            compiler.StartInfo.UseShellExecute = false;
            compiler.StartInfo.RedirectStandardOutput = true;
            compiler.Start();
            compiler.WaitForExit(5000);

        }



        private void Excel(string sArchivo)
        {

            String sLeft;
            String sRigth;
            String sServer;
            String sPackage;
            Char delimiter = '.';
            String sRuta;

            var location = new Uri(Assembly.GetEntryAssembly().GetName().CodeBase);
            sRuta = new FileInfo(location.AbsolutePath).Directory.ToString() + "\\";

            Process compiler = new Process();
            compiler.StartInfo.FileName = "C:\\Program Files (x86)\\Microsoft Office\\Office14\\excel.exe";
            compiler.StartInfo.Arguments = sArchivo;
            compiler.StartInfo.UseShellExecute = false;
            compiler.StartInfo.RedirectStandardOutput = true;
            compiler.Start();
            compiler.WaitForExit(5000);

        }

        private void Project(String sArchivo)
        {

            Process compiler = new Process();
            compiler.StartInfo.FileName = "C:\\Program Files\\ProjectLibre\\ProjectLibre.exe";
            compiler.StartInfo.Arguments = sArchivo;
            compiler.StartInfo.UseShellExecute = false;
            compiler.StartInfo.RedirectStandardOutput = true;
            compiler.Start();
            //compiler.WaitForExit(5000);
        }

        private void Soporte()
        {


            if (1 == 1)
            {
                // Read the files
                foreach (String file in ofdSoporte.FileNames)
                {
                    // Create a PictureBox.

                    dsComponentes(dgSoporte, "Soporte", cboSoporteHarvest.Text, "", file, FechaModificacion(file));

                }
            }
        }
        
        private string Compactar(DataGridView dg, string sRuta, string columna, string pwd)
        {
            using (ZipFile zip = new ZipFile())
            {
                if (pwd != "")
                {
                    zip.Password = pwd;
                    zip.Encryption = EncryptionAlgorithm.PkzipWeak;
                }

                foreach (DataGridViewRow fila in dg.Rows)
                {
                    if (fila.Visible==true & fila.Cells[columna].Value != null)
                    {
                        zip.AddFile(fila.Cells[columna].Value.ToString());
                    }
                }
                zip.Save(sRuta + "\\" + "Ejecutar.zip");
            }

            return sRuta + "\\" + "Ejecutar.zip";
        }

        private string CompactarSoporte(DataGridView dg, string Tipo, string sRuta, string columna, string pwd)
        {
            string sArchivo = "";
            using (ZipFile zip = new ZipFile())
            {
                if (pwd != "")
                {
                    zip.Password = pwd;
                    zip.Encryption = EncryptionAlgorithm.PkzipWeak;
                }

                foreach (DataGridViewRow fila in dg.Rows)
                {
                    if (fila.Visible == true & fila.Cells[columna].Value != null)
                    {
                        if (fila.Cells["Tipo"].Value.ToString() == Tipo)
                        {
                          sArchivo = fila.Cells[columna].Value.ToString();
                          zip.AddFile(sArchivo);
                        }
                    }
                }

                if (sArchivo!="")
                {
                    zip.Save(sRuta + "\\" + "CC-" + Tipo + ".zip");
                    sArchivo=sRuta + "\\" + "CC-" + Tipo + ".zip";
                }
                
            }

            return sArchivo;
        }

        private string Listado(DataGridView dg, string Tipo , string columna)
        {
            string sLista = "";
            string sValor = "";
            foreach (DataGridViewRow fila in dg.Rows)
            {
                if ( (fila.Cells[columna].Value != null) && (fila.Visible) )
                {
                    if (fila.Cells["tipo"].Value.ToString()==Tipo)
                    { 
                        sValor = fila.Cells[columna].Value.ToString();
                        if (sValor.IndexOf("\\") > -1)
                        {
                            sValor = sValor.Substring(sValor.LastIndexOf("\\") + 1);
                        }
                        sLista = sLista + sValor + Environment.NewLine;
                    }
                    
                }
            }

            return sLista;
        }

        private void Explorador(string sArchivo)
        {

            String sLeft;
            String sRigth;
            String sServer;
            String sPackage;
            Char delimiter = '.';
            String sRuta;

            var location = new Uri(Assembly.GetEntryAssembly().GetName().CodeBase);
            sRuta = new FileInfo(location.AbsolutePath).Directory.ToString() + "\\";

            Process compiler = new Process();
            compiler.StartInfo.FileName = "explorer.exe";
            compiler.StartInfo.Arguments = sArchivo;
            compiler.StartInfo.UseShellExecute = false;
            compiler.StartInfo.RedirectStandardOutput = true;
            compiler.Start();            

        }


        private DataTable EjecucionTotalizar(string sMes)
        {
            DataTable dtIncidencias = new DataTable();
            DataTable dtHarvest = new DataTable();
            decimal Horas;

            dtIncidencias = (DataTable)dgBitacora.DataSource;
            DataTable dtEjecucion = new DataTable();

            dtEjecucion.Columns.Add("Id Harvest");
            dtEjecucion.Columns.Add("Horas Consumidas");
            dtEjecucion.Columns.Add("Clasificacion");

            string Id = "", Hh = "", Es = "", sC = "";
            decimal Hs = 0;
            int Fila = 0;

            DataTable sortedDT = new DataTable();
            sortedDT = dtIncidencias.DefaultView.ToTable(true, new string[] { "Id Harvest" });

            foreach (DataRow Harvest in sortedDT.Rows)
            {
                Id = Harvest["Id Harvest"].ToString();
                Hs = 0;
                foreach (DataRow item in dtIncidencias.Select("[Id Harvest]='" + Harvest["Id Harvest"].ToString() + "'"))
                {
                    if (sMes != "")
                    {

                        if (item["Fecha"].ToString().Substring(3, 2) == sMes)
                        {
                            Hs = Hs + Val(item["Horas Consumidas"].ToString());
                        }
                    }
                    else
                    {
                        Hs = Hs + Val(item["Horas Consumidas"].ToString());
                    }
                }

                dtEjecucion.Rows.Add(dtEjecucion.NewRow());
                dtEjecucion.Rows[Fila]["Id Harvest"] = Id;
                dtEjecucion.Rows[Fila]["Horas Consumidas"] = Hs;
                dtEjecucion.Rows[Fila]["Clasificacion"] = Celda(dgPedidos, "Id Harvest", Id, "Clasificacion");
                Fila++;

            }
            dtEjecucion.AcceptChanges();

            bool bEncontro = false;
            dtHarvest = (DataTable)dgPedidos.DataSource;
            foreach (DataRow item in dtEjecucion.Rows)
            {
                Id = item["Id Harvest"].ToString().Replace(" ", "");
                Hs = Val(item["Horas Consumidas"].ToString());

                bEncontro = false;
                for (Fila = 0; Fila < dtHarvest.Rows.Count; Fila++)
                {
                    if (dtHarvest.Rows[Fila]["Id Harvest"].ToString().Replace(" ", "") == Id)
                    {
                        dtHarvest.Rows[Fila]["Horas Consumidas"] = Hs;
                        bEncontro = true;
                        break;
                    }
                }
                if (!(bEncontro))
                {
                    MessageBox.Show("El registro de " + Id + ", no se encontró en la Planificación");
                    dtHarvest.Rows.Add(dtHarvest.NewRow());
                    dtHarvest.Rows[dtHarvest.Rows.Count - 1]["Id Harvest"] = Id;
                }
            }

            return dtHarvest;

        }

        private void Harvest(ComboBox cb)
        {
            if (dgPedidos.DataSource != null)
            {
                DataTable dtPlanificacion = new DataTable();
                dtPlanificacion = (DataTable)dgPedidos.DataSource;

                cb.Items.Clear();
                foreach (DataRow item in dtPlanificacion.Rows)
                {
                    if ((item["Estado"]!=null) && item["Estado"].ToString() != "Cerrada")
                    {
                        cb.Items.Add(item["Id Harvest"]);
                    }
                }

            }
        }

        private void Proyectos(ComboBox cb)
        {
            cb.Items.Clear();
            String sObjeto="";
            String sProyecto="";

            if (dgProyectos.DataSource != null)
            {
                foreach (DataGridViewRow item in dgProyectos.Rows)
                {
                    if (item.Cells["RUTA DES"].Value != null && item.Cells["RUTA DES"].Value.ToString() != "")
                    { 
                    sObjeto=item.Cells["RUTA DES"].Value.ToString();
                    sProyecto= item.Cells["Nombre"].Value.ToString();
                    cb.Items.Add(sProyecto);
                    }
                }
            
            }
        }

        private void Tecnologia(DataGridView dg,  ComboBox cb , string sCol )
        {
            cb.Items.Clear();
            DataTable dtDistinct = new DataTable ();
            DataTable dtPEE = new DataTable ();

            if (dg.DataSource != null)
            {
                dtPEE = (DataTable)dg.DataSource;
                dtDistinct = dtPEE.DefaultView.ToTable(true, new string[] { sCol });

                foreach (DataRow item in dtDistinct.Rows)
                {
                    cb.Items.Add(item[sCol].ToString());                    
                }

            }
        }

        
        private String Configuracion(String sProyecto, String sValor)
        {
            String sDato = "";
            DataTable dtProyectos = new DataTable();
            dtProyectos = (DataTable)dgProyectos.DataSource;

            foreach (DataRow item in dtProyectos.Select("[nombre]='" + sProyecto + "'"))
            {
                sDato = item[sValor].ToString();
                break;
            }

            return sDato;
        }
        private string[] ArchivosFiltrados(string sRuta, string sFecha)
        {
            string[] sArchivos = new string[0];
            string[] sModifica = new string[0];

            int iLargo = 1;
            FileInfo fArchivos;


            sArchivos = Directory.GetFiles(sRuta, "*.*", SearchOption.AllDirectories);

            foreach (var item in sArchivos)
            {
                if (!(Filtro(dgArchivosExcluidos, item)))
                {
                    fArchivos = new FileInfo(item);
                    if (fArchivos.LastWriteTime > Convert.ToDateTime(sFecha))
                    {
                        Array.Resize(ref sModifica, iLargo);
                        sModifica[iLargo - 1] = item;
                        iLargo++;
                    }
                }
            }
            return sModifica;
        }
        private string[] Archivos(string sRuta)
        {
            string[] sArchivos = new string[0];
            sArchivos = Directory.GetFiles(sRuta, "*.*", SearchOption.AllDirectories);
            return sArchivos;
        }
        private void Tfs()
        {
            /*
            compiler.StartInfo.FileName = "C:\\Program Files (x86)\\WinMerge\\WinMergeU.exe";
            compiler.StartInfo.Arguments = sLeft + " " + sRigth;
            compiler.StartInfo.UseShellExecute = false;
            compiler.StartInfo.RedirectStandardOutput = true;
            compiler.Start();
            compiler.WaitForExit(5000);
             * */

        }
        private string ExcelDesa(string sHarvest, string sAmbiente, string sProyecto, string sRuta, string sAdjunto)
        {
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Range formatRange;
            object misValue = System.Reflection.Missing.Value;


            string sOrigen = "C:\\Pendientes\\PLANTILLAS\\MDS.xls";
            string sFecha = DateTime.Now.ToShortDateString();
            string sHora = "";
            string sFecNo = Fecha("yyyyMMdd");
            string sAPNom = "";
            string sAPCor = "";
            string sAPTel = "";

            string sJPNom = "";
            string sJPCor = "";
            string sJPTel = "";

            string sActividad = "";
            string sTipo = "";
            string sSitio = "";

            sAPNom = Celda(dgProyectos, "Nombre", sProyecto, "Desarrollador");
            sAPCor = Celda(dgEquipo, "Nombre", sAPNom, "Correo");
            sAPTel = Celda(dgEquipo, "Nombre", sAPNom, "Fono");

            sJPNom = Celda(dgProyectos, "Nombre", sProyecto, "Jefe Proyecto");
            sJPCor = Celda(dgEquipo, "Nombre", sJPNom, "Correo");
            sJPTel = Celda(dgEquipo, "Nombre", sJPNom, "Fono");
            sTipo = Celda(dgProyectos, "Nombre", sProyecto, "Tipo");
            sSitio = Celda(dgProyectos, "Nombre", sProyecto, "Url Sitio");
            if (sTipo == "ASPX")
            {
                sActividad = "Descomprimir contenido de archivo adjunto: Ejecutar.zip Pwd:explotacion , en el sitio " + sSitio + " , sobre-escribiendo los existentes";
            }
            else
            {
                sActividad = "Favor ejecutar bat adjunto " + NombreBAT(sProyecto,sAmbiente) + " en servidor de " + sAmbiente + ". Reemplazar '_' por 't' en la extensión del archivo adjunto.";
            }


            if (sAmbiente=="Produccion")
            {
                sHora = "20:00";
            }
            else
            {
                sHora = DateTime.Now.AddMinutes(30).ToShortTimeString();
            }
            

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(sOrigen, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            formatRange = xlWorkSheet.get_Range("B9");
            formatRange.NumberFormat = "@";

            formatRange = xlWorkSheet.get_Range("B30");
            formatRange.NumberFormat = "@";

            formatRange = xlWorkSheet.get_Range("D30");
            formatRange.NumberFormat = "@";

            formatRange = xlWorkSheet.get_Range("B43");
            formatRange.NumberFormat = "@";

            formatRange = xlWorkSheet.get_Range("D83");
            formatRange.NumberFormat = "@";

            xlWorkSheet.Range["B8"].Value = sAPNom;
            xlWorkSheet.Range["B9"].Value = sFecha;
            xlWorkSheet.Range["D9"].Value = sHarvest;
            xlWorkSheet.Range["D10"].Value = sJPNom;
            xlWorkSheet.Range["B13"].Value = sProyecto;
            xlWorkSheet.Range["B14"].Value = "Se actualiza versión para incluir los cambios solicitados.";
            xlWorkSheet.Range["B15"].Value = sAmbiente; //Equipos Involucrados
            xlWorkSheet.Range["B25"].Value = sAPNom;
            xlWorkSheet.Range["B29"].Value = ""; //Sistemas Afectados
            xlWorkSheet.Range["D29"].Value = 1;

            xlWorkSheet.Range["B30"].Value = sFecha;
            xlWorkSheet.Range["D30"].Value = sHora;
            xlWorkSheet.Range["B32"].Value = "01:00";

            xlWorkSheet.Range["A43"].Value = 1;
            xlWorkSheet.Range["B43"].Value = sFecha + " " + sHora;
            xlWorkSheet.Range["C43"].Value = sActividad;
            xlWorkSheet.Range["D43"].Value = sAPNom;
            xlWorkSheet.Range["E43"].Value = "00:00:05";


            xlWorkSheet.Range["B71"].Value = sJPNom;
            xlWorkSheet.Range["C71"].Value = sJPTel;
            xlWorkSheet.Range["D71"].Value = sJPCor;

            xlWorkSheet.Range["B72"].Value = sAPNom;
            xlWorkSheet.Range["C72"].Value = sAPTel;
            xlWorkSheet.Range["D72"].Value = sAPCor;

            xlWorkSheet.Range["C83"].Value = sJPNom;
            xlWorkSheet.Range["D83"].Value = sFecha;
            xlWorkSheet.Range["E83"].Value = "Aprobado";

            if (sAdjunto != "")
            {
                Microsoft.Office.Interop.Excel.OLEObjects oleObjects = (Microsoft.Office.Interop.Excel.OLEObjects)xlWorkSheet.OLEObjects(Type.Missing);
                oleObjects.Add(
                    Type.Missing,            // ClassType
                    sAdjunto,// Filename
                    false,                   // Link
                    true,                  // DisplayAsIcon
                    Type.Missing,           // IconFileName
                    Type.Missing,   // IconIndex
                    Type.Missing,   // IconLabel
                    10,   // Left
                    1300,   // Top
                    Type.Missing,   // Width
                    Type.Missing    // Height
                );
            }

            xlWorkBook.SaveAs(sRuta + "\\MDS_SOLPRO_" + sFecNo + "_" + sHarvest + "-" + sProyecto + ".xls");
            xlWorkBook.Close(true, misValue, misValue);

            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            Excel(sRuta + "\\MDS_SOLPRO_" + sFecNo + "_" + sHarvest + "-" + sProyecto + ".xls");

            return sRuta + "\\MDS_SOLPRO_" + sFecNo + "_" + sHarvest + "-" + sProyecto + ".xls";
        }

        private string ExcelSopo(string sHarvest, string sRuta, string sLiberacion, string sALiberacion, string sReversa, string sAReversa)
        {
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Range formatRange;
            object misValue = System.Reflection.Missing.Value;


            string sOrigen = "C:\\Pendientes\\PLANTILLAS\\CC.xls";
            string sFecha = DateTime.Now.ToShortDateString();
            string sHora = DateTime.Now.AddMinutes(30).ToShortTimeString();
            string sFecNo = Fecha("yyyyMMdd"); ;
            string sAPNom = "";
            string sAPCor = "";
            string sAPTel = "";

            string sJPNom = "";
            string sJPCor = "";
            string sJPTel = "";


            sAPNom = Celda(dgEquipo, "Tipo", "AP", "Nombre");
            sAPCor = Celda(dgEquipo, "Nombre", sAPNom, "Correo");
            sAPTel = Celda(dgEquipo, "Nombre", sAPNom, "Fono");

            sJPNom = Celda(dgPedidos, "Id Harvest", sHarvest, "Solicitante");
            sJPCor = Celda(dgEquipo, "Nombre", sJPNom, "Correo");
            sJPTel = Celda(dgEquipo, "Nombre", sJPNom, "Fono");

            sFecNo = DateTime.Now.Year.ToString("0000") + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00");

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(sOrigen, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            formatRange = xlWorkSheet.get_Range("B9");
            formatRange.NumberFormat = "@";

            formatRange = xlWorkSheet.get_Range("B30");
            formatRange.NumberFormat = "@";

            formatRange = xlWorkSheet.get_Range("D30");
            formatRange.NumberFormat = "@";

            formatRange = xlWorkSheet.get_Range("B43");
            formatRange.NumberFormat = "@";

            formatRange = xlWorkSheet.get_Range("D82");
            formatRange.NumberFormat = "@";

            xlWorkSheet.Range["B8"].Value = sAPNom;
            xlWorkSheet.Range["B9"].Value = sFecha;
            xlWorkSheet.Range["D9"].Value = sHarvest;
            xlWorkSheet.Range["D10"].Value = sJPNom;
            xlWorkSheet.Range["B13"].Value = "Se ejecuta script para resolver lo solicitado en el Harvest Asociado";
            xlWorkSheet.Range["B14"].Value = "Se genera CC para aplicar en producción según lo solicitado en Harvest asociado";
            xlWorkSheet.Range["B15"].Value = "SERVIDOR: ANDES – INSTANCIA: ISCTOS – ESQUEMA : AFIL"; //Equipos Involucrados
            xlWorkSheet.Range["B25"].Value = sAPNom;
            xlWorkSheet.Range["B29"].Value = ""; //Sistemas Afectados
            xlWorkSheet.Range["D29"].Value = 1;

            xlWorkSheet.Range["B30"].Value = sFecha;
            xlWorkSheet.Range["D30"].Value = sHora;
            xlWorkSheet.Range["B32"].Value = "01:00";

            xlWorkSheet.Range["A43"].Value = 1;
            xlWorkSheet.Range["B43"].Value = sFecha + " " + sHora;
            xlWorkSheet.Range["C43"].Value = "Favor extraer archivo adjunto ( CC-Liberacion.zip) , y ejecutar en el siguiente orden :" + sLiberacion;
            xlWorkSheet.Range["D43"].Value = sAPNom;
            xlWorkSheet.Range["E43"].Value = "00:05:00";


            xlWorkSheet.Range["B70"].Value = sJPNom;
            xlWorkSheet.Range["C70"].Value = sJPTel;
            xlWorkSheet.Range["D70"].Value = sJPCor;

            xlWorkSheet.Range["B71"].Value = sAPNom;
            xlWorkSheet.Range["C71"].Value = sAPTel;
            xlWorkSheet.Range["D71"].Value = sAPCor;

            xlWorkSheet.Range["C81"].Value = sJPNom;
            xlWorkSheet.Range["D81"].Value2 = sFecha;
            xlWorkSheet.Range["E81"].Value = "Aprobado";

            if (sALiberacion != "")
            {
                Microsoft.Office.Interop.Excel.OLEObjects oleObjects = (Microsoft.Office.Interop.Excel.OLEObjects)xlWorkSheet.OLEObjects(Type.Missing);
                oleObjects.Add(
                    Type.Missing,   // ClassType
                    sALiberacion,       // Filename
                    false,          // Link
                    true,           // DisplayAsIcon
                    Type.Missing,   // IconFileName
                    Type.Missing,   // IconIndex
                    Type.Missing,   // IconLabel
                    10,             // Left
                    1300,           // Top
                    Type.Missing,   // Width
                    Type.Missing    // Height
                );
            }


            if (sAReversa != "")
            {

                xlWorkSheet.Range["B65"].Value = "Favor extraer archivo adjunto ( CC-Reversa.zip) , y ejecutar en el siguiente orden :" + sReversa;
                xlWorkSheet.Range["D65"].Value = "00:05:00";
                Microsoft.Office.Interop.Excel.OLEObjects oleObjects = (Microsoft.Office.Interop.Excel.OLEObjects)xlWorkSheet.OLEObjects(Type.Missing);
                oleObjects.Add(
                    Type.Missing,   // ClassType
                    sAReversa,       // Filename
                    false,          // Link
                    true,           // DisplayAsIcon
                    Type.Missing,   // IconFileName
                    Type.Missing,   // IconIndex
                    Type.Missing,   // IconLabel
                    10,             // Left
                    1300,           // Top
                    Type.Missing,   // Width
                    Type.Missing    // Height
                );
            }

            xlWorkBook.SaveAs(sRuta + "\\CC_" + sHarvest + "_" + sFecNo + "-" + "Isapre_Consalud-Ejecución_Scripts" + ".xls");
            xlWorkBook.Close(true, misValue, misValue);

            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            Excel(sRuta + "\\CC_" + sHarvest + "_" + sFecNo + "-" + "Isapre_Consalud-Ejecución_Scripts" + ".xls");
            return sRuta + "\\CC_" + sHarvest + "_" + sFecNo + "-" + "Isapre_Consalud-Ejecución_Scripts" + ".xls";
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        private void InformeMensual(string sPeriodo)
        {

            DataTable dtClase = new DataTable();
            DataTable dtBitacora = new DataTable();

            int iFila = 1;
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet _xlSheet;
            Microsoft.Office.Interop.Excel.Range chartRangeUno;
            Microsoft.Office.Interop.Excel.Range chartRangeDos;
            Microsoft.Office.Interop.Excel.Range formatRange;

            object misValue = System.Reflection.Missing.Value;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add();
            _xlSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);

            dtBitacora = EjecucionTotalizar(sPeriodo);
            foreach (DataRow item in dtBitacora.Rows)
            {
                if (Val(item["Horas Consumidas"])>0)
                { 
                _xlSheet.Cells[iFila, 1] = item["Id Harvest"];
                _xlSheet.Cells[iFila, 2] = Val(item["Horas Consumidas"]);
                _xlSheet.Cells[iFila, 3] = item["Clasificacion"].ToString();
                iFila++;
                }
            }
            iFila--;
            chartRangeUno = _xlSheet.Range["Hoja2!$A$1:$B$" + iFila];

            iFila = 1;
            dtClase = SubTotal(dtBitacora, "Clasificacion", "Horas Consumidas");
            foreach (DataRow item in dtClase.Rows)
            {
                if (Val(item["Cantidad"]) > 0)
                {
                    _xlSheet.Cells[iFila, 5] = item["Clasificacion"];
                    _xlSheet.Cells[iFila, 6] = Val(item["Cantidad"]);
                    iFila++;
                }
            }
            iFila--;
            chartRangeDos = _xlSheet.Range["Hoja2!$E$1:$F$" + iFila];


            iFila = 1;
            dtBitacora = EjecucionBuscar(sPeriodo);
            dtClase = SubTotal(dtBitacora, "Fecha", "Horas Consumidas");
            DateTime dFecha;
            foreach (DataRow item in dtClase.Rows)
            {
                if (Val(item["Cantidad"]) > 0)
                {
                    dFecha = Convert.ToDateTime(item["Fecha"].ToString());
                    formatRange = _xlSheet.get_Range("H" + iFila);
                    formatRange.NumberFormat = "@";
                    _xlSheet.Cells[iFila, 8] = dFecha.ToString("ddddd, dd/MM/yyyy");
                    _xlSheet.Cells[iFila, 9] = Val(item["Cantidad"]) ;
                    iFila++;
                }
            }

            _xlSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Microsoft.Office.Interop.Excel.ChartObjects xlCharts = (Microsoft.Office.Interop.Excel.ChartObjects)_xlSheet.ChartObjects(Type.Missing);

            Microsoft.Office.Interop.Excel.ChartObject chartObjUno = (Microsoft.Office.Interop.Excel.ChartObject)xlCharts.Add(320, 10, 300, 300);
            Microsoft.Office.Interop.Excel.Chart chartuno = chartObjUno.Chart;

            Microsoft.Office.Interop.Excel.ChartObject chartObjDos = (Microsoft.Office.Interop.Excel.ChartObject)xlCharts.Add(10, 10, 300, 300);
            Microsoft.Office.Interop.Excel.Chart chartdos = chartObjDos.Chart;

            chartuno.SetSourceData(chartRangeUno, misValue);
            chartuno.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnStacked; // Many other types of charts can be speficied
            // It is not enough to set the title; you also have to tell it that it has a title first
            chartuno.HasTitle = true;
            chartuno.ChartTitle.Text = "Detalle Tareas";
            chartuno.ApplyDataLabels(Microsoft.Office.Interop.Excel.XlDataLabelsType.xlDataLabelsShowPercent, Microsoft.Office.Interop.Excel.XlDataLabelsType.xlDataLabelsShowLabel, true, false, false, true, true, true);
            chartuno.HasLegend = false;


            //OTRO GRAFICO?                      
            chartdos.SetSourceData(chartRangeDos, misValue);
            chartdos.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlPie; // Many other types of charts can be speficied
            // It is not enough to set the title; you also have to tell it that it has a title first
            chartdos.HasTitle = true;
            chartdos.ChartTitle.Text = "Resumen Mensual";
            chartdos.ApplyDataLabels(Microsoft.Office.Interop.Excel.XlDataLabelsType.xlDataLabelsShowPercent, Microsoft.Office.Interop.Excel.XlDataLabelsType.xlDataLabelsShowLabel, true, false, false, true, true, true);
            chartdos.HasLegend = false;

            xlWorkBook.SaveAs("c:\\bitacora-" + sPeriodo + ".xls");
            xlWorkBook.Close(true, misValue, misValue);

            xlApp.Quit();

            releaseObject(_xlSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            Excel("c:\\bitacora-" + sPeriodo + ".xls");
        }
        private DataTable SubTotal(DataTable dtDatos, string sColumna, string sValor)
        {
            DataTable dtResult = new DataTable();
            DataTable dtDistinct = new DataTable();

            dtResult.Columns.Add(sColumna);
            dtResult.Columns.Add("Cantidad");
            dtDatos.DefaultView.Sort = sColumna;
            dtDistinct = dtDatos.DefaultView.ToTable(true, new string[] { sColumna });
            
            decimal Hs = 0;
            string Cl = "";

            foreach (DataRow clase in dtDistinct.Rows)
            {
                Hs = 0;
                foreach (DataRow item in dtDatos.Select("[" + sColumna + "]='" + clase[sColumna] + "'"))
                {
                    if (item[sValor].ToString() != "")
                    {
                        Hs = Hs + Convert.ToDecimal(item[sValor].ToString().Replace(".", ","));
                    }
                    Cl = clase[sColumna].ToString();
                }
                dtResult.Rows.Add(dtResult.NewRow());
                dtResult.Rows[(dtResult.Rows.Count - 1)][sColumna] = Cl;
                dtResult.Rows[(dtResult.Rows.Count - 1)]["Cantidad"] = Hs;
            }

            return dtResult;

        }
        private DataTable ExcelLoad(string sRuta)
        {
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            DataTable dtBitacora = new DataTable();


            string sOrigen = sRuta;
            string sFecha = DateTime.Now.ToShortDateString();
            string sFecNo = "";
            string sDato = "";
            int iFila = 1;

            sFecNo = DateTime.Now.Year.ToString("0000") + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00");

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(sOrigen, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            dtBitacora.Columns.Add(xlWorkSheet.Range["A1"].Value);
            dtBitacora.Columns.Add(xlWorkSheet.Range["B1"].Value);
            dtBitacora.Columns.Add(xlWorkSheet.Range["C1"].Value);
            dtBitacora.Columns.Add(xlWorkSheet.Range["D1"].Value);
            dtBitacora.Columns.Add(xlWorkSheet.Range["E1"].Value);
            dtBitacora.Columns.Add(xlWorkSheet.Range["F1"].Value);
            dtBitacora.Columns.Add(xlWorkSheet.Range["G1"].Value);
            dtBitacora.Columns.Add(xlWorkSheet.Range["H1"].Value);
            dtBitacora.Columns.Add(xlWorkSheet.Range["I1"].Value);
            dtBitacora.AcceptChanges();

            iFila++;
            sDato = xlWorkSheet.Range["A1"].Value;
            while (sDato != "")
            {
                dtBitacora.Rows.Add(dtBitacora.NewRow());
                dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["A1"].Value] = xlWorkSheet.Range["A" + iFila].Text;
                dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["B1"].Value] = xlWorkSheet.Range["A" + iFila].Text;
                dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["C1"].Value] = xlWorkSheet.Range["A" + iFila].Text;
                dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["D1"].Value] = xlWorkSheet.Range["A" + iFila].Text;
                dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["E1"].Value] = xlWorkSheet.Range["A" + iFila].Text;
                dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["F1"].Value] = xlWorkSheet.Range["A" + iFila].Text;
                dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["G1"].Value] = xlWorkSheet.Range["A" + iFila].Text;
                dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["H1"].Value] = xlWorkSheet.Range["A" + iFila].Text;
                dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["I1"].Value] = xlWorkSheet.Range["A" + iFila].Text;
                iFila++;
                sDato = xlWorkSheet.Range["A" + iFila].Text;
            }
            dtBitacora.AcceptChanges();
            dgBitacora.DataSource = dtBitacora;

            return dtBitacora;

        }
        private DataTable ExcelHarvest(string sRuta)
        {

            DataTable dtBitacora = new DataTable();

            if (MessageBox.Show("¿Desea cargar la Planilla Harvest?", "Bitacora", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Microsoft.Office.Interop.Excel.Application xlApp;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                string sOrigen = sRuta;
                string sFecha = DateTime.Now.ToShortDateString();
                string sFecNo = "";
                string sDato = "";
                int iFila = 1;

                sFecNo = DateTime.Now.Year.ToString("0000") + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00");

                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(sOrigen, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                dtBitacora.Columns.Add(xlWorkSheet.Range["A1"].Value);
                dtBitacora.Columns.Add(xlWorkSheet.Range["B1"].Value);
                dtBitacora.Columns.Add(xlWorkSheet.Range["C1"].Value);
                dtBitacora.Columns.Add(xlWorkSheet.Range["D1"].Value);
                dtBitacora.Columns.Add(xlWorkSheet.Range["E1"].Value);
                dtBitacora.Columns.Add(xlWorkSheet.Range["F1"].Value);
                dtBitacora.Columns.Add(xlWorkSheet.Range["G1"].Value);
                dtBitacora.Columns.Add(xlWorkSheet.Range["H1"].Value);
                dtBitacora.Columns.Add(xlWorkSheet.Range["I1"].Value);
                dtBitacora.Columns.Add(xlWorkSheet.Range["J1"].Value);
                dtBitacora.Columns.Add(xlWorkSheet.Range["K1"].Value);
                dtBitacora.Columns.Add(xlWorkSheet.Range["L1"].Value);
                dtBitacora.Columns.Add(xlWorkSheet.Range["M1"].Value);
                dtBitacora.Columns.Add(xlWorkSheet.Range["N1"].Value);
                dtBitacora.Columns.Add(xlWorkSheet.Range["O1"].Value);
                //dtBitacora.Columns.Add("HRS. CONSUMIDAS").DefaultValue=0;
                dtBitacora.Columns.Add(xlWorkSheet.Range["P1"].Value);
                dtBitacora.Columns.Add(xlWorkSheet.Range["Q1"].Value);
                dtBitacora.Columns.Add(xlWorkSheet.Range["R1"].Value);
                dtBitacora.Columns.Add(xlWorkSheet.Range["S1"].Value);
                dtBitacora.Columns.Add(xlWorkSheet.Range["T1"].Value);
                dtBitacora.Columns.Add(xlWorkSheet.Range["U1"].Value);
                dtBitacora.Columns.Add(xlWorkSheet.Range["V1"].Value);
                dtBitacora.Columns.Add(xlWorkSheet.Range["W1"].Value);
                dtBitacora.Columns.Add(xlWorkSheet.Range["X1"].Value);
                dtBitacora.Columns.Add(xlWorkSheet.Range["Y1"].Value);
                dtBitacora.Columns.Add(xlWorkSheet.Range["Z1"].Value);
                dtBitacora.AcceptChanges();

                iFila++;
                sDato = xlWorkSheet.Range["F1"].Value;
                while (sDato != "")
                {
                    dtBitacora.Rows.Add(dtBitacora.NewRow());
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["A1"].Value] = xlWorkSheet.Range["A" + iFila].Text;
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["B1"].Value] = xlWorkSheet.Range["B" + iFila].Text;
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["C1"].Value] = xlWorkSheet.Range["C" + iFila].Text;
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["D1"].Value] = xlWorkSheet.Range["D" + iFila].Text;
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["E1"].Value] = xlWorkSheet.Range["E" + iFila].Text;
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["F1"].Value] = xlWorkSheet.Range["F" + iFila].Text.Replace(" ", "");
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["G1"].Value] = xlWorkSheet.Range["G" + iFila].Text;
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["H1"].Value] = xlWorkSheet.Range["H" + iFila].Text;
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["I1"].Value] = xlWorkSheet.Range["I" + iFila].Text;
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["J1"].Value] = xlWorkSheet.Range["J" + iFila].Text;
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["K1"].Value] = xlWorkSheet.Range["K" + iFila].Text;
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["L1"].Value] = xlWorkSheet.Range["L" + iFila].Text;
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["M1"].Value] = xlWorkSheet.Range["M" + iFila].Text;
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["N1"].Value] = xlWorkSheet.Range["N" + iFila].Text;
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["O1"].Value] = xlWorkSheet.Range["O" + iFila].Text;
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["P1"].Value] = xlWorkSheet.Range["P" + iFila].Text;
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["Q1"].Value] = xlWorkSheet.Range["Q" + iFila].Text;
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["R1"].Value] = xlWorkSheet.Range["R" + iFila].Text;
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["S1"].Value] = xlWorkSheet.Range["S" + iFila].Text;
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["T1"].Value] = xlWorkSheet.Range["T" + iFila].Text;
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["U1"].Value] = xlWorkSheet.Range["U" + iFila].Text;
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["V1"].Value] = xlWorkSheet.Range["V" + iFila].Text;
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["W1"].Value] = xlWorkSheet.Range["W" + iFila].Text;
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["X1"].Value] = xlWorkSheet.Range["X" + iFila].Text;
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["Y1"].Value] = xlWorkSheet.Range["Y" + iFila].Text;
                    dtBitacora.Rows[(dtBitacora.Rows.Count - 1)][xlWorkSheet.Range["Z1"].Value] = xlWorkSheet.Range["Z" + iFila].Text;
                    
                    iFila++;
                    sDato = xlWorkSheet.Range["F" + iFila].Text;
                }

                dtBitacora.AcceptChanges();
                dgHarvest.DataSource = dtBitacora;
            }
            else
            {
                Transacciones(dgHarvest, "Harvest", "HAR");
                dtBitacora = (DataTable)dgHarvest.DataSource;
            }

            Persistencia(dgHarvest, "Harvest");
            return dtBitacora;
        }
        private string ArchivoBAT(string sProyecto, string sRuta)
        {
            string sNombre = "";
            string sSitio = "";
            string sArchivo = "";
            string sCFG;
            string sTipo;
            string sAmbiente = cbAmbienteAplicativo.Text;
            DataTable dtBat = new DataTable();

            dtBat = (DataTable)dgBatLiberacion.DataSource;
            sSitio = Celda(dgProyectos, "Nombre", sProyecto, "Nombre Sitio");
            sTipo = Celda(dgProyectos, "Nombre", sProyecto, "Tipo");
            sNombre = NombreBAT(sProyecto, sAmbiente);

            foreach (DataRow item in dtBat.Rows)
            {
                if (item["Tipo"].ToString() == sTipo && item["Ambiente"].ToString() ==sAmbiente )
                {
                    sCFG = item["Objeto"].ToString();
                    sCFG = sCFG.Replace("[]", sSitio);
                    sArchivo = sArchivo + sCFG + Environment.NewLine;
                }
            }

            sRuta = sRuta + "\\" + sNombre;
            File.WriteAllText(sRuta, sArchivo, Encoding.GetEncoding(1252));
            return sRuta;

        }
        private string Mes()
        {
            return (cbBitacoraPeriodo.SelectedIndex + 1).ToString("00");
        }
        private DataTable EjecucionBuscar(string sMes)
        {
            DataTable dtIncidencias = new DataTable();
            DataTable dtHarvest = new DataTable();

            dtIncidencias = (DataTable)dgBitacora.DataSource;
            DataTable dtEjecucion = dtIncidencias.Clone();

            foreach (DataRow Harvest in dtIncidencias.Rows)
            {
                if (sMes != "")
                {
                    if (Harvest["Fecha"].ToString().Substring(3, 2) == sMes)
                    {
                        dtEjecucion.ImportRow(Harvest);
                    }
                }
            }

            dtEjecucion.AcceptChanges();

            return dtEjecucion;

        }
        private void AgregarArchivos(DataGridView dg, string sHarvest, string sArchivo)
        {
            // Read the files
            foreach (String file in ofdSoporte.FileNames)
            {
                dsComponentes(dg, sArchivo, sHarvest, "Liberacion", file, FechaModificacion(file));
            }

        }
        private void AgregarAnalisis(DataGridView dg, string sArchivo)
        {
            string sCompo;
            String sPromedio="";
            String sCodigo="";
            String sGlosa="";
            Decimal sHoras=0;
            Char delimiter = ':';

            sCompo = cbAnalisisComponente.Text;
            if (sCompo!="")
            {            
            sCodigo = sCompo.Split(delimiter)[0];
            sGlosa = sCompo.Split(delimiter)[1];
            }

            sPromedio = Celda(dgParametrosPEE, "CODIGO", sCodigo, "SUGERIDO");
            dsRequerimiento(dg, sArchivo, cbAnalisisHarvest.Text, cbAnalisisProyecto.Text, sCodigo, sGlosa, sPromedio);                      

            //CalcularAltovuelo();

        }

        private Boolean AnalizarEnsamblados()
        {
            string sTemp;
            string sArchivo;
            string sRuta;

            sArchivo = CeldaSeleccionada(dgApp, "Objeto");
            sTemp = sArchivo.Substring(sArchivo.LastIndexOf("\\") + 1);
            sRuta = "C:\\Pendientes\\TOOLS\\";

            try
            {
                Process compiler = new Process();

                compiler.StartInfo.FileName = sRuta + "ildasm.exe";
                compiler.StartInfo.Arguments = "/SOURCE  /VISIBILITY=PUB+PRI " + sArchivo + " /out=" + sRuta + sTemp + ".il";
                compiler.StartInfo.UseShellExecute = false;
                compiler.StartInfo.RedirectStandardOutput = true;
                compiler.Start();
                compiler.WaitForExit(5000);

            }
            catch (Exception ex)
            {

                sArchivo = "";
            }

            String sDependencias = ""; ;
            Char delimiter = '"';

            dgDependencias.DataSource = null;
            string[] lines = System.IO.File.ReadAllLines(sRuta + sTemp + ".il");
            foreach (string line in lines)
            {
                if (line.IndexOf("PCK") > -1)
                {
                    sTemp = line.Substring(line.IndexOf(delimiter));
                    dsConfiguracion(dgDependencias, "Dependencias", "", sTemp);
                }

                if ((line.ToUpper().IndexOf("CLASS PUBLIC") > -1) & (line.ToUpper().IndexOf("SVC") > -1))
                {
                    sTemp = line.Substring(line.LastIndexOf(" "));
                    sTemp = sTemp.Substring(1, sTemp.LastIndexOf(".") - 1);
                    if (sDependencias != sTemp)
                    {
                        dsConfiguracion(dgDependencias, "Dependencias", "", sTemp);
                    }
                    sDependencias = sTemp;
                }
            }


            return true;
        }

        private void CargarHarvest()
        {            
            Harvest(cbDesarrolloHarvest);
            Harvest(cboSoporteHarvest);
            Harvest(cboCorreoHarvest);
            //Harvest(cboCorreoHarvest);
            Harvest(cbAnalisisHarvest);
            Harvest(cbBaseHarvest);
            Harvest(cbPlanificacionHarvest);

            cbPlanificacionHarvest.Items.Insert(0, "(Todos)");

        }

        private string NombreBAT(string sProyecto, string sAmbiente)
        {
            string sTipo;
            string sNombre;
            string sSitio;

            sTipo = Celda(dgProyectos, "Nombre", sProyecto, "Tipo");
            sSitio = Celda(dgProyectos, "Nombre", sProyecto, "Nombre Sitio");

            if (sAmbiente == "Produccion")
            {
                sNombre = "Traspaso_" + sSitio + "_NET.ba_";
            }
            else
            {
                sNombre = "Traspaso_" + sSitio + "_TES.ba_";
            }

            return sNombre;
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            Respaldar(dgProyectos);
        }

        private string Respaldar(DataGridView dg)
        {
            string sFuente="";
            string sServidor="";
            string sProyecto="";
            string sRuta="";

            string[] sArchivos = new string[0];


            foreach (DataGridViewRow fila in dg.Rows)
            {
                if (fila.Cells["Nombre"].Value != null)
                {
                    using (ZipFile zip = new ZipFile())
                    {
                        sFuente = fila.Cells["RUTA FUE"].Value.ToString();
                        sServidor = fila.Cells["RUTA RES"].Value.ToString();
                        sProyecto = fila.Cells["Nombre"].Value.ToString();

                        sRuta=sServidor + "\\" + sProyecto ;
                        if (sServidor != "" && sFuente != "")
                        {
                            sArchivos = Archivos(sFuente);
                            foreach (var item in sArchivos)
                            {
                                zip.AddFile(item);
                            }

                            Directory.CreateDirectory((sServidor + "\\").Substring(1, (sServidor + "\\").LastIndexOf("\\") ));

                            zip.Save(sRuta+ ".zip");

                        }
                    }

                }
            }

            return sServidor + "\\" + sProyecto + ".zip";
        }

        private string Respaldar(string sProyecto , string sHacia)
        {
            string sFuente = "";
            string sServidor = "";
            string[] sArchivos = new string[0];
            
            sFuente = Celda(dgProyectos, "Nombre", sProyecto, "RUTA FUE");
            sServidor = Celda(dgProyectos, "Nombre", sProyecto, "RUTA RES");
                        
            if (sServidor != "" && sFuente != "")
            {
                using (ZipFile zip = new ZipFile())
                {
                    sArchivos = Archivos(sFuente);
                    foreach (var item in sArchivos)
                    {zip.AddFile(item);}
                    zip.Save(sHacia + "\\" + sProyecto + ".zip");
                 }                    
            }

            return sHacia + "\\" + sProyecto + ".zip";
        }

        private void CalcularAltovuelo()
        {
            Decimal dHoras = 0, dHora = 0;
            int iComponentes=0, iElementos = 0;
            string sTipo = "", sObservacion = "", sPesi = "", sOpti = "", sProyecto="";


            Decimal Total = 0;
            Decimal Factor;
            DataTable dtEtapas = new DataTable();
            //dtEtapas = ((DataTable)dgParametrosEtapas.DataSource).Copy();
            dtEtapas.Columns.Add("Resumen");
            dtEtapas.Columns.Add("Cantidad");
            sProyecto = cbAnalisisProyecto.Text;

            for (int fila = 0; fila < dgRequerimientos.Rows.Count; fila++)
            {                
                if (dgRequerimientos.Rows[fila].Cells["Proyecto"].Value == null)
                {
                    break;
                }

                if (dgRequerimientos.Rows[fila].Visible)
                {

                    iElementos++;
                    dgRequerimientos.Rows[fila].Cells["Componente"].Value = iElementos;

                    if (dgRequerimientos.Rows[fila].Cells["Proyecto"].Value.ToString()=="")
                    {
                        dgRequerimientos.Rows[fila].Cells["Proyecto"].Value = sProyecto;
                    }

                    dHora = Val(dgRequerimientos.Rows[fila].Cells["Horas"].Value);
                    sTipo = dgRequerimientos.Rows[fila].Cells["Tipo Componente"].Value.ToString().Trim();

                    if (sTipo!="")
                    {
                        sPesi = Celda(dgParametrosPEE, "Codigo", sTipo, "Pesimista");
                        sOpti = Celda(dgParametrosPEE, "Codigo", sTipo, "Optimista");    
                 

                        //if ((dHora == 0))
                        //{
                            dgRequerimientos.Rows[fila].Cells["Horas"].Value = sPesi;
                            dHora = Val(sPesi);
                        //}
                    

                        if (dHora > 4)
                        {
                            sObservacion = "Existen Componentes con mas de 4 HH";
                        }

                        if (dHora > Val(sPesi) || dHora < Val(sOpti))
                        {
                            dtEtapas.Rows.Add(dtEtapas.NewRow());
                            dtEtapas.Rows[dtEtapas.Rows.Count - 1]["Resumen"] = "Componente fuera de Rango:" + sOpti + "-" + sPesi;
                            dtEtapas.Rows[dtEtapas.Rows.Count - 1]["Cantidad"] = iElementos;
                        }                        

                        iComponentes++;
                        dHoras = dHoras + dHora;
    
                    }                                
                    else
                    {
                        dgRequerimientos.Rows[fila].Cells["Horas"].Value = "";
                    }
                }                
            }            

            dtEtapas.Rows.Add(dtEtapas.NewRow());
            dtEtapas.Rows[dtEtapas.Rows.Count - 1]["Resumen"] = "Total de HH";
            dtEtapas.Rows[dtEtapas.Rows.Count - 1]["Cantidad"] = dHoras;

            dtEtapas.Rows.Add(dtEtapas.NewRow());
            dtEtapas.Rows[dtEtapas.Rows.Count - 1]["Resumen"] = "Total de Componentes";
            dtEtapas.Rows[dtEtapas.Rows.Count - 1]["Cantidad"] = iComponentes;


            if (sObservacion!="")
            {
                dtEtapas.Rows.Add(dtEtapas.NewRow());
                dtEtapas.Rows[dtEtapas.Rows.Count - 1]["Resumen"] = "Errores";
                dtEtapas.Rows[dtEtapas.Rows.Count - 1]["Cantidad"] = sObservacion;
            }

            //Factor = dHoras * 100;
            //Factor = Factor / Val(Celda(dgParametrosEtapas, "Etapa", "Desarrollo", "Factor"));
            //for (int i = 0; i < dtEtapas.Rows.Count; i++)
            //{
            //    dtEtapas.Rows[i]["Sub Total"] = (Factor * Val(dtEtapas.Rows[i]["Factor"].ToString())) / 100;
            //    Total = Total + (Factor * Val(dtEtapas.Rows[i]["Factor"].ToString())) / 100;
            //}

            //dtEtapas.Rows.Add(dtEtapas.NewRow());
            //dtEtapas.Rows[dtEtapas.Rows.Count - 1]["Sub Total"] = Total;
            //dtEtapas.Rows[dtEtapas.Rows.Count - 1]["Factor"] = "TOTAL";
            dtEtapas.AcceptChanges();
            dgAnalisisTotal.DataSource = dtEtapas;

        }

        //private void CalcularAltovuelo()
        //{
        //    Decimal dHoras=0;
        //    for (int fila = 0; fila < dgRequerimientos.Rows.Count; fila++)
        //    {
        //        if (dgRequerimientos.Rows[fila].Cells["Horas"].Value != null)
        //        {
        //            dHoras = dHoras + Val(dgRequerimientos.Rows[fila].Cells["Horas"].Value);
        //        }
        //    }

        //    Decimal Total=0;
        //    Decimal Factor;
        //    DataTable dtEtapas = new DataTable();
        //    dtEtapas=((DataTable)dgParametrosEtapas.DataSource).Copy() ;
        //    dtEtapas.Columns.Add("Sub Total");

        //    Factor = dHoras * 100;
        //    Factor = Factor / Val(Celda(dgParametrosEtapas, "Etapa", "Desarrollo", "Factor"));
        //    for (int i = 0; i < dtEtapas.Rows.Count ; i++)
        //    {   
        //        dtEtapas.Rows[i]["Sub Total"] = (Factor * Val(dtEtapas.Rows[i]["Factor"].ToString()))/100;
        //        Total = Total + (Factor * Val(dtEtapas.Rows[i]["Factor"].ToString())) / 100;
        //    }

        //    dtEtapas.Rows.Add(dtEtapas.NewRow());
        //    dtEtapas.Rows[dtEtapas.Rows.Count-1]["Sub Total"] = Total;
        //    dtEtapas.Rows[dtEtapas.Rows.Count - 1]["Factor"] = "TOTAL";
        //    dtEtapas.AcceptChanges();
        //    dgAnalisisTotal.DataSource = dtEtapas;

        //}

        private void Filtrar(DataGridView dg, string Harvest)
        {
            String mess;
            String mesc;
            //int fila;
            int filas;
            DataGridViewRow dr;

            dg.CurrentCell = null;
            filas = dg.Rows.Count;

            for (int fila = 0; fila < filas; fila++)
            {
                dr = dg.Rows[fila];
                if (dr.Cells["Id Harvest"].Value != null)
                {
                    mesc = dr.Cells["Id Harvest"].Value.ToString();
                    if ((mesc.IndexOf(Harvest) > -1) || (Harvest=="(Todos)"))
                    {
                        dr.Visible = true;
                    }
                    else
                    {
                        dr.Visible = false;
                    }
                }
            }
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void dgCorreosFiltro_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void cbAnalisisTecno_SelectedIndexChanged(object sender, EventArgs e)
        {
            Char delimiter = ':';
            string sTecno;
            string sRango;
            sTecno = cbAnalisisTecno.Text;
            //sTecno = sTecno.Split(delimiter)[0].Trim();

            cbAnalisisComponente.Items.Clear();

            foreach (DataGridViewRow item in dgParametrosPEE.Rows)
            {
                if (item.Cells["TIPO"].Value != null)
                {
                    if (item.Cells["TIPO"].Value.ToString() == sTecno)
                    {
                        sRango = "[" + item.Cells["OPTIMISTA"].Value.ToString() + "/" + item.Cells["PESIMISTA"].Value.ToString() + "]";
                        cbAnalisisComponente.Items.Add(item.Cells["CODIGO"].Value.ToString() + ":" + item.Cells["DESCRIPCION"].Value.ToString() + sRango);
                    }

                }
            }
        }

        private void btnAnalisisAgregar_Click(object sender, EventArgs e)
        {
            AgregarAnalisis(dgRequerimientos, "Soporte");            
            CalcularAltovuelo();

            Filtrar(dgRequerimientos, cbAnalisisHarvest.Text);
            Filtrar(dgAnalisisDocs, cbAnalisisHarvest.Text);
        }

        private void bntAnalisisGuardar_Click(object sender, EventArgs e)
        {
            Persistencia(dgRequerimientos, "Requerimientos");
            Persistencia(dgAnalisisDocs, "Documentos");
        }

        private Decimal Val(object oTexto)
        {
            Decimal salida = 0;
            try
            {
                
                if (!(Decimal.TryParse(oTexto.ToString().Replace(".", ","), out salida)))
                {
                    salida = 0;
                }
            }
            catch (Exception)
            {

                salida = 0;
            }
            

            return salida;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            String sHarv;
            String sSoli;
            String sJP;

            sHarv = CeldaSeleccionada(dgHarvest, "Id Solicitud");
            sSoli = CeldaSeleccionada(dgHarvest, "Descripción");
            sJP = CeldaSeleccionada(dgHarvest, "JEFE PROYECTO");

            sHarv = sHarv.Replace(" ", "");
            dsPlanificacion(dgPedidos, "Planificacion", sHarv, sSoli, sJP);
        }

        private void dgPlanificacion_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void cbAnalisisHarvest_SelectedIndexChanged(object sender, EventArgs e)
        {
            String sDescripcion;

            sDescripcion = Celda(dgPedidos, "Id Harvest", cbAnalisisHarvest.Text, "Solicitud");
            txtAnalisisAltoVuelo.Text = sDescripcion;
            txtAnalisisRuta.Text = "C:\\Analisis\\" + cbAnalisisHarvest.Text;

            Filtrar(dgRequerimientos, cbAnalisisHarvest.Text);
            Filtrar(dgAnalisisDocs, cbAnalisisHarvest.Text);
            
        }

        private void button11_Click(object sender, EventArgs e)
        {
            Char delimiter = '.';
            String sFecha;
            String sArchivo;
            String sLeft;
            String sRigth;
            String sServer;
            String sPackage;
            String sRuta;

            sRuta = "C:\\Pendientes\\Liberaciones\\";
            sArchivo = CeldaSeleccionada(dgObjetos, "Objeto");
            sServer = sArchivo.Split(delimiter)[0];
            sPackage = sArchivo.Split(delimiter)[1];
            sRigth = sRuta + BasePackage(sServer, sPackage, "C:\\Pendientes\\Liberaciones\\");

            sFecha = CeldaSeleccionada(dgObjetos, "Fecha");
            sArchivo = CeldaSeleccionada(dgObjetos, "Objeto");
            sLeft = sRuta + sArchivo + "." + FechaArchivo(sFecha) + ".sql";
            Merger(sLeft, sRigth);
        }

        private void FiltrarBitacora(string Mes)
        {
            String ejer, ejerc;
            String mesc;
            //int fila;
            int filas;
            DataGridViewRow dr;

            dgBitacora.CurrentCell = null;
            ejer = txtEjercicio.Text;
            filas = dgBitacora.Rows.Count;

            for (int fila = 0; fila < filas; fila++)
            {
                dr = dgBitacora.Rows[fila];
                if ( (dr.Cells["fecha"].Value != null) && (dr.Cells["fecha"].Value != "") )
                {
                    mesc = dr.Cells["fecha"].Value.ToString().Substring(3, 2);
                    ejerc = dr.Cells["fecha"].Value.ToString().Substring(6, 4);
                    if ((Convert.ToInt16(Mes) == Convert.ToInt16(mesc)) && (Convert.ToInt16(ejer) == Convert.ToInt16(ejerc)))
                    {
                        dr.Visible = true;
                    }
                    else
                    {
                        dr.Visible = false;
                    }
                }

            }

        }

        private void LiberarArchivos(string sHarvest)
        {
            String sDesde = "";
            String sHasta = "";
            String sObjeto = "";
            DialogResult rRespuesta;           
            sDesde = Celda(dgProyectos, "Nombre", cbDesarrolloProyecto.Text, "RUTA LOC");
            sHasta = Celda(dgProyectos, "Nombre", cbDesarrolloProyecto.Text, "RUTA DES");

            rRespuesta = MessageBox.Show("¿Desea Liberar la Versión al Ambiente de Desarrollo?" + Environment.NewLine + sDesde + Environment.NewLine + sHasta, "Liberar", MessageBoxButtons.YesNo);
            if (rRespuesta == DialogResult.Yes)
            {

            foreach (DataGridViewRow item in dgApp.Rows)
            {
                if ( item.Visible == true  && item.Cells["Objeto"].Value != null && item.Cells["Objeto"].Value.ToString() != "")
                {
                    sObjeto = item.Cells["Objeto"].Value.ToString();
                    sObjeto = sObjeto.Replace(sDesde, "");
                    try
                    {
                        File.Copy(sDesde + sObjeto, sHasta + sObjeto, true);
                    }
                    catch (Exception)
                    {

                        sObjeto ="";
                    }

                    
                }
            }
            }
        }

        private void cbDesarrolloProyecto_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cbDesarrolloHarvest_SelectedIndexChanged(object sender, EventArgs e)
        {
            {
                String mess;
                String mesc;
                //int fila;
                int filas;
                DataGridViewRow dr;

                dgApp.CurrentCell = null;
                mess = (cbDesarrolloHarvest.Text).ToString();
                filas = dgApp.Rows.Count;

                for (int fila = 0; fila < filas; fila++)
                {
                    dr = dgApp.Rows[fila];
                    if (dr.Cells["Id Harvest"].Value != null)
                    {
                        mesc = dr.Cells["Id Harvest"].Value.ToString();
                        if (mess == mesc)
                        {
                            dr.Visible = true;
                        }
                        else
                        {
                            dr.Visible = false;
                        }
                    }

                }
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            CalcularAltovuelo();
        }

        private void tabPage9_Click(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {
            String sAntig = CeldaSeleccionada(dgSoporte, "objeto");
            String sNuevo = Microsoft.VisualBasic.Interaction.InputBox("Renombrar", "Ingrese el nuevo nombre", sAntig, 200, 200);
            if (sNuevo !="")
            {
                File.Move(sAntig, sNuevo);
                CeldaSeleccionada(dgSoporte, "Objeto", sNuevo);
                Persistencia(dgSoporte, "Soporte");
            }
            
        }

        private void button14_Click(object sender, EventArgs e)
        {
            Char delimiter = '.';
            String sFecha;
            String sArchivo;
            String sLeft;
            String sRigth;
            String sServer;
            String sPackage;
            String sRuta;

            sRuta = "C:\\Pendientes\\Liberaciones\\";
            
            sArchivo = CeldaSeleccionada(dgObjetos, "Objeto");
            sServer = sArchivo.Split(delimiter)[0];
            sPackage = sArchivo.Split(delimiter)[1];
            sRigth = sRuta + BasePackage(sServer, sPackage, "C:\\Pendientes\\Liberaciones\\");
            sLeft = sRuta + BasePackage("QISCTOS", sPackage, "C:\\Pendientes\\Liberaciones\\");

            File.Copy( sRigth, txtSoporteRuta.Text + "\\PROD-" + sPackage + ".sql",true);
            File.Copy(sLeft, txtSoporteRuta.Text + "\\ROLL-" + sPackage + ".sql", true);

            sArchivo = txtSoporteRuta.Text + "\\PROD-" + sPackage + ".sql";
            dsComponentes(dgSoporte, "Soporte", cboSoporteHarvest.Text, "Liberacion", sArchivo, FechaModificacion(sArchivo));

            sArchivo = txtSoporteRuta.Text + "\\ROLL-" + sPackage + ".sql";
            dsComponentes(dgSoporte, "Soporte", cboSoporteHarvest.Text, "Reversa", sArchivo, FechaModificacion(sArchivo));
            Persistencia(dgSoporte, "Soporte");            
        }

        private void button15_Click(object sender, EventArgs e)
        {
            SqlPlus(CeldaSeleccionada(dgSoporte, "Objeto"));
        }

        private void button16_Click(object sender, EventArgs e)
        {
            fbdSoporte.SelectedPath = txtAnalisisRuta.Text;
            fbdSoporte.ShowDialog();
            txtAnalisisRuta.Text = fbdSoporte.SelectedPath;

            ofdSoporte.Multiselect = true;
            ofdSoporte.InitialDirectory = fbdSoporte.SelectedPath;
            ofdSoporte.ShowDialog();
            AgregarArchivos(dgAnalisisDocs, cbAnalisisHarvest.Text, "Archivos");
        }

        private void button17_Click(object sender, EventArgs e)
        {
           
        }

      

        private void button18_Click(object sender, EventArgs e)
        {
            string sLiberacion;
            sLiberacion = CeldaSeleccionada(dgLiberaciones, "Archivo");
            Explorador(sLiberacion.Substring(0,sLiberacion.LastIndexOf("\\") + 1));
            Excel(sLiberacion);
        }

        private void button19_Click(object sender, EventArgs e)
        {
            string sOrigen = "C:\\Pendientes\\";
            string sResp,sNuevo;
            foreach (DataGridViewRow fila in dgCambios.Rows)
            {
                if (fila.Cells["Objeto"].Value != null)
                {
                    sResp = BasePackage("DISCTOS", fila.Cells["Objeto"].Value.ToString(), "C:\\Pendientes\\Liberaciones\\");
                    sNuevo = BasePackage("TISCTOS", fila.Cells["Objeto"].Value.ToString(), "C:\\Pendientes\\Liberaciones\\");
                    SqlPlus("DISCTOS", sOrigen + sNuevo);
                }
            }
        }

        private void timBackUp_Tick(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            Filtrar(dgPedidos, textBox1.Text);
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            Filtrar(dgBitacora, textBox3.Text);
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
        //    Filtrar(dgLiberaciones, textBox2.Text);
        }

        private void dgPlanificacion_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void dgPlanificacion_CurrentCellChanged(object sender, EventArgs e)
        {
            txtGlosa.Text = CeldaSeleccionada(dgPedidos, "Solicitud");
        }

        private void dgHarvest_CurrentCellChanged(object sender, EventArgs e)
        {
            txtHarvest.Text = CeldaSeleccionada(dgHarvest, "DESCRIPCIÓN");
        }

        private void button20_Click(object sender, EventArgs e)
        {
            string sLiberacion;
            sLiberacion = CeldaSeleccionada(dgAnalisisDocs, "Objeto");
            Explorador(sLiberacion.Substring(0, sLiberacion.LastIndexOf("\\") + 1));            
        }
       
        private void Gantt()
        {
            string sHead="";
            string sTask="";
            string sAssi="";
            string sCale="";
            string sResso="";
            string sArchivo="";

            sTask = "";
            int filas = 0;
            int Tasks=1, Task = 1, Id=1;
            filas = dgPlanificacion.Rows.Count;
            DataGridViewRow dr;
            bool bVisible;
            DataTable dtPlan;
            DataTable dtDatos;
            
            dtPlan = ((DataTable)dgPlanificacion.DataSource);
            dtPlan.DefaultView.Sort = "Id Componente";
            dtDatos = dtPlan.DefaultView.ToTable(true, new string[] { "Id Harvest"   });
            //dtDatos = dtPlan;

            sTask = sTask + "        <Tasks>";
            foreach (DataRow item in dtDatos.Rows)
            {
                //TAREA PRINCIPAL                
                if (item["Id Harvest"].ToString()!="")
                {                
                sTask = sTask + "        <Task>";
                sTask = sTask + "            <UID>" + Id.ToString() + "</UID>";
                sTask = sTask + "            <ID>" + Id.ToString() + "</ID>";
                sTask = sTask + "            <Name>" + item["Id Harvest"].ToString()  + "</Name>";
                sTask = sTask + "            <Type>0</Type>";
                sTask = sTask + "            <IsNull>0</IsNull>";
                sTask = sTask + "            <CreateDate>" + FechaAMD() + "</CreateDate>";
                sTask = sTask + "            <WBS></WBS>";
                sTask = sTask + "            <OutlineNumber>1</OutlineNumber>";
                sTask = sTask + "            <OutlineLevel>1</OutlineLevel>";
                sTask = sTask + "            <Priority>500</Priority>";
                sTask = sTask + "            <Start>" + FechaAMD() + "</Start>";
                sTask = sTask + "            <Finish></Finish>";
                sTask = sTask + "            <Duration>PT64H0M0S</Duration>";
                sTask = sTask + "            <DurationFormat>39</DurationFormat>";
                sTask = sTask + "            <ResumeValid>0</ResumeValid>";
                sTask = sTask + "            <EffortDriven>1</EffortDriven>";
                sTask = sTask + "            <Recurring>0</Recurring>";
                sTask = sTask + "            <OverAllocated>0</OverAllocated>";
                sTask = sTask + "            <Estimated>1</Estimated>";
                sTask = sTask + "            <Milestone>0</Milestone>";
                sTask = sTask + "            <Summary>1</Summary>";
                sTask = sTask + "            <Critical>1</Critical>";
                sTask = sTask + "            <IsSubproject>0</IsSubproject>";
                sTask = sTask + "            <IsSubprojectReadOnly>0</IsSubprojectReadOnly>";
                sTask = sTask + "            <ExternalTask>0</ExternalTask>";
                sTask = sTask + "            <FixedCostAccrual>2</FixedCostAccrual>";
                sTask = sTask + "            <RemainingDuration>PT64H0M0S</RemainingDuration>";
                sTask = sTask + "            <ConstraintType>0</ConstraintType>";
                sTask = sTask + "            <CalendarUID>-1</CalendarUID>";
                sTask = sTask + "            <ConstraintDate>1970-01-01T00:00:00</ConstraintDate>";
                sTask = sTask + "            <LevelAssignments>0</LevelAssignments>";
                sTask = sTask + "            <LevelingCanSplit>0</LevelingCanSplit>";
                sTask = sTask + "            <LevelingDelay>0</LevelingDelay>";
                sTask = sTask + "            <LevelingDelayFormat>7</LevelingDelayFormat>";
                sTask = sTask + "            <IgnoreResourceCalendar>0</IgnoreResourceCalendar>";
                sTask = sTask + "            <HideBar>0</HideBar>";
                sTask = sTask + "            <Rollup>0</Rollup>";
                sTask = sTask + "            <EarnedValueMethod>0</EarnedValueMethod>";
                sTask = sTask + "            <Baseline>";
                sTask = sTask + "                <Number>0</Number>";
                sTask = sTask + "                <Start>" + FechaAMD() + "</Start>";
                sTask = sTask + "                <Finish></Finish>";
                sTask = sTask + "                <Work>PT6H0M0S</Work>";
                sTask = sTask + "            </Baseline>";
                sTask = sTask + "            <Active>1</Active>";
                sTask = sTask + "            <Manual>0</Manual>";
                sTask = sTask + "        </Task>";

                Task++;
                Tasks=1;
                Id++;
                }

                //
                foreach (DataRow harvest in dtPlan.Select("[Id Harvest]='" + item["Id Harvest"].ToString() + "'"))
            {
                bVisible = true;                

                sTask = sTask + "        <Task>";
                sTask = sTask + "            <UID>" + Id.ToString() + "</UID>";
                sTask = sTask + "            <ID>" + Id.ToString() + "</ID>";
                sTask = sTask + "            <Name>" + harvest["Descripcion"].ToString() + "</Name>";
                sTask = sTask + "            <Type>0</Type>";
                sTask = sTask + "            <IsNull>0</IsNull>";
                sTask = sTask + "            <CreateDate>" + FechaAMD() + "</CreateDate>";
                sTask = sTask + "            <WBS></WBS>";
                sTask = sTask + "            <OutlineNumber>1.1</OutlineNumber>";
                sTask = sTask + "            <OutlineLevel>2</OutlineLevel>";
                sTask = sTask + "            <Priority>500</Priority>";
                sTask = sTask + "            <Start>" + FechaAMD() + "</Start>";
                sTask = sTask + "            <Finish></Finish>";
                sTask = sTask + "            <Duration>" + Duracion(harvest["Horas"].ToString()) + "</Duration>";
                sTask = sTask + "            <DurationFormat>39</DurationFormat>";
                sTask = sTask + "            <ResumeValid>0</ResumeValid>";
                sTask = sTask + "            <EffortDriven>1</EffortDriven>";
                sTask = sTask + "            <Recurring>0</Recurring>";
                sTask = sTask + "            <OverAllocated>0</OverAllocated>";
                sTask = sTask + "            <Estimated>1</Estimated>";
                sTask = sTask + "            <Milestone>0</Milestone>";
                sTask = sTask + "            <Summary>0</Summary>";
                sTask = sTask + "            <Critical>1</Critical>";
                sTask = sTask + "            <IsSubproject>0</IsSubproject>";
                sTask = sTask + "            <IsSubprojectReadOnly>0</IsSubprojectReadOnly>";
                sTask = sTask + "            <ExternalTask>0</ExternalTask>";
                sTask = sTask + "            <FixedCostAccrual>2</FixedCostAccrual>";
                sTask = sTask + "            <RemainingDuration>PT8H0M0S</RemainingDuration>";
                sTask = sTask + "            <ConstraintType>0</ConstraintType>";
                sTask = sTask + "            <CalendarUID>-1</CalendarUID>";
                sTask = sTask + "            <ConstraintDate>1970-01-01T00:00:00</ConstraintDate>";
                sTask = sTask + "            <LevelAssignments>0</LevelAssignments>";
                sTask = sTask + "            <LevelingCanSplit>0</LevelingCanSplit>";
                sTask = sTask + "            <LevelingDelay>0</LevelingDelay>";
                sTask = sTask + "            <LevelingDelayFormat>7</LevelingDelayFormat>";
                sTask = sTask + "            <IgnoreResourceCalendar>0</IgnoreResourceCalendar>";
                sTask = sTask + "            <HideBar>0</HideBar>";
                sTask = sTask + "            <Rollup>0</Rollup>";
                sTask = sTask + "            <EarnedValueMethod>0</EarnedValueMethod>";
                sTask = sTask + "            <Baseline>";
                sTask = sTask + "                <Number>0</Number>";
                sTask = sTask + "                <Start>" + FechaAMD() + "</Start>";
                sTask = sTask + "                <Finish></Finish>";
                sTask = sTask + "                <Work>PT8H0M0S</Work>";
                sTask = sTask + "            </Baseline>";
                sTask = sTask + "            <Active>1</Active>";
                sTask = sTask + "            <Manual>0</Manual>";
                sTask = sTask + "        </Task>";
                Tasks++;
                Id++;
            }
            }

            sTask = sTask + "        </Tasks>";

            sHead = "";
            sHead = sHead + "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>";
            sHead = sHead + "<Project xmlns='http://schemas.microsoft.com/project'>";
            sHead = sHead + "    <SaveVersion>14</SaveVersion>";
            sHead = sHead + "    <Name>Afiliaciones 2018</Name>";
            sHead = sHead + "    <Title>Afiliaciones 2018</Title>";
            sHead = sHead + "    <Manager>Pablo Miranda</Manager>";
            sHead = sHead + "    <ScheduleFromStart>1</ScheduleFromStart>";
            sHead = sHead + "    <StartDate>" + FechaAMD() + "</StartDate>";
            sHead = sHead + "    <FinishDate></FinishDate>";
            sHead = sHead + "    <FYStartDate>1</FYStartDate>";
            sHead = sHead + "    <CriticalSlackLimit>0</CriticalSlackLimit>";
            sHead = sHead + "    <CurrencyDigits>2</CurrencyDigits>";
            sHead = sHead + "    <CurrencySymbol>$</CurrencySymbol>";
            sHead = sHead + "    <CurrencySymbolPosition>0</CurrencySymbolPosition>";
            sHead = sHead + "    <CalendarUID>1</CalendarUID>";
            sHead = sHead + "    <DefaultStartTime>04:00:00</DefaultStartTime>";
            sHead = sHead + "    <DefaultFinishTime>13:00:00</DefaultFinishTime>";
            sHead = sHead + "    <MinutesPerDay>360</MinutesPerDay>";
            sHead = sHead + "    <MinutesPerWeek>1800</MinutesPerWeek>";
            sHead = sHead + "    <DaysPerMonth>20</DaysPerMonth>";
            sHead = sHead + "    <DefaultTaskType>0</DefaultTaskType>";
            sHead = sHead + "    <DefaultFixedCostAccrual>2</DefaultFixedCostAccrual>";
            sHead = sHead + "    <DefaultStandardRate>10</DefaultStandardRate>";
            sHead = sHead + "    <DefaultOvertimeRate>15</DefaultOvertimeRate>";
            sHead = sHead + "    <DurationFormat>7</DurationFormat>";
            sHead = sHead + "    <WorkFormat>2</WorkFormat>";
            sHead = sHead + "    <EditableActualCosts>0</EditableActualCosts>";
            sHead = sHead + "    <HonorConstraints>0</HonorConstraints>";
            sHead = sHead + "    <EarnedValueMethod>0</EarnedValueMethod>";
            sHead = sHead + "    <InsertedProjectsLikeSummary>0</InsertedProjectsLikeSummary>";
            sHead = sHead + "    <MultipleCriticalPaths>0</MultipleCriticalPaths>";
            sHead = sHead + "    <NewTasksEffortDriven>0</NewTasksEffortDriven>";
            sHead = sHead + "    <NewTasksEstimated>1</NewTasksEstimated>";
            sHead = sHead + "    <SplitsInProgressTasks>0</SplitsInProgressTasks>";
            sHead = sHead + "    <SpreadActualCost>0</SpreadActualCost>";
            sHead = sHead + "    <SpreadPercentComplete>0</SpreadPercentComplete>";
            sHead = sHead + "    <TaskUpdatesResource>1</TaskUpdatesResource>";
            sHead = sHead + "    <FiscalYearStart>0</FiscalYearStart>";
            sHead = sHead + "    <WeekStartDay>1</WeekStartDay>";
            sHead = sHead + "    <MoveCompletedEndsBack>0</MoveCompletedEndsBack>";
            sHead = sHead + "    <MoveRemainingStartsBack>0</MoveRemainingStartsBack>";
            sHead = sHead + "    <MoveRemainingStartsForward>0</MoveRemainingStartsForward>";
            sHead = sHead + "    <MoveCompletedEndsForward>0</MoveCompletedEndsForward>";
            sHead = sHead + "    <BaselineForEarnedValue>0</BaselineForEarnedValue>";
            sHead = sHead + "    <AutoAddNewResourcesAndTasks>1</AutoAddNewResourcesAndTasks>";
            sHead = sHead + "    <CurrentDate>2018-07-06T16:39:00</CurrentDate>";
            sHead = sHead + "    <MicrosoftProjectServerURL>1</MicrosoftProjectServerURL>";
            sHead = sHead + "    <Autolink>1</Autolink>";
            sHead = sHead + "    <NewTaskStartDate>0</NewTaskStartDate>";
            sHead = sHead + "    <DefaultTaskEVMethod>0</DefaultTaskEVMethod>";
            sHead = sHead + "    <ProjectExternallyEdited>0</ProjectExternallyEdited>";
            sHead = sHead + "    <ActualsInSync>0</ActualsInSync>";
            sHead = sHead + "    <RemoveFileProperties>0</RemoveFileProperties>";
            sHead = sHead + "    <AdminProject>0</AdminProject>";
            sHead = sHead + "    <ExtendedAttributes/>";

            //
            sCale = sCale + "    <Calendars>";
            sCale = sCale + "        <Calendar>";
            sCale = sCale + "            <UID>1</UID>";
            sCale = sCale + "            <Name>Estandar</Name>";
            sCale = sCale + "            <IsBaseCalendar>1</IsBaseCalendar>";
            sCale = sCale + "            <WeekDays>";
            sCale = sCale + "                <WeekDay>";
            sCale = sCale + "                    <DayType>1</DayType>";
            sCale = sCale + "                    <DayWorking>0</DayWorking>";
            sCale = sCale + "                </WeekDay>";
            sCale = sCale + "                <WeekDay>";
            sCale = sCale + "                    <DayType>2</DayType>";
            sCale = sCale + "                    <DayWorking>1</DayWorking>";
            sCale = sCale + "                    <WorkingTimes>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>10:30:00</FromTime>";
            sCale = sCale + "                            <ToTime>12:30:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>13:30:00</FromTime>";
            sCale = sCale + "                            <ToTime>18:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                    </WorkingTimes>";
            sCale = sCale + "                </WeekDay>";
            sCale = sCale + "                <WeekDay>";
            sCale = sCale + "                    <DayType>3</DayType>";
            sCale = sCale + "                    <DayWorking>1</DayWorking>";
            sCale = sCale + "                    <WorkingTimes>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>10:30:00</FromTime>";
            sCale = sCale + "                            <ToTime>12:30:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>13:30:00</FromTime>";
            sCale = sCale + "                            <ToTime>18:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                    </WorkingTimes>";
            sCale = sCale + "                </WeekDay>";
            sCale = sCale + "                <WeekDay>";
            sCale = sCale + "                    <DayType>4</DayType>";
            sCale = sCale + "                    <DayWorking>1</DayWorking>";
            sCale = sCale + "                    <WorkingTimes>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>10:30:00</FromTime>";
            sCale = sCale + "                            <ToTime>12:30:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>13:30:00</FromTime>";
            sCale = sCale + "                            <ToTime>18:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                    </WorkingTimes>";
            sCale = sCale + "                </WeekDay>";
            sCale = sCale + "                <WeekDay>";
            sCale = sCale + "                    <DayType>5</DayType>";
            sCale = sCale + "                    <DayWorking>1</DayWorking>";
            sCale = sCale + "                    <WorkingTimes>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>10:30:00</FromTime>";
            sCale = sCale + "                            <ToTime>12:30:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>13:30:00</FromTime>";
            sCale = sCale + "                            <ToTime>17:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                    </WorkingTimes>";
            sCale = sCale + "                </WeekDay>";
            sCale = sCale + "                <WeekDay>";
            sCale = sCale + "                    <DayType>6</DayType>";
            sCale = sCale + "                    <DayWorking>1</DayWorking>";
            sCale = sCale + "                    <WorkingTimes>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>10:30:00</FromTime>";
            sCale = sCale + "                            <ToTime>12:30:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>13:30:00</FromTime>";
            sCale = sCale + "                            <ToTime>18:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                    </WorkingTimes>";
            sCale = sCale + "                </WeekDay>";
            sCale = sCale + "                <WeekDay>";
            sCale = sCale + "                    <DayType>7</DayType>";
            sCale = sCale + "                    <DayWorking>0</DayWorking>";
            sCale = sCale + "                </WeekDay>";
            sCale = sCale + "            </WeekDays>";
            sCale = sCale + "        </Calendar>";
            sCale = sCale + "        <Calendar>";
            sCale = sCale + "            <UID>2</UID>";
            sCale = sCale + "            <Name>24 Hours</Name>";
            sCale = sCale + "            <IsBaseCalendar>1</IsBaseCalendar>";
            sCale = sCale + "            <WeekDays>";
            sCale = sCale + "                <WeekDay>";
            sCale = sCale + "                    <DayType>1</DayType>";
            sCale = sCale + "                    <DayWorking>1</DayWorking>";
            sCale = sCale + "                    <WorkingTimes>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>00:00:00</FromTime>";
            sCale = sCale + "                            <ToTime>00:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                    </WorkingTimes>";
            sCale = sCale + "                </WeekDay>";
            sCale = sCale + "                <WeekDay>";
            sCale = sCale + "                    <DayType>2</DayType>";
            sCale = sCale + "                    <DayWorking>1</DayWorking>";
            sCale = sCale + "                    <WorkingTimes>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>00:00:00</FromTime>";
            sCale = sCale + "                            <ToTime>00:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                    </WorkingTimes>";
            sCale = sCale + "                </WeekDay>";
            sCale = sCale + "                <WeekDay>";
            sCale = sCale + "                    <DayType>3</DayType>";
            sCale = sCale + "                    <DayWorking>1</DayWorking>";
            sCale = sCale + "                    <WorkingTimes>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>00:00:00</FromTime>";
            sCale = sCale + "                            <ToTime>00:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                    </WorkingTimes>";
            sCale = sCale + "                </WeekDay>";
            sCale = sCale + "                <WeekDay>";
            sCale = sCale + "                    <DayType>4</DayType>";
            sCale = sCale + "                    <DayWorking>1</DayWorking>";
            sCale = sCale + "                    <WorkingTimes>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>00:00:00</FromTime>";
            sCale = sCale + "                            <ToTime>00:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                    </WorkingTimes>";
            sCale = sCale + "                </WeekDay>";
            sCale = sCale + "                <WeekDay>";
            sCale = sCale + "                    <DayType>5</DayType>";
            sCale = sCale + "                    <DayWorking>1</DayWorking>";
            sCale = sCale + "                    <WorkingTimes>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>00:00:00</FromTime>";
            sCale = sCale + "                            <ToTime>00:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                    </WorkingTimes>";
            sCale = sCale + "                </WeekDay>";
            sCale = sCale + "                <WeekDay>";
            sCale = sCale + "                    <DayType>6</DayType>";
            sCale = sCale + "                    <DayWorking>1</DayWorking>";
            sCale = sCale + "                    <WorkingTimes>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>00:00:00</FromTime>";
            sCale = sCale + "                            <ToTime>00:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                    </WorkingTimes>";
            sCale = sCale + "                </WeekDay>";
            sCale = sCale + "                <WeekDay>";
            sCale = sCale + "                    <DayType>7</DayType>";
            sCale = sCale + "                    <DayWorking>1</DayWorking>";
            sCale = sCale + "                    <WorkingTimes>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>00:00:00</FromTime>";
            sCale = sCale + "                            <ToTime>00:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                    </WorkingTimes>";
            sCale = sCale + "                </WeekDay>";
            sCale = sCale + "            </WeekDays>";
            sCale = sCale + "        </Calendar>";
            sCale = sCale + "        <Calendar>";
            sCale = sCale + "            <UID>3</UID>";
            sCale = sCale + "            <Name>Turno nocturno</Name>";
            sCale = sCale + "            <IsBaseCalendar>1</IsBaseCalendar>";
            sCale = sCale + "            <WeekDays>";
            sCale = sCale + "                <WeekDay>";
            sCale = sCale + "                    <DayType>1</DayType>";
            sCale = sCale + "                    <DayWorking>0</DayWorking>";
            sCale = sCale + "                </WeekDay>";
            sCale = sCale + "                <WeekDay>";
            sCale = sCale + "                    <DayType>2</DayType>";
            sCale = sCale + "                    <DayWorking>1</DayWorking>";
            sCale = sCale + "                    <WorkingTimes>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>23:00:00</FromTime>";
            sCale = sCale + "                            <ToTime>00:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                    </WorkingTimes>";
            sCale = sCale + "                </WeekDay>";
            sCale = sCale + "                <WeekDay>";
            sCale = sCale + "                    <DayType>3</DayType>";
            sCale = sCale + "                    <DayWorking>1</DayWorking>";
            sCale = sCale + "                    <WorkingTimes>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>00:00:00</FromTime>";
            sCale = sCale + "                            <ToTime>03:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>04:00:00</FromTime>";
            sCale = sCale + "                            <ToTime>08:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>23:00:00</FromTime>";
            sCale = sCale + "                            <ToTime>00:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                    </WorkingTimes>";
            sCale = sCale + "                </WeekDay>";
            sCale = sCale + "                <WeekDay>";
            sCale = sCale + "                    <DayType>4</DayType>";
            sCale = sCale + "                    <DayWorking>1</DayWorking>";
            sCale = sCale + "                    <WorkingTimes>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>00:00:00</FromTime>";
            sCale = sCale + "                            <ToTime>03:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>04:00:00</FromTime>";
            sCale = sCale + "                            <ToTime>08:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>23:00:00</FromTime>";
            sCale = sCale + "                            <ToTime>00:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                    </WorkingTimes>";
            sCale = sCale + "                </WeekDay>";
            sCale = sCale + "                <WeekDay>";
            sCale = sCale + "                    <DayType>5</DayType>";
            sCale = sCale + "                    <DayWorking>1</DayWorking>";
            sCale = sCale + "                    <WorkingTimes>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>00:00:00</FromTime>";
            sCale = sCale + "                            <ToTime>03:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>04:00:00</FromTime>";
            sCale = sCale + "                            <ToTime>08:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>23:00:00</FromTime>";
            sCale = sCale + "                            <ToTime>00:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                    </WorkingTimes>";
            sCale = sCale + "                </WeekDay>";
            sCale = sCale + "                <WeekDay>";
            sCale = sCale + "                    <DayType>6</DayType>";
            sCale = sCale + "                    <DayWorking>1</DayWorking>";
            sCale = sCale + "                    <WorkingTimes>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>00:00:00</FromTime>";
            sCale = sCale + "                            <ToTime>03:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>04:00:00</FromTime>";
            sCale = sCale + "                            <ToTime>08:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>23:00:00</FromTime>";
            sCale = sCale + "                            <ToTime>00:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                    </WorkingTimes>";
            sCale = sCale + "                </WeekDay>";
            sCale = sCale + "                <WeekDay>";
            sCale = sCale + "                    <DayType>7</DayType>";
            sCale = sCale + "                    <DayWorking>1</DayWorking>";
            sCale = sCale + "                    <WorkingTimes>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>00:00:00</FromTime>";
            sCale = sCale + "                            <ToTime>03:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                        <WorkingTime>";
            sCale = sCale + "                            <FromTime>04:00:00</FromTime>";
            sCale = sCale + "                            <ToTime>08:00:00</ToTime>";
            sCale = sCale + "                        </WorkingTime>";
            sCale = sCale + "                    </WorkingTimes>";
            sCale = sCale + "                </WeekDay>";
            sCale = sCale + "            </WeekDays>";
            sCale = sCale + "        </Calendar>";
            sCale = sCale + "    </Calendars>";                                               

            sResso = sResso + "    <Resources>";
            sResso = sResso + "        <Resource>";
            sResso = sResso + "            <UID>0</UID>";
            sResso = sResso + "            <ID>0</ID>";
            sResso = sResso + "            <Name>No Asignado</Name>";
            sResso = sResso + "            <Type>1</Type>";
            sResso = sResso + "            <IsNull>0</IsNull>";
            sResso = sResso + "            <Initials>N</Initials>";
            sResso = sResso + "            <Group></Group>";
            sResso = sResso + "            <EmailAddress></EmailAddress>";
            sResso = sResso + "            <MaxUnits>1</MaxUnits>";
            sResso = sResso + "            <PeakUnits>1</PeakUnits>";
            sResso = sResso + "            <OverAllocated>0</OverAllocated>";
            sResso = sResso + "            <Start>" + FechaAMD() + "</Start>";
            sResso = sResso + "            <Finish></Finish>";
            sResso = sResso + "            <CanLevel>0</CanLevel>";
            sResso = sResso + "            <AccrueAt>3</AccrueAt>";
            sResso = sResso + "            <StandardRateFormat>3</StandardRateFormat>";
            sResso = sResso + "            <OvertimeRateFormat>3</OvertimeRateFormat>";
            sResso = sResso + "            <IsGeneric>0</IsGeneric>";
            sResso = sResso + "            <IsInactive>0</IsInactive>";
            sResso = sResso + "            <IsEnterprise>0</IsEnterprise>";
            sResso = sResso + "            <IsBudget>0</IsBudget>";
            sResso = sResso + "            <AvailabilityPeriods/>";
            sResso = sResso + "        </Resource>";
            sResso = sResso + "    </Resources>";

            sAssi = sAssi + "    <Assignments>";
            sAssi = sAssi + "        <Assignment>";
            sAssi = sAssi + "            <UID>1</UID>";
            sAssi = sAssi + "            <TaskUID>1</TaskUID>";
            sAssi = sAssi + "            <ResourceUID>-65535</ResourceUID>";
            sAssi = sAssi + "            <Finish>2018-07-18T17:00:00</Finish>";
            sAssi = sAssi + "            <HasFixedRateUnits>1</HasFixedRateUnits>";
            sAssi = sAssi + "            <FixedMaterial>0</FixedMaterial>";
            sAssi = sAssi + "            <RemainingWork>PT8H0M0S</RemainingWork>";
            sAssi = sAssi + "            <Start>" + FechaAMD() + "</Start>";
            sAssi = sAssi + "            <Stop>1969-12-31T21:00:00</Stop>";
            sAssi = sAssi + "            <Resume>" + FechaAMD() + "</Resume>";
            sAssi = sAssi + "            <Units>1</Units>";
            sAssi = sAssi + "            <Work>PT8H0M0S</Work>";
            sAssi = sAssi + "            <WorkContour>0</WorkContour>";
            sAssi = sAssi + "            <Baseline>";
            sAssi = sAssi + "                <Number>0</Number>";
            sAssi = sAssi + "                <Start>" + FechaAMD() + "</Start>";
            sAssi = sAssi + "                <Finish></Finish>";
            sAssi = sAssi + "            </Baseline>";
            sAssi = sAssi + "            <TimephasedData>";
            sAssi = sAssi + "                <Type>1</Type>";
            sAssi = sAssi + "                <UID>1</UID>";
            sAssi = sAssi + "                <Start>" + FechaAMD() + "</Start>";
            sAssi = sAssi + "                <Finish></Finish>";
            sAssi = sAssi + "                <Unit>3</Unit>";
            sAssi = sAssi + "                <Value>PT8H0M0S</Value>";
            sAssi = sAssi + "            </TimephasedData>";
            sAssi = sAssi + "            <TimephasedData>";
            sAssi = sAssi + "                <Type>4</Type>";
            sAssi = sAssi + "                <UID>1</UID>";
            sAssi = sAssi + "                <Start>" + FechaAMD() + "</Start>";
            sAssi = sAssi + "                <Finish></Finish>";
            sAssi = sAssi + "                <Unit>3</Unit>";
            sAssi = sAssi + "                <Value>PT8H0M0S</Value>";
            sAssi = sAssi + "            </TimephasedData>";
            sAssi = sAssi + "        </Assignment>";

            
            sAssi = sAssi + "    </Assignments>";

            sArchivo = txtAnalisisRuta.Text + "\\" + cbAnalisisHarvest.Text + ".xml";
            File.WriteAllText(sArchivo, sHead + sCale + sTask + sResso + sAssi + "</Project>", Encoding.GetEncoding(1252));
            Project(sArchivo);
        }

        private void Planificar(string sHarvest)
        {
            string sProyecto, sCompo, sDesc, sFecha, sHoras;         
            DataTable dtPlan = new DataTable();
            DataTable sortedDT = new DataTable();


            dtPlan = (DataTable)dgRequerimientos.DataSource;
            dtPlan.DefaultView.Sort = "Componente";

            DataTable dtCloned = dtPlan.Clone();
            dtCloned.Columns["Componente"].DataType = typeof(Int32);
            foreach (DataRow row in dtPlan.Rows)
            {
                if ((row["% Reutilizacion"].ToString() != "100") && (row["Id Harvest"].ToString() == sHarvest))
                {
                    row["Horas"] = Val(row["Horas"]) - ((Val(row["% Reutilizacion"].ToString()) * Val(row["Horas"])) / 100);
                    dtCloned.ImportRow(row);
                }
                
            }
            dtCloned.DefaultView.Sort = "Componente";

            sortedDT = dtCloned.DefaultView.ToTable(true, new string[] { "Id Harvest", "Proyecto", "Id Componente", "Titulo", "Horas" });
            //foreach (DataRow harvest in sortedDT.Select("[Id Harvest]='" + sHarvest + "' AND Horas <>''"))
            foreach (DataRow harvest in sortedDT.Rows)
            {
                if (harvest["Id Harvest"].ToString() == sHarvest && harvest["Horas"].ToString() != "")
                {
                    sProyecto = harvest["Proyecto"].ToString();
                    sCompo = harvest["Id Componente"].ToString();
                    sDesc = harvest["Titulo"].ToString();
                    sHoras = harvest["Horas"].ToString();
                    dsEjecucion(dgPlanificacion, "Ejecucion", sHarvest, sProyecto, sCompo, sDesc, "", sHoras);
                }
                
            }
        }

        private void Planificar(string sHarvest, string sComponente )
        {
            string sProyecto, sCompo, sDesc, sFecha, sHoras;
            DataTable dtPlan = new DataTable();
            DataTable sortedDT = new DataTable();
            int filas = 0;


            dtPlan = (DataTable)dgPlanificacion.DataSource;
            filas = dtPlan.Rows.Count - 1;

            sProyecto = CeldaSeleccionada(dgRequerimientos, "Proyecto");
            sCompo = CeldaSeleccionada(dgRequerimientos, "Id Componente");
            sDesc = CeldaSeleccionada(dgRequerimientos, "Descripcion");
            sHoras = CeldaSeleccionada(dgRequerimientos, "Horas");
            DataRow dr;

            for (int i = 0; i < filas; i++)
            {
                if (dgPlanificacion.Rows[i].Cells["% Avance"].Value.ToString() != "100")
                {
                    dr = dtPlan.NewRow();
                    dr["Id Harvest"] = CeldaSeleccionada(dgRequerimientos, "Id Harvest");
                    dr["Proyecto"] = CeldaSeleccionada(dgRequerimientos, "Proyecto");
                    dr["Id Componente"] = CeldaSeleccionada(dgRequerimientos, "Id Componente");
                    dr["Descripcion"] = CeldaSeleccionada(dgRequerimientos, "Descripcion");                    

                    dr["Iniciar"] = "";
                    dr["Finalizar"] = "";
                    dr["Horas"] = CeldaSeleccionada(dgRequerimientos, "Horas");
                    dr["% Avance"] = "0";

                    dtPlan.Rows.InsertAt(dr, i-1);
                    dtPlan.AcceptChanges();
                    break;
                    dgPlanificacion.DataSource = dtPlan;                    
                }
                
            }            
        }
        private void btnGantt_Click(object sender, EventArgs e)
        {
            Gantt();
        }

        private string FechaGantt(string sFecha, string sHrs)
        {
            Boolean bFechaSEM = true;
            string sActividad;
            DateTime dActividad;
            DateTime dIniAM;
            DateTime dFinAM;
            DateTime dIniPM;
            DateTime dFinPM;
            DateTime dFecha;

            double dHrs;
            double dMin=0;
            sHrs = sHrs.Replace(".", ",");
            dFecha = new DateTime( Convert.ToInt16(sFecha.Substring(6, 4)), Convert.ToInt16(sFecha.Substring(3, 2)), Convert.ToInt16(sFecha.Substring(0, 2)), Convert.ToInt16(sFecha.Substring(11, 2)), Convert.ToInt16(sFecha.Substring(14, 2)), 0, 0);
            dActividad = dFecha.AddMinutes(0);
            if (sHrs.Trim()=="")
            {
                sHrs = "0";
            }

            dHrs = Convert.ToDouble(sHrs) * 60;          
  
            
            while (dMin<=dHrs)
            {
                dActividad = dActividad.AddMinutes(1);
                dIniAM = new DateTime(dActividad.Year, dActividad.Month, dActividad.Day, 09, 30, 0);
                dFinAM = new DateTime(dActividad.Year, dActividad.Month, dActividad.Day, 12, 30, 0);

                dIniPM = new DateTime(dActividad.Year, dActividad.Month, dActividad.Day, 13, 30, 0);
                dFinPM = new DateTime(dActividad.Year, dActividad.Month, dActividad.Day, 21, 00, 0);

                if (dFecha.DayOfWeek == DayOfWeek.Friday)
                {
                    dIniPM = new DateTime(dActividad.Year, dActividad.Month, dActividad.Day, 13, 30, 0);
                    dFinPM = new DateTime(dActividad.Year, dActividad.Month, dActividad.Day, 16, 00, 0);
                }

                if (((dActividad >= dIniAM && dActividad <= dFinAM) || (dActividad >= dIniPM && dActividad <= dFinPM)) && EsLaboral(dActividad))
                {
                    dMin++;
                }         
                

            }

            return dActividad.ToString("dd/MM/yyyy HH:mm");
        }

        private bool EsLaboral(DateTime dFecha)
        {
            if ((dFecha.DayOfWeek == DayOfWeek.Saturday) || (dFecha.DayOfWeek == DayOfWeek.Sunday))
            {
                return  false;
            }

            //FERIADOS
            if (dFecha.Month==8 & dFecha.Day==15)
            {
                return false;
            }
            return true;
        }
        
        private void Formato( DataGridView dg)
        {        
             foreach (DataGridViewColumn c in dg.Columns)
            {
                c.DefaultCellStyle.Font = new Font("Arial", 12F, GraphicsUnit.Pixel);
            }
        }

        private void NoSort(DataGridView dg)
        {
            foreach (DataGridViewColumn column in dg.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            Planificar(cbPlanificacionHarvest.Text);
        }

        private void button22_Click(object sender, EventArgs e)
        {
            Persistencia(dgPlanificacion, "Ejecucion");
        }

        private void button23_Click(object sender, EventArgs e)
        {
            Gantt();
        }

        private void button24_Click(object sender, EventArgs e)
        {
            Planificar(cbAnalisisHarvest.Text);
        }

        private void button17_Click_1(object sender, EventArgs e)
        {
            //no se usa?
            String sHarv;
            String sSoli;
            String sFecha;

            sHarv = CeldaSeleccionada(dgPlanificacion, "Id Harvest");
            sSoli = CeldaSeleccionada(dgPlanificacion, "Descripcion");
            sFecha = FechaCorta(CeldaSeleccionada(dgPlanificacion, "Iniciar"));

            dsBitacora(dgBitacora, "Bitacora", sHarv, "",sSoli,sFecha);
        }

        private void cbPlanificacionHarvest_SelectedIndexChanged(object sender, EventArgs e)
        {
            Filtrar(dgPlanificacion, cbPlanificacionHarvest.Text);
        }

        private void button25_Click(object sender, EventArgs e)
        {
            Persistencia(dgObjetos, "Packages");
        }

        private void cbBaseHarvest_SelectedIndexChanged(object sender, EventArgs e)
        {
            Filtrar(dgObjetos, cbBaseHarvest.Text);
        }

        private void button26_Click(object sender, EventArgs e)
        {
            
            int filas=0;
            string hrs;
            DataGridViewRow dr;
            string sFecha;
            DateTime dFecha;

            string sInicio;
            string sFinal;

            sInicio = txtInicioPlan.Text;
            sFinal=txtFinPlan.Text;
            
            filas = dgPlanificacion.Rows.Count;
            sFecha = txtInicioPlan.Text;
            for (int fila = 0; fila < filas; fila++)
            {

                dr = dgPlanificacion.Rows[fila];                
                if (dr.Visible)
                { 
                    if ((dr.Cells["% Avance"].Value != null) && (dr.Cells["% Avance"].Value.ToString() != "100") )
                    {
                        dr.Cells["Iniciar"].Value = sFecha;
                        hrs = dr.Cells["Horas"].Value.ToString();
                        sFecha = FechaGantt(sFecha, hrs);
                        dr.Cells["Finalizar"].Value = sFecha;
                    }
                }
            }



        }

        private void button27_Click(object sender, EventArgs e)
        {
            String sHarv;
            String sSoli;
            String sFecha;
            String sSolic;

            sHarv = cbAnalisisHarvest.Text;
            sSoli = txtAnalisisAltoVuelo.Text.Trim();
            sFecha = FechaCorta( Fecha() );
            sSolic = Celda(dgPedidos, "Id Harvest", sHarv, "Solicitante");
            dsBitacora(dgBitacora, "Bitacora", sHarv, sSolic,sSoli, sFecha);
        }

        private void button28_Click(object sender, EventArgs e)
        {
                    
            Char delimiter = '.';
            String sFecha;
            String sArchivo;
            String sLeft;
            String sRigth;
            String sServer;
            String sPackage;
            String sRuta;

            sRuta = "C:\\Pendientes\\";
            sArchivo = CeldaSeleccionada(dgSoporte, "Objeto");
            sServer = sArchivo.Split(delimiter)[0];
            sPackage = sArchivo.Split(delimiter)[1];
            sRigth = sRuta + BasePackage(sServer, sPackage, "C:\\Pendientes\\Liberaciones\\");

            sFecha = CeldaSeleccionada(dgSoporte, "Fecha");
            sArchivo = CeldaSeleccionada(dgSoporte, "Objeto");
            sLeft = sRuta + sArchivo + "." + FechaArchivo(sFecha) + ".sql";
            Merger(sLeft, sRigth);
        
        }

        private string FechaModificacion(string sArchivo)
        {
            string sFecha = "";
            try
            {
                FileInfo fi = new FileInfo(sArchivo);
                sFecha =  fi.LastWriteTime.ToString("dd/MM/yyyy HH:mm");;                

            }
            catch (Exception)
            {

                sFecha = "";
            }


            return sFecha;
        }


        static async void getImage()
        {            
            string[] url = new string[32];
            Char delimiter = '|';
            string request;
            string trabajo;
            string requests="";

            url[0] = "80152325|TECNOCOM|D";
            url[1] = "80154140|PREVIRED|D";
            url[2] = "80151275|TECNOCOM|D";
            url[3] = "80154504|MANUAL|D";
                        
            for (int i = 0; i < 32; i++)
            {
            if (url[i] != null)
            {

                request = url[i].Split(delimiter)[1].ToUpper();
                trabajo = url[i].Split(delimiter)[2].ToUpper();

                if (trabajo == "V" || trabajo == "I")
                {
                    request = "http://rptventas.consalud.net/ventas/reportsventas.aspx?REPORTE=GRAL&USUARIO=SSARMIENTO&params=pin_nFolio,pin_destino&paramsvalues=" + url[i].Split(delimiter)[0] + ",ISAP&LOTE=FUN_DIGITAL&COPIAS=1&TIPO_REPORTE=PANT";
                    requests = requests + request;
                    GetUrl(request, "c:\\funes\\" + url[i].Split(delimiter)[0] + "-ISAP.pdf");
                    //Thread.Sleep(5000);

                }

                if (request != "MANUAL" && (trabajo == "D" || trabajo == "P"))
                {
                    request = "http://rptventas.consalud.net/ventas/reportsventas.aspx?REPORTE=GRAL&USUARIO=SSARMIENTO&params=pin_nFolio,pin_destino&paramsvalues=" + url[i].Split(delimiter)[0] + ",NOTI&LOTE=FUN_DIGITAL&COPIAS=1&TIPO_REPORTE=PANT";
                    requests = requests + request;
                    GetUrl(request, "c:\\funes\\" + url[i].Split(delimiter)[0] + "-NOTI.pdf");
                   // Thread.Sleep(5000);
                                
                }

                if (request == "MANUAL" && (trabajo == "D" || trabajo == "P"))
                {
                    request = "http://rptventas.consalud.net/ventas/reportsventas.aspx?REPORTE=GRAL&USUARIO=SSARMIENTO&params=pin_nFolio,pin_destino&paramsvalues=" + url[i].Split(delimiter)[0] + ",ISAP&LOTE=FUN_DIGITAL&COPIAS=1&TIPO_REPORTE=PANT";
                    requests = requests + request;
                    GetUrl(request,"c:\\funes\\" + url[i].Split(delimiter)[0] + "-ISAP.pdf" );
                   // Thread.Sleep(5000);

                    request = "http://rptventas.consalud.net/ventas/reportsventas.aspx?REPORTE=GRAL&USUARIO=SSARMIENTO&params=pin_nFolio,pin_destino&paramsvalues=" + url[i].Split(delimiter)[0] + ",EMPL&LOTE=FUN_DIGITAL&COPIAS=1&TIPO_REPORTE=PANT";
                    requests = requests + request;
                    GetUrl(request, "c:\\funes\\" + url[i].Split(delimiter)[0] + "-EMPL.pdf");
                    //Thread.Sleep(5000);    
                }

            }

        }

            requests = requests;
            //para los que son voluntarios o Independientes necesito solo la copia de la isapre
            //para los dependientes y pensionados que sean eletronicos copia isapre-empresa
            //para los que sean dependientes y pensionados con notificacion manual necesitos copia isapre y empleador por separado

        }


        public void SetearDatos()
        {
            string lsR;
            BaseFont bf;
            int liPag;
            int liPags;

            string[] aCartas = new string[7];
            string[] aMontos = new string[7];


            aCartas[0] = @"C:\FUNES\TODO_104480202.pdf";
            aCartas[1] = @"C:\FUNES\TODO_104371109.pdf";

            aMontos[0] = "0,5";
            aMontos[1] = "0,19";

            lsR = @"C:\FUNES\";

            for (liPag = 0; liPag <= 1; liPag++)
            {
                lsR = aCartas[liPag];
                PdfReader rdrContrato = new PdfReader(lsR);
                MemoryStream mstContrato = new MemoryStream();
                PdfStamper stpContrato = new PdfStamper(rdrContrato, mstContrato);
                PdfContentByte cobContrato;

                iTextSharp.text.Image instanceImg;
                instanceImg = iTextSharp.text.Image.GetInstance(@"C:\Desarrollo\Adecuaciones\4.2 Adecuaciones\cuadro.jpg");

                // Seteamos el Folio            
                bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

                cobContrato = stpContrato.GetOverContent(2);
                instanceImg.SetAbsolutePosition(375, 322);
                cobContrato.AddImage(instanceImg);

                cobContrato = stpContrato.GetOverContent(2);
                instanceImg.SetAbsolutePosition(501, 322);
                cobContrato.AddImage(instanceImg);

                
                bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cobContrato = stpContrato.GetOverContent(2);
                cobContrato.BeginText();
                cobContrato.SetFontAndSize(bf, 6);
                cobContrato.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, aMontos[liPag], 386, 326, 0); // P:
                cobContrato.EndText();

                bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cobContrato = stpContrato.GetOverContent(2);
                cobContrato.BeginText();
                cobContrato.SetFontAndSize(bf, 6);
                cobContrato.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, aMontos[liPag], 514, 326, 0); // P:
                cobContrato.EndText();
                
                stpContrato.Writer.CloseStream = false;
                stpContrato.Close();
                mstContrato.Position = 0;
                                
                FileStream mstContratoCanasta = new FileStream(lsR + ".pdf", FileMode.Create);
                PdfCopyFields tmpContratoCanasta = new PdfCopyFields(mstContratoCanasta);
                tmpContratoCanasta.AddDocument(new PdfReader(mstContrato));
                tmpContratoCanasta.Writer.CloseStream = false;
                tmpContratoCanasta.Close();
                mstContrato.Close();
                rdrContrato.Close();

                cobContrato = null/* TODO Change to default(_) if this is not a reference type */;
                stpContrato = null/* TODO Change to default(_) if this is not a reference type */;                
                mstContrato = null;
                rdrContrato = null/* TODO Change to default(_) if this is not a reference type */;
                                
                tmpContratoCanasta = null/* TODO Change to default(_) if this is not a reference type */;
                mstContratoCanasta.Close();
            }
        }


        private void button28_Click_1(object sender, EventArgs e)
        {
            getImage();
        }

        private string ArchivoComparar(string path1, string path2)
        {
            string sEstado="";
            string sfi = "";
            string sfd = "";
            FileInfo fi ;
            FileInfo fd ;

            try
            {
                 fi = new FileInfo(path1);
                 sfi = fi.LastWriteTime.ToString();
                 if (!(fi.Exists))
                 {
                     sEstado="NEI";
                 }
            }
            catch (Exception)
            {

                sEstado="NEI";
            }

            try
            {
                 fd = new FileInfo(path2);
                 sfd = fd.LastWriteTime.ToString();
                 if (!(fd.Exists))
                 {
                     sEstado = "NED";
                 }
            }
            catch (Exception)
            {

                sEstado = "NED";
            }
            
            if (sEstado=="")
            {
                sEstado = "Identico";
                byte[] file1 = File.ReadAllBytes(path1);
                byte[] file2 = File.ReadAllBytes(path2);
                if (file1.Length == file2.Length)
                {
                    for (int i = 0; i < file1.Length; i++)
                    {
                        if (file1[i] != file2[i])
                        {
                            sEstado = "Diferente," + sfd;
                            break;
                        }
                    }                    
                }
                else
                {
                    sEstado = "Diferente," + sfd;
                }                

                if (path1.ToUpper().IndexOf(".CONFIG")>-1)
                {
                    sEstado = sEstado + ", Arch. Config!";
                }
            }

            return sEstado;
        }

        private string ArchivoValidar(string path1)
        {
            string sEstado = "";            
            string c;
            int numero;
            int linea=1;
            
            byte[] file1 = File.ReadAllBytes(path1);
            for (int i = 0; i < file1.Length; i++)
            {
                c = file1[i].ToString();
                numero=Convert.ToInt16(c) ;
                if ((numero< 32) || (numero> 126) || (numero== 26) || (numero== 123) || (numero== 125))
                {
                    if (numero == 10)
                    {
                        linea++;
                    }

                    if (!((numero== 10) || (numero== 13)))                    
                    {
                        sEstado = sEstado + "Error en la linea " + linea.ToString() + " , " + Convert.ToChar(numero) + " " + numero + Environment.NewLine;
                    }
                    
                }            
            }

            return sEstado;
        }

        private void LeerPEE()
        {
            DataSet dsCarga = new DataSet();
            dgParametrosPEE.DataSource = null;
            dsCarga=dsPEE(dgParametrosPEE, "Estimaciones");
            string[] lines = System.IO.File.ReadAllLines("C:\\pee.txt");
            string[] cols;
            int fila = 0;
            foreach (string line in lines)
            {
                cols = line.Split(Convert.ToChar(9));
                dsCarga.Tables[0].Rows.Add(dsCarga.Tables[0].NewRow());
                dsCarga.Tables[0].Rows[fila]["Codigo"] = cols[0];
                dsCarga.Tables[0].Rows[fila]["Tipo"] = cols[1];
                dsCarga.Tables[0].Rows[fila]["Descripcion"] = cols[2];
                dsCarga.Tables[0].Rows[fila]["Optimista"] = cols[3];
                dsCarga.Tables[0].Rows[fila]["Pesimista"] = cols[4];
                dsCarga.Tables[0].Rows[fila]["Sugerido"] = cols[5];
                fila++;
            }
            dsCarga.AcceptChanges();
            dgParametrosPEE.DataSource = dsCarga.Tables[0];
        }

        static async void GetUrl(string request, string filename)
        {
            try
            {

                    //instance of HTTPClient
                    HttpClient client = new HttpClient();

                    //send  request asynchronously
                    HttpResponseMessage response = await client.GetAsync(request);

                    // Check that response was successful or throw exception
                    response.EnsureSuccessStatusCode();

                    // Read response asynchronously and save asynchronously to file
                    using (FileStream fileStream = new FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        //copy the content from response to filestream
                        await response.Content.CopyToAsync(fileStream);


                    }    
                }                

                
            

            catch (HttpRequestException rex)
            {
                Console.WriteLine(rex.ToString());
            }
            catch (Exception ex)
            {
                // For debugging
                Console.WriteLine(ex.ToString());
            }
                }

            private void button29_Click(object sender, EventArgs e)
            {
                string sOrigen = "C:\\Pendientes\\Respaldos\\";
                string sResp, sNuevo;
                foreach (DataGridViewRow fila in dgCambios.Rows)
                {
                    if (fila.Cells["Objeto"].Value != null)
                    {
                        sNuevo = BaseObjeto(txtServer.Text, fila.Cells["Objeto"].Value.ToString(), sOrigen);                    
                    }
                }
            }

            private void button30_Click(object sender, EventArgs e)
            {
                ExcelHarvest("C:\\Harvest.xls");
            }

            private void button31_Click(object sender, EventArgs e)
            {
                LeerPEE();
            }

            private void button32_Click(object sender, EventArgs e)
            {

                //ACA SE GENERAN EN UNA CARPETA DE PASO
                string sOrigen = "C:\\Pendientes\\Respaldos\\";
                string sResp, sNuevo;
                string sObjeto;
                foreach (DataGridViewRow fila in dgCambios.Rows)
                {
                    if (fila.Cells["Objeto"].Value != null)
                    {
                        sObjeto = fila.Cells["Objeto"].Value.ToString();
                        BasePackage(txtServer.Text, sObjeto, "C:\\Pendientes\\Liberaciones\\");
                        BasePackage("QISCTOS", sObjeto, "C:\\Pendientes\\Liberaciones\\");
                        dsBase(dgObjetos, "Packages", cbBaseHarvest.Text, Fecha(), txtServer.Text + '.' + sObjeto);
                        Persistencia(dgObjetos, "Packages");
                        Proyectos(cbDesarrolloProyecto);
                    }
                }

            }

            private void btnLiberacionesGuardar_Click_1(object sender, EventArgs e)
            {

            }

            private void button3_Click(object sender, EventArgs e)
            {

            }

            private void btnAnalisisInsertar_Click(object sender, EventArgs e)
            {
                int iFilas,iSelec;

                DataTable dt = new DataTable();                
                DataRow dr;

                dt = (DataTable)dgRequerimientos.DataSource;                

                iSelec=dgRequerimientos.SelectedRows[0].Index;                
                dr = (DataRow)dt.Rows[iSelec];
                dt.Rows.InsertAt(dt.NewRow(), iSelec);
                dt.Rows[iSelec].ItemArray = dt.Rows[iSelec + 1].ItemArray;
            }

            private void dgRequerimientos_SortCompare(object sender, DataGridViewSortCompareEventArgs e)
            {
                if (e.Column.Index == 2)
                {
                    e.SortResult = int.Parse(e.CellValue1.ToString()).CompareTo(int.Parse(e.CellValue2.ToString()));
                    e.Handled = true;//pass by the default sorting
                }
            }

            private void label1_Click(object sender, EventArgs e)
            {

            }

                      
   }
   }       
  