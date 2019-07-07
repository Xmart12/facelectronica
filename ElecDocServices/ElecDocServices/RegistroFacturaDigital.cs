using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.IO;
using System.Xml;
using ElecDocServices.Helpers;
using FTPConnection;
using ElecDocServices.IfacereServices;

namespace ElecDocServices
{
    public class RegistroFacturaDigital
    {
        //Clases Globales
        private Utils utl = new Utils();
        private ConMySQL con = null;
        private Errors err = null;
        private FtpAccess ftp = null;


        //Variables Globales
        private string usuario = null;
        private string localpath = "C:\\Projects\\TFS\\Compusal\\Resources\\";
        public string mensaje = null;

        private string TipoDoc = null;
        private string Serie = null;
        private string Resolucion = null;
        private int NoDocumento = 0;

        private DataTable dtEncabezado = null;
        private DataTable dtDetalle = null;


        //Constructor
        public RegistroFacturaDigital(string config, string user)
        {
            if (!File.Exists(config))
            {
                throw new Exception("Config file does not exists");
            }

            string errorPath = utl.convertirString(utl.getConfigValue(config, "SERVICES", "errors"));

            con = new ConMySQL(config, user);
            err = new Errors(user, "ElecDocs", errorPath);
        }


        //Registrar Documento
        public bool RegistrarDocumento(string resol, string doc, string serie, string docno)
        {

           
           

            

            return false;
        }


        //Descargar Documento
        public bool DescargarDocumento(string resol, string doc, string serie, string docno)
        {
            try
            {
                bool resultado = false;

                //Carga de Datos
                this.Resolucion = resol;
                this.TipoDoc = doc;
                this.Serie = serie;
                this.NoDocumento = utl.convertirInt(docno);

                cargarDatos();

                string filename = this.Resolucion + this.TipoDoc + this.Serie + this.NoDocumento + ".pdf";

                if (File.Exists(this.localpath + filename))
                {
                     resultado = utl.openFile(this.localpath + filename);
                }
                else
                {
                   
                    

                    if (resultado)
                    {
                        //resultado = convertPDF(res.pRespuesta);
                        DescargarDocumento(resol, doc, serie, docno);
                    }
                }
            }
            catch (Exception ex)
            {
                err.AddErrors(ex, "Abrir Factura Digital");
            }

            return false;
        }


        //Anular Documento
        public bool AnularDocumento(string resol, string doc, string serie, string docno)
        {
            return false;
        }




        


        


        //Conversion de Respuesta de WS a PDF
        private bool convertPDF(object pdf)
        {
            bool resultado = false;

            try
            {
                string filename = this.Resolucion + this.TipoDoc + this.Serie + this.NoDocumento + ".pdf";

                if (pdf != null)
                {
                    File.WriteAllBytes(this.localpath + filename, (byte[])pdf);

                    resultado = true;
                }
                else
                {
                    resultado = false;
                }
            }
            catch (Exception ex)
            {
                err.AddErrors(ex, "Conversion PDF");
                resultado = false;
            }

            return resultado;
            
        }


        //Carga de Datos
        private void cargarDatos()
        {

            

        }


    }
}
