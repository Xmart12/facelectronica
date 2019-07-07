using System;
using System.Data;
using System.IO;
using ElecDocServices.Helpers;
using ElecDocServices.Providers;
using FTPConnection;

namespace ElecDocServices
{
    class FacturacionElectronica
    {
        //Clases Globales
        private Utils utl = null;
        private Errors err = null;
        private ConMySQL con = null;
        private FtpAccess ftp = null;

        //Variables de Control
        private string UserSystem = null;
        private string ConfigFile = null;

        //Variables de Negocio
        private string Resolucion = null;
        private string TipoDoc = null;
        private string SerieDoc = null;
        private string DocNo = null;

        //Variables de Almacenamiento
        private DataTable DocHeader = null;
        private DataTable DocDetail = null;


        /// <summary>
        /// 
        /// </summary>
        /// <param name="Config"></param>
        /// <param name="User"></param>
        //Constructor de Entrada
        public FacturacionElectronica(string Config, string User)
        {
            //Verificacion de existencia de archivo de configuracion
            if (!File.Exists(Config))
            {
                throw new Exception("Archivo de configuracion no existe");
            }

            //Seteo de varuales de control
            this.ConfigFile = Config;
            this.UserSystem = User;

            //Intanciamiento de clases globales
            this.utl = new Utils();
            string errorPath = utl.convertirString(utl.getConfigValue(this.ConfigFile, "SERVICES", "errors"));
            this.con = new ConMySQL(this.ConfigFile, this.UserSystem);
            this.err = new Errors(this.UserSystem, "FacturacionElectronica", "");

        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="Res"></param>
        /// <param name="Tipo"></param>
        /// <param name="Serie"></param>
        /// <param name="NoDoc"></param>
        public void CargarDocumento(string Res, string Tipo, string Serie, string NoDoc)
        {
            //Seteo de variables de negocio
            this.Resolucion = Res;
            this.TipoDoc = Tipo;
            this.SerieDoc = Serie;
            this.DocNo = NoDoc;
        }

        private void CargarDatos()
        {
            // Encabezado

            string campos = "Resolucion, Documento, Serie, DocNo, ReferenciaDoc, Fecha, Empresa, Sucursal, Caja, ";
            campos += " Usuario, Divisa, TasaCambio, TipoGeneracion, NombreCliente, DireccionCliente, NITCliente, ";
            campos += " ValorNeto, IVA, Descuento, Exento, Total";

            string whr = "Resolucion = '" + this.Resolucion + "' and Documento = '" + this.TipoDoc + "' ";
            //whr += " and Serie = '" + this.Serie + "' and DocNo = '" + this.NoDocumento + "' ";

            //dtEncabezado = con.obtenerDatos("documentheader", campos, whr);

            // Detalle

            campos = "Resolucion, Documento, Serie, DocNo, Linea, Codigo, Descripcion, Tipo, Metrica, ";
            campos += "Cantidad, PrecioUnitario, Importe, IVA, Total, Exento";

            whr = "Resolucion = '" + this.Resolucion + "' and Documento = '" + this.TipoDoc + "' ";
            //whr += " and Serie = '" + this.Serie + "' and DocNo = '" + this.NoDocumento + "' ";

            //dtDetalle = con.obtenerDatos("documentdetail", campos, whr);
        }

    }
}
