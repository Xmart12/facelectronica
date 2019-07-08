using System;
using System.Data;
using System.IO;
using System.Linq;
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
        private string NoDocumento = null;

        //Variables de Almacenamiento
        private DataTable DocHeader = null;
        private DataTable DocDetail = null;
        private DataTable DocProvider = null;

        public string Mensaje = null;


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


        public bool RegistrarDocumento(string Res, string Tipo, string Serie, string DocNo)
        {
            //Seteo de variables de negocio
            this.Resolucion = Res;
            this.TipoDoc = Tipo;
            this.SerieDoc = Serie;
            this.NoDocumento = DocNo;

            CargarDatos();

            

            return false;
        }


        public bool ObtenerDocumento(string Res, string TipoDoc, string Serie, string DocNo)
        {
            return false;
        }


        public bool AnularDocumento(string Res, string TipoDoc, string Serie, string DocNo)
        {
            return false;
        }



        private void CargarDatos()
        {
            // Encabezado

            string campos = "Resolucion, Documento, Serie, DocNo, ReferenciaDoc, Fecha, Empresa, Sucursal, Caja, ";
            campos += " Usuario, Divisa, TasaCambio, TipoGeneracion, NombreCliente, DireccionCliente, NITCliente, ";
            campos += " ValorNeto, IVA, Descuento, Exento, Total, Estado, Observacion";

            string whr = "Resolucion = '" + this.Resolucion + "' and Documento = '" + this.TipoDoc + "' ";
            whr += " and Serie = '" + this.SerieDoc + "' and DocNo = '" + this.NoDocumento + "' ";

            DocHeader = con.obtenerDatos("documentheader", campos, whr);

            // Detalle

            campos = "Resolucion, Documento, Serie, DocNo, Linea, Codigo, Descripcion, Tipo, Metrica, ";
            campos += "Cantidad, PrecioUnitario, Importe, IVA, Total, Exento";

            whr = "Resolucion = '" + this.Resolucion + "' and Documento = '" + this.TipoDoc + "' ";
            whr += " and Serie = '" + this.SerieDoc + "' and DocNo = '" + this.NoDocumento + "' ";

            DocDetail = con.obtenerDatos("documentdetail", campos, whr);

            // Datos Proveedor
            campos = " cd.Resolucion, cd.Fecha, cd.Documento, cd.Serie, cd.Ultimo, df.FormatName, ";
            campos += " sp.Name Proveedor, sp.ClassName, sp.ProviderAuth ";

            string[] join = new string[]
            {
                "left join serviceprovider sp on cd.ProviderID = sp.ProviderID",
                "left join docformat df on cd.Formato = df.FormatID",
            };

            if (DocHeader.Rows.Count.Equals(0) || DocDetail.Rows.Count.Equals(0))
            {
                throw new Exception("No se encontró el documento");
            }
            
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="proveedor"></param>
        /// <returns></returns>
        private IFacElecInterface ObtenerInterface(string proveedor)
        {
            var _className = (from t in System.Reflection.Assembly.GetExecutingAssembly().GetTypes()
                              where t.Name.Equals(proveedor)
                              select t.FullName).Single();

            IFacElecInterface _class = (IFacElecInterface)Activator.CreateInstance(Type.GetType(_className)); 

            return _class;
        }


    }
}
