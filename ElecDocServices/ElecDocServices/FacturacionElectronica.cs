using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using ElecDocServices.Helpers;
using ElecDocServices.Providers;
using FTPConnection;

namespace ElecDocServices
{
    public class FacturacionElectronica
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

        //Variables Bandera
        private enum Modo { None, Registro, Obtencion, Anulacion, };
        private Modo Funcion = Modo.None;

        //Variables de interaccion con usuario
        /// <summary>
        /// Mansaje de estado de proceso
        /// </summary>
        public string Mensaje = null;



        /// <summary>
        /// Constructor de Entrada. Instancia una sesion de facturacion electronica
        /// </summary>
        /// <param name="Config">Path de ubicacion del archivo de configuracion</param>
        /// <param name="User">Usuario de sistema</param>
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
            this.err = new Errors(this.UserSystem, "ElecDocs", errorPath);
            this.ftp = new FtpAccess(this.ConfigFile, this.UserSystem);
        }


        /// <summary>
        /// Funcion de Registro de documento segun datos enviados
        /// </summary>
        /// <param name="Resol">Resolucion de proveedor a la que pertenece el documento</param>
        /// <param name="TipoDoc">TipoDoc de documento a registrar</param>
        /// <param name="Serie">Serie del documento a registrar</param>
        /// <param name="DocNo">Correlativo del documento a registrar</param>
        /// <returns>Respuesta de éxito o fallo de registro. Ver Variable de mensaje</returns>
        //Registro de documento al proveedor
        public bool RegistrarDocumento(string Resol, string TipoDoc, string Serie, string DocNo)
        {
            bool resultado = false;
            string msg = null, cae = null;

            //Seteo de variables de negocio
            this.Resolucion = Resol;
            this.TipoDoc = TipoDoc;
            this.SerieDoc = Serie;
            this.NoDocumento = DocNo;

            //Seteo de modo de funcion
            Funcion = Modo.Registro;

            //Carga de Datos
            CargarDatos();

            try
            {
                //Obtencion de Clase del proveedor para ejecucion del proceso
                string provider = utl.convertirString(DocProvider.Rows[0]["ClassName"]);

                //Seleccion de interface segun clase de proveedor
                IFacElecInterface inter = ObtenerInterface(provider);

                //Carga de elementos en la interface
                inter.DocHeader = this.DocHeader;
                inter.DocDetail = this.DocDetail;
                inter.DocProvider = this.DocProvider;

                //Ejecucion de proceso
                List<Parameter> res = inter.RegistrarDocumento();

                //Bitacora de ejecucion
                GuardarBitacora(res);

                //Obtencion de resultados
                resultado = utl.convertirBoolean(res.FirstOrDefault(f => f.ParameterName.Equals("Resultado")).Value);
                msg = utl.convertirString(res.FirstOrDefault(f => f.ParameterName.Equals("Mensaje")).Value);
                cae = utl.convertirString(res.FirstOrDefault(f => f.ParameterName.Equals("CAE")).Value);
                string extmsg = null;

                //Verificacion de resultado exitoso
                if (resultado)
                {
                    //Actualizacion de registros
                    ActualizarRegistro(res);

                    //Ejecucion de proceso de Documento PDF
                    ProcesoPDF(res, ref extmsg);
                }
                else if (cae != "")
                {
                    //Actualizacion por si CAE no regresa vacio pero dio error
                    ActualizarRegistroCAE(res);
                }

                //Captura de mensajes
                this.Mensaje = msg + extmsg;
            }
            catch (Exception ex)
            {
                //Captura de Excepcion
                err.AddErrors(ex, "Registro de Documento");
                this.Mensaje = "Error en la Ejecución: " + ex.Message;
            }

            return resultado;
        }


        /// <summary>
        /// Funcion de obtencion de documento segun datos enviados
        /// </summary>
        /// <param name="Resol">Resolucion de proveedor a la que pertenece el documento</param>
        /// <param name="TipoDoc">TipoDoc de documento a obtener</param>
        /// <param name="Serie">Serie del documento a obtener</param>
        /// <param name="DocNo">Correlativo del documento a obtener</param>
        /// <returns>Respuesta de éxito o fallo de obtencion de documento. Ver Variable de mensaje</returns>
        //Obtencion del documento del proveedor
        public bool ObtenerDocumento(string Resol, string TipoDoc, string Serie, string DocNo)
        {
            bool resultado = false;
            string msg = null;

            //Seteo de variables de negocio
            this.Resolucion = Resol;
            this.TipoDoc = TipoDoc;
            this.SerieDoc = Serie;
            this.NoDocumento = DocNo;

            //Seteo de modo de funcion
            Funcion = Modo.Obtencion;

            //Carga de Datos
            CargarDatos();

            try
            {
                //Seteo de variables de archivo
                string filename = this.Resolucion + this.TipoDoc + this.SerieDoc + this.NoDocumento + ".pdf";
                string localpath = utl.convertirString(utl.getConfigValue(this.ConfigFile, "DOCPATH", "services"));
                string remotepath = utl.convertirString(utl.getConfigValue(this.ConfigFile, "FTP_REMOTE_PATH", "ftp"));

                
                //Verificacion de existencia de archivo en servidor
                if (ftp.fileExist(remotepath, filename))
                {
                    //Descarga de documento
                    resultado = DescargarDocumento();

                    //Verificacion de resultado
                    if (!resultado)
                    {
                        this.Mensaje = "No se pudo descargar el documento. Revisar log";
                    }
                }
                else
                {
                    //Obtencion de Clase del proveedor para ejecucion del proceso
                    string provider = utl.convertirString(DocProvider.Rows[0]["ClassName"]);

                    //Seleccion de interface segun clase de proveedor
                    IFacElecInterface inter = ObtenerInterface(provider);

                    //Carga de elementos en la interface
                    inter.DocHeader = this.DocHeader;
                    inter.DocDetail = this.DocDetail;
                    inter.DocProvider = this.DocProvider;

                    //Ejecucion de proceso
                    List<Parameter> res = inter.ObtenerDocumento();

                    //Obtencion de resultados
                    resultado = utl.convertirBoolean(res.FirstOrDefault(f => f.ParameterName.Equals("Resultado")).Value);
                    msg = utl.convertirString(res.FirstOrDefault(f => f.ParameterName.Equals("Mensaje")).Value);
                    string extmsg = null;

                    //Verificacion de resultado
                    if (resultado)
                    {
                        //Ejecucion de proceso de documento PDF
                        ProcesoPDF(res, ref extmsg);

                        if (extmsg != null)
                        {
                            //Captura de mensajes
                            this.Mensaje = msg + extmsg;
                            resultado = false;
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                //Captura de Excepcion
                err.AddErrors(ex, "Obtencion de Documento");
                this.Mensaje = "Error en la Ejecución: " + ex.Message;
            }

            return resultado;
        }


        /// <summary>
        /// Funcion de Anulacion de documentos segun datos enviados
        /// </summary>
        /// <param name="Resol">Resolucion de proveedor a la que pertenece el documento</param>
        /// <param name="TipoDoc">TipoDoc de documento a anular</param>
        /// <param name="Serie">Serie del documento a anular</param>
        /// <param name="DocNo">Correlativo del documento a anular</param>
        /// <returns>Respuesta de éxito o fallo de anulacion de documento. Ver Variable de mensaje</returns>
        //Anulacion de documento al proveedor
        public bool AnularDocumento(string Resol, string TipoDoc, string Serie, string DocNo)
        {
            bool resultado = false;
            string msg = null;

            //Seteo de variables de negocio
            this.Resolucion = Resol;
            this.TipoDoc = TipoDoc;
            this.SerieDoc = Serie;
            this.NoDocumento = DocNo;

            //Seteo de modo de funcion
            Funcion = Modo.Anulacion;

            //Carga de Datos
            CargarDatos();

            try
            {
                //Obtencion de Clase del proveedor para ejecucion del proceso
                string provider = utl.convertirString(DocProvider.Rows[0]["ClassName"]);

                //Seleccion de interface segun clase de proveedor
                IFacElecInterface inter = ObtenerInterface(provider);

                //Carga de elementos en la interface
                inter.DocHeader = this.DocHeader;
                inter.DocDetail = this.DocDetail;
                inter.DocProvider = this.DocProvider;

                //Ejecucion de proceso
                List<Parameter> res = inter.AnularDocumento();

                //Bitacora de ejecucion
                GuardarBitacora(res);

                //Obtencion de resultados
                resultado = utl.convertirBoolean(res.FirstOrDefault(f => f.ParameterName.Equals("Resultado")).Value);
                msg = utl.convertirString(res.FirstOrDefault(f => f.ParameterName.Equals("Mensaje")).Value);
                string extmsg = null;

                if (resultado)
                {
                    //Actualizacion de registro de anulacion
                    ActualizarRegsitroAnulacion(res);
                }

                //Captura de mensajes
                this.Mensaje = msg + extmsg;
            }
            catch (Exception ex)
            {
                //Captura de Excepcion
                err.AddErrors(ex, "Obtencion de Documento");
                this.Mensaje = "Error en la Ejecución: " + ex.Message;
            }

            return resultado;

        }



        #region Funciones


        //Carga de datos del documento
        private void CargarDatos()
        {
            // Encabezado

            string campos = " dh.Resolucion, dh.Documento, dh.Serie, dh.DocNo, dh.ReferenciaDoc, dh.Fecha, dh.Empresa, ";
            campos += " dh.Sucursal, dh.Caja, dh.Usuario, dh.Divisa, bg.Name Moneda, dh.TasaCambio, dh.TipoGeneracion, ";
            campos += " dh.NombreCliente, dh.DireccionCliente, dh.NITCliente, dh.ValorNeto, dh.IVA, dh.Descuento,  ";
            campos += " dh.Exento, dh.Total, dh.Estado, dh.Observacion, dh.Estado ";

            string[] join = { "left join Badge bg on dh.Divisa = bg.BadgeID " };

            string whr = "dh.Resolucion = '" + this.Resolucion + "' and dh.Documento = '" + this.TipoDoc + "' ";
            whr += " and dh.Serie = '" + this.SerieDoc + "' and dh.DocNo = '" + this.NoDocumento + "' ";

            DocHeader = con.obtenerDatos("documentheader dh", campos, join, whr);

            // Detalle

            campos = "Resolucion, Documento, Serie, DocNo, Linea, Codigo, Descripcion, Tipo, Metrica, ";
            campos += "Cantidad, PrecioUnitario, Importe, IVA, Total, Exento";

            whr = "Resolucion = '" + this.Resolucion + "' and Documento = '" + this.TipoDoc + "' ";
            whr += " and Serie = '" + this.SerieDoc + "' and DocNo = '" + this.NoDocumento + "' ";

            DocDetail = con.obtenerDatos("documentdetail", campos, whr);

            // Datos Proveedor

            campos = " cd.Resolucion, cd.Fecha, cd.Documento, cd.Serie, cd.Ultimo, cd.Liquidado, ";
            campos += " df.FormatName, sp.Name Proveedor, sp.ClassName, sp.ProviderAuth, sp.ServiceURL ";

            join = new string[]
            {
                "left join serviceprovider sp on cd.ProviderID = sp.ProviderID",
                "left join docformat df on cd.Formato = df.FormatID",
            };

            whr = " cd.Documento = '" + this.TipoDoc + "' and cd.Resolucion = '" + this.Resolucion + "' ";
            whr += " and cd.Serie = '" + this.SerieDoc + "' ";

            DocProvider = con.obtenerDatos("corrdocument cd", campos, join, whr);


            //Validacion de Documento encontrado
            if (DocHeader.Rows.Count.Equals(0) || DocDetail.Rows.Count.Equals(0))
            {
                throw new Exception("No se encontró el documento");
            }
            else if (DocHeader.Rows.Count > 0)
            {
                //Verificacion del modo registro
                if (Funcion.Equals(Modo.Registro))
                {
                    //Verificacion de estado de documento
                    if (utl.convertirInt(DocHeader.Rows[0]["Estado"]).Equals(1))
                    {
                        throw new Exception("Documento se encuentra registrado");
                    }
                    else if (utl.convertirInt(DocHeader.Rows[0]["Estado"]).Equals(2))
                    {
                        throw new Exception("Documento se encuentra anulado");
                    }
                }
            }
            
            //Validacion de Resolucion de proceedor encontrada
            if (DocProvider.Rows.Count.Equals(0))
            {
                throw new Exception("No se encontró resolución de documento");
            }
            else if (DocProvider.Rows.Count > 0)
            {
                //Verificacion del modo registro
                if (Funcion.Equals(Modo.Registro))
                {
                    //Validacion de resolucion de proveedor aun vigente
                    if (utl.convertirInt(DocProvider.Rows[0]["Liquidado"]).Equals(1))
                    {
                        throw new Exception("Resolución de documento está liquidado");
                    }
                }
            }
        }


        //Guardado de datos en bitacora en base de datos
        private void GuardarBitacora(List<Parameter> par)
        {
            //Campos a guardar
            string[] campos = new string[]
            {
                "Fecha", "Usuario", "HostName", "IpAddresses", "MacAddresses", "ServiceUrl", "FunctionName",
                "Resolucion", "Documento", "Serie", "DocNo", "XmlData", "OtherParameters", "Resultado",
                "Descripcion", "Respuesta", "DocTrib", "CAE", "CAEC", "OtherResponse"
            };

            //Obtencion de datos a guardar
            string ips = utl.convertirArraytoString(utl.getIpAddress().ToArray());
            string macs = utl.convertirArraytoString(utl.getMACAddress().ToArray());
            string url = utl.convertirString(this.DocProvider.Rows[0]["ServiceURL"]);
            string function = utl.convertirString(par.FirstOrDefault(f => f.ParameterName.Equals("Function")).Value);
            string xml = utl.convertirString(par.FirstOrDefault(f => f.ParameterName.Equals("XML")).Value);
            bool res = utl.convertirBoolean(par.FirstOrDefault(f => f.ParameterName.Equals("Resultado")).Value);
            string mensaje = utl.convertirString(par.FirstOrDefault(f => f.ParameterName.Equals("Mensaje")).Value);
            string iddoc = utl.convertirString(par.FirstOrDefault(f => f.ParameterName.Equals("IDDoc")).Value);
            string cae = utl.convertirString(par.FirstOrDefault(f => f.ParameterName.Equals("CAE")).Value);
            string caec = utl.convertirString(par.FirstOrDefault(f => f.ParameterName.Equals("CAEC")).Value);

            //Arreglo de datos
            object[] datos = new object[]
            {
                utl.formatoFechaSql(DateTime.Now, true), this.UserSystem, utl.getHostName(), ips, macs, url, function, 
                this.Resolucion, this.TipoDoc, this.SerieDoc, this.NoDocumento, xml, null, res,
                mensaje, null, iddoc, cae, caec, null
            };

            //Insercion de bitacora
            bool resultado = con.execInsert("logelecdocument", campos, datos);

            //Verificacion de resultado
            if (!resultado)
            {
                err.AddErrors("No se ingreso bitacora de documento", "", null, this.UserSystem);
            }
        }


        //Actualizacion de registro de documento
        private void ActualizarRegistro(List<Parameter> par)
        {
            //Obtencion de datos para actualizar
            string doc = utl.convertirString(par.FirstOrDefault(f => f.ParameterName.Equals("IDDoc")).Value);
            string cae = utl.convertirString(par.FirstOrDefault(f => f.ParameterName.Equals("CAE")).Value);
            string caec = utl.convertirString(par.FirstOrDefault(f => f.ParameterName.Equals("CAEC")).Value);
            bool allowFtp = utl.convertirBoolean(utl.getConfigValue(this.ConfigFile, "FTP", "services"));
            string filename = this.Resolucion + this.TipoDoc + this.SerieDoc + this.NoDocumento + ".pdf";
            string remotepath = utl.convertirString(utl.getConfigValue(this.ConfigFile, "FTP_REMOTE_PATH", "ftp"));
            string path = (allowFtp) ? (remotepath + filename) : null;

            //Preparacion de actualizacion de registro
            string[] campos = { "Registrado", "FechaRegistro", "UsuarioRegistro", "RefDocTributario", "DocumentPath", "Estado", "CAE", "CAEC" };
            object[] datos = { 1, utl.formatoFechaSql(DateTime.Now, true), this.UserSystem, doc, path, 1, cae, caec };
            string whr = "Resolucion = '" + this.Resolucion + "' and Documento = '" + this.TipoDoc + "' ";
            whr += " and Serie = '" + this.SerieDoc + "' and DocNo = '" + this.NoDocumento + "' ";

            //Actualizacion de registro
            bool res = con.execUpdate("documentheader", campos, datos, whr);

            //Verificacion de actualizacion de registro
            if (res)
            {
                int num = (utl.convertirInt(this.NoDocumento) + 1);

                //Actualizacion de correlativo
                res = con.execUpdate("corrdocument", "Ultimo", num, "Resolucion = '" + this.Resolucion + "' "); 

                //Verificacion de actualizacion de correlativo
                if (!res)
                {
                    err.AddErrors("No se actualizo registro de correlativo de documento", null, null, this.UserSystem);
                }
            }
            else
            {
                err.AddErrors("No se actualizo registro de documento", null, null, this.UserSystem);
            }
        }


        //Actualizacion de registro de documento si en caso solo se obtiene CAE
        private void ActualizarRegistroCAE(List<Parameter> par)
        {
            string doc = utl.convertirString(par.FirstOrDefault(f => f.ParameterName.Equals("IDDoc")).Value);
            string cae = utl.convertirString(par.FirstOrDefault(f => f.ParameterName.Equals("CAE")).Value);
            string caec = utl.convertirString(par.FirstOrDefault(f => f.ParameterName.Equals("CAEC")).Value);

            string[] campos = { "RefDocTributario", "CAE", "CAEC" };
            object[] datos = { doc,cae, caec };
            string whr = "Resolucion = '" + this.Resolucion + "' and Documento = '" + this.TipoDoc + "' ";
            whr += " and Serie = '" + this.SerieDoc + "' and DocNo = '" + this.NoDocumento + "' ";

            //Actualizacion de registro
            bool res = con.execUpdate("documentheader", campos, datos, whr);
        }


        //Actualizacion de registro (anulacion) de documento
        private void ActualizarRegsitroAnulacion(List<Parameter> par)
        {
            //Preparacion de actualizacion de registro
            string[] campos = { "FechaAnulacion", "UsuarioAnulacion", "Estado" };
            object[] datos = { utl.formatoFechaSql(DateTime.Now, true), this.UserSystem, 2 };
            string whr = "Resolucion = '" + this.Resolucion + "' and Documento = '" + this.TipoDoc + "' ";
            whr += " and Serie = '" + this.SerieDoc + "' and DocNo = '" + this.NoDocumento + "' ";

            //Actualizacion de registro
            bool res = con.execUpdate("documentheader", campos, datos, whr);

            //Verificacion de actualizacion de registro
            if (!res)
            {
                err.AddErrors("No se actualizo registro de documento", null, null, this.UserSystem);
            }
        }


        //Funcion para la preparacion de guardado de documento PDF
        private void ProcesoPDF(List<Parameter> par, ref string extmsg)
        {
            //Obtencion de documento PDF Bytes
            object pdf = par.FirstOrDefault(f => f.ParameterName.Equals("Respuesta")).Value;
            bool allowPdf = utl.convertirBoolean(par.FirstOrDefault(f => f.ParameterName.Equals("PDF")).Value);

            //Verificacion de permiso de guardado de PDF por servicio
            if (allowPdf)
            {
                //Verificacion de Guardado de PDF localmente
                if (GuardarPDF(pdf))
                {
                    //Verificacion de permiso de guardado en servidor
                    if (utl.convertirBoolean(utl.getConfigValue(ConfigFile, "FTP", "services")))
                    {
                        //Verificacion de subida de documento
                        if (!CargarDocumento())
                        {
                            extmsg = " - Error en proceso, ver log.";
                            err.AddErrors("Error en carga de documento a servidor", null, null, this.UserSystem);
                        }
                    }
                }
                else
                {
                    //Guardado de error de proceso
                    extmsg = " - Error en proceso, ver log.";
                    err.AddErrors("Error en guardado de documento", null, null, this.UserSystem);
                }
            }
        }


        //Funcion de conversion de datos a archivo PDF
        private bool GuardarPDF(object data)
        {
            bool resultado = false;

            try
            {
                //Preparacion de ubicacion y nombre de archivo
                string filename = this.Resolucion + this.TipoDoc + this.SerieDoc + this.NoDocumento + ".pdf";
                string localpath = utl.convertirString(utl.getConfigValue(this.ConfigFile, "DOCPATH", "services"));

                //Verificacion de datos nulos
                if (data != null)
                {
                    //Guardado de documento
                    File.WriteAllBytes((localpath + filename), (byte[])data);
                    resultado = true;

                    //Apertura de documento
                    if (!utl.openFile((localpath + filename)))
                    {
                        //Verificacion de modo de proceso
                        if (this.Funcion.Equals(Modo.Obtencion))
                        {
                            //Guardado de log
                            err.AddErrors("Error en apertura de archivo", null, null, this.UserSystem);
                        }
                    }
                }
                else
                {
                    //Guardado de bitacora
                    resultado = false;
                    err.AddErrors("No hay datos que convertir a PDF", null, null, this.UserSystem);
                }
            }
            catch (Exception ex)
            {
                //Captura de excepcion
                err.AddErrors(ex, "Conversion PDF");
                resultado = false;
            }

            return resultado;
            
        }


        //Carga de documento PDF a servidor FTP
        private bool CargarDocumento()
        {
            bool resultado = false;

            try
            {
                //Preparacion de nombre y ubicacion de archivos
                string filename = this.Resolucion + this.TipoDoc + this.SerieDoc + this.NoDocumento + ".pdf";
                string localpath = utl.convertirString(utl.getConfigValue(this.ConfigFile, "DOCPATH", "services"));
                string remotepath = utl.convertirString(utl.getConfigValue(this.ConfigFile, "FTP_REMOTE_PATH", "ftp"));

                //Carga de archivo
                resultado = ftp.uploadWinFtp(remotepath, (localpath + filename));
            }
            catch (Exception ex)
            {
                //Captura de excepcion
                err.AddErrors(ex, "Guardar documento en servidor");
            }

            return resultado;
        }


        //Descarga de documento PDF desde servidor FTP
        private bool DescargarDocumento()
        {
            bool resultado = false;

            try
            {
                //Preparacion de nombre y ubicacion de archivo
                string filename = this.Resolucion + this.TipoDoc + this.SerieDoc + this.NoDocumento + ".pdf";
                string localpath = utl.convertirString(utl.getConfigValue(this.ConfigFile, "DOCPATH", "services"));
                string remotepath = utl.convertirString(utl.getConfigValue(this.ConfigFile, "FTP_REMOTE_PATH", "ftp"));

                //Descarga de documento
                resultado = ftp.downloadWinFtp((remotepath + filename), (localpath + filename));

                //Verificacion de resultado de descarga
                if (resultado)
                {
                    //Apertura de archivo
                    utl.openFile((localpath + filename));
                }
            }
            catch (Exception ex)
            {
                //Captura de excpecion
                err.AddErrors(ex, "Descargar documento de servidor");
            }

            return resultado;
        }


        //Obtencion de interface de servicio de facturacion electronica
        private IFacElecInterface ObtenerInterface(string proveedor)
        {
            //Busqueda de clase proveedor para instanciar interface
            var _className = (from t in System.Reflection.Assembly.GetExecutingAssembly().GetTypes()
                              where t.Name.Equals(proveedor)
                              select t.FullName).Single();

            //Instancia de interface
            IFacElecInterface _class = (IFacElecInterface)Activator.CreateInstance(Type.GetType(_className)); 

            //Retorno de interface
            return _class;
        }



        #endregion

    }
}
 