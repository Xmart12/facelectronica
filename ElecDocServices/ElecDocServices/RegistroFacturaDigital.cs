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

            bool resultado = false;
            List<string> msg = new List<string>();

            try
            {
                //Carga de datos
                this.Resolucion = resol;
                this.TipoDoc = doc;
                this.Serie = serie;
                this.NoDocumento = utl.convertirInt(docno);

                cargarDatos();

                SSO_wsEFactura service = new SSO_wsEFactura();

                string xml = construirXML();

                clsResponseGeneral res = service.RegistraFacturaXML_PDF(xml);
                resultado = res.pResultado;
                this.mensaje = "Respuesta WS: " + res.pDescripcion;


                if (resultado)
                {
                    if (!convertPDF(res.pRespuesta))
                    {
                        err.AddErrors("No se pudo realizar conversion de PDF", "", "", this.usuario);
                    }
                }
               
            }
            catch (Exception ex)
            {
                err.AddErrors(ex, "Registro de Factura Digital");
            }

            return resultado;
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
                    SSO_wsEFactura service = new SSO_wsEFactura();

                    string xml = construirXMLRetorno();

                    clsResponseGeneral res = service.RetornaDatosFacturaXML_PDF(xml);
                    resultado = res.pResultado;

                    if (resultado)
                    {
                        resultado = convertPDF(res.pRespuesta);
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




        //Construccion de XML
        private string construirXML()
        {
            string xml = "";

            try
            {
                XmlDocument doc = new XmlDocument();

                XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "utf-8", null);
                XmlElement root = doc.DocumentElement;
                doc.InsertBefore(xmlDeclaration, root);

                //FACTURA
                XmlElement factura = utl.createXmlNode(doc, "FACTURA");
                doc.AppendChild(factura);

                //ENCABEZADO
                XmlElement encabezado = utl.createXmlNode(doc, "ENCABEZADO");
                factura.AppendChild(encabezado);

                //OPCIONAL
                XmlElement opcional = utl.createXmlNode(doc, "OPCIONAL");
                factura.AppendChild(opcional);

                //DETALLE
                XmlElement detalle = utl.createXmlNode(doc, "DETALLE");
                factura.AppendChild(detalle);

                double conversion = 1;

                //Existencia de Datos de ENCABEZADO
                if (this.dtEncabezado.Rows.Count > 0)
                {
                    DataRow r = dtEncabezado.Rows[0];

                    //Configuracion de Calculos
                    double TasaCambio = utl.convertirDouble(r["TasaCambio"]);
                    conversion = (TasaCambio > 0) ? (1 / TasaCambio) : 1;

                    double total = utl.convertirDouble(r["Total"]);
                    double valorneto = utl.convertirDouble(r["ValorNeto"]);
                    double valoriva = utl.convertirDouble(r["IVA"]);
                    double descuento = utl.convertirDouble(r["Descuento"]);
                    double exento = utl.convertirDouble(r["Exento"]);

                    string nit = (this.TipoDoc == "Ex" || this.TipoDoc == "Ee") ? "CF" : utl.convertirString(r["NITCliente"]);

                    //Etiqueta ENCABEZADO
                    encabezado.AppendChild(utl.createXmlNode(doc, "NOFACTURA", utl.convertirString(this.NoDocumento)));
                    encabezado.AppendChild(utl.createXmlNode(doc, "RESOLUCION", this.Resolucion));
                    encabezado.AppendChild(utl.createXmlNode(doc, "IDSERIE", this.Serie));
                    encabezado.AppendChild(utl.createXmlNode(doc, "EMPRESA", utl.convertirString(r["Empresa"])));
                    encabezado.AppendChild(utl.createXmlNode(doc, "SUCURSAL", utl.convertirString(r["Sucursal"])));
                    encabezado.AppendChild(utl.createXmlNode(doc, "CAJA", utl.convertirString(r["Caja"])));
                    encabezado.AppendChild(utl.createXmlNode(doc, "USUARIO", utl.convertirString(r["Usuario"])));
                    encabezado.AppendChild(utl.createXmlNode(doc, "MONEDA", utl.convertirString(r["Divisa"])));
                    encabezado.AppendChild(utl.createXmlNode(doc, "TASACAMBIO", utl.convertirString(r["TasaCambio"])));
                    encabezado.AppendChild(utl.createXmlNode(doc, "GENERACION", utl.convertirString(r["TipoGeneracion"])));
                    encabezado.AppendChild(utl.createXmlNode(doc, "FECHAEMISION", utl.formatoFecha(r["Fecha"])));
                    encabezado.AppendChild(utl.createXmlNode(doc, "NOMBRECONTRIBUYENTE", utl.convertirString(r["NombreCliente"]), true));
                    encabezado.AppendChild(utl.createXmlNode(doc, "DIRECCIONCONTRIBUYENTE", utl.convertirString(r["DireccionCliente"]), true));
                    encabezado.AppendChild(utl.createXmlNode(doc, "NITCONTRIBUYENTE", nit));
                    encabezado.AppendChild(utl.createXmlNode(doc, "VALORNETO", utl.formatoCurrencySS(valorneto * conversion)));
                    encabezado.AppendChild(utl.createXmlNode(doc, "IVA", utl.formatoCurrencySS(valoriva * conversion)));
                    encabezado.AppendChild(utl.createXmlNode(doc, "TOTAL", utl.formatoCurrencySS(total * conversion)));
                    encabezado.AppendChild(utl.createXmlNode(doc, "DESCUENTO", utl.formatoCurrencySS(descuento * conversion)));
                    encabezado.AppendChild(utl.createXmlNode(doc, "EXENTO", utl.formatoCurrencySS(exento * conversion)));
                    
                    if (this.TipoDoc == "Ex" || this.TipoDoc == "Ee")
                    {
                        opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL5", utl.convertirString(r["NITCliente"]), true));
                        opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL9", utl.convertirString(r["ReferenciaDoc"]), true));
                    }
                    else
                    {
                        opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL5"));
                        opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL6"));
                        opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL7", utl.convertirString(r["ReferenciaDoc"])));
                        opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL8"));
                        opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL9", "SUJETO A RETENCION DEFINITIVA", true));
                    }

                    opcional.AppendChild(utl.createXmlNode(doc, "TOTAL_LETRAS", utl.formatoNumeroALetras((total * conversion), 2, true, "Quetzales", true).ToUpper(), true));
                }

                //Existencia de datos de Detalle de documento
                if (this.dtDetalle.Rows.Count > 0)
                {
                    //Etiqueta DETALLE
                    foreach (DataRow rf in this.dtDetalle.Rows)
                    {
                        string[] desc = utl.convertirString(rf["Descripcion"]).Split('\n');

                        for (int i = 0; i < desc.Length; i++)
                        {
                            //LINEA
                            XmlElement linea = utl.createXmlNode(doc, "LINEA");
                            detalle.AppendChild(linea);

                            string metrica = (i == 0) ? utl.convertirString(rf["Metrica"]) : "";
                            string tipodet = (i == 0) ? utl.convertirString(rf["Tipo"]) : "";
                            string exento = (i == 0) ? ((utl.convertirBoolean(rf["Exento"])) ? "S" : "N") : "N";
                            string codigo = (i == 0) ? utl.convertirString(rf["Codigo"]) : "";

                            double precioUnitario = (i == 0) ? utl.convertirDouble(rf["PrecioUnitario"]) : 0;
                            double valor = (i == 0) ? utl.convertirDouble(rf["Importe"]) : 0;
                            double cantidad = (i == 0) ? utl.convertirDouble(rf["Cantidad"]) : 0;

                            //Etiqueta LINEA
                            linea.AppendChild(utl.createXmlNode(doc, "CANTIDAD", utl.convertirString(cantidad)));
                            linea.AppendChild(utl.createXmlNode(doc, "DESCRIPCION", desc[i], true));
                            linea.AppendChild(utl.createXmlNode(doc, "METRICA", metrica));
                            linea.AppendChild(utl.createXmlNode(doc, "PRECIOUNITARIO", utl.formatoCurrencySS(precioUnitario * conversion)));
                            linea.AppendChild(utl.createXmlNode(doc, "VALOR", utl.formatoCurrencySS(valor * conversion)));
                            linea.AppendChild(utl.createXmlNode(doc, "TIPO_PRODUCTO", tipodet, true));
                            linea.AppendChild(utl.createXmlNode(doc, "EXENTO", exento, true));
                            linea.AppendChild(utl.createXmlNode(doc, "DETALLE1", codigo));
                        }
                    }
                }

                xml = utl.getXmlString(doc);
            }
            catch (Exception ex)
            {
                err.AddErrors(ex, "Construccion de Factura XML");
                xml = "";
            }

            return xml;
        }


        //Construccion de XML para retorno de Datos
        private string construirXMLRetorno()
        {
            string xml = "";

            try
            {
                XmlDocument doc = new XmlDocument();

                XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "utf-8", null);
                XmlElement root = doc.DocumentElement;
                doc.InsertBefore(xmlDeclaration, root);

                //FACTURA
                XmlElement factura = utl.createXmlNode(doc, "FACTURA");
                doc.AppendChild(factura);


                factura.AppendChild(utl.createXmlNode(doc, "NODOCUMENTO", utl.convertirString(this.NoDocumento)));
                factura.AppendChild(utl.createXmlNode(doc, "SERIE", this.Serie));
                factura.AppendChild(utl.createXmlNode(doc, "EMPRESA", "1003"));
                factura.AppendChild(utl.createXmlNode(doc, "SUCURSAL", "1"));

                //Existencia de Datos de ENCABEZADO
                if (this.dtEncabezado.Rows.Count > 0)
                {
                    DataRow r = dtEncabezado.Rows[0];

                }

                xml = utl.getXmlString(doc);
            }
            catch (Exception ex)
            {
                err.AddErrors(ex, "Construccion de Factura XML");
                xml = "";
            }

            return xml;
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

            // Encabezado

            string campos = "Resolucion, Documento, Serie, DocNo, ReferenciaDoc, Fecha, Empresa, Sucursal, Caja, ";
            campos += " Usuario, Divisa, TasaCambio, TipoGeneracion, NombreCliente, DireccionCliente, NITCliente, ";
            campos += " ValorNeto, IVA, Descuento, Exento, Total";

            string whr = "Resolucion = '" + this.Resolucion + "' and Documento = '" + this.TipoDoc + "' "; 
            whr += " and Serie = '" + this.Serie + "' and DocNo = '" + this.NoDocumento + "' ";

            dtEncabezado = con.obtenerDatos("documentheader", campos, whr);

            // Detalle

            campos = "Resolucion, Documento, Serie, DocNo, Linea, Codigo, Descripcion, Tipo, Metrica, ";
            campos += "Cantidad, PrecioUnitario, Importe, IVA, Total, Exento";

            whr = "Resolucion = '" + this.Resolucion + "' and Documento = '" + this.TipoDoc + "' ";
            whr += " and Serie = '" + this.Serie + "' and DocNo = '" + this.NoDocumento + "' ";

            dtDetalle = con.obtenerDatos("documentdetail", campos, whr);

        }


    }
}
