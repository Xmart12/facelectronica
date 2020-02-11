using System;
using System.Collections.Generic;
using System.Data;
using System.Xml;
using ElecDocServices.Helpers;
using ElecDocServices.IfacereServices;

namespace ElecDocServices.Providers
{
    internal class Ifacere : IFacElecInterface
    {
        private Utils utl = new Utils();

        public DataTable DocHeader { get; set; }
        public DataTable DocDetail { get; set; }
        public DataTable DocProvider { get; set; }



        //Funcion de comunicacion con Ifacere para Registro de documento
        public List<Parameter> RegistrarDocumento()
        {
            clsResponseGeneral res = new clsResponseGeneral();
            string xml = null;

            try
            {
                SSO_wsEFactura service = new SSO_wsEFactura();
                xml = ConstruirXMLRegistro();
                res = service.RegistraFacturaXML_PDF(xml);
            }
            catch (Exception ex)
            {
                res.pResultado = false;
                res.pDescripcion = ex.Message;
            }

            return ObtenerDatosResultado(res, xml, "RegistraFacturaXML_PDF");
        }


        //
        public List<Parameter> RegistrarDocNC()
        {
            throw new NotImplementedException();
        }


        //Funcion de comunicacion con Ifacere para Obtencion de datos de documento registrado
        public List<Parameter> ObtenerDocumento()
        {
            clsResponseGeneral res = new clsResponseGeneral();
            string xml = null;

            try
            {
                SSO_wsEFactura service = new SSO_wsEFactura();
                xml = ConstruirXMLRetorno();
                res = service.RetornaDatosFacturaXML_PDF(xml);
            }
            catch (Exception ex)
            {
                res.pResultado = false;
                res.pDescripcion = ex.Message;
            }

            return ObtenerDatosResultado(res, xml, "RetornaDatosFacturaXML_PDF");
        }


        //Funcion de comunicacion con Ifacere para Anulacion de documento
        public List<Parameter> AnularDocumento()
        {
            clsResponseGeneral res = new clsResponseGeneral();
            string xml = null;

            try
            {
                SSO_wsEFactura service = new SSO_wsEFactura();
                xml = ConstruirXMLAnulacion();
                res = service.AnulacionFacturaXML(xml);
            }
            catch (Exception ex)
            {
                res.pResultado = false;
                res.pDescripcion = ex.Message;
            }

            return ObtenerDatosResultado(res, xml, "AnulacionFacturaXML");
        }


        //
        public List<Parameter> AnularDocNC()
        {
            throw new NotImplementedException();
        }



        #region Funciones



        //Construccion de XML
        private string ConstruirXMLRegistro()
        {
            string xml = "";

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

            //Existencia de Datos de ENCABEZADO
            if (DocHeader.Rows.Count > 0)
            {
                DataRow r = DocHeader.Rows[0];

                //Etiqueta ENCABEZADO
                encabezado.AppendChild(utl.createXmlNode(doc, "NOFACTURA", utl.convertirString(r["DocNo"])));
                encabezado.AppendChild(utl.createXmlNode(doc, "RESOLUCION", utl.convertirString(r["Resolucion"])));
                encabezado.AppendChild(utl.createXmlNode(doc, "IDSERIE", utl.convertirString(r["Serie"])));
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
                encabezado.AppendChild(utl.createXmlNode(doc, "NITCONTRIBUYENTE", utl.convertirString(r["NITCliente"])));
                encabezado.AppendChild(utl.createXmlNode(doc, "VALORNETO", utl.formatoCurrencySS(utl.convertirDouble(r["ValorNeto"]))));
                encabezado.AppendChild(utl.createXmlNode(doc, "IVA", utl.formatoCurrencySS(utl.convertirDouble(r["IVA"]))));
                encabezado.AppendChild(utl.createXmlNode(doc, "TOTAL", utl.formatoCurrencySS(utl.convertirDouble(r["Total"]))));
                encabezado.AppendChild(utl.createXmlNode(doc, "DESCUENTO", utl.formatoCurrencySS(utl.convertirDouble(r["Descuento"]))));
                encabezado.AppendChild(utl.createXmlNode(doc, "EXENTO", utl.formatoCurrencySS(utl.convertirDouble(r["Exento"]))));

                //Etiqueta OPCIONAL
                opcional.AppendChild(utl.createXmlNode(doc, "TOTAL_LETRAS", utl.formatoNumeroALetras(utl.convertirDouble(r["Total"]), 2, true, utl.convertirString(r["Moneda"]), true).ToUpper(), true));
            }

            //Existencia de datos de Detalle de documento
            if (DocDetail.Rows.Count > 0)
            {
                //Etiqueta DETALLE
                foreach (DataRow rf in DocDetail.Rows)
                {
                    //LINEA
                    XmlElement linea = utl.createXmlNode(doc, "LINEA");
                    detalle.AppendChild(linea);

                    string tipodet = utl.convertirString(rf["Tipo"]);
                    string exento = (utl.convertirBoolean(rf["Exento"])) ? "S" : "N";
                    string codigo = utl.convertirString(rf["Codigo"]);
                    double precioUnitario = utl.convertirDouble(rf["PrecioUnitario"]);
                    double valor = utl.convertirDouble(rf["Importe"]);
                    double cantidad = utl.convertirDouble(rf["Cantidad"]);

                    //Etiqueta LINEA
                    linea.AppendChild(utl.createXmlNode(doc, "CANTIDAD", utl.convertirString(cantidad)));
                    linea.AppendChild(utl.createXmlNode(doc, "DESCRIPCION", utl.convertirString(rf["Descripcion"]), true));
                    linea.AppendChild(utl.createXmlNode(doc, "METRICA", utl.convertirString(rf["Metrica"])));
                    linea.AppendChild(utl.createXmlNode(doc, "PRECIOUNITARIO", utl.formatoCurrencySS(precioUnitario)));
                    linea.AppendChild(utl.createXmlNode(doc, "VALOR", utl.formatoCurrencySS(valor)));
                    linea.AppendChild(utl.createXmlNode(doc, "TIPO_PRODUCTO", tipodet, true));
                    linea.AppendChild(utl.createXmlNode(doc, "EXENTO", exento, true));
                    linea.AppendChild(utl.createXmlNode(doc, "DETALLE1", codigo));
                }
            }

            xml = utl.getXmlString(doc);

            return xml;
        }


        //Construccion de XML para retorno de Datos
        private string ConstruirXMLRetorno()
        {
            string xml = "";

            XmlDocument doc = new XmlDocument();

            XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "utf-8", null);
            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(xmlDeclaration, root);

            //FACTURA
            XmlElement factura = utl.createXmlNode(doc, "FACTURA");
            doc.AppendChild(factura);

            if (DocHeader.Rows.Count > 0)
            {
                DataRow r = DocHeader.Rows[0];

                factura.AppendChild(utl.createXmlNode(doc, "NODOCUMENTO", utl.convertirString(r["DocNo"])));
                factura.AppendChild(utl.createXmlNode(doc, "SERIE", utl.convertirString(r["Serie"])));
                factura.AppendChild(utl.createXmlNode(doc, "EMPRESA", utl.convertirString(r["Empresa"])));
                factura.AppendChild(utl.createXmlNode(doc, "SUCURSAL", utl.convertirString(r["Sucursal"])));
            }

            xml = utl.getXmlString(doc);

            return xml;
        }


        private string construirXMLNC()
        {
            string xml = "";

            XmlDocument doc = new XmlDocument();

            XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "utf-8", null);
            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(xmlDeclaration, root);

            //NOTA CREDITO
            XmlElement factura = utl.createXmlNode(doc, "NOTACREDITO");
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
            if (this.DocHeader.Rows.Count > 0)
            {
                DataRow r = DocHeader.Rows[0];

                //Documento Referencia de aplicacion de NC
                string[] docref = utl.convertirString(r["Parametros"]).Split('|');

                if (docref.Length != 4)
                {
                    throw new Exception("Documento de referencia de aplicación de Nota de Crédito Inválido");
                }

                string re = utl.convertirString(docref[0]);
                string nf = utl.convertirString(docref[1]);
                string se = utl.convertirString(docref[2]);

                //Etiqueta ENCABEZADO
                encabezado.AppendChild(utl.createXmlNode(doc, "NODOCUMENTO", utl.convertirString(r["DocNo"])));
                encabezado.AppendChild(utl.createXmlNode(doc, "RESOLUCION", utl.convertirString(r["Resolucion"])));
                encabezado.AppendChild(utl.createXmlNode(doc, "IDSERIE", utl.convertirString(r["Serie"])));
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
                encabezado.AppendChild(utl.createXmlNode(doc, "NITCONTRIBUYENTE", utl.convertirString(r["NITCliente"])));
                encabezado.AppendChild(utl.createXmlNode(doc, "VALORNETO", utl.formatoCurrencySS(utl.convertirDouble(r["ValorNeto"]))));
                encabezado.AppendChild(utl.createXmlNode(doc, "IVA", utl.formatoCurrencySS(utl.convertirDouble(r["IVA"]))));
                encabezado.AppendChild(utl.createXmlNode(doc, "TOTAL", utl.formatoCurrencySS(utl.convertirDouble(r["Total"]))));
                encabezado.AppendChild(utl.createXmlNode(doc, "DESCUENTO", utl.formatoCurrencySS(utl.convertirDouble(r["Descuento"]))));
                encabezado.AppendChild(utl.createXmlNode(doc, "EXENTO", utl.formatoCurrencySS(utl.convertirDouble(r["Exento"]))));
                encabezado.AppendChild(utl.createXmlNode(doc, "NOFACTURA", docref[1]));
                encabezado.AppendChild(utl.createXmlNode(doc, "SERIEFACTURA", docref[2]));
                encabezado.AppendChild(utl.createXmlNode(doc, "FECHAFACTURA", docref[3]));
                encabezado.AppendChild(utl.createXmlNode(doc, "CONCEPTO", utl.convertirString(r["Observacion"])));

                //Etiqueta OPCIONAL
                opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL1"));
                opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL2"));
                opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL3"));
                opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL4"));
                opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL5"));
                opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL6"));
                opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL7"));
                opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL8"));
                opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL9"));
                opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL10"));
                opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL11"));
                opcional.AppendChild(utl.createXmlNode(doc, "TOTAL_LETRAS", utl.formatoNumeroALetras(utl.convertirDouble(r["TotalNC"]), 2, true, utl.convertirString(r["Moneda"]), true).ToUpper(), true));
            }

            //Existencia de datos de Detalle de documento
            if (this.DocDetail.Rows.Count > 0)
            {
                //Etiqueta DETALLE
                foreach (DataRow rf in this.DocDetail.Rows)
                {
                    string[] desc = utl.convertirString(rf["Descripcion"]).Split('\n');

                    for (int i = 0; i < desc.Length; i++)
                    {
                        //LINEA
                        XmlElement linea = utl.createXmlNode(doc, "LINEA");
                        detalle.AppendChild(linea);

                        string tipodet = (i == 0) ? ((utl.convertirASCII(rf["ManejaExistencias"]) == "T") ? "BIEN" : "SERVICIO") : "";
                        string exento = (i == 0) ? ((utl.convertirASCII(rf["Exento"]) == "T") ? "S" : "N") : "";
                        string codigo = (i == 0) ? utl.convertirString(rf["Codigo"]) : "";

                        double precioUnitario = (i == 0) ? utl.convertirDouble(rf["PrecioUnitario"]) : 0;
                        double valor = (i == 0) ? utl.convertirDouble(rf["Importe"]) : 0;
                        double cantidad = (i == 0) ? utl.convertirDouble(rf["Cantidad"]) : 0;

                        //Etiqueta LINEA
                        linea.AppendChild(utl.createXmlNode(doc, "CANTIDAD", utl.convertirString(cantidad)));
                        linea.AppendChild(utl.createXmlNode(doc, "DESCRIPCION", desc[i], true));
                        linea.AppendChild(utl.createXmlNode(doc, "METRICA", "UNI"));
                        linea.AppendChild(utl.createXmlNode(doc, "PRECIOUNITARIO", utl.formatoCurrencySS(precioUnitario * conversion)));
                        linea.AppendChild(utl.createXmlNode(doc, "VALOR", utl.formatoCurrencySS(valor * conversion)));
                        linea.AppendChild(utl.createXmlNode(doc, "TIPO_PRODUCTO", tipodet, true));
                        linea.AppendChild(utl.createXmlNode(doc, "EXENTO", exento, true));
                        linea.AppendChild(utl.createXmlNode(doc, "DETALLE1", codigo));
                    }
                }
            }

            xml = utl.getXmlString(doc);

            return xml;
        }


        //Construccion de XML para anulacion
        private string ConstruirXMLAnulacion()
        {
            string xml = "";

            XmlDocument doc = new XmlDocument();

            XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "utf-8", null);
            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(xmlDeclaration, root);

            //FACTURA
            XmlElement documento = utl.createXmlNode(doc, "FACTURA");
            doc.AppendChild(documento);

            if (DocHeader.Rows.Count > 0)
            {
                DataRow r = DocHeader.Rows[0];

                documento.AppendChild(utl.createXmlNode(doc, "NODOCUMENTO", utl.convertirString(r["DocNo"])));
                documento.AppendChild(utl.createXmlNode(doc, "SERIE", utl.convertirString(r["Serie"])));
                documento.AppendChild(utl.createXmlNode(doc, "EMPRESA", utl.convertirString(r["Empresa"])));
                documento.AppendChild(utl.createXmlNode(doc, "SUCURSAL", utl.convertirString(r["Sucursal"])));
                documento.AppendChild(utl.createXmlNode(doc, "RAZONANULACION", utl.convertirString(r["Observacion"]), true));
            }

            xml = utl.getXmlString(doc);

            return xml;
        }


        //Construcion de listado de datos a retornar por las funciones
        private List<Parameter> ObtenerDatosResultado(clsResponseGeneral res, string xml, string function)
        {
            List<Parameter> parameters = new List<Parameter>
            {
                new Parameter() { ParameterName = "XML", Value = xml },
                new Parameter() { ParameterName = "Function", Value = function },
                new Parameter() { ParameterName = "Resultado", Value = res.pResultado },
                new Parameter() { ParameterName = "PDF", Value = true },
                new Parameter() { ParameterName = "Respuesta", Value = res.pRespuesta },
                new Parameter() { ParameterName = "Mensaje", Value = res.pDescripcion },
                new Parameter() { ParameterName = "IDDoc", Value = res.pIdSerieSat },
                new Parameter() { ParameterName = "Serie", Value = res.pIdSerieSat },
                new Parameter() { ParameterName = "DocNo", Value = res.pNoDocumento },
                new Parameter() { ParameterName = "CAE", Value = res.pCAE },
                new Parameter() { ParameterName = "CAEC", Value = res.pCAEC }
            };

            return parameters;
        }



        #endregion


    }
}
