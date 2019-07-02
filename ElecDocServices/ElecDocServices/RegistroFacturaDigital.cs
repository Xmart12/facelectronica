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


        //Variables Globales
        private string usuario = null;

        private string TipoDoc = null;
        private string Serie = null;
        private string Resolucion = null;
        private int NoDocumento = 0;

        private DataTable dtEncabezado = null;
        private DataTable dtDetalle = null;
        private int Estatus = 0;


        //Constructor
        private RegistroFacturaDigital(string config, string user)
        {

        }

        //Formulario Mostradado (Llenado de Datos)
        private void frmEnvioFacturaDigital_Shown(object sender, EventArgs e)
        {
            if (dtEncabezado.Rows.Count > 0)
            {
                DataRow r = dtEncabezado.Rows[0];

                lblTipoDoc.Text = utl.convertirString(r["TipoDocumento"]);
                lblResolucion.Text = utl.convertirString(r["Resolucion"]);
                lblSerie.Text = utl.convertirString(r["Serie"]);
                lblNoDoc.Text = utl.convertirString(r["DocNo"]);
                lblPais.Text = utl.convertirString(r["Pais"]);
                lblFecha.Text = utl.formatoFecha(r["FechaFactura"]);
                lblUsuario.Text = utl.convertirString(r["Usuario"]);
                lblFormaPago.Text = utl.convertirString(r["FormaPago"]).ToUpper();
                lblCodCliente.Text = utl.convertirString(r["CodCliente"]);
                lblNombreCliente.Text = utl.convertirString(r["NombreCliente"]);
                lblTotal.Text = utl.formatoCurrencySS(r["TotalFactura"]);

                bool estatus = utl.convertirBoolean(r["Impreso"]);

                this.Estatus = utl.convertirInt(r["Estatus"]);

                chkRevisado.Checked = utl.convertirBoolean(r["Revisado"]);
                lblUsuarioRevisado.Text = utl.convertirString(r["RevisadoPor"]);

                verificacionRevision();

                if (estatus)
                {
                    lblEstado.Text = "La factura ya ha sido procesada.";
                    btnAceptar.Enabled = false;
                    btnImprimir.Enabled = true;
                }

                if (this.TipoDoc == "Ex" || this.TipoDoc == "Ee")
                {
                    if (utl.convertirASCII(r["Exento"]) == "F")
                    {
                        lblEstado.Text = "Cliente no es exento de impuestos.";
                        btnAceptar.Enabled = false;
                        btnImprimir.Enabled = false;
                        btnPreview.Enabled = false;
                    }
                }
            }
            else
            {
                btnAceptar.Enabled = false;
                btnImprimir.Enabled = false;
                btnPreview.Enabled = false;
                chkRevisado.Enabled = false;
                lblEstado.Text = "No se encontró la factura a registrar";
                utl.MessageBoxGenerico(2, "Resgistro Factura Digital", new string[] { "No se encontró la factura a registrar" }.ToList());
            }
        }


        //Revision de Documento
        private void chkRevisado_Click(object sender, EventArgs e)
        {
            if (chkRevisado.Enabled)
            {
                chkRevisado.Checked = false;

                string msg = "¿Desea marcar como revisado el documento " + this.TipoDoc + " Resolución: " + this.Resolucion;
                msg += " Serie: " + this.Serie + " No: " + this.NoDocumento + " para su registro?";

                DialogResult que = MessageBox.Show(msg, "Registro Factura Digital", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (que == DialogResult.Yes)
                {
                    string whr = "ef.CodPais = '" + this.CodPais + "' and ef.Dealer = " + this.Dealer;
                    whr += " and ef.Tipo_Documento = '" + this.TipoDoc + "' and ef.Resolucion = '" + this.Resolucion + "' ";
                    whr += " and ef.Serie = '" + this.Serie + "' and ef.Factura_No = " + this.NoDocumento + " ";

                    string[] campos = new string[] { "ef.Revisado", "ef.RevisadoPor", "ef.FechaRevision" };
                    object[] datos = new object[] { true, this.usuario, utl.formatoFechaSql(DateTime.Now, true) };

                    bool res = con.execUpdate("ENCABEZADO_FACTURA ef", campos, datos, whr);

                    if (res)
                    {
                        chkRevisado.Checked = true;
                        chkRevisado.Enabled = false;
                        lblUsuarioRevisado.Text = this.usuario;

                        string[] spdatos = new string[] 
                        { 
                            this.CodPais, utl.convertirString(this.Dealer), this.TipoDoc, this.Resolucion, this.Serie,
                            utl.convertirString(this.NoDocumento)
                        };

                        con.execBoolProcedure("sp_ConversionDoc", spdatos);

                        verificacionRevision();
                    }
                    else
                    {
                        utl.MessageBoxGenerico(9, "Registro Factura Digital", new string[] { "No se guardó la revisión" }.ToList());
                    }
                }
                
            }
        }


        //Boton Registrar
        private void btnAceptar_Click(object sender, EventArgs e)
        {
            bool resultado = false;
            List<string> msg = new List<string>();

            try
            {
                SSO_wsEFactura service = new SSO_wsEFactura();

                lblEstado.Text = "Generacion de XML";
                this.Refresh();

                string xml = construirXML();

                lblEstado.Text = "Procesando solicitud de registro";
                this.Refresh();

                clsResponseGeneral res = service.RegistraFacturaXML_PDF(xml);
                resultado = res.pResultado;
                msg.Add("Respuesta WS: " + res.pDescripcion);

                lblEstado.Text = "Insercion de bitacora";
                this.Refresh();
                if (!bitacoraFactura(service.Url, "RegistraFacturaXML_PDF", xml, res))
                {
                    err.AddErrors("No se Guardo registro en la bitacora de la Factura", "", "", this.usuario);
                }

                if (resultado)
                {
                    btnAceptar.Enabled = false;
                    btnImprimir.Enabled = true;

                    lblEstado.Text = "Guardado de PDF";
                    this.Refresh();

                    if (!convertPDF(res.pRespuesta))
                    {
                        err.AddErrors("No se pudo realizar conversion de PDF", "", "", this.usuario);
                    }

                    lblEstado.Text = "Actualizacion de estado de Factura";
                    this.Refresh();

                    //Actualizacion de Datos de la factura
                    string whr = "ef.CodPais = '" + this.CodPais + "' and ef.Dealer = " + this.Dealer;
                    whr += " and ef.Tipo_Documento = '" + this.TipoDoc + "' and ef.Resolucion = '" + this.Resolucion + "' ";
                    whr += " and ef.Serie = '" + this.Serie + "' and ef.Factura_No = " + this.NoDocumento + " ";

                    if (!con.execUpdate("ENCABEZADO_FACTURA ef", new string[] { "ef.Estatus", "ef.Impreso", "ef.FImpreso" }, new object[] { 1, 1, utl.formatoFechaSql(DateTime.Now, true) }, whr))
                    {
                        err.AddErrors("No se actualizo el estatus de la factura.", "", "", this.usuario);
                    }


                    // Actualizaciòn de datos de provision de contrato
                    string whrc = "CodPais = '" + this.CodPais + "' and Dealer = " + this.Dealer;
                    whrc += " and  TipoDoc='" + this.TipoDoc + "' and Serie='" + this.Serie + "' and DocNo= " + this.NoDocumento;

                    if (!con.execUpdate("ProvisionContratos", new string[] { "Impresa", "FechaImp", "HoraImp" }, new object[] { "T", utl.formatoFechaSql(DateTime.Now), utl.formatoHora(DateTime.Now) }, whrc))
                    {
                        err.AddErrors("No se actualizo datos en provision de contrato.", "", "", this.usuario);
                    }


                    utl.MessageBoxGenerico(6, "Registro Factura Digital", msg);

                    btnCancelar.Text = "Salir";

                    lblEstado.Text = "Factura Registrada";
                    this.Refresh();
                }
                else
                {
                    utl.MessageBoxGenerico(9, "Registro Factura Digital", msg);
                }
            }
            catch (Exception ex)
            {
                err.AddErrors(ex, "Registro de Factura Digital");
                utl.MessageBoxGenerico(9, "Registro Factura Digital");

                lblEstado.Text = "";
                this.Refresh();
            }
        }


        //Boton Preview
        private void btnPreview_Click(object sender, EventArgs e)
        {
            try
            {
                ReportData rd = new ReportData();

                rd.DataParameters.Add(new Parameter() { ParameterName = "pEnc", Value = dtEncabezado });
                rd.DataParameters.Add(new Parameter() { ParameterName = "pDet", Value = dtDetalle });

                rd.ReportParameters.Add(new Parameter() { ParameterName = "v_user", Value = this.usuario });
                rd.ReportParameters.Add(new Parameter() { ParameterName = "v_docno", Value = this.NoDocumento });
                rd.ReportParameters.Add(new Parameter() { ParameterName = "v_doctipo", Value = this.TipoDoc });
                rd.ReportParameters.Add(new Parameter() { ParameterName = "v_docserie", Value = this.Serie });
                rd.ReportParameters.Add(new Parameter() { ParameterName = "v_resolucion", Value = this.Resolucion });

                rd.DisplayName = "Factura";
                rd.NamespaceName = "FacturacionDigital";
                rd.IsLocal = true;
                rd.PrintDirect = false;
                rd.btnPrint = true;


                if (this.Dealer == 16)
                {
                    switch (this.TipoDoc)
                    {
                        case "Fe":
                            rd.ReportName = "rpt_factura_fe_gt";
                            break;

                        case "Fd":
                            rd.ReportName = "rpt_facturaC_fd_gt";
                            break;

                        case "Ee":
                            rd.ReportName = "rpt_exportacion_ex_gt";
                            break;

                        case "NC":
                            rd.ReportName = "rpt_notacredito_nc_gt";
                            break;
                    }
                }

                Utilitarios.VistaReporte frm = new Utilitarios.VistaReporte(rd);

                frm.ShowDialog();

            }
            catch (Exception ex)
            {
                err.AddErrors(ex, "Impresion de Factura");
                utl.MessageBoxGenerico(9, "Impresion de Factura");
            }
        }


        //Boton Imprimir
        private void btnImprimir_Click(object sender, EventArgs e)
        {
            try
            {
                bool resultado = false;

                string est = (this.Estatus == 4) ? "ANU" : "";

                string filename = this.TipoDoc + this.Resolucion + this.Serie + this.NoDocumento + est + ".pdf";
                string remotepath = "documentos/FacturasDigitales/";
                remotepath += (this.Dealer == 16) ? "GT" : "SV";

                if (ftp.fileExist(remotepath, filename))
                {
                    lblEstado.Text = "Descargando Archivo";
                    this.Refresh();

                    resultado = ftp.downloadWinFtp(remotepath + "/" + filename, global.filepath + filename);

                    if (resultado)
                    {
                        ftp.openFile(global.filepath + filename);

                        lblEstado.Text = "Archivo Descargado";
                        this.Refresh();
                    }
                    else
                    {
                        lblEstado.Text = "";
                        this.Refresh();
                        utl.MessageBoxGenerico(9, "Registro Factura Digital");
                    }
                }
                else
                {
                    lblEstado.Text = "Solicitanto archivo para almacenamiento";
                    this.Refresh();

                    SSO_wsEFactura service = new SSO_wsEFactura();

                    string xml = construirXMLRetorno();

                    clsResponseGeneral res = service.RetornaDatosFacturaXML_PDF(xml);
                    resultado = res.pResultado;
                    
                    if (resultado)
                    {
                        lblEstado.Text = "Generando y almacenando archivo";
                        this.Refresh();
                        resultado = convertPDF(res.pRespuesta);
                    }

                    if (resultado)
                    {
                        btnImprimir_Click(sender, e);
                    }
                    else
                    {
                        string[] msg = new string[] { "No se pudo obtener el documento digital." };
                        utl.MessageBoxGenerico(9, "Registro de Factura Digital", msg.ToList());
                    }
                }
            }
            catch (Exception ex)
            {
                err.AddErrors(ex, "Abrir Factura Digital");
                utl.MessageBoxGenerico(9, "Registro de Factura Digital");

                lblEstado.Text = "";
                this.Refresh();
            }
        }




        #region funciones


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

                    double total = utl.convertirDouble(r["TotalFactura"]);
                    double tasaiva = utl.convertirDouble(r["TasaIVA"]);
                    double vtagravada = utl.convertirDouble(r["VentasGravadas"]);
                    double totalsiniva = Math.Round((vtagravada / (1 + (tasaiva / 100))), 2);
                    double valoriva = (totalsiniva > 0) ? (vtagravada - totalsiniva) : 0;
                    double descuento = utl.convertirDouble(r["ValorDescuento"]);
                    double exento = utl.convertirDouble(r["VentasExentas"]);

                    string nit = (this.TipoDoc == "Ex" || this.TipoDoc == "Ee") ? "CF" : utl.convertirString(r["NIT"]);

                    //Etiqueta ENCABEZADO
                    encabezado.AppendChild(utl.createXmlNode(doc, "NOFACTURA", utl.convertirString(this.NoDocumento)));
                    encabezado.AppendChild(utl.createXmlNode(doc, "RESOLUCION", this.Resolucion));
                    encabezado.AppendChild(utl.createXmlNode(doc, "IDSERIE", this.Serie));
                    encabezado.AppendChild(utl.createXmlNode(doc, "EMPRESA", utl.convertirString(r["EmpresaDigital"])));
                    encabezado.AppendChild(utl.createXmlNode(doc, "SUCURSAL", utl.convertirString(r["Sucursal"])));
                    encabezado.AppendChild(utl.createXmlNode(doc, "CAJA", utl.convertirString(r["Cajero"])));
                    encabezado.AppendChild(utl.createXmlNode(doc, "USUARIO", utl.convertirString(r["Usuario"])));
                    encabezado.AppendChild(utl.createXmlNode(doc, "MONEDA", utl.convertirString(r["SimboloMoneda"])));
                    encabezado.AppendChild(utl.createXmlNode(doc, "TASACAMBIO", utl.convertirString(r["TasaCambio"])));
                    encabezado.AppendChild(utl.createXmlNode(doc, "GENERACION", "O"));
                    encabezado.AppendChild(utl.createXmlNode(doc, "FECHAEMISION", utl.formatoFecha(r["FechaFactura"])));
                    encabezado.AppendChild(utl.createXmlNode(doc, "NOMBRECONTRIBUYENTE", utl.convertirString(r["NombreCliente"]), true));
                    encabezado.AppendChild(utl.createXmlNode(doc, "DIRECCIONCONTRIBUYENTE", utl.convertirString(r["Direccion"]), true));
                    encabezado.AppendChild(utl.createXmlNode(doc, "NITCONTRIBUYENTE", nit));
                    encabezado.AppendChild(utl.createXmlNode(doc, "VALORNETO", utl.formatoCurrencySS(totalsiniva * conversion)));
                    encabezado.AppendChild(utl.createXmlNode(doc, "IVA", utl.formatoCurrencySS(valoriva * conversion)));
                    encabezado.AppendChild(utl.createXmlNode(doc, "TOTAL", utl.formatoCurrencySS(total * conversion)));
                    encabezado.AppendChild(utl.createXmlNode(doc, "DESCUENTO", utl.formatoCurrencySS(descuento * conversion)));
                    encabezado.AppendChild(utl.createXmlNode(doc, "EXENTO", utl.formatoCurrencySS(exento * conversion)));

                    //Etiqueta OPCIONAL
                    opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL1", utl.convertirString(r["CodCliente"])));
                    opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL2", utl.convertirString(r["Departamento"])));
                    opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL3", utl.convertirString(r["Pais"])));
                    opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL4", utl.convertirString(r["FormaPago"]).ToUpper()));
                    
                    if (this.TipoDoc == "Ex" || this.TipoDoc == "Ee")
                    {
                        opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL5", utl.convertirString(r["NIT"]), true));
                        opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL6", utl.convertirString(r["Peso"]), true));
                        opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL7", utl.convertirString(r["Volumen"]), true));
                        opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL8", utl.convertirString(r["Bultos"]), true));
                        opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL9", "SUJETO A RETENCION DEFINITIVA", true));
                        opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL10", utl.convertirString(r["Opcional1"]), true));
                        opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL11", utl.convertirString(r["Opcional2"]), true));
                    }
                    else
                    {
                        opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL5"));
                        opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL6"));
                        opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL7", utl.convertirString(r["Vendedor"])));
                        opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL8"));
                        opcional.AppendChild(utl.createXmlNode(doc, "OPCIONAL9", "SUJETO A RETENCION DEFINITIVA", true));
                    }

                    opcional.AppendChild(utl.createXmlNode(doc, "TOTAL_LETRAS", utl.formatoNumeroALetras((total * conversion), 2, true, utl.convertirString(r["Moneda"]), true).ToUpper(), true));
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

                            string tipodet = (i == 0) ? ((utl.convertirASCII(rf["ManejaExistencias"]) == "T") ? "BIEN" : "SERVICIO") : "";
                            string exento = (i == 0) ? ((utl.convertirASCII(rf["Exento"]) == "T") ? "S" : "N") : "N";
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

                //Existencia de Datos de ENCABEZADO
                if (this.dtEncabezado.Rows.Count > 0)
                {
                    DataRow r = dtEncabezado.Rows[0];

                    factura.AppendChild(utl.createXmlNode(doc, "NODOCUMENTO", utl.convertirString(this.NoDocumento)));
                    factura.AppendChild(utl.createXmlNode(doc, "SERIE", this.Serie));
                    factura.AppendChild(utl.createXmlNode(doc, "EMPRESA", utl.convertirString(r["EmpresaDigital"])));
                    factura.AppendChild(utl.createXmlNode(doc, "SUCURSAL", utl.convertirString(r["Sucursal"])));
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


        //Bitacora de Factura
        private bool bitacoraFactura(string url, string function, string xml, clsResponseGeneral res)
        {
            bool resultado = false;

            string[] campos = new string[] 
            {
                "Fecha", "Usuario", "HostName", "IpAddresses", "MacAddresses", "ServiceUrl", "FunctionName", "TipoDocDADA", 
                "ResolucionDADA", "SerieDADA", "FacNoDADA", "XmlData", "OtherParameters", "Resultado", "Descripcion", 
                "Respuesta", "idSAT", "CAE", "CAEC", "Serie", "DocNo", "OtherResponse"
            };

            string hostname = utl.getHostName();
            string ips = utl.convertirArraytoString(utl.getIpAddress().ToArray());
            string macs = utl.convertirArraytoString(utl.getMACAddress().ToArray());

            string filename = "documentos/FacturasDigitales/";
            filename += (this.Dealer == 16) ? "GT" : "SV";
            filename += "/" + this.TipoDoc + this.Resolucion + this.Serie + this.NoDocumento + ".pdf";

            var response = new 
            { 
                pResultado = res.pResultado, pDescripcion = res.pDescripcion, pIdSat = res.pIdSat, pCAE = res.pCAE,
                pCAEC = res.pCAEC, pIdSerieSat = res.pIdSerieSat, pNoDocumento = res.pNoDocumento, pSerie = res.pSerie,
            };

            object[] datos = new object[]
            {
                utl.formatoFechaSql(DateTime.Now, true), this.usuario, hostname, ips, macs, url, function, this.TipoDoc, 
                this.Resolucion, this.Serie, this.NoDocumento, xml, null, res.pResultado, res.pDescripcion, 
                filename, res.pIdSat, res.pCAE, res.pCAEC, res.pIdSerieSat, res.pNoDocumento, utl.getJson(response)
            };

            resultado = con.execInsert("tbl_log_facturaDigital", campos, datos);

            return resultado;
        }


        //Conversion de Respuesta de WS a PDF
        private bool convertPDF(object pdf)
        {
            bool resultado = false;

            try
            {
                string est = (this.Estatus == 4) ? "ANU" : "";

                string filename = this.TipoDoc + this.Resolucion + this.Serie + this.NoDocumento + est + ".pdf";
                string remotepath = "documentos/FacturasDigitales/";
                remotepath += (this.Dealer == 16) ? "GT" : "SV";
                string localpath = global.filepath;

                if (pdf != null)
                {
                    File.WriteAllBytes(localpath + filename, (byte[])pdf);

                    resultado = ftp.uploadWinFtp(remotepath, localpath + filename);

                    if (resultado)
                    {
                        File.Delete(localpath + filename);
                    }
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



        #endregion


    }
}
