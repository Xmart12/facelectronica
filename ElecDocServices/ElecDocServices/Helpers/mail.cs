using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Windows.Forms;
using System.Net;
using ERPDesktopDADA.Helpers;

namespace ERPDesktopDADA.Helpers
{

    /// <summary>
    /// Clase que trabaja como motor de envío de correos desde el servidor de la empresa
    /// </summary>
    class mail
    {
        //Clases Globales
        private errors err = new errors();
        private utilidades utl = new utilidades();
        private conexionMySQL con = new conexionMySQL();

        //Variables Globales
        public string user = null;


        /// <summary>
        /// Inicializa una instancia de la clase mail
        /// </summary>
        /// <param name="u">Usuario</param>
        public mail(string u)
        {
            this.user = u;
        }



        /// <summary>
        /// Metodo que permite el envío de correo 
        /// </summary>
        /// <param name="from_mail">Direccion de correo Remitente</param>
        /// <param name="from_name">Nombre del remitente</param>
        /// <param name="to_mail">Direccion de correo Destinatario</param>
        /// <param name="to_name">Nombre del Destinatario</param>
        /// <param name="subject">Asunto del correo</param>
        /// <param name="body">Cuerpo del correo (HTML habilitado)</param>
        /// <returns>Booleano</returns>
        public bool enviarCorreo(string from_mail, string from_name, string to_mail, string to_name, string subject, string body, NetworkCredential credentials = null, MailAddress replyTo = null)
        {
            bool resultado = false;

            try
            {
                SmtpClient smtpMail = new SmtpClient(global.mailserver);
                MailAddress from = new MailAddress(from_mail, from_name);
                MailAddress toUser = new MailAddress(to_mail, to_name);
                MailMessage message = new MailMessage(from, toUser);

                if (replyTo != null)
                {
                    message.ReplyToList.Add(replyTo);
                }

                message.Subject = subject;
                message.Body = body;
                message.IsBodyHtml = true;

                smtpMail.EnableSsl = true;
                smtpMail.UseDefaultCredentials = false;
                smtpMail.Host = global.mailserver;
                //smtpMail.Port = 25;
                smtpMail.Credentials = credentials != null ? credentials : new System.Net.NetworkCredential(global.mailuser, global.mailpass);   

                smtpMail.Send(message);

                resultado = true;

                insertarLog(from_name, credentials.UserName, subject, replyTo.Address, to_mail, null, false);
             
            }
            catch (Exception ex)
            {
                err.AddErrors(ex, "Envío de Correo");
                resultado = false;
            }

            return resultado;
        }



        /// <summary>
        /// Método que permite el envío de correo
        /// </summary>
        /// <param name="from_mail">Direccion de correo Remitente</param>
        /// <param name="from_name">Nombre del remitente</param>
        /// <param name="to_mail">Direccion de correo Destinatario</param>
        /// <param name="to_name">Nombre del Destinatario</param>
        /// <param name="subject">Asunto del correo</param>
        /// <param name="body">Cuerpo del correo (HTML habilitado)</param>
        /// <param name="attachments">Lista de rutas para adjuntos</param>
        /// <returns></returns>
        public bool enviarCorreo(string from_mail, string from_name, string to_mail, string to_name, string subject, string body, List<string> attachments, NetworkCredential credentials = null, MailAddress replyTo = null)
        {
            bool resultado = false;

            try
            {
                SmtpClient smtpMail = new SmtpClient(global.mailserver);
                MailAddress from = new MailAddress(from_mail, from_name);
                MailAddress toUser = new MailAddress(to_mail, to_name);
                MailMessage message = new MailMessage(from, toUser);


                System.Net.Mail.Attachment attachment;
              

                foreach (var item in attachments)
                {
                    attachment = new System.Net.Mail.Attachment(item);
                    message.Attachments.Add(attachment);
                }
                
                if (replyTo != null)
                {
                    message.ReplyToList.Add(replyTo);
                }
                message.Subject = subject;
                message.Body = body;
                message.IsBodyHtml = true;

                smtpMail.EnableSsl = true;
                smtpMail.UseDefaultCredentials = false;
                smtpMail.Host = global.mailserver;
                //smtpMail.Port = 25;
                smtpMail.Credentials = credentials != null ? credentials : new System.Net.NetworkCredential(global.mailuser, global.mailpass);   
                smtpMail.Send(message);

                resultado = true;

                insertarLog(from_name, credentials.UserName, subject, replyTo.Address, to_mail, null, false);
            }
            catch (Exception ex)
            {
                err.AddErrors(ex, "Envío de Correo");
                resultado = false;
            }

            return resultado;
        }



        /// <summary>
        /// Metodo que permite el envío de correo 
        /// </summary>
        /// <param name="from">Objeto de tipo MailAddress con el Remitente</param>
        /// <param name="to">Objeto de tipo MailAddress con el Destinatario</param>
        /// <param name="cc">Lista de tipo MailAddress con los Destinatarios de tipo Copia</param>
        /// <param name="subject">Asunto del correo</param>
        /// <param name="body">Cuerpo del correo (HTML habilitado)</param>
        /// <returns></returns>
        public bool enviarCorreo(MailAddress from, List<MailAddress> to, List<MailAddress> cc, string subject, string body, NetworkCredential credentials = null, MailAddress replyTo = null)
        {
            bool resultado = false;

            try
            {
                SmtpClient smtpMail = new SmtpClient(global.mailserver);

                MailMessage message = new MailMessage();
                message.From = from;

                if (replyTo != null)
                {
                    message.ReplyToList.Add(replyTo);
                }

                foreach (MailAddress m in to)
                {
                    message.To.Add(m);
                }

                foreach (MailAddress m in cc)
                {
                    message.CC.Add(m);
                }

                message.Subject = subject;
                message.Body = body;
                message.IsBodyHtml = true;

                smtpMail.EnableSsl = true;
                smtpMail.UseDefaultCredentials = false;
                smtpMail.Host = global.mailserver;
                //smtpMail.Port = 25;
                smtpMail.Credentials = credentials != null ? credentials : new System.Net.NetworkCredential(global.mailuser, global.mailpass);   

                smtpMail.Send(message);

                resultado = true;

                string enviadoPor = from.DisplayName;
                string credencial = credentials != null ? credentials.UserName : global.mailuser;
                string replyToAddress = replyTo != null ? replyTo.Address : null;
                string toConcat = utl.convertirArraytoString(to.ToArray(), ";");
                string ccConcat = utl.convertirArraytoString(cc.ToArray(), ";");
                bool poseeAdjuntos = false;

                insertarLog(enviadoPor, credencial, subject, replyToAddress, toConcat, ccConcat, poseeAdjuntos);

            }
            catch (Exception ex)
            {
                err.AddErrors(ex, "Envio de Correo");
                resultado = false;
            }

            return resultado;
        }



        /// <summary>
        /// Metodo que permite el envío de correo 
        /// </summary>
        /// <param name="from">Objeto de tipo MailAddress con el Remitente</param>
        /// <param name="to">Objeto de tipo MailAddress con el Destinatario</param>
        /// <param name="cc">Lista de tipo MailAddress con los Destinatarios de tipo Copia</param>
        /// <param name="subject">Asunto del correo</param>
        /// <param name="body">Cuerpo del correo</param>
        /// <param name="attachments">Lista de rutas para adjuntos</param>
        /// <returns></returns>
        public bool enviarCorreo(MailAddress from, List<MailAddress> to, List<MailAddress> cc, string subject, string body, List<string> attachments, NetworkCredential credentials = null, MailAddress replyTo = null)
        {
            bool resultado = false;

            try
            {
                SmtpClient smtpMail = new SmtpClient(global.mailserver);

                MailMessage message = new MailMessage();
                message.From = from;

                if (replyTo != null)
                {
                    message.ReplyToList.Add(replyTo);
                }

                foreach (MailAddress m in to)
                {
                    message.To.Add(m);
                }

                foreach (MailAddress m in cc)
                {
                    message.CC.Add(m);
                }

                System.Net.Mail.Attachment attachment;

                foreach (var item in attachments)
                {
                    attachment = new System.Net.Mail.Attachment(item);
                    message.Attachments.Add(attachment);
                }

                message.Subject = subject;
                message.Body = body;
                message.IsBodyHtml = true;

                smtpMail.EnableSsl = true;
                smtpMail.UseDefaultCredentials = false;
                smtpMail.Host = global.mailserver;
                //smtpMail.Port = 25;          
                smtpMail.Credentials = credentials != null ? credentials : new System.Net.NetworkCredential(global.mailuser, global.mailpass);   
      
                //Envío del correo de manera asincrona
                smtpMail.SendCompleted += (s, e) =>
                {
                    smtpMail.Dispose();
                    message.Dispose();
                    if (e.Error != null)
                    {                     
                        err.AddErrors("Error en envío de correos", utl.convertirString(e.Error), utl.convertirString(this), this.user);                     
                        utl.ventanaNotificacion("Envío de correo", "Error al enviar el correo: " + subject, 3);
                    }
                    else
                    {
                        string enviadoPor = from.DisplayName;
                        string credencial = credentials != null ? credentials.UserName : global.mailuser;
                        string replyToAddress = replyTo != null ? replyTo.Address : null;
                        string toConcat = utl.convertirArraytoString(to.ToArray(), ";");
                        string ccConcat = utl.convertirArraytoString(cc.ToArray(), ";");
                        bool poseeAdjuntos = attachments != null && attachments.Count > 0 ? true : false;

                        insertarLog(enviadoPor, credencial, subject, replyToAddress, toConcat, ccConcat, poseeAdjuntos);

                        utl.ventanaNotificacion("Envío de correo", "Correo enviado correctamente: " + subject, 1);
                    }                   
                };
                smtpMail.SendAsync(message, null);
                //smtpMail.Send(message);        
  
                resultado = true;
            }
            catch (Exception ex)
            {
                err.AddErrors(ex, "Envio de Correo");
                resultado = false;
            }

            return resultado;
        }



        //Metodo para insertar al log los correos enviados
        private void insertarLog(string enviadoPor, string credencial, string asunto, string replyTo, string to, string cc, bool adjuntos)
        {
            try
            {
                string[] campos = { "EnviadoPor", "Credencial", "Usuario", "Asunto", "ReplyTo", "EnviadoA", "CC", "PoseeAdjuntos" };
                object [] datos = { enviadoPor, credencial, this.user, asunto, replyTo, to, cc, adjuntos ? 1 : 0};
                con.execInsert("tbl_correos_enviados", campos, datos);
            }
            catch (Exception ex)
            {
                err.AddErrors(ex, "Error al insertar log de envío de correos");
            }
        }


    }
}
