using System;
using System.Text;
using System.IO;

namespace ElecDocServices.Helpers
{
    /******************************************************
    * En esta clase se definen metodos para  
    * crear o actualizar un archivo con el  
    * log de errores generados en el sistema.
    ******************************************************/

    /// <summary>
    /// Clase para el control de bitacora de errores en el programa
    /// </summary>
    class Errors
    {
        //Variables globales
        private string ErrorPath = null;
        private string NameLog = null;
        private string User = null;


        /// <summary>
        /// Inicializa una instancia de la clase Errors
        /// </summary>
        public Errors (string user, string name, string path)
        {
            this.ErrorPath = path;
            this.NameLog = name;
            this.User = user;
        }


        ///<summary>
        ///Agrega errores de ejecucion del programa al archivo de log que se esta apuntando en la instancia
        ///</summary>
        ///<param name="ex">Objecto de Excepcion de ejecucion</param>
        ///<param name="extramsg">Mensaje adicional a escribir en el log</param>
        //Agregar Error Sistema
        public void AddErrors(Exception ex, string extramsg = "")
        {
            try
            {
                //Comprobamos si existe el fichero Log
                string PathLog = this.PathLog();
                string n = Environment.NewLine;

                string errorDesc = ex.Message;
                string innererror = getExceptionMessage(ex.InnerException);
                string user = this.User;                

                //Preparamos el texto de la línea
                string txtError = n + n +
                                "FECHA: " + DateTime.Now + n +
                                "USER: " + user + n +
                                "ERROR: " + errorDesc + "-" + n +
                                "INNERERROR: " + innererror + "-" + n +
                                "COMENTARIO: " + extramsg + "-" + n +
                                "TRACKING: " + ex.StackTrace + n;

                File.AppendAllText(PathLog, txtError, Encoding.GetEncoding("437"));

            }
            catch
            {
                
            }
        }


        ///<summary>
        ///Agrega errores de ejecucion del programa, causado por error en una petición al servidor de base de datos, 
        ///al archivo de log que se esta apuntando en la instancia
        ///</summary>
        ///<param name="ex">Objecto de Excepcion de ejecucion</param>
        ///<param name="query">Codigo SQL que se ejecutó</param>
        ///<param name="con">Conexion de Base de Datos</param>
        ///<param name="extramsg">Mensaje adicional a escribir en el log</param>
        //Agregar Error SQL
        public void AddErrors(Exception ex, string query, string con, string extramsg = "")
        {
            try
            {
                //Comprobamos si existe el fichero Log
                string PathLog = this.PathLog();
                string n = Environment.NewLine;

                string errorDesc = ex.Message;
                string innererror = getExceptionMessage(ex.InnerException);
                string user = this.User;

                //Preparamos el texto de la línea
                string txtError = n + n +
                                "FECHA: " + DateTime.Now + n +
                                "USER: " + user + n +
                                "ERROR: " + errorDesc + n +
                                "INNERERROR: " + innererror + n +
                                "SQL QUERY: " + query + n +
                                "CONEXION: " + con + n +
                                "COMENTARIO: " + extramsg + n +
                                "TRACKING: " + ex.StackTrace + n;

                File.AppendAllText(PathLog, txtError, Encoding.GetEncoding("437"));

            }
            catch
            {

            }
        }


        ///<summary>
        ///Agrega errores de ejecucion del programa al archivo de log que se esta apuntando en la instancia
        ///</summary>
        ///<param name="errorDesc">Descripcion del error</param>
        ///<param name="innererror">Error anidado</param>
        ///<param name="localizacion">Localizacion del error en el programa</param>
        ///<param name="user">Usuario de sistema</param>
        public void AddErrors(string errorDesc, string innererror, string localizacion, string user)
        {
            try
            {
                //Comprobamos si existe el fichero Log
                string PathLog = this.PathLog();
                string n = Environment.NewLine;

                //Preparamos el texto de la línea
                string txtError = n + n + 
                                "FECHA: " + DateTime.Now + n +
                                "USER: " + user + n +
                                "ERROR: " + errorDesc + "-" + n +
                                "INNERERROR: " + innererror + "-" + n +
                                "LOCALIZACION: " + localizacion + n;

                File.AppendAllText(PathLog, txtError, Encoding.GetEncoding("437"));

            }
            catch
            { 
                
            }
        }


        ///<summary>
        ///Agrega errores de ejecucion del programa, causado por error en una petición al servidor de base de datos, 
        ///al archivo de log que se esta apuntando en la instancia
        ///</summary>
        ///<param name="errorDesc">Descripcion del error</param>
        ///<param name="innererror">Error anidado</param>
        ///<param name="localizacion">Localizacion del error en el programa</param>
        ///<param name="query">Codigo SQL que se ejecutó</param>
        ///<param name="user">Usuario de sistema</param>
        public void AddErrors(string errorDesc, string innererror, string localizacion, string query, string user)
        {
            try
            {
                //Comprobamos si existe el fichero Log
                string PathLog = this.PathLog();
                string n = Environment.NewLine;

                //Preparamos el texto de la línea
                string txtError = n + n +
                                "FECHA: " + DateTime.Now + n +
                                "USER: " + user + n +
                                "ERROR: " + errorDesc + "-" + n +
                                "INNERERROR: " + innererror + "-" + n +
                                "MySQL QUERY: " + query + n +
                                "LOCALIZACION: " + localizacion + n;

                File.AppendAllText(PathLog, txtError, Encoding.GetEncoding("437"));

            }
            catch
            {

            }
        }


        /// <summary>
        /// Se crea el fichero si este no existe en la carpeta correspondiente
        /// </summary>
        /// <returns></returns>
        //Se crea el fichero si este no existe en la carpeta correspondiente
        private string PathLog() 
        {
            string pathLog = this.ErrorPath;
            string filename = (this.NameLog != null) ? "log" + this.NameLog + ".txt" : "logFTP.txt";

            var dir = pathLog;
            var file = Path.Combine(dir, filename);
            
            try
            {
                //Comprobamos si existe el fichero Log
                if (!File.Exists(file))
                {
                    //Si no existe el directorio se crea un directorio LOCAL
                    if (!Directory.Exists(dir))
                    {
                        Directory.CreateDirectory(dir);
                        //Si no existe el fichero log lo creamos
                        StreamWriter NewLog = File.CreateText(file);
                        NewLog.Close();
                    } 
                    else
                    {
                        //Si no existe el fichero log lo creamos
                        StreamWriter NewLog = File.CreateText(file);
                        NewLog.Close();
                    }
                }

                return file;
            }
            catch
            {
                return file;
            }
        }


        ///<summary>
        ///Funcion que obtiene las excepciones anidadas a la excepcion principañ
        ///</summary>
        ///<param name="ex">Excepcion lanzada por un error en el programa</param>
        //Obtiene las excepciones Anidadas a la excepcion
        private string getExceptionMessage(Exception ex)
        {
            string msg = null;
            if (ex != null)
            {
                msg += ex.Message;
                msg += (ex.InnerException != null) ? " -- " + getExceptionMessage(ex.InnerException) : null;
            }

            return msg;
        }


    }
}
