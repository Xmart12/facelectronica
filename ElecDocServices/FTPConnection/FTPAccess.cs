using System;
using System.IO;
using System.Linq;
using WinSCP;

namespace FTPConnection
{

    /******************************************************
    * En esta clase se definen todos lo metodos para  
    * la transferencia de archivos vía ftp
    * ------------------------------------
    * Para realizar las operaciones ftp, se utiliza el ensamblado de WinSCP:
    * El ensamblado winscpnet.dll de WinSCP .NET 
    * permite la conexion a una maquina remota
    * y manipular archivos remotos a traves de sesiones SFTP, SCP y FTP 
    * desde lenguajes .NET, tales como C#, VB.NET y otros.     
        * http://winscp.net/eng/docs/library
        * Para hacer uso de este ensamblado, 
        * es necesario descargarlo de la pagina de WinSCP,
        * copiar el archivo dll en la carpeta del proyecto, 
        * agregar la referencia a este archivo, 
        * y agregarlo en las clases donde se necesite.
    ******************************************************/


    /// <summary>
    /// Clase que contiene los metodos para la transferencia de archivos vía ftp
    /// </summary>
    public class FtpAccess
    {
        //Clases Globales
        private Utils utl = new Utils();
        private Errors err = null;

        //Variables Globales
        private string ConfigPath = null;
        private string ErrorPath = null;
        private string FtpServer = null;
        private string UserServer = null;
        private string PasswordServer = null;
        private string ApplicationPath = null;
        private string FileName = null;
        private string SystemUser = null;



        /// <summary>
        /// Inicializar los valores con los datos por defecto
        /// </summary>
        //Inicializar los valores con los datos por defecto
        public FtpAccess(string config, string user)
        {
            this.ConfigPath = config;
            this.SystemUser = user;

            getConfigParameters();
        }


        ///<summary>
        ///Transferir el archivo local al servidor en la carpeta especificada
        ///</summary>
        ///<param name="remotePath">Carpeta destino del archivo</param>
        ///<param name="localFile">Ubicacion y nombre del archivo a subir</param>
        //Transferir el archivo local al servidor en la carpeta especificada
        public bool uploadWinFtp(string remotePath, string localFile)
        {
            bool resultado = false;

            try
            {
                // Setup session options
                SessionOptions sessionOptions = new SessionOptions
                {
                    Protocol = Protocol.Ftp,
                    HostName = this.FtpServer,
                    UserName = this.UserServer,
                    Password = this.PasswordServer
                };

                using (Session session = new Session())
                {                
                    session.ExecutablePath = this.ApplicationPath;

                    // Connect
                    session.Open(sessionOptions);

                    // Upload files
                    TransferOptions transferOptions = new TransferOptions();
                    transferOptions.TransferMode = TransferMode.Binary;

                    TransferOperationResult transferResult;
                    string remoteFile = remotePath + @"/" + Path.GetFileName(localFile);
                    transferResult = session.PutFiles(localFile, remoteFile, false, transferOptions);

                    // Throw on any error
                    transferResult.Check();

                    this.FileName = utl.convertirArraytoString(transferResult.Transfers.Select(s => s.FileName).ToArray());

                    resultado = true;
                }
            }
            catch (Exception e)
            {
                string c = " EXE: " + Path.GetFullPath(this.ApplicationPath);
                string os = " OS: " + Environment.OSVersion.Platform.ToString();
                err.AddErrors(e, "FTP Param: RemotePath:" + remotePath + " LocalFile:" + localFile + c + os);
                resultado = false;
            }

            return resultado;
        }


        ///<summary>
        ///Transferir el archivo local al servidor en la carpeta especificada y con el nombre proporcionado
        ///</summary>
        ///<param name="remotePath">Carpeta destino del archivo</param>
        ///<param name="localPath">Ubicación y nombre del archivo a subir</param>
        ///<param name="fileName">Nombre con el cual se guardara el archivo en el servidor</param>
        //Transferir el archivo local al servidor en la carpeta especificada y con el nombre proporcionado
        public bool uploadWinFtp(string remotePath, string localPath, string fileName)
        {
            bool resultado = false;

            try
            {
                // Setup session options
                SessionOptions sessionOptions = new SessionOptions
                {
                    Protocol = Protocol.Ftp,
                    HostName = this.FtpServer,
                    UserName = this.UserServer,
                    Password = this.PasswordServer
                };

                //SslCertificate = FtpMode.Active
                using (Session session = new Session())
                {
                    session.ExecutablePath = this.ApplicationPath;

                    // Connect
                    session.Open(sessionOptions);

                    // Upload files
                    TransferOptions transferOptions = new TransferOptions();
                    transferOptions.TransferMode = TransferMode.Binary;

                    TransferOperationResult transferResult;
                    string remoteFile = remotePath + @"/" + fileName;
                    transferResult = session.PutFiles(localPath, remoteFile, false, transferOptions);

                    // Throw on any error
                    transferResult.Check();

                    this.FileName = utl.convertirArraytoString(transferResult.Transfers.Select(s => s.FileName).ToArray());

                    resultado = true;

                }
            }
            catch (Exception e)
            {
                string c = " EXE: " + Path.GetFullPath(this.ApplicationPath);
                string os = " OS: " + Environment.OSVersion.Platform.ToString();
                err.AddErrors(e, "FTP Param: RemotePath:" + remotePath + " LocalFile:" + localPath + " FileName: " + fileName + c + os);
                resultado = false;
            }

            return resultado;
        }


        ///<summary>
        ///Transferir un archivo desde el servidor hasta una ubicacion local
        ///</summary>
        ///<param name="remotePath">Ubicacion y nombre del archivo a descargar</param>
        ///<param name="localPath">Carpeta en la que se copiara el archivo</param>
        //Transferir un archivo desde el servidor hasta una ubicacion local       
        public bool downloadWinFtp(string remotePath, string localPath)
        {
            bool resultado = false;

            try
            {
                // Setup session options
                SessionOptions sessionOptions = new SessionOptions
                {
                    Protocol = Protocol.Ftp,
                    HostName = this.FtpServer,
                    UserName = this.UserServer,
                    Password = this.PasswordServer
                };

                using (Session session = new Session())
                {             
                    // Connect
                    session.ExecutablePath = this.ApplicationPath;
                    session.Open(sessionOptions);

                    TransferOptions transferOptions = new TransferOptions();
                    transferOptions.TransferMode = TransferMode.Binary;

                    TransferOperationResult transferResult;
                    transferResult = session.GetFiles(remotePath, localPath, false, transferOptions);

                    // Throw on any error
                    transferResult.Check();

                    this.FileName = utl.convertirArraytoString(transferResult.Transfers.Select(s => s.FileName).ToArray());

                    resultado = true;

                }

            }
            catch (Exception e)
            {
                string c = " EXE: " + Path.GetFullPath(this.ApplicationPath);
                string os = " OS: " + Environment.OSVersion.Platform.ToString();
                err.AddErrors(e, "FTP Param: RemotePath:" + remotePath + " LocalFile:" + localPath + c + os);
                resultado = false;
            }

            return resultado;

        }


        ///<summary>
        ///Elimina un archivo en el servidor
        ///</summary>
        ///<param name="remotePath">Ubicacion y nombre del archivo a eliminar</param>
        //Elimina un archivo en el servidor
        public bool deleteWinFtp(string remotePath)
        {
            bool resultado = false;

            try
            {
                // Setup session options
                SessionOptions sessionOptions = new SessionOptions
                {
                    Protocol = Protocol.Ftp,
                    HostName = this.FtpServer,
                    UserName = this.UserServer,
                    Password = this.PasswordServer
                };

                using (Session session = new Session())
                {
                    session.ExecutablePath = this.ApplicationPath;
                    
                    // Connect
                    session.Open(sessionOptions);

                    if (session.FileExists(remotePath))
                    {
                        session.RemoveFiles(remotePath);
                    }

                    resultado = true;
                }

            }
            catch (Exception e)
            {
                string c = " EXE: " + Path.GetFullPath(this.ApplicationPath);
                string os = " OS: " + Environment.OSVersion.Platform.ToString();
                err.AddErrors(e, "FTP Param: RemotePath:" + remotePath + c + os);
                resultado = false;
            }

            return resultado;

        }


        ///<summary>
        ///Verifica si existe el archivo en la direccion mandada
        ///</summary>
        ///<param name="remotePath">Carpeta destino del archivo</param>
        ///<param name="fileName">Nombre con el cual se guardara el archivo en el servidor</param>
        //Verifica si existe el archivo en la direccion mandada
        public bool fileExist(string remotePath, string fileName)
        {
            bool resultado = false;

            try
            {
                string remoteFile = remotePath + @"/" + fileName;

                // Setup session options
                SessionOptions sessionOptions = new SessionOptions
                {
                    Protocol = Protocol.Ftp,
                    HostName = this.FtpServer,
                    UserName = this.UserServer,
                    Password = this.PasswordServer
                };

                //SslCertificate = FtpMode.Active
                using (Session session = new Session())
                {
                    // Connect
                    session.ExecutablePath = this.ApplicationPath;
                    session.Open(sessionOptions);

                    resultado = session.FileExists(remoteFile);
                }
            }
            catch (Exception e)
            {
                string c = " EXE: " + Path.GetFullPath(this.ApplicationPath);
                string os = " OS: " + Environment.OSVersion.Platform.ToString();
                err.AddErrors(e, "FTP Param: RemotePath:" + remotePath + " FileName:" + fileName + c + os);
                resultado = false;
            }

            return resultado;
        }


        ///<summary>
        ///Abre un archivo local
        ///</summary>
        ///<param name="localPath">Ubicacion y nombre del archivo local a abrir</param>
        //Abre un archivo local
        public bool openFile(string localPath)
        {
            bool open = false;
            try
            {
                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.StartInfo.FileName = localPath;
                proc.Start();
                proc.Close();
                open = true;
            }
            catch (Exception ex)
            {
                err.AddErrors(ex, "File Param: LocalPath:" + localPath);
                open = false;
            }

            return open;
        }


        /// <summary>
        /// Obtiene el tamaño de un archivo en el servidor
        /// </summary>
        /// <param name="remotePath">Carpeta en la que se encuentra el archivo</param>
        /// <param name="fileName">Nombre del archivo a consultar</param>
        //Obtiene el tamaño de un archivo en el servidor
        public long getFileSizeFtp(string remotePath, string fileName)
        {
            long size = 0;

            try
            {
                // Setup session options
                SessionOptions sessionOptions = new SessionOptions
                {
                    Protocol = Protocol.Ftp,
                    HostName = this.FtpServer,
                    UserName = this.UserServer,
                    Password = this.PasswordServer
                };

                using (Session session = new Session())
                {
                    session.ExecutablePath = ApplicationPath;
                    
                    // Connect
                    session.Open(sessionOptions);
                    RemoteDirectoryInfo directory = session.ListDirectory(remotePath);

                    foreach (RemoteFileInfo fileInfo in directory.Files)
                    {
                        if (fileName == fileInfo.Name)
                        {
                            size = fileInfo.Length;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                string c = " EXE: " + Path.GetFullPath(this.ApplicationPath);
                string os = " OS: " + Environment.OSVersion.Platform.ToString();
                err.AddErrors(e, "FTP Param: RemotePath:" + remotePath + " FileName:" + fileName + c + os);
                size = 0;
            }

            return size;
        }


        /// <summary>
        /// Obtiene los valores de parametros de configuracion de funcionamiento de FTP
        /// </summary>
        //Obtiene los valores de parametros de configuracion de funcionamiento de FTP
        private void getConfigParameters()
        {
            if (!File.Exists(this.ConfigPath))
            {
                throw new Exception("Config File does not exists");
            }

            this.FtpServer = utl.convertirString(utl.getConfigValue(this.ConfigPath, "FTP_SERVER", "ftp"));
            this.UserServer = utl.convertirString(utl.getConfigValue(this.ConfigPath, "FTP_USER", "ftp"));
            this.PasswordServer = utl.convertirString(utl.getConfigValue(this.ConfigPath, "FTP_PASSWORD", "ftp"));
            this.ApplicationPath = utl.convertirString(utl.getConfigValue(this.ConfigPath, "FTP_APP", "ftp"));
            this.ErrorPath = utl.convertirString(utl.getConfigValue(this.ConfigPath, "FTP_LOG", "errors"));

            this.err = new Errors(this.SystemUser, "FTP", this.ErrorPath);
        }

    }
}
