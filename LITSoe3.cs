using System;
using System.Collections.Generic;
using System.Linq;
//using System.Data;
using System.Text;
using System.Data.OleDb;
using System.Data;

using System.Collections.Specialized;
using System.Runtime.InteropServices;

using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Server;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.SOESupport;


//TODO: sign the project (project properties > signing tab > sign the assembly)
//      this is strongly suggested if the dll will be registered using regasm.exe <your>.dll /codebase


namespace LITSoe3
{
    [ComVisible(true)]
    [Guid("2d8c8f41-bdcb-45f6-839b-61e21a3797f0")]
    [ClassInterface(ClassInterfaceType.None)]
    [ServerObjectExtension("MapServer",//use "MapServer" if SOE extends a Map service and "ImageServer" if it extends an Image service.
        AllCapabilities = "",
        DefaultCapabilities = "",
        Description = "",
        DisplayName = "LITSoe3",
        Properties = "",
        SupportsREST = true,
        SupportsSOAP = false)]
    public class LITSoe3 : IServerObjectExtension, IObjectConstruct, IRESTRequestHandler
    {
        private string soe_name;

        private IPropertySet configProps;
        private IServerObjectHelper serverObjectHelper;
        private ServerLogger logger;
        private IRESTRequestHandler reqHandler;

        public LITSoe3()
        {
            soe_name = this.GetType().Name;
            logger = new ServerLogger();
            reqHandler = new SoeRestImpl(soe_name, CreateRestSchema()) as IRESTRequestHandler;
        }

        #region IServerObjectExtension Members

        public void Init(IServerObjectHelper pSOH)
        {
            serverObjectHelper = pSOH;
        }

        public void Shutdown()
        {
        }

        #endregion

        #region IObjectConstruct Members

        public void Construct(IPropertySet props)
        {
            configProps = props;
        }

        #endregion

        #region IRESTRequestHandler Members

        public string GetSchema()
        {
            return reqHandler.GetSchema();
        }

        public byte[] HandleRESTRequest(string Capabilities, string resourceName, string operationName, string operationInput, string outputFormat, string requestProperties, out string responseProperties)
        {
            return reqHandler.HandleRESTRequest(Capabilities, resourceName, operationName, operationInput, outputFormat, requestProperties, out responseProperties);
        }

        #endregion

        private RestResource CreateRestSchema()
        {
            RestResource rootRes = new RestResource(soe_name, false, RootResHandler);

            RestOperation sampleOper = new RestOperation("ExecuteStoredProcedure",
                                                      new string[] { "ParamValuePairs", "Extra" },
                                                      new string[] { "json" },
                                                      ExecuteStoredProcedureHandler);

            rootRes.operations.Add(sampleOper);

            return rootRes;
        }
        private string  UseAOToCreateUpdateFeatures(string sql )
        {
            string retString = "in get UseAOToCreateUpdateFeatures >";
            try
            {
                //get map server  
                IMapServer3 mapServer = (IMapServer3)serverObjectHelper.ServerObject;

                IMapServerDataAccess dataAccess = (IMapServerDataAccess)mapServer;

                IFeatureClass fc = dataAccess.GetDataSource(mapServer.DefaultMapName, 0) as IFeatureClass;
                retString += ">Attempting to get wse";
                IWorkspace ws = (fc as IDataset).Workspace ;
                //string sql = @"exec dbo.apLITSaveChanges @curGeoLocID = 5, @NewGeoLocId = 2, @address = '222W.Pine', @windowsLogin = 'x', @city = 'wern', @zipCode = '12345', @Latitude = 1, @longitude = 2, @facilityname = 'asdf', @appID = 'asdf', @serverName = 'asdf', @status = 'r'";
                ws.ExecuteSQL(sql);
               return "ok";
            }
            catch (Exception ex)
            {
                return "ERROR " + ex.ToString();
            }
        }
        private byte[] RootResHandler(NameValueCollection boundVariables, string outputFormat, string requestProperties, out string responseProperties)
        {
            responseProperties = null;

            JsonObject result = new JsonObject();
            result.AddString("DEQ", "LIT Soe Support");

            return Encoding.UTF8.GetBytes(result.ToJson());
        }

        private byte[] ExecuteStoredProcedureHandler(NameValueCollection boundVariables,
                                                  JsonObject operationInput,
                                                      string outputFormat,
                                                      string requestProperties,
                                                  out string responseProperties)
        {
            responseProperties = null;
            string retString = "";

            try
            {

                
                //return Encoding.UTF8.GetBytes(retStrn);
                //return null;

            string pipeDelimetedStringValuePairsForStoredProc="";
            bool found = operationInput.TryGetString("ParamValuePairs", out pipeDelimetedStringValuePairsForStoredProc);
            if (!found || string.IsNullOrEmpty(pipeDelimetedStringValuePairsForStoredProc))
                throw new ArgumentNullException("ParamValuePairs");

            string extra;
            found = operationInput.TryGetString("Extra", out extra);
            if (!found || string.IsNullOrEmpty(extra))
                throw new ArgumentNullException("extra");

            responseProperties = null;
            IServerEnvironment3 senv = GetServerEnvironment() as IServerEnvironment3;
            JsonObject result = new JsonObject();
            JsonObject suinfoj = new JsonObject();
            //get user info and serialize into JSON 
            IServerUserInfo suinfo = senv.UserInfo;
            if (null != suinfo)
            {
                suinfoj.AddString("currentUser", suinfo.Name);

                IEnumBSTR roles = suinfo.Roles;
                List<string> rolelist = new List<string>();
                if (null != roles)
                {
                    string role = roles.Next();
                    while (!string.IsNullOrEmpty(role))
                    {
                        rolelist.Add(role);
                        role = roles.Next();
                    }
                }

                suinfoj.AddArray("roles", rolelist.ToArray());
                result.AddJsonObject("serverUserInfo", suinfoj);
            }
            else { result.AddJsonObject("serverUserInfo", null); }

            IServerObject so = serverObjectHelper.ServerObject;
            retString = "got so>";

            string progString = "";
            retString += "> Stored Proc via oleDB ";// + ex.Message;
            OleDbConnection con = new OleDbConnection();
            string paramsThatParsed = "";



            con.ConnectionString = @"Provider =SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog="  + extra.Split(',')[0] + ";Data Source=" + extra.Split(',')[1];// @"Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=TestDB;Data Source=PC684"; //
            string storedProcedureName = "dbo.apLITSaveChanges";
                bool isStatusOutput = Convert.ToBoolean(extra.Split(',')[2]);
                //the connection string below uses integrated security which is usually superior to storing credential in visible text
                //con.ConnectionString = @"Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=TestDB;Data Source=PC684";
                //con.Open();
                //retString += ">opened connection";
                string SQLString = "";
            OleDbCommand cmd = new OleDbCommand(storedProcedureName, con);
            cmd.CommandType = System.Data.CommandType.StoredProcedure;
                string userName = "UNKNOWN";
            if(suinfo.Name != null)
            {
                    userName = suinfo.Name;
            }
            cmd.Parameters.AddWithValue("@WindowsLogin",userName);
                SQLString += "@WindowsLogin='" + userName + "'";
                //SQLString += "@WindowsLogin=" + userName ;
                retString += ">created command";
            string[] paramValsForStoredProc = pipeDelimetedStringValuePairsForStoredProc.Split('|');
            foreach (string paramVal in paramValsForStoredProc)
            {
               string param = paramVal.Split(',')[0]; 
                paramsThatParsed += "," + param;
                string val = paramVal.Split(',')[1];
                retString += ">param and value : " + param + " = " + val;

                param = "@" + param;
                if (param.ToUpper().Contains("GEOLOCID"))
                {
                        int i = int.Parse(val);
                    cmd.Parameters.AddWithValue(param, i);
                        SQLString += ", "+ param + "= " + i  ;
                    }
                else if(param.ToUpper() == "@LATITUDE" || param.ToUpper() == "@LONGITUDE")
                {

                        double d = double.Parse(val);
                        cmd.Parameters.AddWithValue(param, d);
                        SQLString += ", " + param + "=  " + d  ;

                    }
                else if (param.ToUpper() == "@STATUS")
                    {
                        if (isStatusOutput)
                        {
                            //cmd.Parameters[param].Direction = ParameterDirection.Output;
                            retString += ">Set direction of status parameter to output";
                            SQLString += ", @STATUS = @localstatus OUTPUT";
                        }
                    }
                else
                {
                    cmd.Parameters.AddWithValue(param, val);
                        //SQLString += ", " + param + "=  " + val ;
                        SQLString += ", " + param + "=  '" + val + "'";
                    }




            }//CurGeoLocID,NewGeoLocID,Address,City,ZipCode,Latitude,Longitude,FacilityName,AppID,WindowsLogin,ServerName,ServerName,Status
                SQLString = "exec dbo.apLITSaveChanges " + SQLString;
                if (isStatusOutput)
                {
                    SQLString = "DECLARE @localstatus varchar(256);" + SQLString;
                }
                string retStrn = UseAOToCreateUpdateFeatures(SQLString);
                return Encoding.UTF8.GetBytes(retStrn);
                return null;
            cmd.Connection = con;
            cmd.ExecuteNonQuery();

            return Encoding.UTF8.GetBytes(result.ToJson() + "  -  " + retString.ToString());
            }
            catch (Exception ex)
            {
                return Encoding.UTF8.GetBytes("ERROR   " + ex.ToString() + "  :  " +  retString.ToString());
            }
        }

        public IServerEnvironment3 GetServerEnvironment()
        {
            UID uid = new UIDClass();
            uid.Value = "{32D4C328-E473-4615-922C-63C108F55E60}";

            //use activator to cocreate singleton
            Type t = Type.GetTypeFromProgID("esriSystem.EnvironmentManager");
            System.Object obj = Activator.CreateInstance(t);
            IEnvironmentManager environmentManager = obj as IEnvironmentManager;
            return (environmentManager.GetEnvironment(uid) as IServerEnvironment3);
        }
    }
}
