using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Oracle.ManagedDataAccess.Client;

namespace DefectUpdater {
    public class OraSession {
        public string Schema { get; set; }
        public string Password { get; set; }
        public string Host { get; set; }
        public string Port { get; set; }
        public string ServiceName { get; set; }
        public OracleConnection OracleConnection { get; set; }

        public OraSession(string host, string port, string schema, string password, string serviceName) {
            Host = host;
            Port = port;
            Schema = schema;
            Password = password;
            ServiceName = serviceName;
        }

        private string CreateConnectionString() {
            return "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(" +
                   "HOST=" + Host + ")(" +
                   "PORT=" + Port + ")))(" +
                   "CONNECT_DATA=(SERVICE_NAME=" + ServiceName + "))); " +
                   "USER ID=" + Schema + "; " +
                   "PASSWORD=" + Password + ";";
        }

        public void OpenConnection() {
            OracleConnection = new OracleConnection { ConnectionString = CreateConnectionString() };
            OracleConnection.Open();
        }

        public void CloseConnection() {
            OracleConnection.Close();
            OracleConnection.Dispose();
        }

        public void UpdateDefectsTable(KnownDefect knownDefect, string userId) {
            string query = "UPDATE VT_DEFECTS " +
                "SET Defect = :defect, " +
                "User_Id = :user_Id, " +
                "Changed_date = :changed_date " +
                "WHERE " +
                "Project = :project and " +
                "Lower_Version = :lower_Version and " +
                "Upper_Version = :upper_Version and " +
                "Master_TransNo = :master_TransNo and " +
                "Test_TransNo = :test_TransNo and " +
                "SecId = :secID and " +
                "Deviation_Column_Name = :deviation_Column_Name";
            OracleCommand cmd = new OracleCommand(query, OracleConnection);
            cmd.Parameters.Add(":defect", OracleDbType.Varchar2).Value = knownDefect.DefectNo;
            cmd.Parameters.Add(":user_Id", OracleDbType.Varchar2).Value = userId;
            cmd.Parameters.Add(":changed_date", OracleDbType.TimeStamp).Value = DateTime.Now;
            cmd.Parameters.Add(":project", OracleDbType.Varchar2).Value = knownDefect.Project;
            cmd.Parameters.Add(":lower_Version", OracleDbType.Double).Value = knownDefect.LowerVersion;
            cmd.Parameters.Add(":upper_Version", OracleDbType.Double).Value = knownDefect.UpperVersion;
            cmd.Parameters.Add(":master_TransNo", OracleDbType.Varchar2).Value = knownDefect.MasterTransNo;
            cmd.Parameters.Add(":test_TransNo", OracleDbType.Varchar2).Value = knownDefect.TestTransNo;
            cmd.Parameters.Add(":secId", OracleDbType.Varchar2).Value = knownDefect.SecId;
            cmd.Parameters.Add(":deviation_Column_Name", OracleDbType.Varchar2).Value = knownDefect.DeviationColumnName;                       
            cmd.ExecuteNonQuery();
        }

        public void InsertIntoDefectsTable(KnownDefect knownDefect, string userId) {           
            string query = "INSERT INTO VT_DEFECTS(Project, Lower_Version, Upper_Version, Defect, Master_TransNo, Test_TransNo, SecId, Deviation_Column_Name, Master_Value, Test_Value, User_Id, Changed_date) " +
                "VALUES(:project, :lower_Version, :upper_Version, :defect, :master_TransNo, :test_TransNo, :secId, :deviation_Column_Name, :master_Value, :test_Value, :user_Id, :changed_date)";
            OracleCommand cmd = new OracleCommand(query, OracleConnection);
            cmd.Parameters.Add(":project", OracleDbType.Varchar2).Value = knownDefect.Project;
            cmd.Parameters.Add(":lower_Version", OracleDbType.Double).Value = knownDefect.LowerVersion;
            cmd.Parameters.Add(":upper_Version", OracleDbType.Double).Value = knownDefect.UpperVersion;
            cmd.Parameters.Add(":defect", OracleDbType.Varchar2).Value = knownDefect.DefectNo;
            cmd.Parameters.Add(":master_TransNo", OracleDbType.Varchar2).Value = knownDefect.MasterTransNo;
            cmd.Parameters.Add(":test_TransNo", OracleDbType.Varchar2).Value = knownDefect.TestTransNo;
            cmd.Parameters.Add(":secId", OracleDbType.Varchar2).Value = knownDefect.SecId;
            cmd.Parameters.Add(":deviation_Column_Name", OracleDbType.Varchar2).Value = knownDefect.DeviationColumnName;
            cmd.Parameters.Add(":master_Value", OracleDbType.Varchar2).Value = knownDefect.MasterValue;
            cmd.Parameters.Add(":test_Value", OracleDbType.Varchar2).Value = knownDefect.TestValue;
            cmd.Parameters.Add(":user_Id", OracleDbType.Varchar2).Value = userId;
            cmd.Parameters.Add(":changed_date", OracleDbType.TimeStamp).Value = DateTime.Now;
            cmd.ExecuteNonQuery();
        }

        public void DeleteDefectsTable(KnownDefect knownDefect) {
            string query = "DELETE FROM VT_DEFECTS " +
                "WHERE " +
                "Project = :project and " +
                "Lower_Version = :lower_Version and " +
                "Upper_Version = :upper_Version and " +
                "Master_TransNo = :master_TransNo and " +
                "Test_TransNo = :test_TransNo and " +
                "SecId = :secId and " +
                "Deviation_Column_Name = :deviation_Column_Name";
            OracleCommand cmd = new OracleCommand(query, OracleConnection);
            cmd.Parameters.Add(":project", OracleDbType.Varchar2).Value = knownDefect.Project;
            cmd.Parameters.Add(":lower_Version", OracleDbType.Double).Value = knownDefect.LowerVersion;
            cmd.Parameters.Add(":upper_Version", OracleDbType.Double).Value = knownDefect.UpperVersion;
            cmd.Parameters.Add(":master_TransNo", OracleDbType.Varchar2).Value = knownDefect.MasterTransNo;
            cmd.Parameters.Add(":test_TransNo", OracleDbType.Varchar2).Value = knownDefect.TestTransNo;
            cmd.Parameters.Add(":secId", OracleDbType.Varchar2).Value = knownDefect.SecId;
            cmd.Parameters.Add(":deviation_Column_Name", OracleDbType.Varchar2).Value = knownDefect.DeviationColumnName;
            cmd.ExecuteNonQuery();
        }

        public string GetDefectNoFromDB(KnownDefect knownDefect) {
            string query = "";
            if (knownDefect.MasterTransNo == "" && knownDefect.TestTransNo == "" && knownDefect.SecId == "") {
                query = "select defect from VT_DEFECTS where PROJECT = :proj and Lower_Version = :lower_Version and Upper_Version = :upper_Version and Deviation_Column_Name = :deviation_Column_Name and Master_Value = :master_Value and Test_Value = :test_Value";
            } else {
                query = "select defect from VT_DEFECTS where PROJECT = :proj and Lower_Version = :lower_Version and Upper_Version = :upper_Version and Master_TransNo = :master_TransNo and Test_TransNo = :test_TransNo and SecId = :secId and Deviation_Column_Name = :deviation_Column_Name";
            }
                OracleCommand cmd = new OracleCommand(query, OracleConnection);
            cmd.Parameters.Add(":proj", OracleDbType.Varchar2).Value = knownDefect.Project;
            cmd.Parameters.Add(":lower_Version", OracleDbType.Double).Value = knownDefect.LowerVersion;
            cmd.Parameters.Add(":upper_Version", OracleDbType.Double).Value = knownDefect.UpperVersion;
            cmd.Parameters.Add(":master_TransNo", OracleDbType.Varchar2).Value = knownDefect.MasterTransNo;
            cmd.Parameters.Add(":test_TransNo", OracleDbType.Varchar2).Value = knownDefect.TestTransNo;
            cmd.Parameters.Add(":secId", OracleDbType.Varchar2).Value = knownDefect.SecId;
            cmd.Parameters.Add(":deviation_Column_Name", OracleDbType.Varchar2).Value = knownDefect.DeviationColumnName;
            cmd.Parameters.Add(":master_Value", OracleDbType.Varchar2).Value = knownDefect.MasterValue;
            cmd.Parameters.Add(":test_Value", OracleDbType.Varchar2).Value = knownDefect.TestValue;
            cmd.CommandType = CommandType.Text;
            using (OracleDataReader dataAdapter = cmd.ExecuteReader()) {
                while (dataAdapter.Read()) {
                    return dataAdapter.GetString(0);
                }
            }
            return "";
        }
    }
}
