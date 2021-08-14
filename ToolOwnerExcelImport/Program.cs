using ExcelDataReader;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;

namespace ToolOwnerExcelImport
{
    class Program
    {
        static void Main(string[] args)
        {
            ReadExcel();
        }
        static void ReadExcel()
        {
            string connection = @"Integrated Security=false;Initial Catalog=TiceCentralServices;Data Source=167.86.118.102,50480;User id=ticeuser;password=!QAZ2wsx";
            FileStream stream = File.Open(@"F:\AZ Tool Owner List for Stakeholders.xlsx", FileMode.Open, FileAccess.Read);
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var excelReader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration { FallbackEncoding = Encoding.GetEncoding(1252) });
            var headers = new List<string>();
            int count = 0;
            List<ToolOwner> toolOwner = new List<ToolOwner>();
            while (excelReader.Read())
            {
                count += 1;
                if (count == 2)
                {
                    for (var i = 0; i < excelReader.FieldCount; i++)
                        headers.Add(Convert.ToString(excelReader.GetString(i)));
                }
                else if (count > 2)
                {
                    List<string> FunctionalArea = new List<string>();
                    for (var j = 6; j < excelReader.FieldCount; j++)
                    {
                        var jdbgjs = excelReader.GetString(j);
                        if (excelReader.GetString(j) != null)
                        {
                            FunctionalArea.Add(headers[j]);
                        }
                    }
                    toolOwner.Add(new ToolOwner
                    {
                        Name = excelReader.GetString(0),
                        CompanyName = excelReader.GetString(1),
                        Role = excelReader.GetString(4),
                        // WWID=excelReader.GetString(2),
                        Email = excelReader.GetString(5),
                        Functionalarea = FunctionalArea
                    }); ;
                }

            }
            //string json = JsonConvert.SerializeObject(toolOwner);
            SqlConnection con = new SqlConnection(connection);
            if (con.State == ConnectionState.Closed)
                con.Open();
            int projectId = 4;
            String projectName = "D1D74";
            String FunctionalAreaName = "AM";
            for (var i = 0; i < toolOwner.Count; i++)
            {
                Console.WriteLine("Inserting Records.......");
                string RoleName = toolOwner[i].Role;
                int companyId = GetCompany(con, toolOwner[i].CompanyName);
                var userId = GetApplicationUser(con, toolOwner[i].Email);

                InsertFunctionalAreaTeams(con, toolOwner[i].Functionalarea, 0);
                int FunctionalAreaId = 0;
                if (toolOwner[i].Functionalarea.Contains("AM"))
                {
                    Console.WriteLine("functional Area Detected ");
                    FunctionalAreaId = GetFunctionalArea(con, FunctionalAreaName);
                }
                if (userId > 0)
                {
                    UpdateUser(con, companyId, 0, userId);
                    var roleId = GetRole(con, toolOwner[i].Role);
                    if (roleId > 0)
                    {
                        var roleAdGroupId = GetRoleProjectMapping(con, roleId, projectId);
                        if (roleAdGroupId > 0)
                        {
                            var userRoleAdgroupId = GetUserRoleMapping(con, roleId, userId, projectId);
                            if (userRoleAdgroupId == 0)
                            {
                                SaveUserRoleMapping(con, roleId, roleAdGroupId, userId, projectId);
                            }
                        }
                        else
                        {
                            SaveRoleProjectMapping(con, roleId, roleAdGroupId, projectName, RoleName, projectId);
                            var AdGroupId = GetRoleProjectMapping(con, projectId, roleId);
                            if (AdGroupId > 0)
                            {
                                SaveUserRoleMapping(con, roleId, AdGroupId, userId, projectId);
                            }
                        }
                    }

                    if (FunctionalAreaId > 0 && toolOwner[i].Functionalarea.Contains("AM"))
                    {
                        int FaId = GetUserFunctionAreaMapping(con, FunctionalAreaId, userId, projectId);
                        if (FaId == 0)
                        {
                            SaveUserFAMapping(con, FunctionalAreaId, userId, projectId);
                        }
                    }
                }
                else
                {
                    SaveUser(con, toolOwner[i].Name, toolOwner[i].Email, projectId, companyId);
                    var newuserId = GetApplicationUser(con, toolOwner[i].Email);
                    if (newuserId > 0)
                    {
                        var roleId = GetRole(con, toolOwner[i].Role);
                        if (roleId > 0)
                        {
                            var roleAdGroupId = GetRoleProjectMapping(con, roleId, projectId);
                            if (roleAdGroupId > 0)
                            {
                                var userRoleAdgroupId = GetUserRoleMapping(con, roleId, userId, projectId);
                                if (userRoleAdgroupId == 0)
                                {
                                    SaveUserRoleMapping(con, roleId, roleAdGroupId, userId, projectId);
                                }
                            }
                            else
                            {
                                SaveRoleProjectMapping(con, roleId, roleAdGroupId, projectName, RoleName, projectId);
                                var AdGroupId = GetRoleProjectMapping(con, projectId, roleId);
                                if (AdGroupId > 0)
                                {
                                    SaveUserRoleMapping(con, roleId, AdGroupId, userId, projectId);
                                }
                            }
                        }
                        if (FunctionalAreaId > 0 && toolOwner[i].Functionalarea.Contains("AM"))
                        {
                            int FaId = GetUserFunctionAreaMapping(con, FunctionalAreaId, userId, projectId);
                            if (FaId == 0)
                            {
                                SaveUserFAMapping(con, FunctionalAreaId, userId, projectId);
                            }
                        }

                    }

                }
                //sql_cmnd.ExecuteNonQuery();
            }
            Console.WriteLine("Records Saved Successfully.......");
            con.Close();
        }
        static int GetRoleProjectMapping(SqlConnection con, int ProjectId, int RoleId)
        {
            string query = " select top 1 adGroupId from sec.ProjectRoleMapping where isActive=1 and projectId=" + ProjectId + " and roleId=" + RoleId;
            SqlCommand Comm = new SqlCommand(query, con);
            object result = Comm.ExecuteScalar();
            if (result != null)
            {
                Console.WriteLine(result);
                return Convert.ToInt32(result);
            }
            else
            {
                return 0;
            }
        }
        static int GetCompany(SqlConnection con, string CompanyName)
        {
            string query = " select top 1 companyId from sec.company where isActive=1 and companyName= '" + CompanyName + "'";
            SqlCommand Comm = new SqlCommand(query, con);
            object result = Comm.ExecuteScalar();
            if (result != null)
            {
                Console.WriteLine(result);
                return Convert.ToInt32(result);
            }
            else
            {
                return 0;
            }
        }
        static int GetApplicationUser(SqlConnection con, string UserName)
        {
            string query = "SELECT Top 1 applicationUserId from sec.ApplicationUser where userEmailAddress ='" + UserName + "'";
            SqlCommand Comm = new SqlCommand(query, con);
            object result = Comm.ExecuteScalar();
            if (result != null)
            {
                Console.WriteLine(result);
                return Convert.ToInt32(result);
            }
            else
            {
                return 0;
            }
        }
        static int GetRole(SqlConnection con, string RoleName)
        {
            string query = "SELECT Top 1 roleId from sec.role where roleName ='" + RoleName + "'";
            SqlCommand Comm = new SqlCommand(query, con);
            object result = Comm.ExecuteScalar();
            if (result != null)
            {
                Console.WriteLine(result);
                return Convert.ToInt32(result);
            }
            else
            {
                return 0;
            }
        }
        static int GetUserRoleMapping(SqlConnection con, int RoleId, int UserId, int Projectid)
        {
            string query = " SELECT top 1 prm.AdGroupId as adGroupId FROM sec.Role r inner JOIN sec.ProjectRoleMapping prm on prm.roleId = r.RoleID " +
                " and prm.projectId = " + Projectid + " and prm.isActive = 1 and r.isActive = 1 inner join sec.ADGroupApplicationUser adau on adau.ADGroupID = prm.AdGroupId " +
               " and adau.applicationUserID = " + UserId + " and prm.projectId = " + Projectid + " and r.roleid = " + RoleId;
            SqlCommand Comm = new SqlCommand(query, con);
            object result = Comm.ExecuteScalar();
            if (result != null)
            {
                Console.WriteLine(result);
                return Convert.ToInt32(result);
            }
            else
            {
                return 0;
            }
        }
        static int GetFunctionalArea(SqlConnection con, string FName)
        {
            string query = "SELECT Top 1 FunctionalAreaId from tool.functionalArea where functionalAreaName ='" + FName + "'";
            SqlCommand Comm = new SqlCommand(query, con);
            object result = Comm.ExecuteScalar();
            if (result != null)
            {
                Console.WriteLine(result);
                return Convert.ToInt32(result);
            }
            else
            {
                return 0;
            }
        }
        static int GetUserFunctionAreaMapping(SqlConnection con, int FunctionalAreaId, int UserId, int Projectid)
        {
            string query = " select top 1 functionalAreaId from sec.ProjectUserFunctionalArea where projectId=" + Projectid + " and applicationUserID=" + UserId + " and functionalAreaId= " + FunctionalAreaId;
            SqlCommand Comm = new SqlCommand(query, con);
            object result = Comm.ExecuteScalar();
            if (result != null)
            {
                Console.WriteLine(result);
                return Convert.ToInt32(result);
            }
            else
            {
                return 0;
            }
        }
        static int SaveUserRoleMapping(SqlConnection con, int RoleId, int roleADGroupID, int UserId, int Projectid)
        {
            using (SqlCommand cmd = new SqlCommand("sec.SaveUserRoleMapping", con))
            {
                cmd.CommandType = CommandType.StoredProcedure;

                //p.Add("@ApplicationUserRoleMappingId", model.ApplicationUserRoleMappingId);
                //p.Add("@ApplicationUserID", model.ApplicationUserID);
                //p.Add("@RoleId", model.RoleId);
                //p.Add("@AdgroupId", model.AdGroupId);
                cmd.Parameters.Add("@ApplicationUserRoleMappingId", SqlDbType.Int).Value = 0;
                cmd.Parameters.Add("@ApplicationUserID", SqlDbType.Int).Value = UserId;
                cmd.Parameters.Add("@RoleId", SqlDbType.Int).Value = RoleId;
                cmd.Parameters.Add("@AdgroupId", SqlDbType.Int).Value = roleADGroupID;
                SqlParameter outputIdParam = new SqlParameter("@ErrorLogID", SqlDbType.Int)
                {
                    Direction = ParameterDirection.Output
                };
                cmd.Parameters.Add(outputIdParam);
                return cmd.ExecuteNonQuery();
            }

        }
        static int SaveUserFAMapping(SqlConnection con, int FunctionalAreaId, int UserId, int Projectid)
        {
            using (SqlCommand cmd = new SqlCommand("[sec].[uspInsertUserAndFunctionalAreaMapping]", con))
            {
                cmd.CommandType = CommandType.StoredProcedure;

                //p.Add("@ApplicationUserRoleMappingId", model.ApplicationUserRoleMappingId);
                //p.Add("@ApplicationUserID", model.ApplicationUserID);
                //p.Add("@RoleId", model.RoleId);
                //p.Add("@AdgroupId", model.AdGroupId);

                cmd.Parameters.Add("@ProjectId", SqlDbType.Int).Value = 0;
                cmd.Parameters.Add("@UserId", SqlDbType.Int).Value = UserId;
                cmd.Parameters.Add("@UpdatedById", SqlDbType.Int).Value = UserId;
                cmd.Parameters.Add("@FunctionalAreaId", SqlDbType.Int).Value = FunctionalAreaId;
                SqlParameter outputIdParam = new SqlParameter("@ErrorLogID", SqlDbType.Int)
                {
                    Direction = ParameterDirection.Output
                };
                cmd.Parameters.Add(outputIdParam);
                return cmd.ExecuteNonQuery();
            }

        }
        static int SaveRoleProjectMapping(SqlConnection con, int RoleId, int roleADGroupID, string ProjectName, string RoleName, int Projectid)
        {
            using (SqlCommand cmd = new SqlCommand("[sec].[SaveProjectRoleMapping]", con))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@projectId", SqlDbType.Int).Value = Projectid;
                cmd.Parameters.Add("@roleId", SqlDbType.Int).Value = RoleId;
                cmd.Parameters.Add("@isChecked", SqlDbType.Bit).Value = true;
                cmd.Parameters.Add("@projectName", SqlDbType.VarChar).Value = ProjectName;
                cmd.Parameters.Add("@roleName", SqlDbType.VarChar).Value = RoleName;
                SqlParameter outputIdParam = new SqlParameter("@ErrorLogID", SqlDbType.Int)
                {
                    Direction = ParameterDirection.Output
                };
                cmd.Parameters.Add(outputIdParam);
                return cmd.ExecuteNonQuery();
            }

        }
        //save application user
        static int SaveUser(SqlConnection con, string UserFirstName, string Email, int Projectid, int CompanyId)
        {
            string[] userList = UserFirstName.Split(" ");
            using (SqlCommand cmd = new SqlCommand("[sec].[AddApplicationUser]", con))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@applicationUserID", SqlDbType.Int).Value = 0;
                cmd.Parameters.Add("@oldAdGroupId", SqlDbType.Int).Value = 0;
                cmd.Parameters.Add("@userDomain", SqlDbType.VarChar).Value = "";
                cmd.Parameters.Add("@userEmail", SqlDbType.VarChar).Value = Email;
                cmd.Parameters.Add("@userFirstName", SqlDbType.VarChar).Value = userList[0];
                cmd.Parameters.Add("@userLastName", SqlDbType.VarChar).Value = userList[1];
                cmd.Parameters.Add("@userJobDescription", SqlDbType.VarChar).Value = "";
                cmd.Parameters.Add("@userLogonName", SqlDbType.VarChar).Value = Email;
                cmd.Parameters.Add("@optionaluserEmail", SqlDbType.VarChar).Value = "";
                cmd.Parameters.Add("@companyId", SqlDbType.Int).Value = CompanyId;
                cmd.Parameters.Add("@projectId", SqlDbType.Int).Value = Projectid;
                cmd.Parameters.Add("@projectCompanyMappingId", SqlDbType.Int).Value = 0;
                cmd.Parameters.Add("@samAccountName", SqlDbType.VarChar).Value = UserFirstName;
                cmd.Parameters.Add("@IsJiraUser", SqlDbType.Bit).Value = false;
                cmd.Parameters.Add("@IsJiraUserId", SqlDbType.VarChar).Value = "";
                SqlParameter outputIdParam = new SqlParameter("@ErrorLogID", SqlDbType.Int)
                {
                    Direction = ParameterDirection.Output
                };
                cmd.Parameters.Add(outputIdParam);
                return cmd.ExecuteNonQuery();
            }

        }
        static int UpdateUser(SqlConnection con, int companyId, int wwid, int applicationUserId)
        {

            string qry = "UPDATE sec.applicationUser SET companyId = @companyId, WWID = @WWID" +
" Where applicationUserId = @applicationUserId";
            using (SqlCommand cmd = new SqlCommand(qry, con))
            {
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.Add("@companyId", SqlDbType.Int).Value = companyId;
                cmd.Parameters.Add("@WWID", SqlDbType.Int).Value = wwid;
                cmd.Parameters.Add("@applicationUserId", SqlDbType.Int).Value = applicationUserId;

                return cmd.ExecuteNonQuery();
            }

        }
        static int InsertFunctionalAreaTeams(SqlConnection con, List<string> teams, int functionaAreaId)
        {

            int teamsId = GetFunctionalAreaTeams(con, teams);
            if (teamsId == 0)
            {
                string qry = "Insert INTO tool.FunctionalAreaTeam (functionalAreaTeamName)values(@TeamsName) ";
                using (SqlCommand cmd = new SqlCommand(qry, con))
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.Parameters.Add("@TeamsName", SqlDbType.VarChar).Value = string.Join(",", teams);
                    return cmd.ExecuteNonQuery();

                }
            }
            else
            {
                return 0;
            }
        }
        static int GetFunctionalAreaTeams(SqlConnection con, List<string> teams)
        {



            string qry = "select top 1 functionalareaTeamId from tool.FunctionalAreaTeam  where functionalAreaTeamName='" + string.Join(",", teams)+"'";
            using (SqlCommand cmd = new SqlCommand(qry, con))
            {
                SqlCommand Comm = new SqlCommand(qry, con);
                object result = Comm.ExecuteScalar();
                if (result != null)
                {
                    Console.WriteLine(result);
                    return Convert.ToInt32(result);
                }
                else
                {
                    return 0;
                }
            }
        }

    }
}
class ToolOwner
{
    public string Name { get; set; }
    public string Role { get; set; }
    public string Email { get; set; }
    public string CompanyName { get; set; }
    public string WWID { get; set; }
    public List<string> Functionalarea { get; set; }
}
