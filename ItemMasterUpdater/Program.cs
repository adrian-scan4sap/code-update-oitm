using ItemMasterUpdater.Model;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ItemMasterUpdater
{
    internal class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 7)
            {
                Console.WriteLine("Usage: [1] csv file location, [2] server name, [3] dbUser, [4] dbPassword, [5] dbName, [6] sapUser, [7] sapPassword");
                return;
            }

            /* Declare the company variable - the connection */
            Company company = null;

            Console.WriteLine("Connecting to SAP...");

            /* Connect returns if connection has been established as bool */
            var isConnected = Connect(ref company, args[1], BoDataServerTypes.dst_MSSQL2016, args[2], args[3], args[4], args[5], args[6]);

            if (isConnected)
            {
                var csvPath = args[0];

                /* CSV to Model */
                var decodedList = DecodeCsv(csvPath);

                Console.WriteLine(string.Format("Processing [{0}] items...", decodedList.Count));
                var counter = 1;
                
                /* SAP Updates */
                foreach (var decodedItem in decodedList)
                {
                    var sapResult = UpdateSapItemMaster(company, decodedItem);
                    Console.WriteLine(string.Format("[{0} or {1}] {2}", counter, decodedList.Count, sapResult));
                    counter += 1;
                }
            }

            Console.WriteLine(Environment.NewLine + "Disconnecting now...");

            /* Disconnect + also release the held memory */
            Disconnect(ref company);

            Console.WriteLine(Environment.NewLine + "Disconnected. Press any key to exit.");
            Console.ReadKey();
        }

        private static string UpdateSapItemMaster(Company company, ItemData itemData)
        {
            try
            {
                var result = string.Format("OK for [{0}]", itemData.ItemCode);

                var sapItem = (Items)company.GetBusinessObject(BoObjectTypes.oItems);
                var isLoaded = sapItem.GetByKey(itemData.ItemCode);

                if (isLoaded)
                {
                    if (!string.IsNullOrEmpty(itemData.Width))
                    {
                        sapItem.SalesUnitWidth1 = double.Parse(itemData.Width);
                    }

                    if (!string.IsNullOrEmpty(itemData.Height))
                    {
                        sapItem.SalesUnitHeight1 = double.Parse(itemData.Height);
                    }

                    if (!string.IsNullOrEmpty(itemData.Length))
                    {
                        sapItem.SalesUnitLength1 = double.Parse(itemData.Length);
                    }

                    if (!string.IsNullOrEmpty(itemData.Weight))
                    {
                        sapItem.SalesUnitWeight1 = double.Parse(itemData.Weight);
                    }

                    if (sapItem.Update() != 0)
                    {
                        result = string.Format("SAP Error: {0}", company.GetLastErrorDescription());
                    }
                }
                else
                {
                    result = string.Format("SAP API could not load [{0}]", itemData.ItemCode);
                }

                return result;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        private static List<ItemData> DecodeCsv(string csvPath)
        {
            var returnedList = new List<ItemData>();            

            foreach (var line in System.IO.File.ReadAllLines(csvPath))
            {
                var splittedLine = line.Split(new char[] { ',' });

                returnedList.Add(new ItemData(splittedLine[0], splittedLine[1], splittedLine[2], splittedLine[3], splittedLine[4]));
            }

            return returnedList;
        }

        static bool Connect(ref Company company, string serverName, BoDataServerTypes serverType, string dbUserName, string dbPassword, string companyName, string sapUser, string sapPassword)
        {
            if (company == null)
            {
                company = new Company();
            }

            if (!company.Connected)
            {
                /* Server connection details */
                company.Server = serverName;
                company.DbServerType = serverType;
                company.DbUserName = dbUserName;
                company.DbPassword = dbPassword;
                company.UseTrusted = false;

                /* SAP connection details: DB, SAP User and SAP Password */
                company.CompanyDB = companyName;
                company.UserName = sapUser;
                company.Password = sapPassword;

                /* In case the SAP license server is kept in a different location (in most cases can be left empty) */
                company.LicenseServer = "";

                var isSuccessful = company.Connect() == 0;

                return isSuccessful;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Disconnects and releases the held memory (RAM)
        /// </summary>
        /// <param name="company"></param>
        static void Disconnect(ref SAPbobsCOM.Company company)
        {
            if (company != null
                && company.Connected)
            {
                company.Disconnect();

                Release(ref company);
            }
        }

        /// <summary>
        /// Re-usable method for releasing COM-held memory
        /// </summary>
        /// <typeparam name="T">Type of object to be released</typeparam>
        /// <param name="obj">The instance to be released</param>
        static void Release<T>(ref T obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                }
            }
            catch (Exception) { }
            finally
            {
                obj = default(T);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }
}
