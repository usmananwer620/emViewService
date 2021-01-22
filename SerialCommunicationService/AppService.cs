using Newtonsoft.Json;
using SerialCommunicationService.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace SerialCommunicationService
{
    public class AppService
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        /// <summary>
        /// Saving test module into database
        /// </summary>
        /// <param name="commandResponse"></param>
        /// <param name="serialCommandResponse"></param>
        public static async void RunCreateValuesAndCellsPostAsync(TestModuleObject testModuleObj, string serviceName)
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Accept.Add(
                        new MediaTypeWithQualityHeaderValue("application/json"));
                    var json = JsonConvert.SerializeObject(testModuleObj);
                    var data = new StringContent(json, Encoding.UTF8, "application/json");
                    DataLoggerControl.Write($"Test Module object JSON: {json}", serviceName);
                    DataLoggerControl.Write($"Test Module object JSON: {data}", serviceName);

                    client.DefaultRequestHeaders.Add("Abp.TenantId", "2");
                    HttpResponseMessage response = await client.PostAsJsonAsync("http://66.70.142.79:85/api/services/app/Module/CreateValuesAndCells",
                                                                                testModuleObj);
                    DataLoggerControl.Write($"RunCreateValuesAndCellsPostAsync: {response.ReasonPhrase}", serviceName);
                    if (response.IsSuccessStatusCode)
                    {
                        var readTask = response.Content.ReadAsAsync<TestModuleObject>();
                        readTask.Wait();
                        DataLoggerControl.Write($"Test Module object has been posted with Serial number: {testModuleObj.serialNumber}, Command Response: {testModuleObj.inputValues}", serviceName);
                    }
                    else
                    {
                        DataLoggerControl.Write($"WARNING! Test Module object could not be posted!", serviceName);
                    }
                }
            }
            catch (Exception ex)
            {
                DataLoggerControl.Write($"{ex}", serviceName);
            }
        }

        public static async void RunUpdateModuleStatusPostAsync(ErrorModuleObject errorModuleObject, string serviceName)
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Accept.Add(
                        new MediaTypeWithQualityHeaderValue("application/json"));
                    var json = JsonConvert.SerializeObject(errorModuleObject);
                    var data = new StringContent(json, Encoding.UTF8, "application/json");
                    DataLoggerControl.Write($"Error Module object JSON: {json}", serviceName);
                    DataLoggerControl.Write($"Error Module object JSON: {data}", serviceName);

                    client.DefaultRequestHeaders.Add("Abp.TenantId", "2");
                    HttpResponseMessage response = await client.PutAsJsonAsync("http://66.70.142.79:85/api/services/app/Module/UpdateModuleStatus",
                                                                                errorModuleObject);
                    DataLoggerControl.Write($"RunUpdateModuleStatusPostAsync: {response.ReasonPhrase}", serviceName);
                    if (response.IsSuccessStatusCode)
                    {
                        var readTask = response.Content.ReadAsAsync<ErrorModuleObject>();
                        readTask.Wait();
                        DataLoggerControl.Write($"Error Module response has been posted with status: {errorModuleObject.status} and serial response: {errorModuleObject.serialNumber}", serviceName);
                    }
                    else
                    {
                        DataLoggerControl.Write($"WARNING! Error Module response object could not be posted!", serviceName);
                    }
                }
            }
            catch (Exception ex)
            {
                DataLoggerControl.Write($"{ex}", serviceName);
            }
        }

        public static async void RunModulePostAsync(ModuleObjects moduleObject, string serviceName)
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Accept.Add(
                        new MediaTypeWithQualityHeaderValue("application/json"));
                    var json = JsonConvert.SerializeObject(moduleObject);
                    var data = new StringContent(json, Encoding.UTF8, "application/json");
                    DataLoggerControl.Write($"Module object JSON: {json}", serviceName);
                    DataLoggerControl.Write($"Module object JSON: {data}", serviceName);

                    client.DefaultRequestHeaders.Add("Abp.TenantId", "2");
                    HttpResponseMessage response = await client.PostAsJsonAsync("http://66.70.142.79:85/api/services/app/Module/Create",
                                                                                moduleObject);
                    DataLoggerControl.Write($"RunModulePostAsync: {response.ReasonPhrase}", serviceName);
                    if (response.IsSuccessStatusCode)
                    {
                        var readTask = response.Content.ReadAsAsync<ModuleObjects>();
                        readTask.Wait();
                        DataLoggerControl.Write($"Module object has been Posted with Serial number: {moduleObject.serialNumber}, Name: {moduleObject.name}," +
                            $" Site Name: {moduleObject.sitesName}, GET_FW_DATE Response: {moduleObject.geT_FW_DATE}, " +
                            $"GET_TND Response: {moduleObject.geT_TND}, GET_FW Response: {moduleObject.geT_FW}, COM Port: {moduleObject.comport}, Baudrate:  {moduleObject.boudrate}," +
                $" Status: {moduleObject.status}, Capacity: {moduleObject.capacity} ", serviceName);
                    }
                    else
                    {
                        DataLoggerControl.Write($"WARNING! Module object failed to save. {response.StatusCode}", serviceName);
                    }
                }
            }
            catch (Exception ex)
            {
                DataLoggerControl.Write($"{ex}", serviceName);
            }

        }
    }
}
