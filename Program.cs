using System;
using System.IO;
using ExcelDataReader;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace AzureFridayToJson
{
    using System.Linq;

    class Program
    {
        static void Main(string[] args)
        {
            //createJson();
            updateMainJson();

            
        }

        static void createJson(){
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var cd = Environment.CurrentDirectory;
            Console.WriteLine(cd);

            var inFilePath = @"./Habrok486FxRate.xlsx";
            var outFilePath = @"./FxRate_Habrok486.json";

            using (var inFile = File.Open(inFilePath, FileMode.Open, FileAccess.Read))
            using (var outFile = File.CreateText(outFilePath))
            {
                using (var reader = ExcelReaderFactory.CreateReader(inFile, new ExcelReaderConfiguration()
                    { FallbackEncoding = Encoding.GetEncoding(1252) }))
                using (var writer = new JsonTextWriter(outFile))
                {
                    writer.Formatting = Formatting.Indented; //I likes it tidy
                    writer.WriteStartArray();
                    reader.Read(); //SKIP FIRST ROW, it's TITLES.
                    do
                    {
                        while (reader.Read())
                        {
                            //peek ahead? Bail before we start anything so we don't get an empty object
                            var ssid = reader.GetString(0);
                            if (string.IsNullOrEmpty(ssid)) break; 

                            writer.WriteStartObject();
                                writer.WritePropertyName("SourceSystemReference");
                                writer.WriteValue(reader.GetString(0));

                                writer.WritePropertyName("SettleToBaseFXRate");
                                writer.WriteValue(reader.GetDouble(1));

                                writer.WritePropertyName("SettleToSystemFXRate");
                                writer.WriteValue(reader.GetDouble(2));

                                writer.WritePropertyName("BaseToLocalFXRate");
                                writer.WriteValue(reader.GetDouble(3));

                                /*
                                <iframe src="https://channel9.msdn.com/Shows/Azure-Friday/Erich-Gamma-introduces-us-to-Visual-Studio-Online-integrated-with-the-Windows-Azure-Portal-Part-1/player" width="960" height="540" allowFullScreen frameBorder="0"></iframe>
                                 */

                            writer.WriteEndObject();
                        }
                    } while (reader.NextResult());
                    writer.WriteEndArray();
                }
            }

        }

        static void updateMainJson(){
            
            var path = @"./Habrok_434_PE.json";
            JArray PEStrings = getFromFile(path);
            JArray fxValuesJson = getFromFile(@"./FxRate_Habrok486.json");
            List<UpdatedFxValues> items = fxValuesJson.ToObject<List<UpdatedFxValues>>();
            
            JArray PE = new JArray() {};

            foreach (var positionEvent in PEStrings)
            {
                var matchedItem = items.FirstOrDefault(e =>
                    e.SourceSystemReference == Convert.ToString(positionEvent["SourceSystemReference"]));

                    positionEvent["ReprocessEvent"] = true;
                    //positionEvent["UtcAsOfDateTime"] = "2021-05-24T00:00:00.0000005Z"; // System UTC
                    positionEvent["UtcAsOfDateTime"] = DateTime.UtcNow.ToString("o");
                    

                    foreach (var allocation in positionEvent["Allocations"])
                    {
                        if(matchedItem == null)
                        {
                            var a = Convert.ToString(positionEvent["SourceSystemReference"]);
                        }
                    
                        allocation["SettleToBaseFXRate"] = matchedItem.SettleToBaseFXRate;
                        allocation["SettleToSystemFXRate"] = matchedItem.SettleToSystemFXRate;
                        allocation["BaseToLocalFXRate"] = matchedItem.BaseToLocalFXRate;

                        if(Convert.ToDecimal(allocation["LocalToSettleFXRate"]) == 1M)
                        {
                        allocation["NetAmountBase"] = Convert.ToDecimal(allocation["NetAmount"]) * matchedItem.SettleToBaseFXRate;
                        allocation["NetAmountSystem"] = Convert.ToDecimal(allocation["NetAmount"]) * matchedItem.SettleToSystemFXRate;
                        allocation["SettleAmountBase"] = Convert.ToDecimal(allocation["SettleAmount"]) * matchedItem.SettleToBaseFXRate;
                        allocation["SettleAmountSystem"] = Convert.ToDecimal(allocation["SettleAmount"]) * matchedItem.SettleToSystemFXRate;
                        allocation["PrincipalAmountBase"] = Convert.ToDecimal(allocation["PrincipalAmount"]) * matchedItem.SettleToBaseFXRate;
                        allocation["PrincipalAmountSystem"] = Convert.ToDecimal(allocation["PrincipalAmount"]) * matchedItem.SettleToSystemFXRate;
                        allocation["CommissionBase"] = Convert.ToDecimal(allocation["Commission"]) * matchedItem.SettleToBaseFXRate;
                        allocation["CommissionSystem"] = Convert.ToDecimal(allocation["Commission"]) * matchedItem.SettleToSystemFXRate;
                        allocation["FeesBase"] = Convert.ToDecimal(allocation["Fees"]) * matchedItem.SettleToBaseFXRate;
                        allocation["FeesSystem"] = Convert.ToDecimal(allocation["Fees"]) * matchedItem.SettleToSystemFXRate;
                        }

                    else{
                        allocation["NetAmountBase"] = (Convert.ToDecimal(allocation["NetAmount"]) * (1/ matchedItem.BaseToLocalFXRate));
                        allocation["NetAmountSystem"] = (Convert.ToDecimal(allocation["NetAmount"]) * (1/ Convert.ToDecimal(allocation["SystemToLocalFXRate"])));
                        allocation["SettleAmountBase"] = Convert.ToDecimal(allocation["SettleAmount"]) * matchedItem.SettleToBaseFXRate;
                        allocation["SettleAmountSystem"] = Convert.ToDecimal(allocation["SettleAmount"]) * matchedItem.SettleToSystemFXRate;
                        allocation["PrincipalAmountBase"] = (Convert.ToDecimal(allocation["PrincipalAmount"]) * (1/ matchedItem.BaseToLocalFXRate));
                        allocation["PrincipalAmountSystem"] = (Convert.ToDecimal(allocation["PrincipalAmount"]) * (1/ Convert.ToDecimal(allocation["SystemToLocalFXRate"])));
                        allocation["CommissionBase"] = (Convert.ToDecimal(allocation["Commission"]) * (1/ matchedItem.BaseToLocalFXRate));
                        allocation["CommissionSystem"] = (Convert.ToDecimal(allocation["Commission"]) * (1/ Convert.ToDecimal(allocation["SystemToLocalFXRate"])));
                        allocation["FeesBase"] = (Convert.ToDecimal(allocation["Fees"]) * (1/ matchedItem.BaseToLocalFXRate));
                        allocation["FeesSystem"] = (Convert.ToDecimal(allocation["Fees"]) * (1/ Convert.ToDecimal(allocation["SystemToLocalFXRate"])));
                    }
                }

                PE.Add(positionEvent);
            }
            using (StreamWriter file = File.CreateText(@"./Habrok_434_updatedPE.json"))
            using (JsonTextWriter writer = new JsonTextWriter(file))
                {
                    PE.WriteTo(writer);
                }
           // Console.WriteLine(PE.ToString());
        }

            public static JArray getFromFile(string path) {
            using (StreamReader file = System.IO.File.OpenText(path))
            {
                var jArray = JArray.Parse(file.ReadToEnd());
                return (jArray);
            }
        }
        public static void LogFile(string infoName, string infoValue)
        {
            StreamWriter log;
            if (!File.Exists("logfileDataV.txt"))
            {
                log = new StreamWriter("logfileDataV.txt");
            }
            else
            {
                log = File.AppendText("logfileDataV.txt");
            }
            // Write to the file:
            
            log.WriteLine(infoName +":" + infoValue);
            
            // Close the stream:
            log.Close();
        }
    }
}
