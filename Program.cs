using System;
using System.IO;
using System.IO.Pipes;
using Button;
using Newtonsoft.Json;
using System.Collections.Generic;
using ClosedXML.Excel;


class Program
{
    static void Main()
    {


        try
        {
            // all your code
            Console.WriteLine("TIA Bridge started");

            TiaClass VICOTia = new TiaClass();   // ✅ persistent
            ExcelClass VICOExcel = new ExcelClass();
            string Engineering = "";        // ✅ persistent
            int PLC = 0;
            List<string> FolderListNames = new List<string>();
            List<string> FolderListLayer = new List<string>();
            List<string> BlockListNames = new List<string>();
            List<string> BlockListLayer = new List<string>();
            //VICOTia.ConnectToTia(@"C:\Program Files\Siemens\Automation\Portal " + "V18" + @"\PublicAPI\" + "V18" + @"\Siemens.Engineering.dll");
            //VICOTia.ImportBlocks(0, "0", @"C:\_Projekte\Block.xml", @"C:\Program Files\Siemens\Automation\Portal " + "V18" + @"\PublicAPI\" + "V18" + @"\Siemens.Engineering.dll");

            //Console.Read();

            while (true) // ✅ accept new clients
            {
                using var server = new NamedPipeServerStream(
                    "TIA_PIPE",
                    PipeDirection.InOut,
                    1,
                    PipeTransmissionMode.Byte);

                Console.WriteLine("Waiting for client...");
                server.WaitForConnection();
                Console.WriteLine("Client connected");


                using var reader = new StreamReader(server);
                using var writer = new StreamWriter(server) { AutoFlush = true };


                // ✅ INNER LOOP: multiple commands per client
                while (server.IsConnected)
                {

                    var request = reader.ReadLine();
                    if (request == null)
                    {
                        Console.WriteLine("no request shutdown");
                        break;
                    }


                    Console.WriteLine($"Command: {request}");


                    var parts = request.Split(new[] { '+' }, 3);

                    var command = parts[0];
                    var argument = parts.Length > 1 ? parts[1] : null;
                    var argument2 = parts.Length > 2 ? parts[2] : null;

                    switch (command)
                    {
                        case "Online?":

                            writer.WriteLine(JsonConvert.SerializeObject(new
                            {
                                success = true,
                                data = "Online!"
                            }));
                            Console.WriteLine("Online!");
                            break;

                        case "Version":
                            Engineering = @"C:\Program Files\Siemens\Automation\Portal " + argument + @"\PublicAPI\" + argument + @"\Siemens.Engineering.dll";
                            writer.WriteLine(JsonConvert.SerializeObject(new
                            {
                                success = true,
                                data = argument + " activated"
                            }));
                            break;

                        case "Connect_to_Tia":
                            bool connected = VICOTia.ConnectToTia(Engineering);
                            if (connected)
                            {
                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "Connected"
                                }));
                            }
                            else
                            {
                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "error"
                                }));
                            }

                            break;


                        case "Get_all_PLCs":
                            Execute(writer, () =>
                                VICOTia.Find_Device());
                            break;

                        case "PLC_Number":
                            PLC = Convert.ToInt16(argument);
                            writer.WriteLine(JsonConvert.SerializeObject(new
                            {
                                success = true,
                                data = "PLCNumberActive"
                            }));
                            break;

                        case "ListFoldersAndProgramBlocks":
                            try
                            {
                                FolderListNames.Clear();
                                FolderListLayer.Clear();
                                BlockListNames.Clear();
                                BlockListLayer.Clear();
                                VICOTia.ListFoldersAndProgramBlocks(PLC, Engineering, out FolderListNames, out FolderListLayer, out BlockListNames, out BlockListLayer);
                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "DataReady"
                                }));
                                Console.WriteLine("DataReady");
                                break;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Import error: {ex}");

                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "error"
                                }));
                                Console.WriteLine("error");
                                break;

                            }

                        case "ListFoldersAndDataType":
                            try
                            {
                                FolderListNames.Clear();
                                FolderListLayer.Clear();
                                BlockListNames.Clear();
                                BlockListLayer.Clear();
                                VICOTia.ListFoldersAndProgramTypes(PLC, Engineering, out FolderListNames, out FolderListLayer, out BlockListNames, out BlockListLayer);
                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "DataReady"
                                }));
                                Console.WriteLine("DataReady");
                                break;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Import error: {ex}");

                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "error"
                                }));
                                Console.WriteLine("error");
                                break;

                            }

                        case "FolderListNames":

                            writer.WriteLine(JsonConvert.SerializeObject(new
                            {
                                success = true,
                                data = FolderListNames
                            }));
                            Console.WriteLine("FolderListNames");
                            break;

                        case "FolderListLayer":

                            writer.WriteLine(JsonConvert.SerializeObject(new
                            {
                                success = true,
                                data = FolderListLayer
                            }));
                            Console.WriteLine("FolderListLayer");
                            break;

                        case "BlockListNames":

                            writer.WriteLine(JsonConvert.SerializeObject(new
                            {
                                success = true,
                                data = BlockListNames
                            }));
                            Console.WriteLine("BlockListNames");
                            break;

                        case "BlockListLayer":

                            writer.WriteLine(JsonConvert.SerializeObject(new
                            {
                                success = true,
                                data = BlockListLayer
                            }));
                            Console.WriteLine("BlockListLayer");
                            break;

                        case "FindFolder":

                            Execute(writer, () =>
                            VICOTia.FindFolderLayer(argument, FolderListNames, FolderListLayer));
                            break;

                        case "Save":
                            try
                            {

                                VICOTia.SaveProjectSafe();
                                //VICOTia.CompilePlc(PLC, Engineering);
                                //VICOTia.SaveProjectSafe();
                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "Save"
                                }));
                                Console.WriteLine("Save");
                                break;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Import error: {ex}");

                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "error"
                                }));
                                Console.WriteLine("error");
                                break;

                            }

                        case "ExportBlock":
                            try
                            {

                                VICOTia.ExportBlock(PLC, VICOTia.FindBlockLayer(argument, BlockListNames, BlockListLayer), argument, argument2, Engineering);
                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "BlockExported"
                                }));
                                Console.WriteLine("BlockExported");
                                break;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Import error: {ex}");

                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "error"
                                }));
                                Console.WriteLine("error");
                                break;

                            }

                        case "ImportBlock":
                            try
                            {


                                VICOTia.ImportBlocks(PLC, argument, argument2, Engineering, BlockListNames);
                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "ImportBlock"
                                }));
                                Console.WriteLine("ImportBlock");
                                break;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Import error: {ex}");

                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "error"
                                }));
                                Console.WriteLine("error");
                                break;

                            }

                        case "ExportDatatype":
                            try
                            {

                                VICOTia.ExportDataType(PLC, VICOTia.FindBlockLayer(argument, BlockListNames, BlockListLayer), argument, argument2, Engineering);
                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "ExportDatatype"
                                }));
                                Console.WriteLine("ExportDatatype");
                                break;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Import error: {ex}");

                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "error"
                                }));
                                Console.WriteLine("error");
                                break;

                            }

                        case "ImportDatatype":
                            try
                            {

                                VICOTia.ImportDataTypes(PLC, argument, argument2, Engineering, BlockListNames);
                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "ImportDatatype"
                                }));
                                Console.WriteLine("ImportDatatype");
                                break;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Import error: {ex}");

                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "error"
                                }));
                                Console.WriteLine("error");
                                break;

                            }

                        case "FolderBlock":
                            try
                            {

                                VICOTia.CreateBlockFolder(PLC, argument, argument2, Engineering);
                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "FolderBlock"
                                }));
                                Console.WriteLine("FolderBlock");
                                break;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Import error: {ex}");

                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "error"
                                }));
                                Console.WriteLine("error");
                                break;

                            }

                        case "FolderDatatype":
                            try
                            {

                                VICOTia.CreateDataTypeFolder(PLC, argument, argument2, Engineering);
                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "FolderDatatype"
                                }));
                                Console.WriteLine("FolderDatatype");
                                break;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Import error: {ex}");

                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "error"
                                }));
                                Console.WriteLine("error");
                                break;

                            }

                        case "AxisBasicConfig":
                            try
                            {
                                List<string> Axis = new List<string>();
                                Axis = VICOTia.PosAxisBasicGen(PLC, Engineering);
                                Console.WriteLine("allAxis");
                                for (int i = 0; i < Axis.Count; i++)
                                {
                                    Console.WriteLine(Axis[i]);
                                }
                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = Axis
                                }));

                                break;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Import error: {ex}");

                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "error"
                                }));
                                Console.WriteLine("error");
                                break;
                            }
                        case "ScanAxis":
                            try
                            {
                                List<string> Axis = new List<string>();
                                Axis = VICOTia.ScanAxis(PLC, Engineering);
                                Console.WriteLine("allAxis");
                                for (int i = 0; i < Axis.Count; i++)
                                {
                                    Console.WriteLine(Axis[i]);
                                }
                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = Axis
                                }));

                                break;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Import error: {ex}");

                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "error"
                                }));
                                Console.WriteLine("error");
                                break;
                            }
                        case "AxisAdress":
                            {
                                List<ExcelClass.ExcelFile> File = new List<ExcelClass.ExcelFile>();
                                File = VICOExcel.readExcelFile(argument);
                                VICOExcel.WriteExcelFile(File, "AxisAdress");
                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = argument + " activated"
                                }));
                                break;
                            }

                        case "AxisCreateDBImportInterfaceExcel":
                            {
                                try
                                {
                                    List<string> Axis = new List<string>();
                                    Axis = VICOTia.ScanAxis(PLC, Engineering);
                                    List<ExcelClass.ExcelFile> File = new List<ExcelClass.ExcelFile>();
                                    VICOExcel.CreateBasicTaglist(File);
                                    for (int i = 0; i < Axis.Count; i++)
                                    {
                                        VICOExcel.AddtoTaglist(File, i + 2, "viCo_Axes_DB." + Axis[i]);
                                    }

                                    VICOExcel.WriteExcelFile(File, "AxisValueTags");
                                    writer.WriteLine(JsonConvert.SerializeObject(new
                                    {
                                        success = true,
                                        data = "AxisCreateDBImportInterfaceExcel"
                                    }));
                                    Console.WriteLine("AxisCreateDBImportInterfaceExcel");
                                    break;
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"AxisCreateDBImportInterfaceExcel error: {ex}");

                                    writer.WriteLine(JsonConvert.SerializeObject(new
                                    {
                                        success = true,
                                        data = "error"
                                    }));
                                    Console.WriteLine("error");
                                    break;
                                }
                            }

                        case "ExportTO":
                            try
                            {

                                VICOTia.AxisExportToConfig(PLC, Engineering, argument);
                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "ExportTO"
                                }));
                                Console.WriteLine("ExportTO");
                                break;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Import error: {ex}");

                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "error"
                                }));
                                Console.WriteLine("error");
                                break;

                            }

                        case "ImportTO":
                            try
                            {

                                VICOTia.AxisImportToConfig(PLC, Engineering, argument);
                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "ImportTO"
                                }));
                                Console.WriteLine("ImportTO");
                                break;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Import error: {ex}");

                                writer.WriteLine(JsonConvert.SerializeObject(new
                                {
                                    success = true,
                                    data = "error"
                                }));
                                Console.WriteLine("error");
                                break;

                            }

                        case "CLOSE":
                            writer.WriteLine(JsonConvert.SerializeObject(new
                            {
                                success = true,
                                data = "closing"
                            }));
                            return; // ✅ shutdown bridge

                        default:
                            Console.WriteLine($"Unknown command: {request}");
                            writer.WriteLine(JsonConvert.SerializeObject(new
                            {

                                success = true,
                                data = "error"

                            }));
                            break;
                    }


                }

            }
        }

        catch (Exception ex)

        {

            Console.WriteLine("FATAL ERROR:");

            Console.WriteLine(ex);

            Console.ReadLine(); // keep console open

        }



        Console.WriteLine("Client disconnected");
    }
    static void Execute(StreamWriter writer, Func<object> action)
    {
        try
        {
            var result = action();
            writer.WriteLine(JsonConvert.SerializeObject(new
            {
                success = true,
                data = result
            }));
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
            writer.WriteLine(JsonConvert.SerializeObject(new
            {
                success = true,
                data = "error"

            }));
        }
    }
}



