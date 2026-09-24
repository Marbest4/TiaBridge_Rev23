namespace Button;

using System.Security;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Net;

using Color = System.Drawing.Color;
using File = System.IO.File;
using System.Threading.Tasks;

using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Windows.Forms;
using System.Xml.Linq;
using ClosedXML.Excel;


public class TiaClass
{
    public static dynamic instanceOfTia;
    public static dynamic projectTIA;
    public static dynamic TiaHW;
    public static dynamic TiaSW;
    public static dynamic SWBlock;
    public static dynamic BibPath;
    public static dynamic SWFolder;
    public static dynamic SWType;
    public static dynamic SWFolderType;

    private bool IsConfigured { get; set; }

    #region basic function
    private Assembly LoadEngineeringAssembly(string path)
    {
        return Assembly.LoadFrom(path);
    }

    private dynamic CreateInstance(Assembly assembly, string typeName, params object[] args)
    {
        Type type = assembly.GetType(typeName);
        return Activator.CreateInstance(type, args);
    }




    //Tia connecten
    public bool ConnectToTia(string dllPath)
    {

        Assembly engineeringAssembly = LoadEngineeringAssembly(dllPath);
        string tiaPortalTypeName = "Siemens.Engineering.TiaPortal";
        string tiaPortalProcessTypeName = "Siemens.Engineering.TiaPortalProcess";

        Type tiaPortalType = engineeringAssembly.GetType(tiaPortalTypeName);
        if (tiaPortalType == null)
        {
            MessageBox.Show($"Type '{tiaPortalTypeName}' not found in the assembly.");

            return false;
        }

        MethodInfo getProcessesMethod = tiaPortalType.GetMethod("GetProcesses", BindingFlags.Static | BindingFlags.Public);
        if (getProcessesMethod == null)
        {
            MessageBox.Show("Method 'GetProcesses' not found.");

            return false;
        }

        List<string> foundInstances = new List<string>();
        IEnumerable<dynamic> tiaPortalProcesses = (IEnumerable<dynamic>)getProcessesMethod.Invoke(null, null);

        foreach (dynamic tiaPortalProcess in tiaPortalProcesses)
        {
            foundInstances.Add(Convert.ToString(tiaPortalProcess.ProjectPath));
        }

        List<dynamic> openInstances = new List<dynamic>(tiaPortalProcesses);

        if (foundInstances.Count == 1)
        {
            instanceOfTia = openInstances[0].Attach();
        }
        else
        {
            if (openInstances.Count > 1)
            {
                MessageBox.Show("There is more then one Tia instance");

                return false;
            }
            else
            {
                MessageBox.Show("NO Open Tia found");

                instanceOfTia = null;
                return false;
            }
        }



        if (instanceOfTia != null)
        {
            if (instanceOfTia.Projects.Count > 0)
            {
                projectTIA = instanceOfTia.Projects[0];
            }
            else
            {
                MessageBox.Show("Cant acess Project (GSD missing?)");

                return false;
            }
        }
        return true;
    }

    public bool SaveProjectSafe()
    {
        try
        {
            if (instanceOfTia == null)
            {
                MessageBox.Show(
                    "TIA is not connected. Please connect first.");
                return false;
            }

            if (instanceOfTia.Projects.Count == 0)
            {
                MessageBox.Show(
                    "No project is open in the connected TIA instance.");
                return false;
            }

            projectTIA = instanceOfTia.Projects[0];
            projectTIA.Save();

            Console.WriteLine("TIA project saved");
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Error saving TIA project:\n{ex.Message}");
            return false;
        }
    }

    private Type FindEngineeringType(string fullTypeName)
    {
        foreach (Assembly asm in AppDomain.CurrentDomain.GetAssemblies())
        {
            if (!asm.FullName.StartsWith("Siemens.Engineering"))
                continue;

            Type t = asm.GetType(fullTypeName, false);
            if (t != null)
                return t;
        }

        return null;
    }

    #endregion


    public List<string> Find_Device()
    {
        List<string> FoundPLCS = new List<string>();

        if (instanceOfTia.Projects[0] != null)
        {
            if (instanceOfTia.Projects[0].Devices.Count > 0)
            {
                for (int i = 0; i < instanceOfTia.Projects[0].Devices.Count; i++)
                {
                    if (instanceOfTia.Projects[0].Devices[i].TypeIdentifier != null)
                    {
                        if (instanceOfTia.Projects[0].Devices[i].TypeIdentifier.Contains("71500"))
                        {
                            FoundPLCS.Add(instanceOfTia.Projects[0].Devices[i].Name);
                        }
                        if (instanceOfTia.Projects[0].Devices[i].TypeIdentifier.Contains("6AV7252"))
                        {
                            FoundPLCS.Add(instanceOfTia.Projects[0].Devices[i].Name);
                        }
                        else
                        {
                            FoundPLCS.Add("NoPLC");
                        }
                    }
                    else
                    {
                        FoundPLCS.Add("NoPLC");
                    }
                }
                if (FoundPLCS.Count < 1)
                {
                    FoundPLCS.Clear();
                    for (int i = 0; i < instanceOfTia.Projects[0].Devices.Count; i++)
                    {
                        FoundPLCS.Add(instanceOfTia.Projects[0].Devices[i].Name);
                    }
                }


            }
            else
            {
                MessageBox.Show("PLC in HW not found");

            }
        }
        return FoundPLCS;
    }

    private dynamic GetSoftware(int deviceIndex, string dllPath)
    {
        Assembly engineeringAssembly = LoadEngineeringAssembly(dllPath);
        string typeName = "Siemens.Engineering.HW.Features.SoftwareContainer";

        foreach (dynamic deviceItem in instanceOfTia.Projects[0]
                                               .Devices[deviceIndex]
                                               .DeviceItems)
        {
            var method = deviceItem.GetType()
                .GetMethod("GetService")
                ?.MakeGenericMethod(engineeringAssembly.GetType(typeName));

            if (method == null) continue;

            dynamic container = method.Invoke(deviceItem, null);
            if (container != null)
                return container.Software;
        }

        return null;
    }

    private string AppendLayer(string baseLayer, int index)
    {
        return string.IsNullOrEmpty(baseLayer)
            ? index.ToString()
            : $"{baseLayer};{index}";
    }
    private void TraverseBlockGroup(
        dynamic blockGroup,
        string layer,
        List<string> folderNames,
        List<string> folderLayers,
        List<string> blockNames,
        List<string> blockLayers)
    {
        // Blocks in this folder
        for (int i = 0; i < blockGroup.Blocks.Count; i++)
        {
            blockNames.Add(blockGroup.Blocks[i].Name);
            blockLayers.Add(AppendLayer(layer, i));
        }

        // Subfolders
        for (int i = 0; i < blockGroup.Groups.Count; i++)
        {
            var group = blockGroup.Groups[i];
            var newLayer = AppendLayer(layer, i);

            folderNames.Add(group.Name);
            folderLayers.Add(newLayer);

            // Recursive call ✅
            TraverseBlockGroup(
                group,
                newLayer,
                folderNames,
                folderLayers,
                blockNames,
                blockLayers);
        }
    }
    public void ListFoldersAndProgramBlocks(
        int deviceIndex,
        string dllPath,
        out List<string> folderNames,
        out List<string> folderLayers,
        out List<string> blockNames,
        out List<string> blockLayers)
    {
        folderNames = new();
        folderLayers = new();
        blockNames = new();
        blockLayers = new();

        dynamic software = GetSoftware(deviceIndex, dllPath);
        if (software == null)
            return;

        TraverseBlockGroup(
            software.BlockGroup,
            "",
            folderNames,
            folderLayers,
            blockNames,
            blockLayers);
    }

    private void TraverseTypeGroup(
        dynamic typeGroup,
        string layer,
        List<string> folderNames,
        List<string> folderLayers,
        List<string> typeNames,
        List<string> typeLayers)
    {
        // Types in current folder
        for (int i = 0; i < typeGroup.Types.Count; i++)
        {
            typeNames.Add(typeGroup.Types[i].Name);
            typeLayers.Add(AppendLayer(layer, i));

            Console.WriteLine(typeGroup.Types[i].Name);
            Console.WriteLine(AppendLayer(layer, i));
        }

        // Subfolders
        for (int i = 0; i < typeGroup.Groups.Count; i++)
        {
            var group = typeGroup.Groups[i];
            var newLayer = AppendLayer(layer, i);

            folderNames.Add(group.Name);
            folderLayers.Add(newLayer);

            Console.WriteLine(group.Name);
            Console.WriteLine(newLayer);

            // ✅ recursion
            TraverseTypeGroup(
                group,
                newLayer,
                folderNames,
                folderLayers,
                typeNames,
                typeLayers);
        }
    }
    public void ListFoldersAndProgramTypes(
        int deviceIndex,
        string dllPath,
        out List<string> folderNames,
        out List<string> folderLayers,
        out List<string> typeNames,
        out List<string> typeLayers)
    {
        folderNames = new List<string>();
        folderLayers = new List<string>();
        typeNames = new List<string>();
        typeLayers = new List<string>();

        dynamic software = GetSoftware(deviceIndex, dllPath);
        if (software == null)
            return;

        TraverseTypeGroup(
            software.TypeGroup,
            "",
            folderNames,
            folderLayers,
            typeNames,
            typeLayers);
    }
    public string FindFolderLayer(string FolderNameSearch, List<string> FolderListNames, List<string> FolderListLayer)
    {
        string FolderPosition = "";
        for (int i = 0; i < FolderListNames.Count; i++)
        {
            Console.WriteLine(FolderListNames[i]);
            Console.WriteLine(FolderListLayer[i]);
            if (FolderListNames[i] == FolderNameSearch)
            {

                FolderPosition = FolderListLayer[i].ToString();
                return FolderPosition;
            }
        }
        return "Not Found";

    }

    public string FindBlockLayer(string BlockNameSearch, List<string> BlockListNames, List<string> BlockListLayer)
    {
        string FolderPosition = "";
        for (int i = 0; i < BlockListNames.Count; i++)
        {

            if (BlockListNames[i] == BlockNameSearch)
            {

                FolderPosition = BlockListLayer[i].ToString();



                var parts = FolderPosition.Split(';');

                string result = parts.Length > 1
                    ? string.Join(";", parts.Take(parts.Length - 1))
                    : FolderPosition;




                return result;
            }
        }
        return "";

    }

    public bool lookforBlock(string BlockNameSearch, List<string> BlockListNames)
    {
        string FolderPosition = "";
        for (int i = 0; i < BlockListNames.Count; i++)
        {

            if (BlockListNames[i] == BlockNameSearch)
            {




                return true;
            }
        }
        return false;

    }


    private dynamic ResolveBlockGroup(dynamic rootGroup, string folderPosition)
    {
        if (string.IsNullOrWhiteSpace(folderPosition))
            return rootGroup;

        string[] indices = folderPosition.Split(';');
        dynamic currentGroup = rootGroup;

        foreach (string indexText in indices)
        {
            if (!int.TryParse(indexText, out int index))
                return null;

            if (index < 0 || index >= currentGroup.Groups.Count)
                return null;

            currentGroup = currentGroup.Groups[index];
        }

        return currentGroup;
    }


    private string GetBlockName(string filePath)
    {
        XDocument doc = XDocument.Load(filePath);

        XElement blockNameElement = doc
            .Descendants()
            .FirstOrDefault(x =>
                x.Name.LocalName == "AttributeList")
            ?.Elements()
            .FirstOrDefault(x =>
                x.Name.LocalName == "Name");

        return blockNameElement?.Value;
    }

    public void ImportBlocks(
        int deviceIndex,
        string folderPosition,
        string filePath,
        string dllPath,
        List<string> BlockListNames)
    {

        if (!File.Exists(filePath))
        {
            MessageBox.Show("Following file does not exist: " + filePath);
            return;
        }

        dynamic software = GetSoftware(deviceIndex, dllPath);
        if (software == null)
        {
            MessageBox.Show("Software not found");
            return;
        }

        // Find target BlockGroup based on folderPosition
        dynamic targetBlockGroup = ResolveBlockGroup(
            software.BlockGroup,
            folderPosition);

        if (targetBlockGroup == null)
        {
            MessageBox.Show("No position for integration found: " + folderPosition);
            return;
        }

        string blockName = GetBlockName(filePath);

        bool found = lookforBlock(blockName, BlockListNames);
        if (found)
        {
            Console.WriteLine($"Block '{blockName}' already exists.");
            return;

        }






        // Import options
        Assembly engineeringAssembly = LoadEngineeringAssembly(dllPath);
        var importOptionsType =
            engineeringAssembly.GetType("Siemens.Engineering.ImportOptions");

        dynamic importOptions = Activator.CreateInstance(importOptionsType);

        var overwriteProp =
            importOptionsType.GetProperty("OverwriteExisting")
            ?? importOptionsType.GetProperty("OverrideExisting");

        overwriteProp?.SetValue(importOptions, true);

        importOptionsType
            .GetProperty("KeepOriginalName")
            ?.SetValue(importOptions, true);

        // Import blocks
        targetBlockGroup.Blocks.Import(
            new FileInfo(filePath),
            importOptions);





    }


    public void ExportBlock(
        int deviceIndex,
        string folderPosition,
        string blockName,
        string exportFilePath,
        string dllPath)
    {

        if (string.IsNullOrWhiteSpace(blockName))
            throw new ArgumentException("Block name must be provided");

        dynamic software = GetSoftware(deviceIndex, dllPath);
        if (software == null)
            throw new Exception("Software not found");

        // Resolve BlockGroup from folder position (unlimited depth)
        dynamic targetBlockGroup = ResolveBlockGroup(
            software.BlockGroup,
            folderPosition);

        if (targetBlockGroup == null)
            throw new Exception($"Invalid folder position: {folderPosition}");

        // Find block by name
        dynamic block = null;



        foreach (dynamic b in targetBlockGroup.Blocks)
        {


            if (string.Equals(
                b.Name?.Trim(),
                blockName,
                StringComparison.OrdinalIgnoreCase))
            {
                block = b;
            }
        }

        if (block == null)
            throw new Exception("Block not found in folder " + folderPosition + " blockname = " + blockName);
        if (block == null)
            throw new Exception($"Block '{blockName}' not found");

        // ✅ Ensure overwrite
        if (File.Exists(exportFilePath))
        {
            File.Delete(exportFilePath);
        }

        // Export options
        Assembly engineeringAssembly = LoadEngineeringAssembly(dllPath);
        var exportOptionsType =
            engineeringAssembly.GetType("Siemens.Engineering.ExportOptions");

        dynamic exportOptions = Activator.CreateInstance(exportOptionsType);

        // Export
        block.Export(new FileInfo(exportFilePath), exportOptions);

        Console.WriteLine($"Block '{blockName}' exported to {exportFilePath}");

    }

    public dynamic CreateBlockFolder(
        int deviceIndex,
        string parentFolderPosition,
        string newFolderName,
        string dllPath)
    {
        if (string.IsNullOrWhiteSpace(newFolderName))
            throw new ArgumentException("Folder name must be provided");

        dynamic software = GetSoftware(deviceIndex, dllPath);
        if (software == null)
            throw new Exception("Software not found");

        // Resolve parent BlockGroup
        dynamic parentGroup = ResolveBlockGroup(
            software.BlockGroup,
            parentFolderPosition);

        if (parentGroup == null)
            throw new Exception(
                $"Invalid block folder position: {parentFolderPosition}");

        // Check if folder already exists
        foreach (dynamic g in parentGroup.Groups)
        {
            if (string.Equals(
                g.Name,
                newFolderName,
                StringComparison.OrdinalIgnoreCase))
            {
                return g; // already exists
            }
        }

        // ✅ Create folder
        dynamic newGroup = parentGroup.Groups.Create(newFolderName);
        return newGroup;
    }
    private dynamic ResolveDataTypeGroup(
        dynamic rootGroup,
        string folderPosition)
    {
        if (string.IsNullOrWhiteSpace(folderPosition))
            return rootGroup;

        dynamic currentGroup = rootGroup;
        var pathParts = folderPosition.Split(';');

        foreach (var part in pathParts)
        {
            int index = int.Parse(part);

            if (index < 0 || index >= currentGroup.Groups.Count)
                return null;

            currentGroup = currentGroup.Groups[index];
        }

        return currentGroup;
    }
    public void ImportDataTypes(
        int deviceIndex,
        string folderPosition,
        string filePath,
        string dllPath,
        List<string> BlockListNames)
    {

        if (!File.Exists(filePath))
        {
            MessageBox.Show("Following file does not exist: " + filePath);
            return;
        }

        dynamic software = GetSoftware(deviceIndex, dllPath);
        if (software == null)
        {
            MessageBox.Show("Software not found");
            return;
        }

        // Resolve target DataTypeGroup (same concept as BlockGroup)
        dynamic targetDataTypeGroup = ResolveDataTypeGroup(
      software.TypeGroup,
      folderPosition);

        if (targetDataTypeGroup == null)
        {
            MessageBox.Show("No position for integration found: " + folderPosition);
            return;
        }

        string blockName = GetBlockName(filePath);

        bool found = lookforBlock(blockName, BlockListNames);
        if (found)
        {
            Console.WriteLine($"Block '{blockName}' already exists.");
            return;

        }

        Assembly engineeringAssembly = LoadEngineeringAssembly(dllPath);
        var importOptionsType =
            engineeringAssembly.GetType("Siemens.Engineering.ImportOptions");

        dynamic importOptions = Activator.CreateInstance(importOptionsType);

        // Overwrite existing UDTs
        var overwriteProp =
            importOptionsType.GetProperty("OverwriteExisting")
         ?? importOptionsType.GetProperty("OverrideExisting");

        overwriteProp?.SetValue(importOptions, true);

        importOptionsType
            .GetProperty("KeepOriginalName")
            ?.SetValue(importOptions, true);

        // ✅ Import UDTs
        targetDataTypeGroup.Types.Import(
            new FileInfo(filePath),
            importOptions);

    }

    public void ExportDataType(
     int deviceIndex,
     string folderPosition,
     string dataTypeName,
     string exportFilePath,
     string dllPath)
    {
        if (string.IsNullOrWhiteSpace(dataTypeName))
            throw new ArgumentException("Data type name must be provided");

        dynamic software = GetSoftware(deviceIndex, dllPath);
        if (software == null)
            throw new Exception("Software not found");

        // Resolve DataTypeGroup from folder position (unlimited depth)
        dynamic targetDataTypeGroup = ResolveDataTypeGroup(
     software.TypeGroup,
     folderPosition);
        if (targetDataTypeGroup == null)
            throw new Exception($"Invalid folder position: {folderPosition}");

        // Find datatype by name
        dynamic dataType = null;

        foreach (dynamic dt in targetDataTypeGroup.Types)
        {
            if (string.Equals(
                dt.Name?.Trim(),
                dataTypeName,

                StringComparison.OrdinalIgnoreCase))
            {
                dataType = dt;
            }
        }

        if (dataType == null)
            throw new Exception("Data type not found in folder " + folderPosition + " datatypename = " + dataTypeName);

        if (dataType == null)
            throw new Exception($"Data type '{dataTypeName}' not found");

        // ✅ Ensure overwrite
        if (File.Exists(exportFilePath))
        {
            File.Delete(exportFilePath);
        }

        // Export options
        Assembly engineeringAssembly = LoadEngineeringAssembly(dllPath);
        var exportOptionsType =
            engineeringAssembly.GetType("Siemens.Engineering.ExportOptions");

        dynamic exportOptions = Activator.CreateInstance(exportOptionsType);

        // Export
        dataType.Export(new FileInfo(exportFilePath), exportOptions);

        Console.WriteLine($"Data type '{dataTypeName}' exported to {exportFilePath}");
    }
    public dynamic CreateDataTypeFolder(
        int deviceIndex,
        string parentFolderPosition,
        string newFolderName,
        string dllPath)
    {
        if (string.IsNullOrWhiteSpace(newFolderName))
            throw new ArgumentException("Folder name must be provided");

        dynamic software = GetSoftware(deviceIndex, dllPath);
        if (software == null)
            throw new Exception("Software not found");

        // Resolve parent DataTypeGroup
        dynamic parentGroup = ResolveDataTypeGroup(
            software.TypeGroup,
            parentFolderPosition);

        if (parentGroup == null)
            throw new Exception(
                $"Invalid data type folder position: {parentFolderPosition}");

        // Check if folder already exists
        foreach (dynamic g in parentGroup.Groups)
        {
            if (string.Equals(
                g.Name,
                newFolderName,
                StringComparison.OrdinalIgnoreCase))
            {
                return g; // already exists
            }
        }

        // ✅ Create folder
        dynamic newGroup = parentGroup.Groups.Create(newFolderName);
        return newGroup;
    }
    public void ListDevicesAndGSD(string dllPath, out List<string> DeviceName, out List<string> DeviceType)
    {


        Assembly engineeringAssembly = LoadEngineeringAssembly(dllPath);
        DeviceName = new List<string>();
        DeviceType = new List<string>();



        if (0 != instanceOfTia.Projects[0].UngroupedDevicesGroup.Devices.Count)
        {
            for (int i = 0; i < instanceOfTia.Projects[0].UngroupedDevicesGroup.Devices.Count; i++)
            {
                if (null != instanceOfTia.Projects[0].UngroupedDevicesGroup.Devices[i].TypeIdentifier)
                {

                    if (null != instanceOfTia.Projects[0].UngroupedDevicesGroup.Devices[i].GetAttribute("TypeIdentifier").ToString())
                    {
                        #region DeviceType

                        DeviceName.Add(instanceOfTia.Projects[0].UngroupedDevicesGroup.Devices[i].GetAttribute("Name").ToString());
                        DeviceType.Add(instanceOfTia.Projects[0].UngroupedDevicesGroup.Devices[i].GetAttribute("TypeIdentifier").ToString());
                        #endregion


                    }

                }
            }
        }

    }

    public List<string> PosAxisBasicGen(int Device, string dllPath)
    {
        List<string> ConfiguratedAxis = new List<string>();

        Console.WriteLine("Start axis run");


        Assembly engineeringAssembly = LoadEngineeringAssembly(dllPath);

        string softwareContainerTypeName = "Siemens.Engineering.HW.Features.SoftwareContainer";



        dynamic softwareContainer = null;
        for (int i = 0; i < instanceOfTia.Projects[0].Devices[Device].DeviceItems.Count; i++)
        {
            dynamic deviceItem = instanceOfTia.Projects[0].Devices[Device].DeviceItems[i];

            // Verwende GetService<T> mit dem spezifischen Typ
            MethodInfo getServiceMethod = deviceItem.GetType().GetMethod("GetService").MakeGenericMethod(engineeringAssembly.GetType(softwareContainerTypeName));
            softwareContainer = getServiceMethod.Invoke(deviceItem, null);
            if (softwareContainer != null)
            {
                break;
            }
        }

        if (softwareContainer != null)
        {
            Console.WriteLine("softwarecontainer = true");
            dynamic softwarePath = softwareContainer.Software as dynamic;
            string name = softwarePath.Name;


            if ((softwarePath?.TechnologicalObjectGroup?.TechnologicalObjects?.Count ?? 0) > 0)
            {

                int ElementsLevel0 = softwarePath.TechnologicalObjectGroup.TechnologicalObjects.Count;


                // Abfrage ob element > 0
                if (ElementsLevel0 > 0)
                {




                    for (int i = 0; i < ElementsLevel0; i++)
                    {

                        //ordner namentlich aufzählen
                        var AxisType = softwarePath.TechnologicalObjectGroup.TechnologicalObjects[i].OfSystemLibElement;
                        var AxisName = softwarePath.TechnologicalObjectGroup.TechnologicalObjects[i].Name;
                        var Linear_Round = 0;
                        Console.WriteLine(AxisName);


                        if (AxisType.Contains("Axis"))
                        {
                            Console.WriteLine(AxisType);
                            ConfiguratedAxis.Add(AxisName);
                            if (softwarePath.TechnologicalObjectGroup.TechnologicalObjects[i].Parameters.Count > 0)
                            {
                                if (AxisName.Contains("X") || AxisName.Contains("Y") || AxisName.Contains("Z") || AxisName.Contains("F"))
                                {
                                    Linear_Round = 0;
                                    Console.WriteLine("linear");
                                }
                                else
                                {
                                    Linear_Round = 1;
                                    Console.WriteLine("round");
                                }
                                var parameters = softwarePath.TechnologicalObjectGroup.TechnologicalObjects[i].Parameters;
                                foreach (var parameter in parameters)
                                {
                                    try
                                    {
                                        var paramName = parameter.GetAttribute("Name").ToString();

                                        if (paramName == "_Properties.MotionType")
                                        {

                                            parameter.Value = Linear_Round; //Linear / Rotor
                                        }
                                        if (paramName == "Modulo.Enable")
                                        {
                                            parameter.Value = 0; //Modulachse
                                        }
                                        if (paramName == "Actor.DataAdaption")
                                        {
                                            parameter.Value = 0; //Offline Achse
                                        }
                                        if (paramName == "Sensor[1].DataAdaption")
                                        {
                                            parameter.Value = 0; //Offline Achse
                                        }
                                        if (paramName == "Sensor[1].MountingMode")
                                        {
                                            parameter.Value = Linear_Round;//Linear / Rotor
                                        }
                                        if (paramName == "Simulation.Mode")
                                        {
                                            parameter.Value = 1; //sim mode
                                        }
                                        if (paramName == "Sensor[1].Type")
                                        {
                                            parameter.Value = 2; //Zyklisch Absolut
                                        }
                                        if (paramName == "TorqueLimiting.PositionBasedMonitorings")
                                        {
                                            parameter.Value = 0; //deaktivieren der Schleppüberwachung
                                        }
                                        if (paramName == "FollowingError.EnableMonitoring")
                                        {
                                            parameter.Value = 0; //deaktivieren der Schleppüberwachung
                                        }
                                        if (paramName == "PositionControl.EnableDSC")
                                        {
                                            parameter.Value = 0; //deaktivieren der Schleppüberwachung
                                        }
                                    }
                                    catch (System.Exception)
                                    {


                                    }

                                }


                            }

                        }


                    }
                }
            }



            var techGroup = softwarePath?.TechnologicalObjectGroup;

            if (techGroup != null && HasProperty(techGroup, "Groups"))
            {

                if ((softwarePath?.TechnologicalObjectGroup?.Groups?.Count ?? 0) > 0)
                {
                    int groups = softwarePath.TechnologicalObjectGroup.Groups.Count;
                    // Abfrage ob element > 0
                    if (groups > 0)
                    {

                        for (int i = 0; i < groups; i++)
                        {

                            int ElementsLevel0 = softwarePath.TechnologicalObjectGroup.Groups[i].TechnologicalObjects.Count;

                            for (int counter1 = 0; counter1 < ElementsLevel0; counter1++)
                            {

                                //ordner namentlich aufzählen
                                var AxisType = softwarePath.TechnologicalObjectGroup.Groups[i].TechnologicalObjects[counter1].OfSystemLibElement;
                                var AxisName = softwarePath.TechnologicalObjectGroup.Groups[i].TechnologicalObjects[counter1].Name;
                                var Linear_Round = 0;
                                Console.WriteLine(AxisName);


                                if (AxisType.Contains("Axis"))
                                {
                                    Console.WriteLine(AxisType);
                                    ConfiguratedAxis.Add(AxisName);
                                    if (softwarePath.TechnologicalObjectGroup.Groups[i].TechnologicalObjects[counter1].Parameters.Count > 0)
                                    {
                                        if (AxisName.Contains("X") || AxisName.Contains("Y") || AxisName.Contains("Z") || AxisName.Contains("F"))
                                        {
                                            Linear_Round = 0;
                                            Console.WriteLine("linear");
                                        }
                                        else
                                        {
                                            Linear_Round = 1;
                                            Console.WriteLine("round");
                                        }

                                        var parameters = softwarePath.TechnologicalObjectGroup.Groups[i].TechnologicalObjects[counter1].Parameters;
                                        foreach (var parameter in parameters)
                                        {
                                            try
                                            {
                                                var paramName = parameter.GetAttribute("Name").ToString();

                                                if (paramName == "_Properties.MotionType")
                                                {

                                                    parameter.Value = Linear_Round; //Linear / Rotor
                                                }
                                                if (paramName == "Modulo.Enable")
                                                {
                                                    parameter.Value = 0; //Modulachse
                                                }
                                                if (paramName == "Actor.DataAdaption")
                                                {
                                                    parameter.Value = 0; //Offline Achse
                                                }
                                                if (paramName == "Sensor[1].DataAdaption")
                                                {
                                                    parameter.Value = 0; //Offline Achse
                                                }
                                                if (paramName == "Sensor[1].MountingMode")
                                                {
                                                    parameter.Value = Linear_Round;//Linear / Rotor
                                                }
                                                if (paramName == "Simulation.Mode")
                                                {
                                                    parameter.Value = 1; //sim mode
                                                }
                                                if (paramName == "Sensor[1].Type")
                                                {
                                                    parameter.Value = 2; //Zyklisch Absolut
                                                }
                                                if (paramName == "TorqueLimiting.PositionBasedMonitorings")
                                                {
                                                    parameter.Value = 0; //deaktivieren der Schleppüberwachung
                                                }
                                                if (paramName == "FollowingError.EnableMonitoring")
                                                {
                                                    parameter.Value = 0; //deaktivieren der Schleppüberwachung
                                                }
                                                if (paramName == "PositionControl.EnableDSC")
                                                {
                                                    parameter.Value = 0; //deaktivieren der Schleppüberwachung
                                                }
                                            }
                                            catch (System.Exception)
                                            {


                                            }

                                        }



                                    }

                                }


                            }

                            var techGroup2 = softwarePath?.TechnologicalObjectGroup.Groups[i];
                            if (techGroup2 != null && HasProperty(techGroup, "Groups"))
                            {
                                if ((softwarePath?.TechnologicalObjectGroup?.Groups[i].Groups?.Count ?? 0) > 0)
                                {
                                    for (int j = 0; j < techGroup2.Groups.Count; j++)
                                    {

                                        ElementsLevel0 = softwarePath.TechnologicalObjectGroup.Groups[i].Groups[j].TechnologicalObjects.Count;

                                        for (int counter1 = 0; counter1 < ElementsLevel0; counter1++)
                                        {

                                            //ordner namentlich aufzählen
                                            var AxisType = softwarePath.TechnologicalObjectGroup.Groups[i].Groups[j].TechnologicalObjects[counter1].OfSystemLibElement;
                                            var AxisName = softwarePath.TechnologicalObjectGroup.Groups[i].Groups[j].TechnologicalObjects[counter1].Name;
                                            var Linear_Round = 0;
                                            Console.WriteLine(AxisName);


                                            if (AxisType.Contains("Axis"))
                                            {
                                                Console.WriteLine(AxisType);
                                                ConfiguratedAxis.Add(AxisName);
                                                if (softwarePath.TechnologicalObjectGroup.Groups[i].Groups[j].TechnologicalObjects[counter1].Parameters.Count > 0)
                                                {
                                                    if (AxisName.Contains("X") || AxisName.Contains("Y") || AxisName.Contains("Z") || AxisName.Contains("F"))
                                                    {
                                                        Linear_Round = 0;
                                                        Console.WriteLine("linear");
                                                    }
                                                    else
                                                    {
                                                        Linear_Round = 1;
                                                        Console.WriteLine("round");
                                                    }

                                                    var parameters = softwarePath.TechnologicalObjectGroup.Groups[i].Groups[j].TechnologicalObjects[counter1].Parameters;
                                                    foreach (var parameter in parameters)
                                                    {
                                                        try
                                                        {
                                                            var paramName = parameter.GetAttribute("Name").ToString();

                                                            if (paramName == "_Properties.MotionType")
                                                            {

                                                                parameter.Value = Linear_Round; //Linear / Rotor
                                                            }
                                                            if (paramName == "Modulo.Enable")
                                                            {
                                                                parameter.Value = 0; //Modulachse
                                                            }
                                                            if (paramName == "Actor.DataAdaption")
                                                            {
                                                                parameter.Value = 0; //Offline Achse
                                                            }
                                                            if (paramName == "Sensor[1].DataAdaption")
                                                            {
                                                                parameter.Value = 0; //Offline Achse
                                                            }
                                                            if (paramName == "Sensor[1].MountingMode")
                                                            {
                                                                parameter.Value = Linear_Round;//Linear / Rotor
                                                            }
                                                            if (paramName == "Simulation.Mode")
                                                            {
                                                                parameter.Value = 1; //sim mode
                                                            }
                                                            if (paramName == "Sensor[1].Type")
                                                            {
                                                                parameter.Value = 2; //Zyklisch Absolut
                                                            }
                                                            if (paramName == "TorqueLimiting.PositionBasedMonitorings")
                                                            {
                                                                parameter.Value = 0; //deaktivieren der Schleppüberwachung
                                                            }
                                                            if (paramName == "FollowingError.EnableMonitoring")
                                                            {
                                                                parameter.Value = 0; //deaktivieren der Schleppüberwachung
                                                            }
                                                            if (paramName == "PositionControl.EnableDSC")
                                                            {
                                                                parameter.Value = 0; //deaktivieren der Schleppüberwachung
                                                            }
                                                        }
                                                        catch (System.Exception)
                                                        {


                                                        }

                                                    }



                                                }

                                            }
                                        }
                                    }
                                }


                            }


                        }



                    }

                }

            }
        }


        return ConfiguratedAxis;
    }

    public List<string> ScanAxis(int Device, string dllPath)
    {
        List<string> ConfiguratedAxis = new List<string>();

        Console.WriteLine("Start axis run");


        Assembly engineeringAssembly = LoadEngineeringAssembly(dllPath);

        string softwareContainerTypeName = "Siemens.Engineering.HW.Features.SoftwareContainer";



        dynamic softwareContainer = null;
        for (int i = 0; i < instanceOfTia.Projects[0].Devices[Device].DeviceItems.Count; i++)
        {
            dynamic deviceItem = instanceOfTia.Projects[0].Devices[Device].DeviceItems[i];

            // Verwende GetService<T> mit dem spezifischen Typ
            MethodInfo getServiceMethod = deviceItem.GetType().GetMethod("GetService").MakeGenericMethod(engineeringAssembly.GetType(softwareContainerTypeName));
            softwareContainer = getServiceMethod.Invoke(deviceItem, null);
            if (softwareContainer != null)
            {
                break;
            }
        }

        if (softwareContainer != null)
        {
            Console.WriteLine("softwarecontainer = true");
            dynamic softwarePath = softwareContainer.Software as dynamic;
            string name = softwarePath.Name;


            if ((softwarePath?.TechnologicalObjectGroup?.TechnologicalObjects?.Count ?? 0) > 0)
            {

                int ElementsLevel0 = softwarePath.TechnologicalObjectGroup.TechnologicalObjects.Count;


                // Abfrage ob element > 0
                if (ElementsLevel0 > 0)
                {




                    for (int i = 0; i < ElementsLevel0; i++)
                    {

                        //ordner namentlich aufzählen
                        var AxisType = softwarePath.TechnologicalObjectGroup.TechnologicalObjects[i].OfSystemLibElement;
                        var AxisName = softwarePath.TechnologicalObjectGroup.TechnologicalObjects[i].Name;
                        var Linear_Round = 0;
                        Console.WriteLine(AxisName);


                        if (AxisType.Contains("Axis"))
                        {
                            Console.WriteLine(AxisType);
                            ConfiguratedAxis.Add(AxisName);

                        }


                    }
                }
            }



            var techGroup = softwarePath?.TechnologicalObjectGroup;

            if (techGroup != null && HasProperty(techGroup, "Groups"))
            {

                if ((softwarePath?.TechnologicalObjectGroup?.Groups?.Count ?? 0) > 0)
                {
                    int groups = softwarePath.TechnologicalObjectGroup.Groups.Count;
                    // Abfrage ob element > 0
                    if (groups > 0)
                    {

                        for (int i = 0; i < groups; i++)
                        {

                            int ElementsLevel0 = softwarePath.TechnologicalObjectGroup.Groups[i].TechnologicalObjects.Count;

                            for (int counter1 = 0; counter1 < ElementsLevel0; counter1++)
                            {

                                //ordner namentlich aufzählen
                                var AxisType = softwarePath.TechnologicalObjectGroup.Groups[i].TechnologicalObjects[counter1].OfSystemLibElement;
                                var AxisName = softwarePath.TechnologicalObjectGroup.Groups[i].TechnologicalObjects[counter1].Name;
                                var Linear_Round = 0;
                                Console.WriteLine(AxisName);


                                if (AxisType.Contains("Axis"))
                                {
                                    Console.WriteLine(AxisType);
                                    ConfiguratedAxis.Add(AxisName);

                                }


                            }

                            var techGroup2 = softwarePath?.TechnologicalObjectGroup.Groups[i];
                            if (techGroup2 != null && HasProperty(techGroup, "Groups"))
                            {
                                if ((softwarePath?.TechnologicalObjectGroup?.Groups[i].Groups?.Count ?? 0) > 0)
                                {
                                    for (int j = 0; j < techGroup2.Groups.Count; j++)
                                    {

                                        ElementsLevel0 = softwarePath.TechnologicalObjectGroup.Groups[i].Groups[j].TechnologicalObjects.Count;

                                        for (int counter1 = 0; counter1 < ElementsLevel0; counter1++)
                                        {

                                            //ordner namentlich aufzählen
                                            var AxisType = softwarePath.TechnologicalObjectGroup.Groups[i].Groups[j].TechnologicalObjects[counter1].OfSystemLibElement;
                                            var AxisName = softwarePath.TechnologicalObjectGroup.Groups[i].Groups[j].TechnologicalObjects[counter1].Name;
                                            var Linear_Round = 0;
                                            Console.WriteLine(AxisName);


                                            if (AxisType.Contains("Axis"))
                                            {
                                                Console.WriteLine(AxisType);
                                                ConfiguratedAxis.Add(AxisName);

                                            }
                                        }
                                    }
                                }


                            }


                        }



                    }

                }

            }
        }


        return ConfiguratedAxis;
    }

    public class ToConfig
    {
        public string Name { get; set; } = "";
        public List<ToParameter> Parameters { get; set; }

    }
    public class ToParameter
    {
        public string Name { get; set; } = "";
        public string Value { get; set; } = "";

    }
    public void AxisExportToConfig(int Device, string dllPath, string folderPath)
    {
        List<ToConfig> Axis = new List<ToConfig>();

        Console.WriteLine("Start Export To");


        Assembly engineeringAssembly = LoadEngineeringAssembly(dllPath);

        string softwareContainerTypeName = "Siemens.Engineering.HW.Features.SoftwareContainer";



        dynamic softwareContainer = null;
        for (int i = 0; i < instanceOfTia.Projects[0].Devices[Device].DeviceItems.Count; i++)
        {
            dynamic deviceItem = instanceOfTia.Projects[0].Devices[Device].DeviceItems[i];

            // Verwende GetService<T> mit dem spezifischen Typ
            MethodInfo getServiceMethod = deviceItem.GetType().GetMethod("GetService").MakeGenericMethod(engineeringAssembly.GetType(softwareContainerTypeName));
            softwareContainer = getServiceMethod.Invoke(deviceItem, null);
            if (softwareContainer != null)
            {
                break;
            }
        }

        if (softwareContainer != null)
        {
            Console.WriteLine("softwarecontainer = true");
            dynamic softwarePath = softwareContainer.Software as dynamic;
            string name = softwarePath.Name;


            if ((softwarePath?.TechnologicalObjectGroup?.TechnologicalObjects?.Count ?? 0) > 0)
            {

                int ElementsLevel0 = softwarePath.TechnologicalObjectGroup.TechnologicalObjects.Count;


                // Abfrage ob element > 0
                if (ElementsLevel0 > 0)
                {




                    for (int i = 0; i < ElementsLevel0; i++)
                    {

                        //ordner namentlich aufzählen
                        var AxisType = softwarePath.TechnologicalObjectGroup.TechnologicalObjects[i].OfSystemLibElement;
                        var AxisName = softwarePath.TechnologicalObjectGroup.TechnologicalObjects[i].Name;
                        var Linear_Round = 0;
                        Console.WriteLine(AxisName);


                        if (AxisType.Contains("Axis"))
                        {


                            if (softwarePath.TechnologicalObjectGroup.TechnologicalObjects[i].Parameters.Count > 0)
                            {
                                List<ToParameter> Parameters = new List<ToParameter>();

                                var parameters = softwarePath.TechnologicalObjectGroup.TechnologicalObjects[i].Parameters;
                                foreach (var parameter in parameters)
                                {
                                    try
                                    {
                                        Parameters.Add(new ToParameter
                                        {
                                            Name = parameter.GetAttribute("Name").ToString(),
                                            Value = parameter.GetAttribute("Value").ToString()

                                        });


                                    }
                                    catch (System.Exception)
                                    {


                                    }

                                }
                                Axis.Add(new ToConfig
                                {
                                    Name = AxisName,
                                    Parameters = Parameters

                                });


                            }

                        }


                    }
                }
            }



            var techGroup = softwarePath?.TechnologicalObjectGroup;

            if (techGroup != null && HasProperty(techGroup, "Groups"))
            {

                if ((softwarePath?.TechnologicalObjectGroup?.Groups?.Count ?? 0) > 0)
                {
                    int groups = softwarePath.TechnologicalObjectGroup.Groups.Count;
                    // Abfrage ob element > 0
                    if (groups > 0)
                    {

                        for (int i = 0; i < groups; i++)
                        {

                            int ElementsLevel0 = softwarePath.TechnologicalObjectGroup.Groups[i].TechnologicalObjects.Count;

                            for (int counter1 = 0; counter1 < ElementsLevel0; counter1++)
                            {

                                //ordner namentlich aufzählen
                                var AxisType = softwarePath.TechnologicalObjectGroup.Groups[i].TechnologicalObjects[counter1].OfSystemLibElement;
                                var AxisName = softwarePath.TechnologicalObjectGroup.Groups[i].TechnologicalObjects[counter1].Name;
                                var Linear_Round = 0;
                                Console.WriteLine(AxisName);


                                if (AxisType.Contains("Axis"))
                                {

                                    if (softwarePath.TechnologicalObjectGroup.Groups[i].TechnologicalObjects[counter1].Parameters.Count > 0)
                                    {

                                        List<ToParameter> Parameters = new List<ToParameter>();
                                        var parameters = softwarePath.TechnologicalObjectGroup.Groups[i].TechnologicalObjects[counter1].Parameters;
                                        foreach (var parameter in parameters)
                                        {
                                            try
                                            {
                                                Parameters.Add(new ToParameter
                                                {
                                                    Name = parameter.GetAttribute("Name").ToString(),
                                                    Value = parameter.GetAttribute("Value").ToString()

                                                });


                                            }
                                            catch (System.Exception)
                                            {


                                            }

                                        }
                                        Axis.Add(new ToConfig
                                        {
                                            Name = AxisName,
                                            Parameters = Parameters

                                        });



                                    }

                                }


                            }

                            var techGroup2 = softwarePath?.TechnologicalObjectGroup.Groups[i];
                            if (techGroup2 != null && HasProperty(techGroup, "Groups"))
                            {
                                if ((softwarePath?.TechnologicalObjectGroup?.Groups[i].Groups?.Count ?? 0) > 0)
                                {
                                    for (int j = 0; j < techGroup2.Groups.Count; j++)
                                    {

                                        ElementsLevel0 = softwarePath.TechnologicalObjectGroup.Groups[i].Groups[j].TechnologicalObjects.Count;

                                        for (int counter1 = 0; counter1 < ElementsLevel0; counter1++)
                                        {

                                            //ordner namentlich aufzählen
                                            var AxisType = softwarePath.TechnologicalObjectGroup.Groups[i].Groups[j].TechnologicalObjects[counter1].OfSystemLibElement;
                                            var AxisName = softwarePath.TechnologicalObjectGroup.Groups[i].Groups[j].TechnologicalObjects[counter1].Name;
                                            var Linear_Round = 0;
                                            Console.WriteLine(AxisName);


                                            if (AxisType.Contains("Axis"))
                                            {

                                                if (softwarePath.TechnologicalObjectGroup.Groups[i].Groups[j].TechnologicalObjects[counter1].Parameters.Count > 0)
                                                {
                                                    List<ToParameter> Parameters = new List<ToParameter>();

                                                    var parameters = softwarePath.TechnologicalObjectGroup.Groups[i].Groups[j].TechnologicalObjects[counter1].Parameters;
                                                    foreach (var parameter in parameters)
                                                    {
                                                        try
                                                        {
                                                            Parameters.Add(new ToParameter
                                                            {
                                                                Name = parameter.GetAttribute("Name").ToString(),
                                                                Value = parameter.GetAttribute("Value").ToString()

                                                            });


                                                        }
                                                        catch (System.Exception)
                                                        {


                                                        }

                                                    }
                                                    Axis.Add(new ToConfig
                                                    {
                                                        Name = AxisName,
                                                        Parameters = Parameters

                                                    });


                                                }

                                            }
                                        }
                                    }
                                }


                            }


                        }



                    }

                }

            }
        }

        if (!Directory.Exists(folderPath + @"\ToConfig"))
        {
            try
            {
                DirectoryInfo di = Directory.CreateDirectory(folderPath + @"\ToConfig");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fehler beim Erstellen des Ordners: {ex.Message}");
            }
        }
        foreach (var Ax in Axis)
        {
            List<string> AxisList = new List<string>();
            AxisList.Add(Ax.Name);
            foreach (var Axparam in Ax.Parameters)
            {

                AxisList.Add(Axparam.Name);
                AxisList.Add(Axparam.Value);

            }
            if (!Directory.Exists(folderPath + @"\ToConfig\" + AxisList[0]))
            {
                try
                {
                    DirectoryInfo di = Directory.CreateDirectory(folderPath + @"\ToConfig\" + AxisList[0]);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Fehler beim Erstellen des Ordners: {ex.Message}");
                }
            }
            SaveTXTItem(folderPath + @"\ToConfig\" + AxisList[0] + @"\" + AxisList[0] + ".txt", AxisList);

        }

        Process.Start(new ProcessStartInfo()
        {
            FileName = "explorer.exe",
            Arguments = folderPath,
            UseShellExecute = true
        });



    }

    public void AxisImportToConfig(int Device, string dllPath, string folderPath)
    {
        List<ToConfig> Axis = new List<ToConfig>();

        if (Directory.Exists(folderPath))
        {


            // Get all .txt files in the folder
            string[] txtFiles = Directory.GetFiles(folderPath, "*.txt", SearchOption.AllDirectories);

            List<string> Axes = new List<string>();
            foreach (string filePath in txtFiles)
            {


                // Read all lines from the current file
                List<ToParameter> Parameters = new List<ToParameter>();

                LoadTXTItem(filePath.ToString(), out Axes);

                for (int i = 1; i < (Axes.Count - 1); i = i + 2)
                {
                    Parameters.Add(new ToParameter
                    {
                        Name = Axes[i],
                        Value = Axes[i + 1]

                    });
                }
                Axis.Add(new ToConfig
                {
                    Name = Axes[0],
                    Parameters = Parameters

                });


                //Console.WriteLine(CanbanizeClass.DecryptString(licenseEntry[0], App.key));
                //Console.WriteLine(CanbanizeClass.DecryptString(licenseEntry[1], App.key));



            }
        }
        else
        {
            Console.WriteLine("Folder does not exist: " + folderPath);
            return;
        }

        Console.WriteLine("Start Export To");


        Assembly engineeringAssembly = LoadEngineeringAssembly(dllPath);

        string softwareContainerTypeName = "Siemens.Engineering.HW.Features.SoftwareContainer";



        dynamic softwareContainer = null;
        for (int i = 0; i < instanceOfTia.Projects[0].Devices[Device].DeviceItems.Count; i++)
        {
            dynamic deviceItem = instanceOfTia.Projects[0].Devices[Device].DeviceItems[i];

            // Verwende GetService<T> mit dem spezifischen Typ
            MethodInfo getServiceMethod = deviceItem.GetType().GetMethod("GetService").MakeGenericMethod(engineeringAssembly.GetType(softwareContainerTypeName));
            softwareContainer = getServiceMethod.Invoke(deviceItem, null);
            if (softwareContainer != null)
            {
                break;
            }
        }

        if (softwareContainer != null)
        {
            Console.WriteLine("softwarecontainer = true");
            dynamic softwarePath = softwareContainer.Software as dynamic;
            string name = softwarePath.Name;


            if ((softwarePath?.TechnologicalObjectGroup?.TechnologicalObjects?.Count ?? 0) > 0)
            {

                int ElementsLevel0 = softwarePath.TechnologicalObjectGroup.TechnologicalObjects.Count;


                // Abfrage ob element > 0
                if (ElementsLevel0 > 0)
                {




                    for (int i = 0; i < ElementsLevel0; i++)
                    {

                        //ordner namentlich aufzählen
                        var AxisType = softwarePath.TechnologicalObjectGroup.TechnologicalObjects[i].OfSystemLibElement;
                        var AxisName = softwarePath.TechnologicalObjectGroup.TechnologicalObjects[i].Name;
                        var Linear_Round = 0;
                        Console.WriteLine(AxisName);
                        int AxisIntCatch = 999;
                        for (int h = 0; h < Axis.Count; h++)
                        {
                            if (Axis[h].Name == AxisName)
                            {
                                AxisIntCatch = h;
                                break;
                            }
                        }
                        if (AxisIntCatch == 999)
                        {
                            Console.WriteLine(AxisName + " not found in config skip");
                            break;
                        }


                        if (AxisType.Contains("Axis"))
                        {


                            if (softwarePath.TechnologicalObjectGroup.TechnologicalObjects[i].Parameters.Count > 0)
                            {


                                var parameters = softwarePath.TechnologicalObjectGroup.TechnologicalObjects[i].Parameters;
                                foreach (var parameter in parameters)
                                {
                                    try
                                    {


                                        for (int h = 0; h < Axis[AxisIntCatch].Parameters.Count; h++)
                                        {
                                            var paramName = parameter.GetAttribute("Name").ToString();

                                            if (paramName == Axis[AxisIntCatch].Parameters[h].Name)
                                            {
                                                Console.WriteLine("Config: " + paramName);
                                                parameter.Value = Axis[AxisIntCatch].Parameters[h].Value;
                                            }
                                        }
                                    }
                                    catch (System.Exception)
                                    {


                                    }

                                }



                            }

                        }


                    }
                }
            }



            var techGroup = softwarePath?.TechnologicalObjectGroup;

            if (techGroup != null && HasProperty(techGroup, "Groups"))
            {

                if ((softwarePath?.TechnologicalObjectGroup?.Groups?.Count ?? 0) > 0)
                {
                    int groups = softwarePath.TechnologicalObjectGroup.Groups.Count;
                    // Abfrage ob element > 0
                    if (groups > 0)
                    {

                        for (int i = 0; i < groups; i++)
                        {

                            int ElementsLevel0 = softwarePath.TechnologicalObjectGroup.Groups[i].TechnologicalObjects.Count;

                            for (int counter1 = 0; counter1 < ElementsLevel0; counter1++)
                            {

                                //ordner namentlich aufzählen
                                var AxisType = softwarePath.TechnologicalObjectGroup.Groups[i].TechnologicalObjects[counter1].OfSystemLibElement;
                                var AxisName = softwarePath.TechnologicalObjectGroup.Groups[i].TechnologicalObjects[counter1].Name;
                                var Linear_Round = 0;
                                Console.WriteLine(AxisName);
                                int AxisIntCatch = 999;
                                for (int h = 0; h < Axis.Count; h++)
                                {
                                    if (Axis[h].Name == AxisName)
                                    {
                                        AxisIntCatch = h;
                                        break;
                                    }
                                }
                                if (AxisIntCatch == 999)
                                {
                                    Console.WriteLine(AxisName + " not found in config skip");
                                    break;
                                }



                                if (AxisType.Contains("Axis"))
                                {

                                    if (softwarePath.TechnologicalObjectGroup.Groups[i].TechnologicalObjects[counter1].Parameters.Count > 0)
                                    {


                                        var parameters = softwarePath.TechnologicalObjectGroup.Groups[i].TechnologicalObjects[counter1].Parameters;
                                        foreach (var parameter in parameters)
                                        {
                                            try
                                            {
                                                for (int h = 0; h < Axis[AxisIntCatch].Parameters.Count; h++)
                                                {
                                                    var paramName = parameter.GetAttribute("Name").ToString();

                                                    if (paramName == Axis[AxisIntCatch].Parameters[h].Name)
                                                    {
                                                        Console.WriteLine("Config: " + paramName);
                                                        parameter.Value = Axis[AxisIntCatch].Parameters[h].Value;
                                                    }
                                                }



                                            }
                                            catch (System.Exception)
                                            {


                                            }

                                        }




                                    }

                                }


                            }

                            var techGroup2 = softwarePath?.TechnologicalObjectGroup.Groups[i];
                            if (techGroup2 != null && HasProperty(techGroup, "Groups"))
                            {
                                if ((softwarePath?.TechnologicalObjectGroup?.Groups[i].Groups?.Count ?? 0) > 0)
                                {
                                    for (int j = 0; j < techGroup2.Groups.Count; j++)
                                    {

                                        ElementsLevel0 = softwarePath.TechnologicalObjectGroup.Groups[i].Groups[j].TechnologicalObjects.Count;

                                        for (int counter1 = 0; counter1 < ElementsLevel0; counter1++)
                                        {

                                            //ordner namentlich aufzählen
                                            var AxisType = softwarePath.TechnologicalObjectGroup.Groups[i].Groups[j].TechnologicalObjects[counter1].OfSystemLibElement;
                                            var AxisName = softwarePath.TechnologicalObjectGroup.Groups[i].Groups[j].TechnologicalObjects[counter1].Name;
                                            var Linear_Round = 0;
                                            Console.WriteLine(AxisName);
                                            int AxisIntCatch = 999;
                                            for (int h = 0; h < Axis.Count; h++)
                                            {
                                                if (Axis[h].Name == AxisName)
                                                {
                                                    AxisIntCatch = h;
                                                    break;
                                                }
                                            }
                                            if (AxisIntCatch == 999)
                                            {
                                                Console.WriteLine(AxisName + " not found in config skip");
                                                break;
                                            }



                                            if (AxisType.Contains("Axis"))
                                            {

                                                if (softwarePath.TechnologicalObjectGroup.Groups[i].Groups[j].TechnologicalObjects[counter1].Parameters.Count > 0)
                                                {


                                                    var parameters = softwarePath.TechnologicalObjectGroup.Groups[i].Groups[j].TechnologicalObjects[counter1].Parameters;
                                                    foreach (var parameter in parameters)
                                                    {
                                                        try
                                                        {
                                                            for (int h = 0; h < Axis[AxisIntCatch].Parameters.Count; h++)
                                                            {
                                                                var paramName = parameter.GetAttribute("Name").ToString();

                                                                if (paramName == Axis[AxisIntCatch].Parameters[h].Name)
                                                                {
                                                                    Console.WriteLine("Config: " + paramName);
                                                                    parameter.Value = Axis[AxisIntCatch].Parameters[h].Value;
                                                                }
                                                            }



                                                        }
                                                        catch (System.Exception)
                                                        {


                                                        }

                                                    }



                                                }

                                            }
                                        }
                                    }
                                }


                            }


                        }



                    }

                }

            }
        }





    }

    public static void SaveTXTItem(string filePath, List<string> ItemList)
    {


        using (StreamWriter sw = new StreamWriter(filePath, false, Encoding.UTF8))
        {
            for (int i = 0; i < ItemList.Count; i++)
            {
                sw.WriteLine(ItemList[i]);
            }
        }

    }

    public static void LoadTXTItem(string filePath, out List<string> ItemList)
    {
        ItemList = new List<string>();
        #region History


        string content = "";
        if (File.Exists(filePath))
        {

            // Den gesamten Inhalt der Datei lesen
            content = File.ReadAllText(filePath);
            List<string> CompleteList = new List<string>();
            using (StreamReader sr = new StreamReader(filePath, Encoding.UTF8))
            {
                while (sr.Peek() != -1)
                    CompleteList.Add(sr.ReadLine());
            }

            for (int i = 0; i < CompleteList.Count; i++)
            {
                ItemList.Add(CompleteList[i]);

            }

        }
        #endregion
    }


    public List<string> AxisSearcheAdress(string dllPath)
    {
        List<string> AxisAdresses = new List<string>();

        Assembly engineeringAssembly = LoadEngineeringAssembly(dllPath);

        var Axisname = "";
        var ALMAmount = 0;

        if (0 != instanceOfTia.Projects[0].UngroupedDevicesGroup.Devices.Count)
        {
            for (int i = 0; i < instanceOfTia.Projects[0].UngroupedDevicesGroup.Devices.Count; i++)
            {
                if (null != instanceOfTia.Projects[0].UngroupedDevicesGroup.Devices[i].TypeIdentifier)
                {
                    if (instanceOfTia.Projects[0].UngroupedDevicesGroup.Devices[i].TypeIdentifier.Contains("SINAMICS"))
                    {
                        if (0 != instanceOfTia.Projects[0].UngroupedDevicesGroup.Devices[i].DeviceItems.Count)
                        {
                            for (int j = 0; j < instanceOfTia.Projects[0].UngroupedDevicesGroup.Devices[i].DeviceItems.Count; j++)
                            {
                                if (null != instanceOfTia.Projects[0].UngroupedDevicesGroup.Devices[i].DeviceItems[j].Name)
                                {
                                    if (instanceOfTia.Projects[0].UngroupedDevicesGroup.Devices[i].DeviceItems[j].GetAttribute("TypeName").ToString() == "PROFIsafe telegr 902" &&
                                        instanceOfTia.Projects[0].UngroupedDevicesGroup.Devices[i].DeviceItems[j].GetAttribute("IsPlugged").ToString() == "True"
                                        )

                                    {
                                        Axisname = instanceOfTia.Projects[0].UngroupedDevicesGroup.Devices[i].DeviceItems[j].GetAttribute("Name").ToString();


                                        Axisname = Axisname.Replace("Profisafe", "");

                                        Axisname = Axisname.Replace("Axis", "");

                                        Axisname = Axisname.Replace("_", "");

                                        AxisAdresses.Add(Axisname);
                                        if ((instanceOfTia.Projects[0].UngroupedDevicesGroup.Devices[i].DeviceItems[j].Addresses.Count > 0))
                                        {
                                            AxisAdresses.Add(Convert.ToString(instanceOfTia.Projects[0].UngroupedDevicesGroup.Devices[i].DeviceItems[j].Addresses[0].StartAddress));
                                        }

                                    }

                                    string Name = instanceOfTia.Projects[0].UngroupedDevicesGroup.Devices[i].DeviceItems[j].GetAttribute("Name").ToString();
                                    string Type = instanceOfTia.Projects[0].UngroupedDevicesGroup.Devices[i].DeviceItems[j].GetAttribute("TypeName").ToString();
                                    if (Name.Contains("Telegramm") && Type.Contains("370") &&
                                        instanceOfTia.Projects[0].UngroupedDevicesGroup.Devices[i].DeviceItems[j].GetAttribute("IsPlugged").ToString() == "True"
                                        )
                                    {

                                        ALMAmount = ALMAmount + 1;

                                        Axisname = "ALM" + ALMAmount;

                                        AxisAdresses.Add(Axisname);

                                        AxisAdresses.Add(Convert.ToString(instanceOfTia.Projects[0].UngroupedDevicesGroup.Devices[i].DeviceItems[j].Addresses[0].StartAddress));

                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        return AxisAdresses;
    }


    private bool HasProperty(object obj, string propertyName)
    {
        if (obj == null)
            return false;

        return obj.GetType().GetProperty(propertyName) != null;
    }



}
