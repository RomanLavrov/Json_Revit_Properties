#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Json_From_Revit.Data_Model;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
#endregion

namespace Json_From_Revit
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        // Document linkedDocument;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Document linkedDocument = null;

            foreach (Document item in uiapp.Application.Documents)
            {
                linkedDocument = item;
            }


            ProjectData projectData = new ProjectData();
            projectData.VersionName = uiapp.Application.VersionName;
            projectData.Architecture_Document = doc.ProjectInformation.Author;
            projectData.Document_Information = GetDocumentInformation(doc);

            List<ElementData> electroDataList = new List<ElementData>();

            var electroList = GetElectricalElements(linkedDocument);
            foreach (Element device in electroList)
            {
                ElementData electroData = new ElementData();

                electroData = GetElectroData(linkedDocument, device);

                var roomList = GetRoomElement(doc);
                var wallList = GetWallElement(doc);

                foreach (Element room in roomList)
                {
                    if (GetHost(doc, room, device))
                    {
                        electroData.Raum = GetRoomData(doc, room);
                    }
                }

                foreach (var wall in wallList)
                {
                    if (GetHost(doc, wall, device))
                    {
                        electroData.Wand = GetWallData(doc, wall);
                    }
                }

                electroDataList.Add(electroData);
            }

            projectData.elements = electroDataList;
            var JSONdata = string.Empty;
            string path = @"D:\Test\Schulweg Oberwil.json";

            JSONdata = JsonConvert.SerializeObject(projectData);


            // TaskDialog.Show("JSON", JSONdata);
            System.IO.File.WriteAllText(path, JSONdata);
            Process.Start(path);

            return Result.Succeeded;
        }

        bool GetHost(Document doc, Element host, Element device)
        {
            BoundingBoxXYZ deviceBox = device.get_BoundingBox(doc.ActiveView);
            BoundingBoxXYZ hostBox = host.get_BoundingBox(doc.ActiveView);

            if (hostBox == null)
            {
                return false;
            }

            if (deviceBox.Min.X > hostBox.Min.X && deviceBox.Min.X < hostBox.Max.X ||
                    deviceBox.Max.X > hostBox.Min.X && deviceBox.Max.X < hostBox.Max.X)
            {
                if (deviceBox.Min.Y > hostBox.Min.Y && deviceBox.Min.Y < hostBox.Max.Y ||
                        deviceBox.Max.Y > hostBox.Min.Y && deviceBox.Max.Y < hostBox.Max.Y)
                {
                    if (deviceBox.Min.Z > hostBox.Min.Z && deviceBox.Min.Z < hostBox.Max.Z ||
                            deviceBox.Max.Z > hostBox.Min.Z && deviceBox.Max.Z < hostBox.Max.Z)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        List<Element> GetElectricalElements(Document linkedDocument)
        {
            List<Element> Equipment = new List<Element>();

            if (linkedDocument != null)
            {
                var equipment =
                new FilteredElementCollector(linkedDocument).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_ElectricalEquipment);
                var fixtures = new FilteredElementCollector(linkedDocument).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_ElectricalFixtures);
                var lighting = new FilteredElementCollector(linkedDocument).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_LightingDevices);
                var switsches = new FilteredElementCollector(linkedDocument).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_LightingFixtures);
                var data = new FilteredElementCollector(linkedDocument).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_DataDevices);
                var fire = new FilteredElementCollector(linkedDocument).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_FireAlarmDevices);
                var emergency = new FilteredElementCollector(linkedDocument).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_NurseCallDevices);
                var telephone = new FilteredElementCollector(linkedDocument).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_TelephoneDevices);

                Equipment = equipment.ToList();

                Equipment.AddRange(fixtures.ToList());
                Equipment.AddRange(lighting.ToList());
                Equipment.AddRange(switsches.ToList());
                Equipment.AddRange(data.ToList());
                Equipment.AddRange(fire.ToList());
                Equipment.AddRange(emergency.ToList());
                Equipment.AddRange(telephone.ToList());
            }

            return Equipment;
        }

        ElementData GetElectroData(Document linkedDocument, Element electroItem)
        {
            string temp = String.Empty;
            ElementData elementData = new ElementData();
            elementData.objectId = electroItem.Id.ToString();
            elementData.name = electroItem.Name;
            elementData.externalId = electroItem.UniqueId;

            Parameters identityData = new Parameters();

            identityData.Bauteil_ID = GetParameter(electroItem, "Bauteil-ID");
            identityData.Beschreibung = GetParameter(electroItem, "Description");
            identityData.Installations_Medium = GetParameter(electroItem, "Installations Medium");
            identityData.Installationsort = GetParameter(electroItem, "Installationsort");
            identityData.Installationsart = GetParameter(electroItem, "Installationsart");
            identityData.Fabrikat = GetParameter(electroItem, "Fabrikat");
            identityData.Produkt = GetParameter(electroItem, "Produkt");
            identityData.Produkte_Nr = GetParameter(electroItem, "Produkte-Nr.");
            identityData.E_Nummer = GetParameter(electroItem, "E-Nummer");

            temp +=
                $"Bauteil-ID: {GetParameter(electroItem, "Bauteil-ID")}, " +
                $"Beschreibung: {GetParameter(electroItem, "Description")}, " +
                $"Installations Medium: {GetParameter(electroItem, "Installations Medium")}, " +
                $"Installationsort: {GetParameter(electroItem, "Installationsort")}, " +
                $"Installationsart: {GetParameter(electroItem, "Installationsart")}, " +
                $"Fabrikat: {GetParameter(electroItem, "Fabrikat")}, " +
                $"Produkt: {GetParameter(electroItem, "Produkt")}, " +
                $"Produkte-Nr.: {GetParameter(electroItem, "Produkte-Nr.")}, " +
                $"E-Nummer: {GetParameter(electroItem, "E-Nummer")}" + Environment.NewLine;

            //TaskDialog.Show("electro", temp);           
            elementData.Identity_Data = identityData;
            return elementData;
        }

        List<Element> GetRoomElement(Document doc)
        {
            List<Element> roomList = new List<Element>();

            FilteredElementCollector roomCollector = new FilteredElementCollector(doc);
            roomCollector.OfClass(typeof(SpatialElement));

            foreach (Element room in roomCollector)
            {
                roomList.Add(room);
            }

            return roomList;
        }

        RoomData GetRoomData(Document doc, Element room)
        {
            string roomDataText = String.Empty;
            string[] rData = room.Name.Split(' ');
            RoomData roomData = new RoomData();
            roomData.RaumEbene = doc.GetElement(room.LevelId).Name;
            roomData.RaumName = rData[0];
            roomData.RaumNummer = rData[1];

            roomDataText += $"Room Name: {roomData.RaumName}, Number: {roomData.RaumNummer}, Level: {roomData.RaumEbene}" + Environment.NewLine;
            //TaskDialog.Show("Room", roomDataText);

            return roomData;
        }

        List<Element> GetWallElement(Document doc)
        {
            List<Element> wallList = new List<Element>();

            FilteredElementCollector wallCollector = new FilteredElementCollector(doc);
            wallCollector.OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType();

            foreach (Element wall in wallCollector)
            {
                wallList.Add(wall);
            }

            return wallList;
        }

        WallData GetWallData(Document doc, Element wall)
        {
            string wallsData = string.Empty;

            string GUID = string.Empty;
            foreach (Parameter parameter in wall.Parameters)
            {
                if (parameter.Definition.Name.Equals("IfcGUID"))
                {
                    GUID = parameter.AsString();
                }
            }
            var wallMaterial = string.Empty;
            try
            {
                var wallSymbol = doc.GetElement(wall.GetTypeId());
                foreach (Parameter param in wallSymbol.Parameters)
                {
                    if (param.Definition.Name.Equals("Material"))
                    {
                        wallMaterial = param.AsString();
                    }
                }
            }
            catch (Exception)
            { }

            WallData wallData = new WallData();
            wallData.Wand_ID = GUID;
            wallData.Material = wallMaterial;
            wallData.Wand_Name = wall.Name;

            wallsData += $"Wall Name:{wall.Name}, GUID:{GUID}, Material:{wallMaterial}" + Environment.NewLine;

            return wallData;
        }

        string GetParameter(Element element, string ParameterName)
        {
            FamilyInstance instance = element as FamilyInstance;
            FamilySymbol symbol = instance.Symbol;

            var value = string.Empty;
            foreach (Parameter parameter in symbol.Parameters)
            {
                if (parameter.Definition.Name.Equals(ParameterName))
                {
                    value = parameter.AsString();
                    if (string.IsNullOrEmpty(value))
                    {
                        value = parameter.AsValueString();
                    }
                }
            }
            return value;
        }

        DocumentInformation GetDocumentInformation (Document doc)
        {
            DocumentInformation docInfo = new DocumentInformation();
            
            docInfo.Name = doc.ProjectInformation.Document.PathName;
            docInfo.Number = doc.ProjectInformation.Number;
            docInfo.Status = doc.ProjectInformation.Status;
            docInfo.Address = doc.ProjectInformation.Address;
            docInfo.ClientName = doc.ProjectInformation.ClientName;

            return docInfo;
        }
    }
}
