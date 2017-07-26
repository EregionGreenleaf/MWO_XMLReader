using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace MWO_XMLReader
{
    public static class Worker
    {

        public static bool PrintExcel(List<MechStats> list)
        {
            ExcelOps.WriteQuirks(list);
            return true;
        }

        public static List<MechStats> SortMechs (List<MechStats> list)
        {
            // SORTING SECTION ###
            List<MechStats> light = new List<MechStats>();
            List<MechStats> medium = new List<MechStats>();
            List<MechStats> heavy = new List<MechStats>();
            List<MechStats> assault = new List<MechStats>();
            list.ForEach(x =>
            {
                switch (x.Class)
                {
                    case 1:
                        light.Add(x);
                        break;
                    case 2:
                        medium.Add(x);
                        break;
                    case 3:
                        heavy.Add(x);
                        break;
                    case 4:
                        assault.Add(x);
                        break;
                    default:
                        break;

                }
            });
            light = light.OrderBy(x => x.Variant).ToList().OrderBy(y => y.Chassis).ToList();
            medium = medium.OrderBy(x => x.Variant).ToList().OrderBy(y => y.Chassis).ToList();
            heavy = heavy.OrderBy(x => x.Variant).ToList().OrderBy(y => y.Chassis).ToList();
            assault = assault.OrderBy(x => x.Variant).ToList().OrderBy(y => y.Chassis).ToList();
            list = new List<MechStats>();
            list.AddRange(light);
            list.AddRange(medium);
            list.AddRange(heavy);
            list.AddRange(assault);
            // ############### ###
            list.OrderBy(x => x.Class).OrderBy(y => y.Chassis).OrderBy(z => z.Variant);
            return list;
        }

        /// <summary>
        /// Loads a list of 
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static List<MechStats> LoadQuirks(string path)
        {
            List<MechStats> mechList = new List<MechStats>();
            List<string> files = Directory.GetFiles(path).ToList();
            foreach (string file in files)
            {
                FileInfo fileI = new FileInfo(file);
                if(fileI.Extension != ".mdf")
                    continue;
                // declaration of new 'Mech
                MechStats mech = new MechStats();
                mech.Variant = "";

                // new xdoc instance 
                XmlDocument xDoc = new XmlDocument();
                //load up the xml from the location 
                xDoc.Load(file);
                foreach (XmlNode node in xDoc.DocumentElement.ChildNodes)
                {
                    int x;
                    if(node.Name == "Mech")
                    {
                        XmlAttributeCollection attributes = node.Attributes;
                        foreach (XmlAttribute att in attributes)
                        {
                            int temp = 0;
                            switch (att.Name)
                            {
                                case "Variant":
                                    mech.Variant = att.Value;
                                    break;
                                case "MaxTons":
                                    mech.MaxTons = int.TryParse(att.Value, out temp) ? temp: 0;
                                    mech.Class = mech.MaxTons <= 35 ? 1 : mech.MaxTons <= 55 ? 2 : mech.MaxTons <= 75 ? 3 : 4;
                                    break;
                                case "MaxJumpJets":
                                    mech.MaxJumpJets = int.TryParse(att.Value, out temp) ? temp : 0;
                                    break;
                                case "CanEquipECM":
                                    mech.CanEquipECM = int.TryParse(att.Value, out temp) ? (temp != 0 ? true : false) : false;
                                    break;
                                case "MinEngineRating":
                                    mech.MinEngineRating = int.TryParse(att.Value, out temp) ? temp : 0;
                                    break;
                                case "MaxEngineRating":
                                    mech.MaxEngineRating = int.TryParse(att.Value, out temp) ? temp : 0;
                                    break;
                                default:
                                    break;
                            }
                        }
                    }

                    if (node.Name == "Cockpit")
                    {
                        XmlAttributeCollection attributes = node.Attributes;
                        foreach (XmlAttribute att in attributes)
                        {
                            if (att.Name == "startup")
                            {
                                mech.Chassis = att.Value.Replace("pilot_startup_", string.Empty).Trim();
                            }
                        }
                    }

                    if (node.Name == "QuirkList")
                    {
                        // we search each note in search of what we need
                        foreach (XmlNode locNode in node)
                        {
                            // if in quirk section
                            if (locNode.Name == "Quirk")
                            {
                                string quirk = string.Empty;
                                double value = 0.0;
                                XmlAttributeCollection attributes = locNode.Attributes;
                                foreach (XmlAttribute att in attributes)
                                {
                                    // we save the quirk values
                                    if (att.Name == "name")
                                        quirk = att.Value;
                                    if (att.Name == "value")
                                    {
                                        string val = att.Value.Contains('.') ? att.Value.Replace('.', ',') : att.Value;
                                        double temp;
                                        value = double.TryParse(val, out temp) ? temp : 0.0;
                                    }
                                }
                                // if quirk is not null, we add it to the 'Mech
                                if (!string.IsNullOrWhiteSpace(quirk) && value != 0.0)
                                {
                                    Quirk q = new Quirk() { Name = quirk, State = true, Value = value };
                                    mech.QuirkList.Add(q);
                                }

                            }

                        }
                    }
                }
                if (!string.IsNullOrWhiteSpace(mech.Variant))
                    mechList.Add(mech);
            }

            // DEBUG ###########################
            //List<Quirk> quirkList = new List<Quirk>();
            //mechList.ForEach(x => quirkList.AddRange(x.QuirkList));
            //List<string> quirkListString = quirkList.Select(x=>x.Name).Distinct().ToList();
            //quirkListString.Sort();
            //FileInfo info = new FileInfo("C:\\TEST\\quirkList.txt");
            //if (info.Exists)
            //    info.Delete();
            //foreach (string quirk in quirkListString)
            //{
            //    Logger.PrintF(info.FullName, quirk, false);
            //}
            // #################################

            return mechList;
        }


    }
}
