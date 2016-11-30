using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace Pavlik {
    class Program {
        static void Main(string[] args) {
            string filename = @"2016-11-28_MEAMD NETMOD Configs.xlsx";
            var path = string.Format("C:\\Users\\John\\Desktop\\{0}", filename);
            var connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0; data source={0}; Extended Properties=Excel 12.0;", path);
            var adapter = new OleDbDataAdapter("SELECT * FROM [EAS$]", connectionString);
            var ds = new DataSet();
            adapter.Fill(ds, "anyNameHere");
            var data = ds.Tables["anyNameHere"].AsEnumerable();

            string singlePattern = @"^interface Gi(\d)/(\d)?/(\d[1-2])";
            string rangePattern = @"^interface? range (.+)";
            string accessVlanPattern = @"switchport access vlan (\d+)";
            string hostnamePattern = @"hostname (.+)";
            Regex singleRegex = new Regex(singlePattern, RegexOptions.IgnoreCase);
            Regex rangeRegex = new Regex(rangePattern, RegexOptions.IgnoreCase);
            Regex accessVlanRegex = new Regex(accessVlanPattern, RegexOptions.IgnoreCase);
            Regex hostnameRegex = new Regex(hostnamePattern, RegexOptions.IgnoreCase);
            var columnCount = 219;
            StringBuilder outputTxt = new StringBuilder();
            for(int columnIndex = 0; columnIndex < columnCount; ++columnIndex) {
                var interfaceMap = new Dictionary<string, List<int>>();
                var currentInterfaces = new List<string>();
                string hostname = "UNKNOWN HOSTNAME";
                foreach(var row in data) {
                    string value = row.Field<string>(columnIndex);
                    if(row.IsNull(columnIndex)) continue;
                    var match = hostnameRegex.Match(value);
                    if(match.Success) {
                        hostname = match.Groups[1].ToString();
                    }
                    match = accessVlanRegex.Match(value);
                    if(match.Success) {
                        string vlanString = match.Groups[1].ToString();
                        int vlan = int.Parse(vlanString);
                        if(vlan == 1000) continue;
                        foreach(string iface in currentInterfaces) {
                            interfaceMap[iface].Add(vlan);
                        }
                    }
                    match = singleRegex.Match(value);
                    if(match.Success) {
                        string sw = match.Groups[1].ToString();
                        string mod = match.Groups[2].ToString();
                        string port = match.Groups[3].ToString();
                        string key = "Gi/" + sw + "/" + mod + "/" + port;
                        currentInterfaces.Clear();
                        currentInterfaces.Add(key);
                        if(!interfaceMap.ContainsKey(key)) {
                            interfaceMap.Add(key, new List<int>());
                        }
                    }
                    match = rangeRegex.Match(value);
                    if(match.Success) {
                        currentInterfaces.Clear();
                        string text = match.Groups[1].ToString();
                        var groups = text.Split(',');
                        foreach(string s in groups) {
                            if(s.Contains("-")) {
                                string[] splitArrays = s.Split('/');
                                string range = splitArrays[2];
                                string[] rangeSplit = range.Split('-');
                                int startRange = int.Parse(rangeSplit[0]);
                                int endRange = int.Parse(rangeSplit[1]);
                                for(int i = startRange; i <= endRange; ++i) {
                                    string key = splitArrays[0].Trim() + "/" + splitArrays[1] + "/" + i;
                                    currentInterfaces.Add(key);
                                    if(!interfaceMap.ContainsKey(key)) {
                                        interfaceMap.Add(key, new List<int>());
                                    }
                                }
                            } else {
                                string key = s.Trim();
                                currentInterfaces.Add(key);
                                if(!interfaceMap.ContainsKey(key)) {
                                    interfaceMap.Add(key, new List<int>());
                                }
                            }
                        }
                    }
                }
                outputTxt.Append(hostname);
                outputTxt.Append("\n");
                foreach(var s in interfaceMap) {
                    if(s.Value.Count > 1) {
                        outputTxt.Append(s.Key);
                        outputTxt.Append(": ");
                        bool first = true;
                        foreach(var t in s.Value) {
                            if(first) {
                                first = false;
                            } else {
                                outputTxt.Append(", ");
                            }
                            outputTxt.Append(t.ToString());
                        }
                        outputTxt.Append("\n");
                    }
                }
            }
            File.WriteAllText(@"C:\Users\John\Desktop\Output.txt", outputTxt.ToString());
        }
    }
}
