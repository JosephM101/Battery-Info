using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Windows.Devices.Enumeration;
using Windows.Devices.Power;
using Windows.System.Power;

namespace Battery_Info
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            refreshToolStripMenuItem.Click += RefreshToolStripMenuItem_Click;
            saveFileToolStripMenuItem.Click += SaveFileToolStripMenuItem_Click;
            exitToolStripMenuItem.Click += ExitToolStripMenuItem_Click;
            toolStripMenuItem2.Click += ToolStripMenuItem2_Click;
            RefreshTreeView();
        }

        private void ToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            new About().ShowDialog();
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        async void RefreshTreeView()
        {
            int index = 0;
            treeView.BeginUpdate();
            var deviceInfo = await DeviceInformation.FindAllAsync(Battery.GetDeviceSelector());
            treeView.Nodes.Clear();
            foreach (DeviceInformation device in deviceInfo)
            {
                try
                {
                    var battery = await Battery.FromIdAsync(device.Id);
                    var report = battery.GetReport();
                    int batteryCharge = (int)Math.Round(((decimal)report.RemainingCapacityInMilliwattHours / (decimal)report.FullChargeCapacityInMilliwattHours) * 100, 0);
                    int healthPercentage = (int)Math.Round((decimal)report.FullChargeCapacityInMilliwattHours / (decimal)report.DesignCapacityInMilliwattHours * 100, 0);
                    treeView.Nodes.Add("Battery" + index.ToString(), "Battery " + index.ToString() + " (" + batteryCharge + "%)");
                    treeView.Nodes[index].Nodes.Add("DeviceID", "Device ID: " + device.Id);
                    treeView.Nodes[index].Nodes.Add("BatteryStatus", "Status: " + StatusToString(report.Status));
                    if (report.Status == BatteryStatus.Charging)
                    {
                        treeView.Nodes[index].Nodes.Add("ChrgRate", "Charge rate: " + report.ChargeRateInMilliwatts.ToString() + " mW");
                    }
                    else if (report.Status == BatteryStatus.Discharging)
                    {
                        treeView.Nodes[index].Nodes.Add("ChrgRate", "Discharge rate: " + Math.Abs((decimal)report.ChargeRateInMilliwatts).ToString() + " mW");
                    }
                    treeView.Nodes[index].Nodes.Add("DesignCapacity", "Design capacity: " + report.DesignCapacityInMilliwattHours.ToString() + " mWh");
                    treeView.Nodes[index].Nodes.Add("ChargeCapacity", "Full charge capacity: " + report.FullChargeCapacityInMilliwattHours.ToString() + " mWh");
                    string batteryHealthSuffix = "";
                    if (healthPercentage > 80)
                    {
                        batteryHealthSuffix = "Healthy";
                    }
                    if (healthPercentage <= 80)
                    {
                        batteryHealthSuffix = "Battery may not perform as well.";
                    }
                    if (healthPercentage <= 60)
                    {
                        batteryHealthSuffix = "Battery health is not good. Consider replacing soon.";
                    }
                    if (healthPercentage <= 55)
                    {
                        batteryHealthSuffix = "Battery is not healthy. Replace it soon.";
                    }
                    if (healthPercentage <= 45)
                    {
                        batteryHealthSuffix = "It's not recommended that you continue using this battery.";
                    }
                    if (healthPercentage <= 30)
                    {
                        batteryHealthSuffix = "Battery is bad. Replace now.";
                    }
                    treeView.Nodes[index].Nodes.Add("RemainingCapacity", "Remaining capacity: " + report.RemainingCapacityInMilliwattHours.ToString() + " mWh (" + batteryCharge + "%)");
                    treeView.Nodes[index].Nodes.Add("BatHealth", "Battery health: " + healthPercentage.ToString() + "% (" + batteryHealthSuffix + ")");
                    index++;
                }
                catch { /* Add error handling, as applicable */ }
            }
            treeView.ExpandAll();
            treeView.EndUpdate();
        }

        private void SaveNodes(TreeNodeCollection nodesCollection, XmlWriter textWriter)
        {
            foreach (var node in nodesCollection.OfType<TreeNode>().Where(x => x.Nodes.Count == 0))
                textWriter.WriteAttributeString(node.Name, node.Text);

            foreach (var node in nodesCollection.OfType<TreeNode>().Where(x => x.Nodes.Count > 0))
            {
                textWriter.WriteStartElement(node.Name);
                SaveNodes(node.Nodes, textWriter);
                textWriter.WriteEndElement();
            }
        }

        private void SaveNodes(TreeNodeCollection nodesCollection, StreamWriter htmlDocumentWriter)
        {
            foreach (var node in nodesCollection.OfType<TreeNode>().Where(x => x.Nodes.Count == 0))
            {
                // Write the node value as a paragraph
                htmlDocumentWriter.WriteLine("<p style=\"margin: 0px; padding: 0px;\">" + node.Text + "</p>");
            }

            // Write info for each battery
            foreach (var node in nodesCollection.OfType<TreeNode>().Where(x => x.Nodes.Count > 0))
            {
                // Battery name
                htmlDocumentWriter.WriteLine("<h2 style=\"margin: 4px; padding: 0px;\">" + node.Text + "</h2>");
                SaveNodes(node.Nodes, htmlDocumentWriter);
            }
        }

        string StatusToString(BatteryStatus batteryStatus)
        {
            switch (batteryStatus)
            {
                case BatteryStatus.Charging:
                    return "Charging";

                case BatteryStatus.Discharging:
                    return "Discharging";

                case BatteryStatus.Idle:
                    return "Idle";

                case BatteryStatus.NotPresent:
                    return "Not Present";

                default:
                    return "Unknown";
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (autoRefreshToolStripMenuItem.Checked == true)
            {
                RefreshTreeView();
            }
        }

        private void RefreshToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            RefreshTreeView();
        }

        private void SaveFileToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            SaveToFile();
        }

        void SaveToFile()
        {
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string extension = Path.GetExtension(saveFileDialog.FileName).Substring(1);
                switch (extension)
                {
                    case "xml":
                        XmlWriterSettings settings = new XmlWriterSettings();
                        settings.WriteEndDocumentOnClose = true;
                        settings.Indent = true;
                        settings.NewLineOnAttributes = true;
                        XmlWriter xmlWriter = XmlWriter.Create(saveFileDialog.FileName, settings);
                        SaveNodes(treeView.Nodes, xmlWriter);
                        xmlWriter.Close();
                        break;

                    case "html":
                        StreamWriter htmlDocument = new StreamWriter(saveFileDialog.FileName);
                        htmlDocument.WriteLine("<!DOCTYPE html>");
                        htmlDocument.WriteLine("<html>");
                        htmlDocument.WriteLine("<head>");
                        htmlDocument.WriteLine("<title>Battery Summary</title>");
                        htmlDocument.WriteLine("</head>");
                        htmlDocument.WriteLine("<body>");
                        // Write current date and time
                        htmlDocument.WriteLine("<p>Generated on " + DateTime.Now.ToString() + "</p>");
                        htmlDocument.WriteLine("<hr>");
                        htmlDocument.WriteLine("<h1>Battery Summary</h1>");
                        htmlDocument.WriteLine("<hr>");
                        SaveNodes(treeView.Nodes, htmlDocument);
                        htmlDocument.WriteLine("</body>");
                        htmlDocument.WriteLine("</html>");
                        htmlDocument.Close();
                        break;

                    default:
                        MessageBox.Show("Unsupported file type.");
                        break;
                }
            }
        }
    }
}