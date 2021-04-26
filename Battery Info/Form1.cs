using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Windows.Devices.Enumeration;
using Windows.Devices.Power;
using Windows.System.Power;

namespace Battery_Info
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        async void RefreshTreeView()
        {
            int index = 0;
            treeView1.BeginUpdate();
            var deviceInfo = await DeviceInformation.FindAllAsync(Battery.GetDeviceSelector());
            treeView1.Nodes.Clear();
            foreach (DeviceInformation device in deviceInfo)
            {
                try
                {
                    var battery = await Battery.FromIdAsync(device.Id);
                    var report = battery.GetReport();
                    int batteryCharge = (int)Math.Round(((decimal)report.RemainingCapacityInMilliwattHours / (decimal)report.FullChargeCapacityInMilliwattHours) * 100, 0);
                    int healthPercentage = (int)Math.Round((decimal)report.FullChargeCapacityInMilliwattHours / (decimal)report.DesignCapacityInMilliwattHours * 100, 0);
                    treeView1.Nodes.Add("Battery " + index.ToString() + " (" + batteryCharge + "%)");
                    treeView1.Nodes[index].Nodes.Add("Device ID: " + device.Id);
                    treeView1.Nodes[index].Nodes.Add("Status: " + StatusToString(report.Status));
                    if (report.Status == BatteryStatus.Charging)
                    {
                        treeView1.Nodes[index].Nodes.Add("Charge rate: " + report.ChargeRateInMilliwatts.ToString() + " mW");
                    }
                    else if (report.Status == BatteryStatus.Discharging)
                    {
                        treeView1.Nodes[index].Nodes.Add("Discharge rate: " + Math.Abs((decimal)report.ChargeRateInMilliwatts).ToString() + " mW");
                    }
                    treeView1.Nodes[index].Nodes.Add("Design capacity: " + report.DesignCapacityInMilliwattHours.ToString() + " mWh");
                    treeView1.Nodes[index].Nodes.Add("Full charge capacity: " + report.FullChargeCapacityInMilliwattHours.ToString() + " mWh");
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
                        batteryHealthSuffix = "Battery health is nearing dangerous levels. Replace battery.";
                    }
                    if (healthPercentage <= 30)
                    {
                        batteryHealthSuffix = "Battery is bad. Replace now.";
                    }
                    treeView1.Nodes[index].Nodes.Add("Remaining capacity: " + report.RemainingCapacityInMilliwattHours.ToString() + " mWh (" + batteryCharge + "%)");
                    treeView1.Nodes[index].Nodes.Add("Battery health: " + healthPercentage.ToString() + "% (" + batteryHealthSuffix + ")");
                    index++;
                }
                catch { /* Add error handling, as applicable */ }
            }
            treeView1.ExpandAll();
            treeView1.EndUpdate();
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

        private void RefreshToolStripMenuItem1_Click(object sender, System.EventArgs e)
        {
            RefreshTreeView();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if(autoRefreshToolStripMenuItem.Checked == true)
            {
                RefreshTreeView();
            }
        }
    }
}