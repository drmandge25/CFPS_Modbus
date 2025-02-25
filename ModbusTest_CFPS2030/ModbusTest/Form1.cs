using EasyModbus;
using S7.Net;
using System;
using System.Data.SqlClient;
using System.Drawing;
using System.IO.Ports;
//using System.Threading.Tasks;
using System.Windows.Forms;

namespace ModbusTest
{
    public partial class Read : Form
    {



        ModbusClient mc1;

        Plc plc1;

        bool SetTimer = false;
        bool plc_connected = false;
        int FilamentMinutes;
        //int i = 0;
        //bool Ramp_Complete = false;
        //int ratio = 1; // 1 watt per second
        string actualFilamentPower;
        int RBFilamentPower;
        double RBGyrotronCurrent;
        double RBFilamentCurrent;
        string connectionString;
        SqlConnection con;
        bool sqlConnectionEstablished=false;




        void accessModbus()
        {
            var values = mc1.ReadHoldingRegisters(0, 9);
            // Task.Delay(200).Wait();


            //txbReadFC.Text = Convert.ToString((float)values[2] / 10);
            //txbReadGC.Text = Convert.ToString((float)values[3] / 10);
            //txbReadPower.Text = Convert.ToString(values[5]);
            //txbRBIbs.Text = Convert.ToString(values[7]);
            //txbRBDPf.Text = Convert.ToString(values[6]);

            actualFilamentPower = Convert.ToString(values[5]); 
            lblReadPower.Text = actualFilamentPower+" W";
            lblReadFC.Text = Convert.ToString((float)values[2] / 10)+" A";
            lblReadGC.Text = Convert.ToString((float)values[3] / 10)+" A";
            lblRBIbs.Text = Convert.ToString(values[7])+" A";

            RBFilamentPower = values[5];
            RBGyrotronCurrent = (float)values[3] / 10;
            RBFilamentCurrent = (float)values[2] / 10;


            System.Collections.BitArray Status_Register = new System.Collections.BitArray(new int[] { values[0] });
            System.Collections.BitArray Error_Register = new System.Collections.BitArray(new int[] { values[1] });


            if (Status_Register[4])
            {
                label7.BackColor = Color.Green;
                label7.Text = "CFPS ON";
            }
            else
            {
                label7.BackColor = Color.Red;
                label7.Text = "CFPS  OFF";
            }

            if (Status_Register[5])
            {
                label8.BackColor = Color.Green;
                label8.Text = "Forcing Mode Enabled";
            }
            else
            {
                label8.BackColor = Color.Red;
                label8.Text = "Forcing Mode Disabled";
            }

            if (Status_Register[6])
            {
                label15.BackColor = Color.Green;
                label15.Text = "Remote Mode Enabled";
            }
            else
            {
                label15.BackColor = Color.Red;
                label15.Text = "Remote Mode Disabled";
            }
            if (Status_Register[7])
            {
                label16.BackColor = Color.Green;
                label16.Text = "Front Panel Lock Mode Enabled";
            }
            else
            {
                label16.BackColor = Color.Red;
                label16.Text = "Front Panel Lock Mode Disabled";
            }

            if (Status_Register[11])
            {
                label17.BackColor = Color.Green;
                label17.Text = "Stabilzation Mode Enabled";
            }
            else
            {
                label17.BackColor = Color.Red;
                label17.Text = "Stabilzation Mode Disabled";
            }

            if (Error_Register[4])
            {
                label18.BackColor = Color.Red;
                label18.Text = "Uf Max Protection Triggered";
            }
            else
            {
                label18.BackColor = Color.Green;
                label18.Text = "Uf Max Protection";
            }

            if (Error_Register[5])
            {
                label19.BackColor = Color.Red;
                label19.Text = "If Max Protection Triggered";
            }
            else
            {
                label19.BackColor = Color.Green;
                label19.Text = "If Max Protection";
            }


            if (Error_Register[6])
            {
                label20.BackColor = Color.Red;
                label20.Text = "Internal Protection Triggered";
            }
            else
            {
                label20.BackColor = Color.Green;
                label20.Text = "Internal Protection";
            }


            if (Error_Register[7])
            {
                label21.BackColor = Color.Red;
                label21.Text = "Power Signal change with Filament ON";
            }
            else
            {
                label21.BackColor = Color.Green;
                label21.Text = "No Change";
            }


            //PLC Communication starts here



            ushort Filament_Power = (ushort)plc1.Read("DB470.DBW2");


            ushort Set_GyrotronCurrent = (ushort)plc1.Read("DB470.DBW4");


            ushort Filament_Control = (ushort)plc1.Read("DB470.DBW6");

            ushort Error_Reset = (ushort)plc1.Read("DB470.DBW8");

            ushort Enable_Forcing = (ushort)plc1.Read("DB470.DBW10");

            ushort Enable_Stabilization = (ushort)plc1.Read("DB470.DBW12");

            ushort Enable_FrontPanelLock = (ushort)plc1.Read("DB470.DBW14");

            label24.Text = Convert.ToString(Filament_Power);
            label25.Text = Convert.ToString(Set_GyrotronCurrent);
            label26.Text = Convert.ToString(Filament_Control);
            label27.Text = Convert.ToString(Error_Reset);
            label28.Text = Convert.ToString(Enable_Forcing);
            label29.Text = Convert.ToString(Enable_Stabilization);
            label30.Text = Convert.ToString(Enable_FrontPanelLock);





            //txbPower.Text = Convert.ToString(Filament_Power);
            //txbSetCurrent.Text = Convert.ToString(Set_GyrotronCurrent);
            /*
            int count = (Filament_Power / ratio);
            while( i < count)
            {
                i = i + 1;
                Task.Delay(1000).Wait();
                textBox1.Text= i.ToString();
                
            }
            count = -1;
            i = 0;
            */

            mc1.WriteSingleRegister(5, Filament_Power);
            mc1.WriteSingleRegister(7, Set_GyrotronCurrent);
            mc1.WriteSingleRegister(10, Filament_Control);
            mc1.WriteSingleRegister(13, Enable_Stabilization);
            mc1.WriteSingleRegister(14, Enable_FrontPanelLock);

            ushort Status_ILK_Register = (ushort)values[0];
            ushort Error_Word = (ushort)values[1];
            plc1.Write("DB470.DBW68", Status_ILK_Register);
            plc1.Write("DB470.DBW70", Error_Word);

            // IF CFPS is ON, Read actual Filament current, Gyrotron current and power. Then, transfer this power to PLC register to display on WinCC
            if (Status_Register[4])
            {
                ushort ReadBack_FilamentCurrent = (ushort)values[2];
                plc1.Write("DB470.DBW16", ReadBack_FilamentCurrent);

                ushort ReadBack_GyrotronCurrent = (ushort)values[3];
                plc1.Write("DB470.DBW18", ReadBack_GyrotronCurrent);

                ushort ReadBack_Power = (ushort)values[5];
                plc1.Write("DB470.DBW20", ReadBack_Power);



            }

        }


        public Read()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {


        }

        private void Start_Click(object sender, EventArgs e)
        {


        }

        private void Stop_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {


        }

        private void bON_Click(object sender, EventArgs e)
        {
            //int x = Convert.ToInt32(textBox1.Text);
            //if (x > 500)
            //{
            //    MessageBox.Show("Enter values between 0 to 500");
            //    x = 500;
            //}

            //mc1.WriteSingleCoil(0, true);
            //mc1.WriteSingleRegister(0, x);
            //mc1.WriteSingleRegister(1, 889);
            //mc1.WriteMultipleRegisters(2, new int[3] { 345, 567, 23 });

        }

        private void bOFF_Click(object sender, EventArgs e)
        {
            //mc1.WriteSingleCoil(0, false);
            //mc1.WriteMultipleRegisters(0, new int[5] { 22, 22, 22, 22, 22 });
        }

        private void bGET_Click(object sender, EventArgs e)
        {

            //var values = mc1.ReadInputRegisters(0, 3);
            //MessageBox.Show(Convert.ToString(values[0]));
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void ErrorReset_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {

        }

        private void SetFilament_Click(object sender, EventArgs e)
        {

            mc1.WriteSingleRegister(10, 1);
            SetFilament.BackColor = Color.Green;
            FilamentOFF.BackColor = Color.White;
        }

        private void label4_Click_1(object sender, EventArgs e)
        {

        }

        private void txbPower_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnSend_Click(object sender, EventArgs e)
        {
            try
            {
                short x = Convert.ToInt16(txbPower.Text);

                short z = Convert.ToInt16(txbSetCurrent.Text);
                if (x > 500)
                {
                    MessageBox.Show("Enter values between 0 to 500 for Filament Power");
                }
                else
                {
                    mc1.WriteSingleRegister(5, x);

                }
                /*
                if (y > 15 || y < 3)
                {
                    MessageBox.Show("Enter values between 3 to 15 for Delta Power");
                }
                else
                {

                    mc1.WriteSingleRegister(6, y);
                }*/
                if (z > 550)
                {
                    MessageBox.Show("Enter values between 0 and 550 for Set Gyrotron Current");
                }
                else
                {

                    mc1.WriteSingleRegister(7, z);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }

        }

        private void FilamentOFF_Click(object sender, EventArgs e)
        {
            mc1.WriteSingleRegister(10, 0);
            FilamentOFF.BackColor = Color.Orange;
            SetFilament.BackColor = Color.White;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            mc1.WriteSingleRegister(10, 1);
            button3.BackColor = Color.Green;
            button2.BackColor = Color.White;
        }

        private void button2_Click_2(object sender, EventArgs e)
        {
            mc1.WriteSingleRegister(10, 0);
            button3.BackColor = Color.White;
            button2.BackColor = Color.Orange;
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            mc1.WriteSingleRegister(11, 1);
            button5.BackColor = Color.Green;
            button4.BackColor = Color.White;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            mc1.WriteSingleRegister(11, 0);
            button4.BackColor = Color.Orange;
            button5.BackColor = Color.White;

        }

        private void button6_Click(object sender, EventArgs e)
        {
            mc1.WriteSingleRegister(14, 0);
            button6.BackColor = Color.Orange;
            button7.BackColor = Color.White;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            mc1.WriteSingleRegister(14, 1);
            button7.BackColor = Color.Green;
            button6.BackColor = Color.White;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            mc1.WriteSingleRegister(13, 1);
            button9.BackColor = Color.Green;
            button8.BackColor = Color.White;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            mc1.WriteSingleRegister(13, 0);
            button8.BackColor = Color.Orange;
            button9.BackColor = Color.White;
        }

        private void txbGetData_Click(object sender, EventArgs e)
        {

        }


        private void timer1_Tick(object sender, EventArgs e)
        {
            if (SetTimer)
            {
                try
                {

                    accessModbus();
                    label4.Visible = false;

                }
                catch (Exception ex)
                {
                    label4.Text = ex.Message;
                    label4.Visible = true;

                }



            }





        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void button10_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click_2(object sender, EventArgs e)
        {

        }

        private void lblRead3_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click_1(object sender, EventArgs e)
        {

        }

        private void label7_Click_2(object sender, EventArgs e)
        {

        }

        private void button11_Click(object sender, EventArgs e)
        {

            mc1 = new ModbusClient(this.cbPorts.GetItemText(this.cbPorts.SelectedItem));
            mc1.UnitIdentifier = 3;
            mc1.Baudrate = 9600;
            mc1.Parity = System.IO.Ports.Parity.None;
            mc1.StopBits = System.IO.Ports.StopBits.One;
            try
            {
                mc1.Connect();
                if (mc1.Connected)
                {
                    label6.Text = "Modbus Device Connection Successful";
                    label6.Visible = true;
                    SetTimer = true;
                    button11.Text = "Connected";
                    button11.ForeColor = Color.Yellow;
                    button11.BackColor = Color.Green;
                    label7.ForeColor = Color.Yellow;
                    label8.ForeColor = Color.Yellow;
                    label15.ForeColor = Color.Yellow;
                    label16.ForeColor = Color.Yellow;
                    label17.ForeColor = Color.Yellow;
                    label18.ForeColor = Color.Yellow;
                    label19.ForeColor = Color.Yellow;
                    label20.ForeColor = Color.Yellow;
                    label21.ForeColor = Color.Yellow;
                    timer4.Enabled = true;


                }
            }
            catch
            {
                label6.Text = "Modbus Device Not available";
                label6.Visible = true;

            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            mc1.Disconnect();
            label6.Text = "Modbus device disconnected";
            button11.Text = "Connect";
            button11.BackColor = Color.White;
            if (!mc1.Connected)
            {
                button12.Text = "Disconnected";
            }
            timer3.Enabled = false;
            label7.ForeColor = Color.DarkViolet;
            label8.ForeColor = Color.DarkViolet;
            label15.ForeColor = Color.DarkViolet;
            label16.ForeColor = Color.DarkViolet;
            label17.ForeColor = Color.DarkViolet;
            label18.ForeColor = Color.DarkViolet;
            label19.ForeColor = Color.DarkViolet;
            label20.ForeColor = Color.DarkViolet;
            label21.ForeColor = Color.DarkViolet;
            timer4.Enabled = false;
            try
            {
                con.Close();
            }
            catch (Exception)
            {

                label4.Text = "Error";
            }
           

        }

        private void txbRBDPf_TextChanged(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void label22_Click(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {

        }

        private void label26_Click(object sender, EventArgs e)
        {

        }

        private void notifyIcon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Show();
            this.WindowState = FormWindowState.Normal;
            notifyIcon.Visible = false;
        }

        private void Read_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                Hide();
                notifyIcon.Visible = true;
                notifyIcon.ShowBalloonTip(1000);
            }
        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void Read_Load(object sender, EventArgs e)
        {
            FilamentMinutes = Properties.Settings.Default.FilamentMinutes;
            string hrsminutes= convertMinutes(FilamentMinutes);
            lblFilamentMin.Text = hrsminutes;
            string[] ports = SerialPort.GetPortNames();
            cbPorts.Items.AddRange(ports);
            cbPorts.SelectedIndex = 0;
           
            plc1 = new Plc(CpuType.S7300, "10.136.120.21", 0, 2);

            try
            {
                plc1.Open();
                plc_status.Text = "PLC Mode Activated";
                plc_status.Visible = true;
                plc_connected = true;

            }
            catch
            {
                plc_status.Text = "Standalone Mode Activated";
                plc_status.Visible = true;
                plc_connected = false;
            }

            try
            {
                connectionString = @"Data Source=192.168.23.103,49170;Network Library = DBMSSOCN;Initial Catalog=ECH_Campaign_2;User ID=sa;Password=123";
                con = new SqlConnection(connectionString);
                con.Open();
                sqlConnectionEstablished = true;
            }
            catch (Exception ex)
            {

                label22.Text = ex.Message;
            }

        }

        private string convertMinutes(int filamentMinutes)
        {
            int totalHours = filamentMinutes / 60;
            int remainingMinutes = filamentMinutes % 60;
            string th = totalHours.ToString()+" hrs ";
            string rm = remainingMinutes.ToString()+" min";
            string outputString = th + rm;
            return outputString;
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            try
            {
                int power = Convert.ToInt32(actualFilamentPower);
                // int power = Convert.ToInt32(txbTemp.Text);
                if (power > 150)
                {
                    timer3.Enabled = true;

                }
                else
                {
                    timer3.Enabled = false;
                }
            }
            catch
            {
                timer3.Enabled = false;
            }


        }

        private void timer3_Tick(object sender, EventArgs e)
        {
           
            FilamentMinutes = FilamentMinutes + 1;
            //txbTotalMinutes.Text = FilamentMinutes.ToString();
            lblFilamentMin.Text = convertMinutes(FilamentMinutes);
        }

        private void Read_FormClosed(object sender, FormClosedEventArgs e)
        {
            Properties.Settings.Default.FilamentMinutes = FilamentMinutes;
            Properties.Settings.Default.Save();
            try
            {
                con.Close();
            }
            catch (Exception)
            {

                label4.Text = "Error in closing the DB";
            }
        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void timer4_Tick(object sender, EventArgs e)
        {
            try
            {
                if (sqlConnectionEstablished)
                {
                    string queryString = "insert into CFPS_Data (FilamentPower, FilamentCurrent, GyrotronCurrent) values ('" + RBFilamentPower + "','" + RBFilamentCurrent + "','" + RBGyrotronCurrent + "')";
                    SqlCommand cmd = new SqlCommand(queryString, con);
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception)
            {

                label22.Text = "Database Connection not available";
            }


        }
    }
}


