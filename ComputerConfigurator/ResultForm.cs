using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ComputerConfigurator
{

    public partial class ResultForm : Form
    {        
        public ResultForm(Settings settings)
        {
            InitializeComponent();
            int _b = 60;

            label1.Location = new System.Drawing.Point(28, _b);
            MBlabel.Location = new System.Drawing.Point(144, _b);
            MBsitelabel.Location= new System.Drawing.Point(611, _b);

            label2.Location = new System.Drawing.Point(28, _b+=35);
            CPUlabel.Location = new System.Drawing.Point(144, _b);
            CPUsitelabel.Location = new System.Drawing.Point(611, _b);

            label3.Location = new System.Drawing.Point(28, _b+=35);
            GClabel.Location = new System.Drawing.Point(144, _b);
            GCsitelabel.Location = new System.Drawing.Point(611, _b);

            label4.Location = new System.Drawing.Point(28, _b+=35);
            RAMlabel.Location = new System.Drawing.Point(144, _b);
            RAMsitelabel.Location = new System.Drawing.Point(611, _b);

            label5.Location = new System.Drawing.Point(28, _b+=35);
            PSlabel.Location = new System.Drawing.Point(144, _b);
            PSsitelabel.Location = new System.Drawing.Point(611, _b);

            label6.Location = new System.Drawing.Point(28, _b+=35);
            Clabel.Location = new System.Drawing.Point(144, _b);
            Csitelabel.Location = new System.Drawing.Point(611, _b);

            label7.Location = new System.Drawing.Point(28, _b+=35);
            HDDlabel.Location = new System.Drawing.Point(144, _b);
            HDDsitelabel.Location = new System.Drawing.Point(611, _b);

            label8.Location = new System.Drawing.Point(28, _b+=35);
            SSDlabel.Location = new System.Drawing.Point(144, _b);
            SSDsitelabel.Location = new System.Drawing.Point(611, _b);

            label9.Location = new System.Drawing.Point(28, _b += 35);
            TotalPricelabel.Location = new System.Drawing.Point(144, _b);

            Excel.Range range;

            // Получить объект приложения Excel.
            Excel.Application excel_app = new Excel.Application();

            Excel.Workbook MBworkbook = excel_app.Workbooks.Open(
                Path.GetFullPath("../../Resources/MotherBoard"),
                Type.Missing, true, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            // Получить первый рабочий лист.
            Excel.Worksheet MBsheet = (Excel.Worksheet)MBworkbook.Sheets[1];

            int MB_c = 2;
            range = (Excel.Range)MBsheet.Cells[MB_c, 1];
            string MB_x = range.Value2;

            //////////////////////////////////////
            Excel.Workbook CPUworkbook = excel_app.Workbooks.Open(
                        Path.GetFullPath("../../Resources/CPU"),
                        Type.Missing, true, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);

            // Получить первый рабочий лист.
            Excel.Worksheet CPUsheet = (Excel.Worksheet)CPUworkbook.Sheets[1];

            int CPU_c = 2;
            range = (Excel.Range)CPUsheet.Cells[CPU_c, 1];
            string CPU_x = range.Value2;
            ////////////////////////////////////////////////
            
            ///////////////////////////////////////////////////
            Excel.Workbook CorpsWorkbook = excel_app.Workbooks.Open(
                        Path.GetFullPath("../../Resources/Corps"),
                        Type.Missing, true, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);

            // Получить первый рабочий лист.
            Excel.Worksheet CorpsSheet = (Excel.Worksheet)CorpsWorkbook.Sheets[1];

            int Corps_c = 2;
            range = (Excel.Range)CorpsSheet.Cells[Corps_c, 1];
            string Corps_x = range.Value2;
            ////////////////////////////////////////////////////

            ////////////////////////////////////////////////////
            Excel.Workbook GCworkbook = excel_app.Workbooks.Open(
                        Path.GetFullPath("../../Resources/GraphicsCard"),
                        Type.Missing, true, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);

            // Получить первый рабочий лист.
            Excel.Worksheet GCsheet = (Excel.Worksheet)GCworkbook.Sheets[1];

            int GC_c = 2;
            range = (Excel.Range)GCsheet.Cells[GC_c, 1];
            string GC_x = range.Value2;
            //////////////////////////////////////////////////

            //////////////////////////////////////////////////
            Excel.Workbook PSworkbook = excel_app.Workbooks.Open(
                        Path.GetFullPath("../../Resources/PowerSupply"),
                        Type.Missing, true, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);

            // Получить первый рабочий лист.
            Excel.Worksheet PSsheet = (Excel.Worksheet)PSworkbook.Sheets[1];

            int PS_c = 2;
            range = (Excel.Range)PSsheet.Cells[PS_c, 1];
            string PS_x = range.Value2;
            //////////////////////////////////////////////////

            /////////////////////////////////////////////////
            Excel.Workbook RAMworkbook = excel_app.Workbooks.Open(
                       Path.GetFullPath("../../Resources/RAM"),
                       Type.Missing, true, Type.Missing, Type.Missing,
                       Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                       Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                       Type.Missing, Type.Missing);

            // Получить первый рабочий лист.
            Excel.Worksheet RAMsheet = (Excel.Worksheet)RAMworkbook.Sheets[1];

            int RAM_c = 2;
            range = (Excel.Range)RAMsheet.Cells[RAM_c, 1];
            string RAM_x = range.Value2;
            ////////////////////////////////////////////////

            /////////////////////////////////////////////////
            Excel.Workbook HDDworkbook = excel_app.Workbooks.Open(
                        Path.GetFullPath("../../Resources/HDD"),
                        Type.Missing, true, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);

            // Получить первый рабочий лист.
            Excel.Worksheet HDDsheet = (Excel.Worksheet)HDDworkbook.Sheets[1];

            int HDD_c = 2;
            range = (Excel.Range)HDDsheet.Cells[HDD_c, 1];
            string HDD_x = range.Value2;
            ////////////////////////////////////////////////

            ///////////////////////////////////////////////
            Excel.Workbook SSDworkbook = excel_app.Workbooks.Open(
                        Path.GetFullPath("../../Resources/SSD"),
                        Type.Missing, true, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);

            // Получить первый рабочий лист.
            Excel.Worksheet SSDsheet = (Excel.Worksheet)SSDworkbook.Sheets[1];

            int SSD_c = 2;
            range = (Excel.Range)SSDsheet.Cells[SSD_c, 1];
            string SSD_x = range.Value2;
            /////////////////////////////////////////////////

            MotherBoard mb = new MotherBoard();

            while (MB_x != null)
            {
                PC computer = new PC();

                mb.Name = MB_x;
                range = (Excel.Range)MBsheet.Cells[MB_c, 2];
                mb.Price = range.Value2;
                range = (Excel.Range)MBsheet.Cells[MB_c, 3];
                mb.Fabricator = range.Value2;
                range = (Excel.Range)MBsheet.Cells[MB_c, 4];
                mb.CPUType = range.Value2;
                range = (Excel.Range)MBsheet.Cells[MB_c, 5];
                mb.Socket = range.Value2;
                range = (Excel.Range)MBsheet.Cells[MB_c, 6];
                mb.Chipset = range.Value2;
                range = (Excel.Range)MBsheet.Cells[MB_c, 7];
                mb.NumberOfPCIEx16Slots = Convert.ToString(range.Value2);
                range = (Excel.Range)MBsheet.Cells[MB_c, 8];
                mb.NumberOfM2Slots = Convert.ToString(range.Value2);
                range = (Excel.Range)MBsheet.Cells[MB_c, 9];
                mb.WiFiAdapter = range.Value2;
                range = (Excel.Range)MBsheet.Cells[MB_c, 10];
                mb.BuildInCPU = range.Value2;
                range = (Excel.Range)MBsheet.Cells[MB_c, 11];
                mb.ForGamingPC = range.Value2;
                range = (Excel.Range)MBsheet.Cells[MB_c, 12];
                mb.MemoryType = range.Value2;
                range = (Excel.Range)MBsheet.Cells[MB_c, 13];
                mb.NumberOfMemorySlots = Convert.ToString(range.Value2);
                range = (Excel.Range)MBsheet.Cells[MB_c, 14];
                mb.Site = range.Value2;

                if (SettingsCheck(mb, settings))
                {
                    computer.motherBoard = mb;

                    //ищем cpu
                    if ((computer.motherBoard.Name != "NULL") && (computer.motherBoard.BuildInCPU == "Нет"))
                    {
                        CPU cpu = new CPU();
                        while (CPU_x != null)
                        {
                            cpu.Name = CPU_x;
                            range = (Excel.Range)CPUsheet.Cells[CPU_c, 2];
                            cpu.Price = range.Value2;
                            range = (Excel.Range)CPUsheet.Cells[CPU_c, 3];
                            cpu.Fabricator = range.Value2;
                            range = (Excel.Range)CPUsheet.Cells[CPU_c, 4];
                            cpu.NumberOfCores = Convert.ToString(range.Value2);
                            range = (Excel.Range)CPUsheet.Cells[CPU_c, 5];
                            cpu.GraphicCore = range.Value2;
                            range = (Excel.Range)CPUsheet.Cells[CPU_c, 6];
                            cpu.MemoryType = range.Value2;
                            range = (Excel.Range)CPUsheet.Cells[CPU_c, 7];
                            cpu.BaseFrequency = range.Value2;
                            range = (Excel.Range)CPUsheet.Cells[CPU_c, 8];
                            cpu.Multithreading = range.Value2;
                            range = (Excel.Range)CPUsheet.Cells[CPU_c, 9];
                            cpu.Socket = range.Value2;
                            range = (Excel.Range)CPUsheet.Cells[CPU_c, 10];
                            cpu.ForGamingPC = range.Value2;
                            range = (Excel.Range)CPUsheet.Cells[CPU_c, 11];
                            cpu.Site = range.Value2;

                            if ((SettingsCheck(cpu, settings)) && (computer.motherBoard.Socket == cpu.Socket) && (computer.motherBoard.CPUType == cpu.Fabricator))
                            {
                                computer.processor = cpu;
                                break;
                            }

                            CPU_c++;
                            range = (Excel.Range)CPUsheet.Cells[CPU_c, 1];
                            CPU_x = range.Value2;
                        }
                    }
                    
                    //ищем graphicscard
                    if ((computer.processor.Name != "NULL") | ((computer.motherBoard.BuildInCPU == "Есть") && (computer.motherBoard.Name != "NULL")))
                    {
                        
                        GraphicsCard gc = new GraphicsCard();

                        while (GC_x != null)
                        {
                            gc.Name = GC_x;
                            range = (Excel.Range)GCsheet.Cells[GC_c, 2];
                            gc.Price = range.Value2;
                            range = (Excel.Range)GCsheet.Cells[GC_c, 3];
                            gc.Fabricator = range.Value2;
                            range = (Excel.Range)GCsheet.Cells[GC_c, 4];
                            gc.RecommendedEnergy = range.Value2;
                            range = (Excel.Range)GCsheet.Cells[GC_c, 5];
                            gc.Memory = Convert.ToString(range.Value2);
                            range = (Excel.Range)GCsheet.Cells[GC_c, 6];
                            gc.MemoryType = range.Value2;
                            range = (Excel.Range)GCsheet.Cells[GC_c, 7];
                            gc.FabricatorOfGPU = range.Value2;
                            range = (Excel.Range)GCsheet.Cells[GC_c, 8];
                            gc.NumberOfMonitors = Convert.ToString(range.Value2);
                            range = (Excel.Range)GCsheet.Cells[GC_c, 9];
                            gc.PCIExpress = Convert.ToString(range.Value2);
                            range = (Excel.Range)GCsheet.Cells[GC_c, 10];
                            gc.MemoryBusWidth = range.Value2;
                            range = (Excel.Range)GCsheet.Cells[GC_c, 11];
                            gc.ForGamingPC = range.Value2;
                            range = (Excel.Range)GCsheet.Cells[GC_c, 12];
                            gc.ProfessionalGraphicsCard = range.Value2;
                            range = (Excel.Range)GCsheet.Cells[GC_c, 13];
                            gc.Site = range.Value2;

                            if (SettingsCheck(gc, settings))
                            {
                                computer.graphicsCard = gc;
                                break;
                            }

                            GC_c++;
                            range = (Excel.Range)GCsheet.Cells[GC_c, 1];
                            GC_x = range.Value2;
                        }
                    }

                    //ищем powersupply
                    if (computer.graphicsCard.Name != "NULL")
                    {
                        

                        PowerSupply ps = new PowerSupply();

                        while (PS_x != null)
                        {
                            ps.Name = PS_x;
                            range = (Excel.Range)PSsheet.Cells[PS_c, 2];
                            ps.Price = range.Value2;
                            range = (Excel.Range)PSsheet.Cells[PS_c, 3];
                            ps.Fabricator = range.Value2;
                            range = (Excel.Range)PSsheet.Cells[PS_c, 4];
                            ps.Energy = range.Value2;
                            range = (Excel.Range)PSsheet.Cells[PS_c, 5];
                            ps.WireBraiding = range.Value2;
                            range = (Excel.Range)PSsheet.Cells[PS_c, 6];
                            ps.Backlight = range.Value2;
                            range = (Excel.Range)PSsheet.Cells[PS_c, 7];
                            ps.DetachableCables = range.Value2;
                            range = (Excel.Range)PSsheet.Cells[PS_c, 8];
                            ps.Site = range.Value2;

                            if ((SettingsCheck(ps, settings)) && (computer.graphicsCard.RecommendedEnergy <= ps.Energy))
                            {
                                computer.powerSupply = ps;
                                break;
                            }

                            PS_c++;
                            range = (Excel.Range)PSsheet.Cells[PS_c, 1];
                            PS_x = range.Value2;
                        }

                    }

                    //ищем ram
                    if (computer.powerSupply.Name != "NULL")
                    {
                       

                        RAM ram = new RAM();
                        while (RAM_x != null)
                        {
                            ram.Name = RAM_x;
                            range = (Excel.Range)RAMsheet.Cells[RAM_c, 2];
                            ram.Price = range.Value2;
                            range = (Excel.Range)RAMsheet.Cells[RAM_c, 3];
                            ram.Fabricator = range.Value2;
                            range = (Excel.Range)RAMsheet.Cells[RAM_c, 4];
                            ram.Backlight = range.Value2;
                            range = (Excel.Range)RAMsheet.Cells[RAM_c, 5];
                            ram.Memory = Convert.ToString(range.Value2);
                            range = (Excel.Range)RAMsheet.Cells[RAM_c, 6];
                            ram.MemoryType = range.Value2;
                            range = (Excel.Range)RAMsheet.Cells[RAM_c, 7];
                            ram.ForGamingPC = range.Value2;
                            range = (Excel.Range)RAMsheet.Cells[RAM_c, 8];
                            ram.Site = range.Value2;

                            if (SettingsCheck(ram, settings) && (computer.motherBoard.MemoryType == ram.MemoryType))
                            {
                                computer.ram = ram;
                                break;
                            }

                            RAM_c++;
                            range = (Excel.Range)RAMsheet.Cells[RAM_c, 1];
                            RAM_x = range.Value2;
                        }
                    }

                    //ищем hdd
                    if (computer.ram.Name != "NULL" && (settings.HDDRequired == true))
                    {
                        

                        HDD hdd = new HDD();

                        while (HDD_x != null)
                        {
                            hdd.Name = HDD_x;
                            range = (Excel.Range)HDDsheet.Cells[HDD_c, 2];
                            hdd.Price = range.Value2;
                            range = (Excel.Range)HDDsheet.Cells[HDD_c, 3];
                            hdd.Fabricator = range.Value2;
                            range = (Excel.Range)HDDsheet.Cells[HDD_c, 4];
                            hdd.Memory = Convert.ToString(range.Value2);
                            range = (Excel.Range)HDDsheet.Cells[HDD_c, 5];
                            hdd.LevelOfNoise = range.Value2;
                            range = (Excel.Range)HDDsheet.Cells[HDD_c, 6];
                            hdd.DataExchangeRate = range.Value2;
                            range = (Excel.Range)HDDsheet.Cells[HDD_c, 7];
                            hdd.BufferSize = Convert.ToString(range.Value2);
                            range = (Excel.Range)HDDsheet.Cells[HDD_c, 8];
                            hdd.Site = range.Value2;

                            if (SettingsCheck(hdd, settings))
                            {
                                computer.hdd = hdd;
                                break;
                            }

                            HDD_c++;
                            range = (Excel.Range)HDDsheet.Cells[HDD_c, 1];
                            HDD_x = range.Value2;
                        }
                    }

                    //ищем ssd
                    if (computer.ram.Name != "NULL" && (settings.SSDRequired == true))
                    {
                        

                        SSD ssd = new SSD();

                        while (SSD_x != null)
                        {
                            ssd.Name = SSD_x;
                            range = (Excel.Range)SSDsheet.Cells[SSD_c, 2];
                            ssd.Price = range.Value2;
                            range = (Excel.Range)SSDsheet.Cells[SSD_c, 3];
                            ssd.Fabricator = range.Value2;
                            range = (Excel.Range)SSDsheet.Cells[SSD_c, 4];
                            ssd.Memory = range.Value2;
                            range = (Excel.Range)SSDsheet.Cells[SSD_c, 5];
                            ssd.WriteSpeed = range.Value2;
                            range = (Excel.Range)SSDsheet.Cells[SSD_c, 6];
                            ssd.ReadSpeed = range.Value2;
                            range = (Excel.Range)SSDsheet.Cells[SSD_c, 7];
                            ssd.Site = range.Value2;


                            if (SettingsCheck(ssd, settings))
                            {
                                computer.ssd = ssd;
                                break;
                            }

                            SSD_c++;
                            range = (Excel.Range)SSDsheet.Cells[SSD_c, 1];
                            SSD_x = range.Value2;
                        }
                    }

                    //ищем corps
                    if (settings.SSDRequired|settings.HDDRequired)
                    {
                        Corps corp = new Corps();

                        while (Corps_x != null)
                        {
                            corp.Name = Corps_x;
                            range = (Excel.Range)CorpsSheet.Cells[Corps_c, 2];
                            corp.Price = range.Value2;
                            range = (Excel.Range)CorpsSheet.Cells[Corps_c, 3];
                            corp.Fabricator = range.Value2;
                            range = (Excel.Range)CorpsSheet.Cells[Corps_c, 4];
                            corp.MainColor = range.Value2;
                            range = (Excel.Range)CorpsSheet.Cells[Corps_c, 5];
                            corp.Window = range.Value2;
                            range = (Excel.Range)CorpsSheet.Cells[Corps_c, 6];
                            corp.Backlight = range.Value2;
                            range = (Excel.Range)CorpsSheet.Cells[Corps_c, 7];
                            corp.FrameSize = range.Value2;
                            range = (Excel.Range)CorpsSheet.Cells[Corps_c, 8];
                            corp.ForGamingPC = range.Value2;
                            range = (Excel.Range)CorpsSheet.Cells[Corps_c, 9];
                            corp.Site = range.Value2;

                            if (SettingsCheck(corp, settings))
                            {
                                computer.corps = corp;
                                break;
                            }

                            Corps_c++;
                            range = (Excel.Range)CorpsSheet.Cells[Corps_c, 1];
                            Corps_x = range.Value2;
                        }
                    }

                    bool IsTotalPriceRight;
                    if((computer.TotalPrice() >= settings.Price[0]) && (computer.TotalPrice() <= settings.Price[1])) { IsTotalPriceRight = true; }
                    else { IsTotalPriceRight = false; }
                    bool IsCPURight;
                    if((((settings.MotherBoardBuildInCPU[0] == "Нет") && (computer.processor.Name != "NULL")) | ((settings.MotherBoardBuildInCPU[1] == "Есть") && (computer.processor.Name == "NULL")))) { IsCPURight = true; }
                    else { IsCPURight = false; }
                    bool IsSSDRight;
                    if ((((settings.SSDRequired == true) && (computer.ssd.Name != "NULL")) | ((settings.SSDRequired == false) && (computer.ssd.Name == "NULL")))) { IsSSDRight = true; }
                    else { IsSSDRight = false; }
                    bool IsHDDRight;
                    if ((((settings.HDDRequired == true) && (computer.hdd.Name != "NULL")) | ((settings.HDDRequired == false) && (computer.hdd.Name == "NULL")))) { IsHDDRight = true; }
                    else { IsHDDRight = false; }

                    if ((IsTotalPriceRight==true)&&(IsCPURight==true) && (IsSSDRight==true) && (IsHDDRight==true) && (computer.corps.Name!="NULL") && (computer.ram.Name!="NULL") && (computer.graphicsCard.Name!="NULL") && (computer.powerSupply.Name!="NULL"))//комп собран, надо проверить ценник
                    {
                        if (computers.Count() < settings.NumberOfPCs)
                        {
                            computers.Add(computer);
                            Corps_c++;
                            range = (Excel.Range)CorpsSheet.Cells[Corps_c, 1];
                            Corps_x = range.Value2;
                        }
                        else
                        {
                            MB_x = null;
                        }
                    }
                    else//комп не собран, выясняем чего не хватает
                    {
                        
                        if((computer.processor.Name=="NULL")&&(settings.MotherBoardBuildInCPU[0] == "Есть"))
                        {
                            MB_c++;
                            range = (Excel.Range)MBsheet.Cells[MB_c, 1];
                            MB_x = range.Value2;

                            CPU_c = 2;
                            range = (Excel.Range)CPUsheet.Cells[CPU_c, 1];
                            CPU_x = range.Value2;

                            Corps_c = 2;
                            range = (Excel.Range)CorpsSheet.Cells[Corps_c, 1];
                            Corps_x = range.Value2;

                            GC_c = 2;
                            range = (Excel.Range)GCsheet.Cells[GC_c, 1];
                            GC_x = range.Value2;

                            PS_c = 2;
                            range = (Excel.Range)PSsheet.Cells[PS_c, 1];
                            PS_x = range.Value2;

                            RAM_c = 2;
                            range = (Excel.Range)RAMsheet.Cells[RAM_c, 1];
                            RAM_x = range.Value2;

                            HDD_c = 2;
                            range = (Excel.Range)HDDsheet.Cells[HDD_c, 1];
                            HDD_x = range.Value2;

                            SSD_c = 2;
                            range = (Excel.Range)SSDsheet.Cells[SSD_c, 1];
                            SSD_x = range.Value2;
                        }
                        else if(computer.graphicsCard.Name=="NULL")
                        {
                            if(settings.MotherBoardBuildInCPU[0] == "Есть")
                            {
                                CPU_c++;
                                range = (Excel.Range)CPUsheet.Cells[CPU_c, 1];
                                CPU_x = range.Value2;

                                Corps_c = 2;
                                range = (Excel.Range)CorpsSheet.Cells[Corps_c, 1];
                                Corps_x = range.Value2;

                                GC_c = 2;
                                range = (Excel.Range)GCsheet.Cells[GC_c, 1];
                                GC_x = range.Value2;

                                PS_c = 2;
                                range = (Excel.Range)PSsheet.Cells[PS_c, 1];
                                PS_x = range.Value2;

                                RAM_c = 2;
                                range = (Excel.Range)RAMsheet.Cells[RAM_c, 1];
                                RAM_x = range.Value2;

                                HDD_c = 2;
                                range = (Excel.Range)HDDsheet.Cells[HDD_c, 1];
                                HDD_x = range.Value2;

                                SSD_c = 2;
                                range = (Excel.Range)SSDsheet.Cells[SSD_c, 1];
                                SSD_x = range.Value2;
                            }
                            else if(settings.MotherBoardBuildInCPU[1] == "Нет")
                            {
                                MB_c++;
                                range = (Excel.Range)MBsheet.Cells[MB_c, 1];
                                MB_x = range.Value2;                                

                                Corps_c = 2;
                                range = (Excel.Range)CorpsSheet.Cells[Corps_c, 1];
                                Corps_x = range.Value2;

                                GC_c = 2;
                                range = (Excel.Range)GCsheet.Cells[GC_c, 1];
                                GC_x = range.Value2;

                                PS_c = 2;
                                range = (Excel.Range)PSsheet.Cells[PS_c, 1];
                                PS_x = range.Value2;

                                RAM_c = 2;
                                range = (Excel.Range)RAMsheet.Cells[RAM_c, 1];
                                RAM_x = range.Value2;

                                HDD_c = 2;
                                range = (Excel.Range)HDDsheet.Cells[HDD_c, 1];
                                HDD_x = range.Value2;

                                SSD_c = 2;
                                range = (Excel.Range)SSDsheet.Cells[SSD_c, 1];
                                SSD_x = range.Value2;
                            }
                        }
                        else if(computer.powerSupply.Name=="NULL")
                        {
                            GC_c++;
                            range = (Excel.Range)GCsheet.Cells[GC_c, 1];
                            GC_x = range.Value2;

                            Corps_c = 2;
                            range = (Excel.Range)CorpsSheet.Cells[Corps_c, 1];
                            Corps_x = range.Value2;                            

                            PS_c = 2;
                            range = (Excel.Range)PSsheet.Cells[PS_c, 1];
                            PS_x = range.Value2;

                            RAM_c = 2;
                            range = (Excel.Range)RAMsheet.Cells[RAM_c, 1];
                            RAM_x = range.Value2;

                            HDD_c = 2;
                            range = (Excel.Range)HDDsheet.Cells[HDD_c, 1];
                            HDD_x = range.Value2;

                            SSD_c = 2;
                            range = (Excel.Range)SSDsheet.Cells[SSD_c, 1];
                            SSD_x = range.Value2;
                        }
                        else if(computer.ram.Name=="NULL")
                        {
                            PS_c++;
                            range = (Excel.Range)PSsheet.Cells[PS_c, 1];
                            PS_x = range.Value2;

                            Corps_c = 2;
                            range = (Excel.Range)CorpsSheet.Cells[Corps_c, 1];
                            Corps_x = range.Value2;

                            RAM_c = 2;
                            range = (Excel.Range)RAMsheet.Cells[RAM_c, 1];
                            RAM_x = range.Value2;

                            HDD_c = 2;
                            range = (Excel.Range)HDDsheet.Cells[HDD_c, 1];
                            HDD_x = range.Value2;

                            SSD_c = 2;
                            range = (Excel.Range)SSDsheet.Cells[SSD_c, 1];
                            SSD_x = range.Value2;
                        }
                        else if((computer.hdd.Name=="NULL")&&(settings.HDDRequired))
                        {
                            RAM_c++;
                            range = (Excel.Range)RAMsheet.Cells[RAM_c, 1];
                            RAM_x = range.Value2;                            
                           
                            HDD_c = 2;
                            range = (Excel.Range)HDDsheet.Cells[HDD_c, 1];
                            HDD_x = range.Value2;

                            if (settings.SSDRequired)
                            {
                                SSD_c = 2;
                                range = (Excel.Range)SSDsheet.Cells[SSD_c, 1];
                                SSD_x = range.Value2;
                            }

                            Corps_c = 2;
                            range = (Excel.Range)CorpsSheet.Cells[Corps_c, 1];
                            Corps_x = range.Value2;
                        }
                        else if((computer.ssd.Name == "NULL") && (settings.SSDRequired))
                        {
                            if(settings.HDDRequired)
                            {
                                HDD_c++;
                                range = (Excel.Range)HDDsheet.Cells[HDD_c, 1];
                                HDD_x = range.Value2;

                                SSD_c = 2;
                                range = (Excel.Range)SSDsheet.Cells[SSD_c, 1];
                                SSD_x = range.Value2;

                                Corps_c = 2;
                                range = (Excel.Range)CorpsSheet.Cells[Corps_c, 1];
                                Corps_x = range.Value2;
                            }
                            else
                            {
                                RAM_c++;
                                range = (Excel.Range)RAMsheet.Cells[RAM_c, 1];
                                RAM_x = range.Value2;

                                SSD_c = 2;
                                range = (Excel.Range)SSDsheet.Cells[SSD_c, 1];
                                SSD_x = range.Value2;

                                Corps_c = 2;
                                range = (Excel.Range)CorpsSheet.Cells[Corps_c, 1];
                                Corps_x = range.Value2;
                            }
                        }
                        else if(computer.corps.Name=="NULL")
                        {
                            if(settings.SSDRequired)
                            {
                                SSD_c++;
                                range = (Excel.Range)SSDsheet.Cells[SSD_c, 1];
                                SSD_x = range.Value2;

                                Corps_c = 2;
                                range = (Excel.Range)CorpsSheet.Cells[Corps_c, 1];
                                Corps_x = range.Value2;
                            }
                            else
                            {
                                HDD_c++;
                                range = (Excel.Range)HDDsheet.Cells[HDD_c, 1];
                                HDD_x = range.Value2;

                                Corps_c = 2;
                                range = (Excel.Range)CorpsSheet.Cells[Corps_c, 1];
                                Corps_x = range.Value2;
                            }
                        }
                        else if (!IsTotalPriceRight)
                        {
                            Corps_c++;
                            range = (Excel.Range)CorpsSheet.Cells[Corps_c, 1];
                            Corps_x = range.Value2;
                        }
                    }
                }
                else
                {
                    MB_c++;
                    range = (Excel.Range)MBsheet.Cells[MB_c, 1];
                    MB_x = range.Value2;
                }

                

                
            }
            

            if (computers.Count > 0)
            {
                for (int i = 0; i < computers.Count; i++)
                {
                    VariantsComboBox.Items.Add("Вариант №" + (i+1));
                }
            }
            else
            {
                throw new Exception("По вашим параметрам не найдено ни одно сборки ПК!");
            }
        }

        private List<Label> ls = new List<Label>();

        /// <summary>
        /// Список вариантов сборок ПК(если конечно я успею реализовать функционал)
        /// </summary>
        private List<PC> computers= new List<PC>();
        

        /// <summary>
        /// Проверка комплектующего на соответсвие настройкам
        /// </summary>
        /// <param name="component">Комплектющее на проверку</param>
        /// <param name="settings">Настройки</param>
        /// <returns></returns>
        private bool SettingsCheck(ComputerComponent component, Settings settings)
        {
            if(component is CPU)
            {
                var cpu = component as CPU;
                for(int i =0;i<settings.CPUCores.Length; i++)
                {
                    if(settings.CPUCores[i]==cpu.NumberOfCores)
                    {
                        break;
                    }
                    if(i== settings.CPUCores.Length-1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.CPUFabricator.Length; i++)
                {
                    if (settings.CPUFabricator[i] == cpu.Fabricator)
                    {
                        break;
                    }
                    if (i == settings.CPUFabricator.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.CPUGraphicCore.Length; i++)
                {
                    if (settings.CPUGraphicCore[i] == cpu.GraphicCore)
                    {
                        break;
                    }
                    if (i == settings.CPUGraphicCore.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.CPUMemoryType.Length; i++)
                {
                    if (settings.CPUMemoryType[i] == cpu.MemoryType)
                    {
                        break;
                    }
                    if (i == settings.CPUMemoryType.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.CPUMultithreading.Length; i++)
                {
                    if (settings.CPUMultithreading[i] == cpu.Multithreading)
                    {
                        break;
                    }
                    if (i == settings.CPUMultithreading.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.ForGamingPC.Length; i++)
                {
                    if (settings.ForGamingPC[i] == cpu.ForGamingPC)
                    {
                        break;
                    }
                    if (i == settings.ForGamingPC.Length - 1)
                    {
                        return false;
                    }
                }
                if ((settings.CPUBaseFrequency[0]>cpu.BaseFrequency) | (settings.CPUBaseFrequency[1] < cpu.BaseFrequency))
                {
                    return false;
                }

                return true;
            }

            if (component is Corps)
            {
                var corp = component as Corps;
                for (int i = 0; i < settings.CorpsMainColor.Length; i++)
                {
                    if (settings.CorpsMainColor[i] == corp.MainColor)
                    {
                        break;
                    }
                    if (i == settings.CorpsMainColor.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.CorpsWindow.Length; i++)
                {
                    if (settings.CorpsWindow[i] == corp.Window)
                    {
                        break;
                    }
                    if (i == settings.CorpsWindow.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.CorpsFrameSize.Length; i++)
                {
                    if (settings.CorpsFrameSize[i] == corp.FrameSize)
                    {
                        break;
                    }
                    if (i == settings.CorpsFrameSize.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.CorpsFabricator.Length; i++)
                {
                    if (settings.CorpsFabricator[i] == corp.Fabricator)
                    {
                        break;
                    }
                    if (i == settings.CorpsFabricator.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.CorpsBacklight.Length; i++)
                {
                    if (settings.CorpsBacklight[i] == corp.Backlight)
                    {
                        break;
                    }
                    if (i == settings.CorpsBacklight.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.ForGamingPC.Length; i++)
                {
                    if (settings.ForGamingPC[i] == corp.ForGamingPC)
                    {
                        break;
                    }
                    if (i == settings.ForGamingPC.Length - 1)
                    {
                        return false;
                    }
                }
                return true;
            }

            if (component is GraphicsCard)
            {
                var card = component as GraphicsCard;

                for (int i = 0; i < settings.GraphicsCardFabricator.Length; i++)
                {
                    if (settings.GraphicsCardFabricator[i] == card.Fabricator)
                    {
                        break;
                    }
                    if (i == settings.GraphicsCardFabricator.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.GraphicsCardFabricatorOfGPU.Length; i++)
                {
                    if (settings.GraphicsCardFabricatorOfGPU[i] == card.FabricatorOfGPU)
                    {
                        break;
                    }
                    if (i == settings.GraphicsCardFabricatorOfGPU.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.GraphicsCardMemory.Length; i++)
                {
                    if (settings.GraphicsCardMemory[i] == card.Memory)
                    {
                        break;
                    }
                    if (i == settings.GraphicsCardMemory.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.GraphicsCardMemoryType.Length; i++)
                {
                    if (settings.GraphicsCardMemoryType[i] == card.MemoryType)
                    {
                        break;
                    }
                    if (i == settings.GraphicsCardMemoryType.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.GraphicsCardNumberOfMonitors.Length; i++)
                {
                    if (settings.GraphicsCardNumberOfMonitors[i] == card.NumberOfMonitors)
                    {
                        break;
                    }
                    if (i == settings.GraphicsCardNumberOfMonitors.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.GraphicsCardPCIExpress.Length; i++)
                {
                    if (settings.GraphicsCardPCIExpress[i] == card.PCIExpress)
                    {
                        break;
                    }
                    if (i == settings.GraphicsCardPCIExpress.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.ForGamingPC.Length; i++)
                {
                    if (settings.ForGamingPC[i] == card.ForGamingPC)
                    {
                        break;
                    }
                    if (i == settings.ForGamingPC.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.ProfCard.Length; i++)
                {
                    if (settings.ProfCard[i] == card.ProfessionalGraphicsCard)
                    {
                        break;
                    }
                    if (i == settings.ProfCard.Length - 1)
                    {
                        return false;
                    }
                }
                if ((settings.GraphicsCardMemoryBusWidth[0] > card.MemoryBusWidth) | (settings.GraphicsCardMemoryBusWidth[1] < card.MemoryBusWidth))
                {
                    return false;
                }
                return true;
            }

            if (component is HDD)
            {
                var hdd = component as HDD;
                for (int i = 0; i < settings.HDDMemory.Length; i++)
                {
                    if (settings.HDDMemory[i] == hdd.Memory)
                    {
                        break;
                    }
                    if (i == settings.HDDMemory.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.HDDFabricator.Length; i++)
                {
                    if (settings.HDDFabricator[i] == hdd.Fabricator)
                    {
                        break;
                    }
                    if (i == settings.HDDFabricator.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.HDDBufferSize.Length; i++)
                {
                    if (settings.HDDBufferSize[i] == hdd.BufferSize)
                    {
                        break;
                    }
                    if (i == settings.HDDBufferSize.Length - 1)
                    {
                        return false;
                    }
                }
                if ((settings.HDDDataExchangeRate[0] > hdd.DataExchangeRate) | (settings.HDDDataExchangeRate[1] < hdd.DataExchangeRate))
                {
                    return false;
                }
                if ((settings.HDDLevelOfNoise[0] > hdd.LevelOfNoise) | (settings.HDDLevelOfNoise[1] < hdd.LevelOfNoise))
                {
                    return false;
                }
                return true;
            }

            if (component is MotherBoard)
            {
                var mb = component as MotherBoard;

                for (int i = 0; i < settings.MotherBoardBuildInCPU.Length; i++)
                {
                    if (settings.MotherBoardBuildInCPU[i] == mb.BuildInCPU)
                    {
                        if (settings.MotherBoardBuildInCPU[1] == "Есть") { return true; }
                        break;
                    }
                    if (i == settings.MotherBoardBuildInCPU.Length - 1)
                    {
                        return false;
                    }
                }                
                for (int i = 0; i < settings.MotherBoardCPUType.Length; i++)
                {
                    if (settings.MotherBoardCPUType[i] == mb.CPUType)
                    {
                        break;
                    }
                    if (i == settings.MotherBoardCPUType.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.MotherBoardFabricator.Length; i++)
                {
                    if (settings.MotherBoardFabricator[i] == mb.Fabricator)
                    {
                        break;
                    }
                    if (i == settings.MotherBoardFabricator.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.MotherBoardMemoryType.Length; i++)
                {
                    if (settings.MotherBoardMemoryType[i] == mb.MemoryType)
                    {
                        break;
                    }
                    if (i == settings.MotherBoardMemoryType.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.MotherBoardNumberOfM2Slots.Length; i++)
                {
                    if (settings.MotherBoardNumberOfM2Slots[i] == mb.NumberOfM2Slots)
                    {
                        break;
                    }
                    if (i == settings.MotherBoardNumberOfM2Slots.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.MotherBoardNumberOfMemorySlots.Length; i++)
                {
                    if (settings.MotherBoardNumberOfMemorySlots[i] == mb.NumberOfMemorySlots)
                    {
                        break;
                    }
                    if (i == settings.MotherBoardNumberOfMemorySlots.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.MotherBoardNumberOfPCIEx16Slots.Length; i++)
                {
                    if (settings.MotherBoardNumberOfPCIEx16Slots[i] == mb.NumberOfPCIEx16Slots)
                    {
                        break;
                    }
                    if (i == settings.MotherBoardNumberOfPCIEx16Slots.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.MotherBoardWiFiAdapter.Length; i++)
                {
                    if (settings.MotherBoardWiFiAdapter[i] == mb.WiFiAdapter)
                    {
                        break;
                    }
                    if (i == settings.MotherBoardWiFiAdapter.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.ForGamingPC.Length; i++)
                {
                    if (settings.ForGamingPC[i] == mb.ForGamingPC)
                    {
                        break;
                    }
                    if (i == settings.ForGamingPC.Length - 1)
                    {
                        return false;
                    }
                }
                return true;
            }

            if (component is PowerSupply)
            {
                var ps = component as PowerSupply;
                for (int i = 0; i < settings.PowerSupplyFabricator.Length; i++)
                {
                    if (settings.PowerSupplyFabricator[i] == ps.Fabricator)
                    {
                        break;
                    }
                    if (i == settings.PowerSupplyFabricator.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.PowerSupplyBacklight.Length; i++)
                {
                    if (settings.PowerSupplyBacklight[i] == ps.Backlight)
                    {
                        break;
                    }
                    if (i == settings.PowerSupplyBacklight.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.PowerSupplyDetachableCables.Length; i++)
                {
                    if (settings.PowerSupplyDetachableCables[i] == ps.DetachableCables)
                    {
                        break;
                    }
                    if (i == settings.PowerSupplyDetachableCables.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.PowerSupplyWireBraiding.Length; i++)
                {
                    if (settings.PowerSupplyWireBraiding[i] == ps.WireBraiding)
                    {
                        break;
                    }
                    if (i == settings.PowerSupplyWireBraiding.Length - 1)
                    {
                        return false;
                    }
                }
                return true;
            }

            if (component is RAM)
            {
                var ram = component as RAM;
                for (int i = 0; i < settings.RAMBacklight.Length; i++)
                {
                    if (settings.RAMBacklight[i] == ram.Backlight)
                    {
                        break;
                    }
                    if (i == settings.RAMBacklight.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.RAMFabricator.Length; i++)
                {
                    if (settings.RAMFabricator[i] == ram.Fabricator)
                    {
                        break;
                    }
                    if (i == settings.RAMFabricator.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.RAMMemory.Length; i++)
                {
                    if (settings.RAMMemory[i] == ram.Memory)
                    {
                        break;
                    }
                    if (i == settings.RAMMemory.Length - 1)
                    {
                        return false;
                    }
                }
                for (int i = 0; i < settings.RAMMemoryType.Length; i++)
                {
                    if (settings.RAMMemoryType[i] == ram.MemoryType)
                    {
                        break;
                    }
                    if (i == settings.RAMMemoryType.Length - 1)
                    {
                        return false;
                    }
                }
                return true;
            }

            if (component is SSD)
            {
                var ssd = component as SSD;

                for (int i = 0; i < settings.SSDFabricator.Length; i++)
                {
                    if (settings.SSDFabricator[i] == ssd.Fabricator)
                    {
                        break;
                    }
                    if (i == settings.SSDFabricator.Length - 1)
                    {
                        return false;
                    }
                }
                if ((settings.SSDMemory[0] > ssd.Memory) | (settings.SSDMemory[1] < ssd.Memory))
                {
                    return false;
                }
                if ((settings.SSDReadSpeed[0] > ssd.ReadSpeed) | (settings.SSDReadSpeed[1] < ssd.ReadSpeed))
                {
                    return false;
                }
                if ((settings.SSDWriteSpeed[0] > ssd.WriteSpeed) | (settings.SSDWriteSpeed[1] < ssd.WriteSpeed))
                {
                    return false;
                }

                return true;
            }

            else { throw new Exception("Непонятная штука."); }
        }

        private void VariantsComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            MBlabel.Text = computers[VariantsComboBox.SelectedIndex].motherBoard.Name;
            MBsitelabel.Text = computers[VariantsComboBox.SelectedIndex].motherBoard.Site;
            if (computers[VariantsComboBox.SelectedIndex].processor.Name == "NULL")
            {
                CPUlabel.Text = "Не требуется";
            }
            else
            {
                CPUlabel.Text = computers[VariantsComboBox.SelectedIndex].processor.Name;
                CPUsitelabel.Text = computers[VariantsComboBox.SelectedIndex].processor.Site;
            }
            GClabel.Text = computers[VariantsComboBox.SelectedIndex].graphicsCard.Name;
            GCsitelabel.Text = computers[VariantsComboBox.SelectedIndex].graphicsCard.Site;
            RAMlabel.Text = computers[VariantsComboBox.SelectedIndex].ram.Name;
            RAMsitelabel.Text = computers[VariantsComboBox.SelectedIndex].ram.Site;
            PSlabel.Text = computers[VariantsComboBox.SelectedIndex].powerSupply.Name;
            PSsitelabel.Text = computers[VariantsComboBox.SelectedIndex].powerSupply.Site;
            Clabel.Text = computers[VariantsComboBox.SelectedIndex].corps.Name;
            Csitelabel.Text = computers[VariantsComboBox.SelectedIndex].corps.Site;
            if (computers[VariantsComboBox.SelectedIndex].hdd.Name == "NULL")
            {
                HDDlabel.Text = "Не требуется";
            }
            else
            {
                HDDlabel.Text = computers[VariantsComboBox.SelectedIndex].hdd.Name;
                HDDsitelabel.Text = computers[VariantsComboBox.SelectedIndex].hdd.Site;
            }
            if (computers[VariantsComboBox.SelectedIndex].ssd.Name == "NULL")
            {
                SSDlabel.Text = "Не требуется";                
            }
            else
            {
                SSDlabel.Text = computers[VariantsComboBox.SelectedIndex].ssd.Name;
                SSDsitelabel.Text = computers[VariantsComboBox.SelectedIndex].ssd.Site;
            }
            TotalPricelabel.Text = Convert.ToString(computers[VariantsComboBox.SelectedIndex].TotalPrice());
        }
    }
}
