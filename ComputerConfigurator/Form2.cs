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
    public partial class Form2 : Form
    {
        public Form2(Settings settings)
        {
            InitializeComponent();

            // Получить объект приложения Excel.
            Excel.Application excel_app = new Excel.Application();
            Excel.Workbook workbook = excel_app.Workbooks.Open(
                Path.GetFullPath("../../Resources/CPU"),
                Type.Missing, true, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            // Получить первый рабочий лист.
            Excel.Worksheet sheet = (Excel.Worksheet)workbook.Sheets[1];

            Excel.Range range;

            CPU cpu = new CPU();

            range = (Excel.Range)sheet.Cells[2, 1];
            cpu.Name = range.Value2;
            range = (Excel.Range)sheet.Cells[2, 2];
            cpu.Price = range.Value2;
            range = (Excel.Range)sheet.Cells[2, 3];
            cpu.Fabricator = range.Value2;
            range = (Excel.Range)sheet.Cells[2, 4];
            cpu.NumberOfCores = Convert.ToString(range.Value2);
            range = (Excel.Range)sheet.Cells[2, 5];
            cpu.GraphicCore = range.Value2;
            range = (Excel.Range)sheet.Cells[2, 6];
            cpu.MemoryType = range.Value2;
            range = (Excel.Range)sheet.Cells[2, 7];
            cpu.BaseFrequency = Convert.ToInt32(range.Value2);
            range = (Excel.Range)sheet.Cells[2, 8];
            cpu.Multithreading = range.Value2;
            range = (Excel.Range)sheet.Cells[2, 9];
            cpu.Socket = range.Value2;
            range = (Excel.Range)sheet.Cells[2, 10];
            cpu.ForGamingPC= range.Value2;
            range = (Excel.Range)sheet.Cells[2, 11];
            cpu.Site = range.Value2;

            if(SettingsCheck(cpu,settings))
            {
                label1.Text = "Работает";
            }
            /*
            int i = 2;
            string x=" ";
            while(x!="end")
            {
                range = (Excel.Range)sheet.Cells[i, 1];

            }*/



        }

        private PC computer = new PC();
        //private List<Computer> computers;

        private void example()
        {
            // Получить объект приложения Excel.
            Excel.Application excel_app = new Excel.Application();

            // Сделать Excel видимым (необязательно).
            //excel_app.Visible = true;

            // Откройте рабочую книгу только для чтения.
            Excel.Workbook workbook = excel_app.Workbooks.Open(
                Path.GetFullPath("../../Resources/CPU"),
                Type.Missing, true, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            // Получить первый рабочий лист.
            Excel.Worksheet sheet = (Excel.Worksheet)workbook.Sheets[1];

            // Получить заголовки и значения.
            Excel.Range range;
            range = (Excel.Range)sheet.Cells[2, 2];
            double A= range.Value2;
            label1.Text = Convert.ToString(A);
            //label1.Text = (string)range.Value2;

            // Закройте книгу без сохранения изменений.
            workbook.Close(false, Type.Missing, Type.Missing);

            // Закройте сервер Excel.
            excel_app.Quit();
        }

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

    }
}
