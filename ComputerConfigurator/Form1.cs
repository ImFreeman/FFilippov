using System;
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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            AdvancedSettings.Visible = false;
            settings.ForGamingPC[0] = "Да";
            settings.ForGamingPC[1] = "Нет";
            settings.ProfCard[0] = "Да";
            settings.ProfCard[1] = "Нет";
            for (int i=0;i<32;i++)
            {
                if (i < 2)
                {
                    CPUCoresList.SetItemChecked(i, true);
                    GraphicsCardFabricatorList.SetItemChecked(i, true);
                    RAMFabricatorList.SetItemChecked(i, true);
                    RAMMemoryList.SetItemChecked(i, true);
                    PowerSupplyFabricatorList.SetItemChecked(i, true);
                    CorpsFabricatorList.SetItemChecked(i, true);
                    CorpsMainColorList.SetItemChecked(i, true);
                    CorpsBacklightList.SetItemChecked(i, true);
                    CorpsFrameSizeList.SetItemChecked(i, true);
                    SSDFabricatorList.SetItemChecked(i, true);
                    MotherBoardFabricatorList.SetItemChecked(i, true);
                    MotherBoardNumberOfM2SlotsList.SetItemChecked(i, true);
                    MotherBoardNumberOfPCIEx16SlotsList.SetItemChecked(i, true);
                    GraphicsCardMemoryList.SetItemChecked(i, true);
                    RAMMemoryTypeList.SetItemChecked(i, true);
                    HDDMemoryList.SetItemChecked(i, true);
                    MotherBoardMemoryTypeList.SetItemChecked(i, true);
                    GraphicsCardMemoryTypeList.SetItemChecked(i, true);
                    HDDBufferSizeList.SetItemChecked(i, true);
                    MotherBoardNumberOfMemorySlotsList.SetItemChecked(i, true);
                    GraphicsCardNumberOfMonitorsList.SetItemChecked(i, true);
                    GraphicsCardPCIExpressList.SetItemChecked(i, true);
                    HDDFabricatorList.SetItemChecked(i, true);
                    MotherBoardCPUTypeList.SetItemChecked(i, true);
                    GraphicsCardFabricatorOfGPUList.SetItemChecked(i, true);
                }
                else if (i < 3)
                {
                    CPUCoresList.SetItemChecked(i, true);
                    GraphicsCardFabricatorList.SetItemChecked(i, true);
                    RAMFabricatorList.SetItemChecked(i, true);
                    RAMMemoryList.SetItemChecked(i, true);
                    PowerSupplyFabricatorList.SetItemChecked(i, true);
                    CorpsFabricatorList.SetItemChecked(i, true);
                    CorpsMainColorList.SetItemChecked(i, true);
                    CorpsBacklightList.SetItemChecked(i, true);
                    CorpsFrameSizeList.SetItemChecked(i, true);
                    SSDFabricatorList.SetItemChecked(i, true);
                    MotherBoardFabricatorList.SetItemChecked(i, true);
                    MotherBoardNumberOfM2SlotsList.SetItemChecked(i, true);
                    MotherBoardNumberOfPCIEx16SlotsList.SetItemChecked(i, true);
                    GraphicsCardMemoryList.SetItemChecked(i, true);
                    RAMMemoryTypeList.SetItemChecked(i, true);
                    HDDMemoryList.SetItemChecked(i, true);
                    MotherBoardMemoryTypeList.SetItemChecked(i, true);
                    GraphicsCardMemoryTypeList.SetItemChecked(i, true);
                    HDDBufferSizeList.SetItemChecked(i, true);
                    MotherBoardNumberOfMemorySlotsList.SetItemChecked(i, true);
                    GraphicsCardNumberOfMonitorsList.SetItemChecked(i, true);
                    GraphicsCardPCIExpressList.SetItemChecked(i, true);
                    HDDFabricatorList.SetItemChecked(i, true);
                }
                else if (i < 4)
                {
                    CPUCoresList.SetItemChecked(i, true);
                    GraphicsCardFabricatorList.SetItemChecked(i, true);
                    RAMFabricatorList.SetItemChecked(i, true);
                    RAMMemoryList.SetItemChecked(i, true);
                    PowerSupplyFabricatorList.SetItemChecked(i, true);
                    CorpsFabricatorList.SetItemChecked(i, true);
                    CorpsMainColorList.SetItemChecked(i, true);
                    CorpsBacklightList.SetItemChecked(i, true);
                    CorpsFrameSizeList.SetItemChecked(i, true);
                    SSDFabricatorList.SetItemChecked(i, true);
                    MotherBoardFabricatorList.SetItemChecked(i, true);
                    MotherBoardNumberOfM2SlotsList.SetItemChecked(i, true);
                    MotherBoardNumberOfPCIEx16SlotsList.SetItemChecked(i, true);
                    GraphicsCardMemoryList.SetItemChecked(i, true);
                    RAMMemoryTypeList.SetItemChecked(i, true);
                    HDDMemoryList.SetItemChecked(i, true);
                    MotherBoardMemoryTypeList.SetItemChecked(i, true);
                    GraphicsCardMemoryTypeList.SetItemChecked(i, true);
                    HDDBufferSizeList.SetItemChecked(i, true);
                }
                else if (i < 5)
                {
                    CPUCoresList.SetItemChecked(i, true);
                    GraphicsCardFabricatorList.SetItemChecked(i, true);
                    RAMFabricatorList.SetItemChecked(i, true);
                    RAMMemoryList.SetItemChecked(i, true);
                    PowerSupplyFabricatorList.SetItemChecked(i, true);
                    CorpsFabricatorList.SetItemChecked(i, true);
                    CorpsMainColorList.SetItemChecked(i, true);
                    CorpsBacklightList.SetItemChecked(i, true);
                    CorpsFrameSizeList.SetItemChecked(i, true);
                    SSDFabricatorList.SetItemChecked(i, true);
                    MotherBoardFabricatorList.SetItemChecked(i, true);
                    MotherBoardNumberOfM2SlotsList.SetItemChecked(i, true);
                    MotherBoardNumberOfPCIEx16SlotsList.SetItemChecked(i, true);
                    GraphicsCardMemoryList.SetItemChecked(i, true);
                    RAMMemoryTypeList.SetItemChecked(i, true);
                    HDDMemoryList.SetItemChecked(i, true);

                }
                else if (i < 6)
                {
                    CPUCoresList.SetItemChecked(i, true);
                    GraphicsCardFabricatorList.SetItemChecked(i, true);
                    RAMFabricatorList.SetItemChecked(i, true);
                    RAMMemoryList.SetItemChecked(i, true);
                    PowerSupplyFabricatorList.SetItemChecked(i, true);
                    CorpsFabricatorList.SetItemChecked(i, true);
                    CorpsMainColorList.SetItemChecked(i, true);
                    CorpsBacklightList.SetItemChecked(i, true);
                    CorpsFrameSizeList.SetItemChecked(i, true);
                    SSDFabricatorList.SetItemChecked(i, true);
                    MotherBoardFabricatorList.SetItemChecked(i, true);
                    MotherBoardNumberOfM2SlotsList.SetItemChecked(i, true);
                    MotherBoardNumberOfPCIEx16SlotsList.SetItemChecked(i, true);
                    GraphicsCardMemoryList.SetItemChecked(i, true);

                }
                else if (i<7)
                {
                    CPUCoresList.SetItemChecked(i, true);
                    GraphicsCardFabricatorList.SetItemChecked(i, true);
                    RAMFabricatorList.SetItemChecked(i, true);
                    RAMMemoryList.SetItemChecked(i, true);
                    PowerSupplyFabricatorList.SetItemChecked(i, true);
                    CorpsFabricatorList.SetItemChecked(i, true);
                    CorpsMainColorList.SetItemChecked(i, true);
                    CorpsBacklightList.SetItemChecked(i, true);
                    CorpsFrameSizeList.SetItemChecked(i, true);
                    SSDFabricatorList.SetItemChecked(i, true);


                }
                else if(i<8)
                {
                    GraphicsCardFabricatorList.SetItemChecked(i, true);
                    RAMFabricatorList.SetItemChecked(i, true);
                    RAMMemoryList.SetItemChecked(i, true);
                    PowerSupplyFabricatorList.SetItemChecked(i, true);
                    CorpsFabricatorList.SetItemChecked(i, true);
                    CorpsBacklightList.SetItemChecked(i, true);
                    CorpsFrameSizeList.SetItemChecked(i, true);
                    SSDFabricatorList.SetItemChecked(i, true);
                }
                else if (i < 9)
                {
                    GraphicsCardFabricatorList.SetItemChecked(i, true);
                    RAMFabricatorList.SetItemChecked(i, true);
                    PowerSupplyFabricatorList.SetItemChecked(i, true);
                    CorpsFabricatorList.SetItemChecked(i, true);
                    CorpsBacklightList.SetItemChecked(i, true);
                    CorpsFrameSizeList.SetItemChecked(i, true);
                    SSDFabricatorList.SetItemChecked(i, true);
                }
                else if(i<10)
                {
                    GraphicsCardFabricatorList.SetItemChecked(i, true);
                    RAMFabricatorList.SetItemChecked(i, true);
                    PowerSupplyFabricatorList.SetItemChecked(i, true);
                    CorpsFabricatorList.SetItemChecked(i, true);
                    CorpsBacklightList.SetItemChecked(i, true);
                    SSDFabricatorList.SetItemChecked(i, true);
                }
                else if(i<11)
                {
                    RAMFabricatorList.SetItemChecked(i, true);
                    PowerSupplyFabricatorList.SetItemChecked(i, true);
                    CorpsFabricatorList.SetItemChecked(i, true);
                    SSDFabricatorList.SetItemChecked(i, true);
                }
                else if (i < 21)
                {
                    RAMFabricatorList.SetItemChecked(i, true);
                    PowerSupplyFabricatorList.SetItemChecked(i, true);
                    CorpsFabricatorList.SetItemChecked(i, true);
                    SSDFabricatorList.SetItemChecked(i, true);
                }
                else if(i<25)
                {
                    PowerSupplyFabricatorList.SetItemChecked(i, true);
                    CorpsFabricatorList.SetItemChecked(i, true);
                    SSDFabricatorList.SetItemChecked(i, true);
                }
                else if(i<31)
                {
                    CorpsFabricatorList.SetItemChecked(i, true);
                    SSDFabricatorList.SetItemChecked(i, true);
                }
                else 
                {
                    SSDFabricatorList.SetItemChecked(i, true);
                }
            }

        }

        private Settings settings = new Settings();       
        
        private void AdvancedSettingsButton_Click(object sender, EventArgs e)
        {
            if(!AdvancedSettings.Visible)
            {
                AdvancedSettings.Visible = true;
            }
            else
            {
                AdvancedSettings.Visible = false;
            }
        }

        private void HDDcheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if(!HDDbox.Visible)
            {
                HDDbox.Visible = true;
            }
            else 
            {
                HDDbox.Visible = false;
                if(!SSDcheckBox.Checked)
                {
                    SSDcheckBox.Checked = true;
                }
            }
        }

        private void SSDcheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (!SSDbox.Visible)
            {
                SSDbox.Visible = true;
            }
            else
            {
                SSDbox.Visible = false;
                if (!HDDcheckBox.Checked)
                {
                    HDDcheckBox.Checked = true;
                }
            }
        }

        private void ListSet(string[] set, CheckedListBox listBox)
        {
            for(int i = 0;i<set.Length;i++)
            {
                if (listBox.GetItemChecked(i) == true) { set[i] = listBox.GetItemText(listBox.Items[i]); } else { set[i] = "NULL"; }
            }
        }

        private void SetSettings(Settings set)
        {
            set.Price[0] = int.Parse(PriceMin.Text);
            set.Price[1] = int.Parse(PriceMax.Text);
            //CPU
            //CPUFabricator
            if (CPUFabricatorAMD.Checked == true) { set.CPUFabricator[0] = "AMD"; } else { set.CPUFabricator[0] = "NULL"; }
            if (CPUFabricatorIntel.Checked == true) { set.CPUFabricator[1] = "Intel"; } else { set.CPUFabricator[1] = "NULL"; }
            //CPUCores
            ListSet(set.CPUCores, CPUCoresList);
            //CPUGraphicCore
            if (CPUGraphicCoreYes.Checked == true) { set.CPUGraphicCore[0] = "Да"; } else { set.CPUGraphicCore[0] = "NULL"; }
            if (CPUGraphicCoreNo.Checked == true) { set.CPUGraphicCore[1] = "Нет"; } else { set.CPUGraphicCore[1] = "NULL"; }
            //CPUMemoryType
            if (CPUMemoryTypeDDR3.Checked == true) { set.CPUMemoryType[0] = "DDR3"; } else { set.CPUMemoryType[0] = "NULL"; }
            if (CPUMemoryTypeDDR4.Checked == true) { set.CPUMemoryType[1] = "DDR4"; } else { set.CPUMemoryType[1] = "NULL"; }
            //CPUBaseFrequency
            set.CPUBaseFrequency[0] = int.Parse(CPUBaseFrequencyMin.Text);
            set.CPUBaseFrequency[1] = int.Parse(CPUBaseFrequencyMax.Text);
            //CPUMultithreading
            if (CPUMultithreadingYes.Checked == true) { set.CPUMultithreading[0] = "Да"; } else { set.CPUMultithreading[0] = "NULL"; }
            if (CPUMultithreadingNo.Checked == true) { set.CPUMultithreading[1] = "Нет"; } else { set.CPUMultithreading[1] = "NULL"; }

            //MotherBoard
            //MotherBoardFabricator
            ListSet(set.MotherBoardFabricator, MotherBoardFabricatorList);
            //MotherBoardCPUType
            ListSet(set.MotherBoardCPUType, MotherBoardCPUTypeList);
            //MotherBoardMemoryType
            ListSet(set.MotherBoardMemoryType, MotherBoardMemoryTypeList);
            //MotherBoardNumberOfMemorySlots
            ListSet(set.MotherBoardNumberOfMemorySlots, MotherBoardNumberOfMemorySlotsList);
            //MotherBoardNumberOfPCIEx16Slots
            ListSet(set.MotherBoardNumberOfPCIEx16Slots, MotherBoardNumberOfPCIEx16SlotsList);
            //MotherBoardNumberOfM2Slots
            ListSet(set.MotherBoardNumberOfM2Slots, MotherBoardNumberOfM2SlotsList);
            //MotherBoardWiFiAdapter
            if (MotherBoardWiFiAdapterYes.Checked == true) { set.MotherBoardWiFiAdapter[0] = "Да"; } else { set.MotherBoardWiFiAdapter[0] = "NULL"; }
            if (MotherBoardWiFiAdapterNo.Checked == true) { set.MotherBoardWiFiAdapter[1] = "Нет"; } else { set.MotherBoardWiFiAdapter[1] = "NULL"; }
            //MotherBoardBuildInCPU
            if (CPUcheckBox.Checked) { set.MotherBoardBuildInCPU[0] = "Да"; set.MotherBoardBuildInCPU[1] = "NULL"; } else { set.MotherBoardBuildInCPU[0] = "NULL"; set.MotherBoardBuildInCPU[1] = "Нет"; }
            /*if (MotherBoardBuildInCPUYes.Checked == true) { set.MotherBoardBuildInCPU[0] = "Да"; } else { set.MotherBoardBuildInCPU[0] = "NULL"; }
            if (MotherBoardBuildInCPUNo.Checked == true) { set.MotherBoardBuildInCPU[1] = "Нет"; } else { set.MotherBoardBuildInCPU[1] = "NULL"; }*/

            //GraphicsCard
            //GraphicsCardFabricator
            ListSet(set.GraphicsCardFabricator, GraphicsCardFabricatorList);
            //GraphicsCardMemory
            ListSet(set.GraphicsCardMemory, GraphicsCardMemoryList);
            //GraphicsCardMemoryType
            ListSet(set.GraphicsCardMemoryType, GraphicsCardMemoryTypeList);
            //GraphicsCardFabricatorOfGPU
            ListSet(set.GraphicsCardFabricatorOfGPU, GraphicsCardFabricatorOfGPUList);
            //GraphicsCardNumberOfMonitors
            ListSet(set.GraphicsCardNumberOfMonitors, GraphicsCardNumberOfMonitorsList);
            //GraphicsCardPCIExpress
            ListSet(set.GraphicsCardPCIExpress, GraphicsCardPCIExpressList);
            //GraphicsCardMemoryBusWidth
            set.GraphicsCardMemoryBusWidth[0] = int.Parse(GraphicsCardMemoryBusWidthMin.Text);
            set.GraphicsCardMemoryBusWidth[1] = int.Parse(GraphicsCardMemoryBusWidthMax.Text);

            //RAM
            //RAMFabricator
            ListSet(set.RAMFabricator, RAMFabricatorList);
            //RAMBacklight
            if (RAMBacklightYes.Checked == true) { set.RAMBacklight[0] = "Да"; } else { set.RAMBacklight[0] = "NULL"; }
            if (RAMBacklightNo.Checked == true) { set.RAMBacklight[1] = "Нет"; } else { set.RAMBacklight[1] = "NULL"; }
            //RAMMemory
            ListSet(set.RAMMemory, RAMMemoryList);
            //RAMMemoryType
            ListSet(set.RAMMemoryType, RAMMemoryTypeList);

            //PowerSupply
            //PowerSupplyFabricator
            ListSet(set.PowerSupplyFabricator, PowerSupplyFabricatorList);
            //PowerSupplyWireBraiding
            if (PowerSupplyWireBraidingYes.Checked == true) { set.PowerSupplyWireBraiding[0] = "Да"; } else { set.PowerSupplyWireBraiding[0] = "NULL"; }
            if (PowerSupplyWireBraidingNo.Checked == true) { set.PowerSupplyWireBraiding[1] = "Нет"; } else { set.PowerSupplyWireBraiding[1] = "NULL"; }
            //PowerSupplyBacklight
            if (PowerSupplyBacklightYes.Checked == true) { set.PowerSupplyBacklight[0] = "Да"; } else { set.PowerSupplyBacklight[0] = "NULL"; }
            if (PowerSupplyBacklightNo.Checked == true) { set.PowerSupplyBacklight[1] = "Нет"; } else { set.PowerSupplyBacklight[1] = "NULL"; }
            //PowerSupplyDetachableCables
            if (PowerSupplyDetachableCablesYes.Checked == true) { set.PowerSupplyDetachableCables[0] = "Да"; } else { set.PowerSupplyDetachableCables[0] = "NULL"; }
            if (PowerSupplyDetachableCablesNo.Checked == true) { set.PowerSupplyDetachableCables[1] = "Нет"; } else { set.PowerSupplyDetachableCables[1] = "NULL"; }

            //Corps
            //CorpsFabricator
            ListSet(set.CorpsFabricator, CorpsFabricatorList);
            //CorpsWindow
            if (CorpsWindowYes.Checked == true) { set.CorpsWindow[0] = "Да"; } else { set.CorpsWindow[0] = "NULL"; }
            if (CorpsWindowNo.Checked == true) { set.CorpsWindow[1] = "Нет"; } else { set.CorpsWindow[1] = "NULL"; }
            //CorpsMainColor
            ListSet(set.CorpsMainColor, CorpsMainColorList);
            //CorpsBacklight
            ListSet(set.CorpsBacklight, CorpsBacklightList);
            //CorpsFrameSize
            ListSet(set.CorpsFrameSize, CorpsFrameSizeList);

            if(HDDcheckBox.Checked)
            {
                set.HDDRequired = true;
                //HDD
                //HDDMemory
                ListSet(set.HDDMemory, HDDMemoryList);
                //HDDLevelOfNoise
                set.HDDLevelOfNoise[0] = int.Parse(HDDLevelOfNoiseMin.Text);
                set.HDDLevelOfNoise[1] = int.Parse(HDDLevelOfNoiseMax.Text);
                //HDDDataExchangeRate
                set.HDDDataExchangeRate[0] = int.Parse(HDDDataExchangeRateMin.Text);
                set.HDDDataExchangeRate[1] = int.Parse(HDDDataExchangeRateMax.Text);
                //HDDFabricator
                ListSet(set.HDDFabricator, HDDFabricatorList);
                //HDDBufferSize
                ListSet(set.HDDBufferSize, HDDBufferSizeList);
            }

            if (SSDcheckBox.Checked)
            {
                set.SSDRequired = true;
                //SSD
                //SSDFabricator
                ListSet(set.SSDFabricator, SSDFabricatorList);
                //SSDMemory
                set.SSDMemory[0] = int.Parse(SSDMemoryMin.Text);
                set.SSDMemory[1] = int.Parse(SSDMemoryMax.Text);
                //SSDWriteSpeed
                set.SSDWriteSpeed[0] = int.Parse(SSDWriteSpeedMin.Text);
                set.SSDWriteSpeed[1] = int.Parse(SSDWriteSpeedMax.Text);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            SetSettings(settings);
            Form2 form2 = new Form2(settings);
            this.Hide();
            form2.ShowDialog();
            Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            settings.ForGamingPC[0] = "Да";
            //settings.ForGamingPC[0] = "NULL";
        }

        private void CPUcheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if(CPUBox.Visible)
            {
                CPUBox.Visible = false;
            }
            else
            {
                CPUBox.Visible = true;
            }
        }
    }
}
