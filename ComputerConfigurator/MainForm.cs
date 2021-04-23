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
    public partial class MainForm : Form
    {
        public MainForm()
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

        /// <summary>
        /// Настройки пользователя
        /// </summary>
        private Settings settings = new Settings();       
        
        /// <summary>
        /// Нажатие кнопки "Расширенные настройки"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        /// <summary>
        /// Изменение положения чекбокса HDD
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        /// <summary>
        /// Изменение положения чекбокса SSD
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        /// <summary>
        /// Сохранение выделенных значений в CheckedListBox
        /// </summary>
        /// <param name="set"></param>
        /// <param name="listBox"></param>
        private void ListSet(string[] set, CheckedListBox listBox)
        {
            for(int i = 0;i<set.Length;i++)
            {
                if (listBox.GetItemChecked(i) == true) { set[i] = listBox.GetItemText(listBox.Items[i]); } else { set[i] = "NULL"; }
            }
        }

        /// <summary>
        /// Сохранение настроек пользователя
        /// </summary>
        /// <param name="set"></param>
        private void SetSettings(Settings set)
        {            
            set.Price[0] = int.Parse(PriceMin.Text);
            set.Price[1] = int.Parse(PriceMax.Text);
            set.NumberOfPCs = int.Parse(NumberOfPCcTextBox.Text);
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
            if (CPUcheckBox.Checked) { set.MotherBoardBuildInCPU[0] = "Нет"; set.MotherBoardBuildInCPU[1] = "NULL"; } else { set.MotherBoardBuildInCPU[0] = "NULL"; set.MotherBoardBuildInCPU[1] = "Есть"; }
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
                //SSDReadSpeed
                set.SSDReadSpeed[0] = int.Parse(SSDReadSpeedMin.Text);
                set.SSDReadSpeed[1] = int.Parse(SSDReadSpeedMax.Text);
            }
        }
        
        

        /// <summary>
        /// Изменение положения чекбокса CPU
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CPUcheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if(CPUBox.Visible)
            {
                CPUBox.Visible = false;
                MotherBoardBox.Visible = false;
            }
            else
            {
                CPUBox.Visible = true;
                MotherBoardBox.Visible = true;
            }
        }

        /// <summary>
        /// Открытие следующей формы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void NextFormButton_Click(object sender, EventArgs e)
        {
            SetSettings(settings);
            ResultForm form2 = new ResultForm(settings);
            this.Hide();
            form2.ShowDialog();
            Close();
        }

        /// <summary>
        /// Нажата кнопка "Офис"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OfficeButton_Click(object sender, EventArgs e)
        {
            OfficeButton.BackColor = SystemColors.ActiveCaption;
            GamesButton.BackColor = SystemColors.Control;
            HomeButton.BackColor = SystemColors.Control;
            GraphicsButton.BackColor = SystemColors.Control;
            DevButton.BackColor = SystemColors.Control;
            AnyButton.BackColor = SystemColors.Control;

            settings.ForGamingPC[0] = "NULL";
            settings.ForGamingPC[1] = "Нет";
            settings.ProfCard[0] = "NULL";
            settings.ProfCard[1] = "Нет";

            CPUcheckBox.Checked = false;
            HDDcheckBox.Checked = true;
            SSDcheckBox.Checked = false;

            GraphicsCardMemoryList.SetItemChecked(0, true);
            for(int i=0;i<10;i++)
            {
                GraphicsCardFabricatorList.SetItemChecked(i, true);
            }
            GraphicsCardFabricatorOfGPUList.SetItemChecked(0, true);
            GraphicsCardFabricatorOfGPUList.SetItemChecked(1, true);
            for(int i =0;i<4;i++)
            {
                GraphicsCardMemoryTypeList.SetItemChecked(i, true);
            }
            GraphicsCardPCIExpressList.SetItemChecked(0, true);
            GraphicsCardPCIExpressList.SetItemChecked(1, true);
            GraphicsCardPCIExpressList.SetItemChecked(2, true);
            for (int i= 1; i < 6;i++)
            {
                GraphicsCardMemoryList.SetItemChecked(i, false);
            }
            GraphicsCardNumberOfMonitorsList.SetItemChecked(0, true);
            GraphicsCardNumberOfMonitorsList.SetItemChecked(1, false);
            GraphicsCardNumberOfMonitorsList.SetItemChecked(2, false);
            GraphicsCardMemoryBusWidthMax.Text = "64";
            GraphicsCardMemoryBusWidthMin.Text = "32";

            RAMBacklightNo.Checked = true;
            RAMBacklightYes.Checked = false;
            RAMMemoryList.SetItemChecked(0, true);
            RAMMemoryList.SetItemChecked(1, true);
            RAMMemoryList.SetItemChecked(2, true);
            for(int i=3;i<8; i++)
            {
                RAMMemoryList.SetItemChecked(i, false);
            }
            for(int i=0;i<5;i++)
            {
                RAMMemoryTypeList.SetItemChecked(i, true);
            }
            for(int i=0;i<21;i++)
            {
                RAMFabricatorList.SetItemChecked(i, true);
            }

            PowerSupplyBacklightNo.Checked = true;
            PowerSupplyBacklightYes.Checked = false;
            PowerSupplyDetachableCablesNo.Checked = true;
            PowerSupplyDetachableCablesYes.Checked = false;
            PowerSupplyWireBraidingNo.Checked = true;
            PowerSupplyWireBraidingYes.Checked = false;
            for(int i=0;i<25;i++)
            {
                PowerSupplyFabricatorList.SetItemChecked(i, true);
            }

            CorpsWindowNo.Checked = true;
            CorpsWindowYes.Checked = false;
            CorpsBacklightList.SetItemChecked(1, true);
            for(int i = 0;i<31;i++)
            {
                CorpsFabricatorList.SetItemChecked(i, true);
            }
            for(int i = 0; i<7;i++)
            {
                CorpsMainColorList.SetItemChecked(i, true);
            }
            for(int i=0;i<10;i++)
            {
                if (i != 7)
                {
                    CorpsBacklightList.SetItemChecked(i, false);
                }
                else
                {
                    CorpsBacklightList.SetItemChecked(i, true);
                }
            }
            for (int i = 0; i < 9; i++)
            {
                if (i != 4)
                {
                    CorpsFrameSizeList.SetItemChecked(i, false);
                }
                else
                {
                    CorpsFrameSizeList.SetItemChecked(i, true);
                }
            }

            HDDLevelOfNoiseMin.Text = "21";
            HDDLevelOfNoiseMax.Text = "27";
            HDDDataExchangeRateMin.Text = "126";
            HDDDataExchangeRateMax.Text = "160";
            for(int i=0;i<4;i++)
            {
                HDDBufferSizeList.SetItemChecked(i, true);
            }
            for(int i=0;i<3;i++)
            {
                HDDFabricatorList.SetItemChecked(i, true);
            }
            HDDMemoryList.SetItemChecked(0, true);
            HDDMemoryList.SetItemChecked(1, true);
            HDDMemoryList.SetItemChecked(2, false);
            HDDMemoryList.SetItemChecked(3, false);
            HDDMemoryList.SetItemChecked(4, false);

        }

        /// <summary>
        /// Нажата кнопка "Видеоигры"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GamesButton_Click(object sender, EventArgs e)
        {
            OfficeButton.BackColor = SystemColors.Control;
            GamesButton.BackColor = SystemColors.ActiveCaption;
            HomeButton.BackColor = SystemColors.Control;
            GraphicsButton.BackColor = SystemColors.Control;
            DevButton.BackColor = SystemColors.Control;
            AnyButton.BackColor = SystemColors.Control;

            settings.ForGamingPC[0] = "Да";
            settings.ForGamingPC[1] = "NULL";
            settings.ProfCard[0] = "NULL";
            settings.ProfCard[1] = "Нет";
            HDDcheckBox.Checked = true;
            SSDcheckBox.Checked = true;

            CPUcheckBox.Checked = true;
            CPUFabricatorAMD.Checked = true;
            CPUFabricatorIntel.Checked = true;
            CPUGraphicCoreYes.Checked = true;
            CPUGraphicCoreNo.Checked = true;
            CPUCoresList.SetItemChecked(0, false);
            for(int i=1;i<7;i++)
            {
                CPUCoresList.SetItemChecked(i, true);
            }
            CPUMultithreadingNo.Checked = false;
            CPUMultithreadingYes.Checked = true;
            CPUMemoryTypeDDR3.Checked = true;
            CPUMemoryTypeDDR4.Checked = true;
            CPUBaseFrequencyMin.Text = "3000";
            CPUBaseFrequencyMax.Text = "4200";

            for(int i=0;i<6;i++)
            {
                MotherBoardFabricatorList.SetItemChecked(i, true);
            }
            MotherBoardCPUTypeList.SetItemChecked(0, true);
            MotherBoardCPUTypeList.SetItemChecked(1, true);
            for(int i=0;i<3;i++)
            {
                MotherBoardNumberOfMemorySlotsList.SetItemChecked(i, true);
                MotherBoardMemoryTypeList.SetItemChecked(i, false);
            }
            MotherBoardMemoryTypeList.SetItemChecked(3, true);
            for(int i=0;i<6;i++)
            {
                MotherBoardNumberOfPCIEx16SlotsList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 6; i++)
            {
                MotherBoardNumberOfM2SlotsList.SetItemChecked(i, true);
            }
            MotherBoardWiFiAdapterNo.Checked = true;
            MotherBoardWiFiAdapterYes.Checked = false;

            GraphicsCardMemoryList.SetItemChecked(0, false);
            GraphicsCardMemoryList.SetItemChecked(1, false);
            for (int i = 2; i < 6; i++)
            {
                GraphicsCardMemoryList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 10; i++)
            {
                GraphicsCardFabricatorList.SetItemChecked(i, true);
            }
            GraphicsCardFabricatorOfGPUList.SetItemChecked(0, true);
            GraphicsCardFabricatorOfGPUList.SetItemChecked(1, true);
            for (int i = 0; i < 4; i++)
            {
                GraphicsCardMemoryTypeList.SetItemChecked(i, true);
            }
            GraphicsCardPCIExpressList.SetItemChecked(0, true);
            GraphicsCardPCIExpressList.SetItemChecked(1, true);
            GraphicsCardPCIExpressList.SetItemChecked(2, true);            
            GraphicsCardNumberOfMonitorsList.SetItemChecked(0, true);
            GraphicsCardNumberOfMonitorsList.SetItemChecked(1, true);
            GraphicsCardNumberOfMonitorsList.SetItemChecked(2, true);
            GraphicsCardMemoryBusWidthMax.Text = "384";
            GraphicsCardMemoryBusWidthMin.Text = "32";

            RAMBacklightNo.Checked = true;
            RAMBacklightYes.Checked = true;
            RAMMemoryList.SetItemChecked(0, false);
            RAMMemoryList.SetItemChecked(1, false);
            RAMMemoryList.SetItemChecked(2, false);
            for (int i = 3; i < 8; i++)
            {
                RAMMemoryList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 4; i++)
            {
                RAMMemoryTypeList.SetItemChecked(i, false);
            }
            RAMMemoryTypeList.SetItemChecked(4, true);
            for (int i = 0; i < 21; i++)
            {
                RAMFabricatorList.SetItemChecked(i, true);
            }

            PowerSupplyBacklightNo.Checked = true;
            PowerSupplyBacklightYes.Checked = true;
            PowerSupplyDetachableCablesNo.Checked = true;
            PowerSupplyDetachableCablesYes.Checked = true;
            PowerSupplyWireBraidingNo.Checked = true;
            PowerSupplyWireBraidingYes.Checked = true;
            for (int i = 0; i < 25; i++)
            {
                PowerSupplyFabricatorList.SetItemChecked(i, true);
            }

            CorpsWindowNo.Checked = true;
            CorpsWindowYes.Checked = true;            
            for (int i = 0; i < 31; i++)
            {
                CorpsFabricatorList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 7; i++)
            {
                CorpsMainColorList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 10; i++)
            {
                CorpsBacklightList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 9; i++)
            {
                if (i != 4)
                {
                    CorpsFrameSizeList.SetItemChecked(i, true);
                }
                else
                {
                    CorpsFrameSizeList.SetItemChecked(i, false);
                }
            }

            HDDLevelOfNoiseMin.Text = "21";
            HDDLevelOfNoiseMax.Text = "27";
            HDDDataExchangeRateMin.Text = "126";
            HDDDataExchangeRateMax.Text = "160";
            for (int i = 0; i < 4; i++)
            {
                HDDBufferSizeList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 3; i++)
            {
                HDDFabricatorList.SetItemChecked(i, true);
            }
            HDDMemoryList.SetItemChecked(0, false);
            HDDMemoryList.SetItemChecked(1, true);
            HDDMemoryList.SetItemChecked(2, true);
            HDDMemoryList.SetItemChecked(3, true);
            HDDMemoryList.SetItemChecked(4, true);

            for(int i=0;i<32;i++)
            {
                SSDFabricatorList.SetItemChecked(i, true);
            }
            SSDMemoryMin.Text = "120";
            SSDMemoryMax.Text = "7680";
            SSDWriteSpeedMin.Text = "290";
            SSDWriteSpeedMax.Text = "540";
            SSDReadSpeedMin.Text = "450";
            SSDReadSpeedMax.Text = "3480";

        }

        /// <summary>
        /// Нажата кнопка "Домашний досуг"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void HomeButton_Click(object sender, EventArgs e)
        {
            OfficeButton.BackColor = SystemColors.Control;
            GamesButton.BackColor = SystemColors.Control;
            HomeButton.BackColor = SystemColors.ActiveCaption;
            GraphicsButton.BackColor = SystemColors.Control;
            DevButton.BackColor = SystemColors.Control;
            AnyButton.BackColor = SystemColors.Control;

            settings.ForGamingPC[0] = "NULL";
            settings.ForGamingPC[1] = "Нет";
            settings.ProfCard[0] = "NULL";
            settings.ProfCard[1] = "Нет";
            HDDcheckBox.Checked = true;
            SSDcheckBox.Checked = false;

            CPUcheckBox.Checked = true;
            CPUFabricatorAMD.Checked = true;
            CPUFabricatorIntel.Checked = true;
            CPUGraphicCoreYes.Checked = true;
            CPUGraphicCoreNo.Checked = true;
            CPUCoresList.SetItemChecked(0, true);
            CPUCoresList.SetItemChecked(1, true);
            CPUCoresList.SetItemChecked(2, true);
            for (int i = 3; i < 7; i++)
            {
                CPUCoresList.SetItemChecked(i, false);
            }
            CPUMultithreadingNo.Checked = true;
            CPUMultithreadingYes.Checked = true;
            CPUMemoryTypeDDR3.Checked = true;
            CPUMemoryTypeDDR4.Checked = true;
            CPUBaseFrequencyMin.Text = "2500";
            CPUBaseFrequencyMax.Text = "4200";

            for (int i = 0; i < 6; i++)
            {
                MotherBoardFabricatorList.SetItemChecked(i, true);
            }
            MotherBoardCPUTypeList.SetItemChecked(0, true);
            MotherBoardCPUTypeList.SetItemChecked(1, true);
            for (int i = 0; i < 3; i++)
            {
                MotherBoardNumberOfMemorySlotsList.SetItemChecked(i, true);
                MotherBoardMemoryTypeList.SetItemChecked(i, true);
            }
            MotherBoardNumberOfMemorySlotsList.SetItemChecked(2, false);
            MotherBoardMemoryTypeList.SetItemChecked(3, true);
            for (int i = 0; i < 6; i++)
            {
                MotherBoardNumberOfPCIEx16SlotsList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 6; i++)
            {
                MotherBoardNumberOfM2SlotsList.SetItemChecked(i, true);
            }
            MotherBoardWiFiAdapterNo.Checked = true;
            MotherBoardWiFiAdapterYes.Checked = true;

            GraphicsCardMemoryList.SetItemChecked(0, true);
            GraphicsCardMemoryList.SetItemChecked(1, true);
            GraphicsCardMemoryList.SetItemChecked(2, true);
            for (int i = 3; i < 6; i++)
            {
                GraphicsCardMemoryList.SetItemChecked(i, false);
            }
            for (int i = 0; i < 10; i++)
            {
                GraphicsCardFabricatorList.SetItemChecked(i, true);
            }
            GraphicsCardFabricatorOfGPUList.SetItemChecked(0, true);
            GraphicsCardFabricatorOfGPUList.SetItemChecked(1, true);
            for (int i = 0; i < 4; i++)
            {
                GraphicsCardMemoryTypeList.SetItemChecked(i, true);
            }
            GraphicsCardPCIExpressList.SetItemChecked(0, true);
            GraphicsCardPCIExpressList.SetItemChecked(1, true);
            GraphicsCardPCIExpressList.SetItemChecked(2, true);
            GraphicsCardNumberOfMonitorsList.SetItemChecked(0, true);
            GraphicsCardNumberOfMonitorsList.SetItemChecked(1, true);
            GraphicsCardNumberOfMonitorsList.SetItemChecked(2, false);
            GraphicsCardMemoryBusWidthMax.Text = "384";
            GraphicsCardMemoryBusWidthMin.Text = "32";

            RAMBacklightNo.Checked = true;
            RAMBacklightYes.Checked = false;
            RAMMemoryList.SetItemChecked(0, true);
            RAMMemoryList.SetItemChecked(1, true);
            RAMMemoryList.SetItemChecked(2, true);
            for (int i = 3; i < 8; i++)
            {
                RAMMemoryList.SetItemChecked(i, false);
            }
            for (int i = 0; i < 4; i++)
            {
                RAMMemoryTypeList.SetItemChecked(i, true);
            }
            RAMMemoryTypeList.SetItemChecked(4, true);
            for (int i = 0; i < 21; i++)
            {
                RAMFabricatorList.SetItemChecked(i, true);
            }

            PowerSupplyBacklightNo.Checked = true;
            PowerSupplyBacklightYes.Checked = false;
            PowerSupplyDetachableCablesNo.Checked = true;
            PowerSupplyDetachableCablesYes.Checked = false;
            PowerSupplyWireBraidingNo.Checked = true;
            PowerSupplyWireBraidingYes.Checked = false;
            for (int i = 0; i < 25; i++)
            {
                PowerSupplyFabricatorList.SetItemChecked(i, true);
            }

            CorpsWindowNo.Checked = true;
            CorpsWindowYes.Checked = false;
            for (int i = 0; i < 31; i++)
            {
                CorpsFabricatorList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 7; i++)
            {
                CorpsMainColorList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 10; i++)
            {
                if (i != 7)
                {
                    CorpsBacklightList.SetItemChecked(i, false);
                }
                else { CorpsBacklightList.SetItemChecked(i, true); }
            }
            for (int i = 0; i < 9; i++)
            {
                CorpsFrameSizeList.SetItemChecked(i, true);
            }

            HDDLevelOfNoiseMin.Text = "21";
            HDDLevelOfNoiseMax.Text = "27";
            HDDDataExchangeRateMin.Text = "126";
            HDDDataExchangeRateMax.Text = "160";
            for (int i = 0; i < 4; i++)
            {
                HDDBufferSizeList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 3; i++)
            {
                HDDFabricatorList.SetItemChecked(i, true);
            }
            HDDMemoryList.SetItemChecked(0, true);
            HDDMemoryList.SetItemChecked(1, true);
            HDDMemoryList.SetItemChecked(2, false);
            HDDMemoryList.SetItemChecked(3, false);
            HDDMemoryList.SetItemChecked(4, false);
        }

        /// <summary>
        /// Нажатие на кнопку "Любая сфера"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AnyButton_Click(object sender, EventArgs e)
        {
            OfficeButton.BackColor = SystemColors.Control;
            GamesButton.BackColor = SystemColors.Control;
            HomeButton.BackColor = SystemColors.Control;
            GraphicsButton.BackColor = SystemColors.Control;
            DevButton.BackColor = SystemColors.Control;
            AnyButton.BackColor = SystemColors.ActiveCaption;

            settings.ForGamingPC[0] = "Да";
            settings.ForGamingPC[1] = "Нет";
            settings.ProfCard[0] = "Да";
            settings.ProfCard[1] = "Нет";
            HDDcheckBox.Checked = true;
            SSDcheckBox.Checked = true;

            CPUcheckBox.Checked = true;
            CPUFabricatorAMD.Checked = true;
            CPUFabricatorIntel.Checked = true;
            CPUGraphicCoreYes.Checked = true;
            CPUGraphicCoreNo.Checked = true;
            CPUCoresList.SetItemChecked(0, true);
            for (int i = 1; i < 7; i++)
            {
                CPUCoresList.SetItemChecked(i, true);
            }
            CPUMultithreadingNo.Checked = true;
            CPUMultithreadingYes.Checked = true;
            CPUMemoryTypeDDR3.Checked = true;
            CPUMemoryTypeDDR4.Checked = true;
            CPUBaseFrequencyMin.Text = "2500";
            CPUBaseFrequencyMax.Text = "4200";

            for (int i = 0; i < 6; i++)
            {
                MotherBoardFabricatorList.SetItemChecked(i, true);
            }
            MotherBoardCPUTypeList.SetItemChecked(0, true);
            MotherBoardCPUTypeList.SetItemChecked(1, true);
            for (int i = 0; i < 3; i++)
            {
                MotherBoardNumberOfMemorySlotsList.SetItemChecked(i, true);
                MotherBoardMemoryTypeList.SetItemChecked(i, true);
            }
            MotherBoardMemoryTypeList.SetItemChecked(3, true);
            for (int i = 0; i < 6; i++)
            {
                MotherBoardNumberOfPCIEx16SlotsList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 6; i++)
            {
                MotherBoardNumberOfM2SlotsList.SetItemChecked(i, true);
            }
            MotherBoardWiFiAdapterNo.Checked = true;
            MotherBoardWiFiAdapterYes.Checked = true;

            GraphicsCardMemoryList.SetItemChecked(0, true);
            GraphicsCardMemoryList.SetItemChecked(1, true);
            for (int i = 2; i < 6; i++)
            {
                GraphicsCardMemoryList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 10; i++)
            {
                GraphicsCardFabricatorList.SetItemChecked(i, true);
            }
            GraphicsCardFabricatorOfGPUList.SetItemChecked(0, true);
            GraphicsCardFabricatorOfGPUList.SetItemChecked(1, true);
            for (int i = 0; i < 4; i++)
            {
                GraphicsCardMemoryTypeList.SetItemChecked(i, true);
            }
            GraphicsCardPCIExpressList.SetItemChecked(0, true);
            GraphicsCardPCIExpressList.SetItemChecked(1, true);
            GraphicsCardPCIExpressList.SetItemChecked(2, true);
            GraphicsCardNumberOfMonitorsList.SetItemChecked(0, true);
            GraphicsCardNumberOfMonitorsList.SetItemChecked(1, true);
            GraphicsCardNumberOfMonitorsList.SetItemChecked(2, true);
            GraphicsCardMemoryBusWidthMax.Text = "384";
            GraphicsCardMemoryBusWidthMin.Text = "32";

            RAMBacklightNo.Checked = true;
            RAMBacklightYes.Checked = true;
            RAMMemoryList.SetItemChecked(0, true);
            RAMMemoryList.SetItemChecked(1, true);
            RAMMemoryList.SetItemChecked(2, true);
            for (int i = 3; i < 8; i++)
            {
                RAMMemoryList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 4; i++)
            {
                RAMMemoryTypeList.SetItemChecked(i, true);
            }
            RAMMemoryTypeList.SetItemChecked(4, true);
            for (int i = 0; i < 21; i++)
            {
                RAMFabricatorList.SetItemChecked(i, true);
            }

            PowerSupplyBacklightNo.Checked = true;
            PowerSupplyBacklightYes.Checked = true;
            PowerSupplyDetachableCablesNo.Checked = true;
            PowerSupplyDetachableCablesYes.Checked = true;
            PowerSupplyWireBraidingNo.Checked = true;
            PowerSupplyWireBraidingYes.Checked = true;
            for (int i = 0; i < 25; i++)
            {
                PowerSupplyFabricatorList.SetItemChecked(i, true);
            }

            CorpsWindowNo.Checked = true;
            CorpsWindowYes.Checked = true;
            for (int i = 0; i < 31; i++)
            {
                CorpsFabricatorList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 7; i++)
            {
                CorpsMainColorList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 10; i++)
            {
                CorpsBacklightList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 9; i++)
            {
                if (i != 4)
                {
                    CorpsFrameSizeList.SetItemChecked(i, true);
                }
                else
                {
                    CorpsFrameSizeList.SetItemChecked(i, true);
                }
            }

            HDDLevelOfNoiseMin.Text = "21";
            HDDLevelOfNoiseMax.Text = "27";
            HDDDataExchangeRateMin.Text = "126";
            HDDDataExchangeRateMax.Text = "160";
            for (int i = 0; i < 4; i++)
            {
                HDDBufferSizeList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 3; i++)
            {
                HDDFabricatorList.SetItemChecked(i, true);
            }
            HDDMemoryList.SetItemChecked(0, true);
            HDDMemoryList.SetItemChecked(1, true);
            HDDMemoryList.SetItemChecked(2, true);
            HDDMemoryList.SetItemChecked(3, true);
            HDDMemoryList.SetItemChecked(4, true);

            for (int i = 0; i < 32; i++)
            {
                SSDFabricatorList.SetItemChecked(i, true);
            }
            SSDMemoryMin.Text = "120";
            SSDMemoryMax.Text = "7680";
            SSDWriteSpeedMin.Text = "290";
            SSDWriteSpeedMax.Text = "540";
            SSDReadSpeedMin.Text = "450";
            SSDReadSpeedMax.Text = "3480";
        }

        /// <summary>
        /// Нажатие кнопки "Разработка"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DevButton_Click(object sender, EventArgs e)
        {
            OfficeButton.BackColor = SystemColors.Control;
            GamesButton.BackColor = SystemColors.Control;
            HomeButton.BackColor = SystemColors.Control;
            GraphicsButton.BackColor = SystemColors.Control;
            DevButton.BackColor = SystemColors.ActiveCaption;
            AnyButton.BackColor = SystemColors.Control;

            settings.ForGamingPC[0] = "NULL";
            settings.ForGamingPC[1] = "Нет";
            settings.ProfCard[0] = "NULL";
            settings.ProfCard[1] = "Нет";
            HDDcheckBox.Checked = true;
            SSDcheckBox.Checked = true;

            CPUcheckBox.Checked = true;
            CPUFabricatorAMD.Checked = true;
            CPUFabricatorIntel.Checked = true;
            CPUGraphicCoreYes.Checked = true;
            CPUGraphicCoreNo.Checked = true;
            CPUCoresList.SetItemChecked(0, true);
            for (int i = 1; i < 7; i++)
            {
                CPUCoresList.SetItemChecked(i, true);
            }
            CPUMultithreadingNo.Checked = true;
            CPUMultithreadingYes.Checked = true;
            CPUMemoryTypeDDR3.Checked = true;
            CPUMemoryTypeDDR4.Checked = true;
            CPUBaseFrequencyMin.Text = "2500";
            CPUBaseFrequencyMax.Text = "4200";

            for (int i = 0; i < 6; i++)
            {
                MotherBoardFabricatorList.SetItemChecked(i, true);
            }
            MotherBoardCPUTypeList.SetItemChecked(0, true);
            MotherBoardCPUTypeList.SetItemChecked(1, true);
            for (int i = 0; i < 3; i++)
            {
                MotherBoardNumberOfMemorySlotsList.SetItemChecked(i, true);
                MotherBoardMemoryTypeList.SetItemChecked(i, true);
            }
            MotherBoardMemoryTypeList.SetItemChecked(3, true);
            for (int i = 0; i < 6; i++)
            {
                MotherBoardNumberOfPCIEx16SlotsList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 6; i++)
            {
                MotherBoardNumberOfM2SlotsList.SetItemChecked(i, true);
            }
            MotherBoardWiFiAdapterNo.Checked = true;
            MotherBoardWiFiAdapterYes.Checked = false;

            GraphicsCardMemoryList.SetItemChecked(0, true);
            GraphicsCardMemoryList.SetItemChecked(1, true);
            GraphicsCardMemoryList.SetItemChecked(2, true);
            GraphicsCardMemoryList.SetItemChecked(3, false);
            GraphicsCardMemoryList.SetItemChecked(4, false);
            GraphicsCardMemoryList.SetItemChecked(5, false);
            for (int i = 0; i < 10; i++)
            {
                GraphicsCardFabricatorList.SetItemChecked(i, true);
            }
            GraphicsCardFabricatorOfGPUList.SetItemChecked(0, true);
            GraphicsCardFabricatorOfGPUList.SetItemChecked(1, true);
            for (int i = 0; i < 4; i++)
            {
                GraphicsCardMemoryTypeList.SetItemChecked(i, true);
            }
            GraphicsCardPCIExpressList.SetItemChecked(0, true);
            GraphicsCardPCIExpressList.SetItemChecked(1, true);
            GraphicsCardPCIExpressList.SetItemChecked(2, true);
            GraphicsCardNumberOfMonitorsList.SetItemChecked(0, true);
            GraphicsCardNumberOfMonitorsList.SetItemChecked(1, true);
            GraphicsCardNumberOfMonitorsList.SetItemChecked(2, true);
            GraphicsCardMemoryBusWidthMax.Text = "384";
            GraphicsCardMemoryBusWidthMin.Text = "32";

            RAMBacklightNo.Checked = true;
            RAMBacklightYes.Checked = false;
            for (int i = 0; i < 8; i++)
            {
                RAMMemoryList.SetItemChecked(i, true);
            }
            RAMMemoryTypeList.SetItemChecked(0, true);
            RAMMemoryTypeList.SetItemChecked(1, true);
            for (int i = 2; i < 4; i++)
            {
                RAMMemoryTypeList.SetItemChecked(i, true);
            }
            RAMMemoryTypeList.SetItemChecked(4, true);
            for (int i = 0; i < 21; i++)
            {
                RAMFabricatorList.SetItemChecked(i, true);
            }

            PowerSupplyBacklightNo.Checked = true;
            PowerSupplyBacklightYes.Checked = false;
            PowerSupplyDetachableCablesNo.Checked = true;
            PowerSupplyDetachableCablesYes.Checked = false;
            PowerSupplyWireBraidingNo.Checked = true;
            PowerSupplyWireBraidingYes.Checked = false;
            for (int i = 0; i < 25; i++)
            {
                PowerSupplyFabricatorList.SetItemChecked(i, true);
            }

            CorpsWindowNo.Checked = true;
            CorpsWindowYes.Checked = false;
            for (int i = 0; i < 31; i++)
            {
                CorpsFabricatorList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 7; i++)
            {
                CorpsMainColorList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 10; i++)
            {
                CorpsBacklightList.SetItemChecked(i, false);
            }
            CorpsBacklightList.SetItemChecked(7, true);
            for (int i = 0; i < 9; i++)
            {
                if (i != 4)
                {
                    CorpsFrameSizeList.SetItemChecked(i, true);
                }
                else
                {
                    CorpsFrameSizeList.SetItemChecked(i, false);
                }
            }

            HDDLevelOfNoiseMin.Text = "21";
            HDDLevelOfNoiseMax.Text = "27";
            HDDDataExchangeRateMin.Text = "126";
            HDDDataExchangeRateMax.Text = "160";
            for (int i = 0; i < 4; i++)
            {
                HDDBufferSizeList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 3; i++)
            {
                HDDFabricatorList.SetItemChecked(i, true);
            }
            HDDMemoryList.SetItemChecked(0, true);
            HDDMemoryList.SetItemChecked(1, true);
            HDDMemoryList.SetItemChecked(2, true);
            HDDMemoryList.SetItemChecked(3, true);
            HDDMemoryList.SetItemChecked(4, true);

            for (int i = 0; i < 32; i++)
            {
                SSDFabricatorList.SetItemChecked(i, true);
            }
            SSDMemoryMin.Text = "120";
            SSDMemoryMax.Text = "7680";
            SSDWriteSpeedMin.Text = "290";
            SSDWriteSpeedMax.Text = "540";
            SSDReadSpeedMin.Text = "450";
            SSDReadSpeedMax.Text = "3480";
        }

        private void GraphicsButton_Click(object sender, EventArgs e)
        {
            OfficeButton.BackColor = SystemColors.Control;
            GamesButton.BackColor = SystemColors.Control;
            HomeButton.BackColor = SystemColors.Control;
            GraphicsButton.BackColor = SystemColors.ActiveCaption;
            DevButton.BackColor = SystemColors.Control;
            AnyButton.BackColor = SystemColors.Control;

            settings.ForGamingPC[0] = "Да";
            settings.ForGamingPC[1] = "Нет";
            settings.ProfCard[0] = "Да";
            settings.ProfCard[1] = "Нет";
            HDDcheckBox.Checked = true;
            SSDcheckBox.Checked = true;

            CPUcheckBox.Checked = true;
            CPUFabricatorAMD.Checked = true;
            CPUFabricatorIntel.Checked = true;
            CPUGraphicCoreYes.Checked = true;
            CPUGraphicCoreNo.Checked = true;
            CPUCoresList.SetItemChecked(0, false);
            for (int i = 1; i < 7; i++)
            {
                CPUCoresList.SetItemChecked(i, true);
            }
            CPUMultithreadingNo.Checked = false;
            CPUMultithreadingYes.Checked = true;
            CPUMemoryTypeDDR3.Checked = true;
            CPUMemoryTypeDDR4.Checked = true;
            CPUBaseFrequencyMin.Text = "3000";
            CPUBaseFrequencyMax.Text = "4200";

            for (int i = 0; i < 6; i++)
            {
                MotherBoardFabricatorList.SetItemChecked(i, true);
            }
            MotherBoardCPUTypeList.SetItemChecked(0, true);
            MotherBoardCPUTypeList.SetItemChecked(1, true);
            for (int i = 0; i < 3; i++)
            {
                MotherBoardNumberOfMemorySlotsList.SetItemChecked(i, true);
                MotherBoardMemoryTypeList.SetItemChecked(i, false);
            }
            MotherBoardMemoryTypeList.SetItemChecked(3, true);
            for (int i = 0; i < 6; i++)
            {
                MotherBoardNumberOfPCIEx16SlotsList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 6; i++)
            {
                MotherBoardNumberOfM2SlotsList.SetItemChecked(i, true);
            }
            MotherBoardWiFiAdapterNo.Checked = true;
            MotherBoardWiFiAdapterYes.Checked = false;

            GraphicsCardMemoryList.SetItemChecked(0, false);
            GraphicsCardMemoryList.SetItemChecked(1, false);
            for (int i = 2; i < 6; i++)
            {
                GraphicsCardMemoryList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 10; i++)
            {
                GraphicsCardFabricatorList.SetItemChecked(i, true);
            }
            GraphicsCardFabricatorOfGPUList.SetItemChecked(0, true);
            GraphicsCardFabricatorOfGPUList.SetItemChecked(1, true);
            for (int i = 0; i < 4; i++)
            {
                GraphicsCardMemoryTypeList.SetItemChecked(i, true);
            }
            GraphicsCardPCIExpressList.SetItemChecked(0, true);
            GraphicsCardPCIExpressList.SetItemChecked(1, true);
            GraphicsCardPCIExpressList.SetItemChecked(2, true);
            GraphicsCardNumberOfMonitorsList.SetItemChecked(0, true);
            GraphicsCardNumberOfMonitorsList.SetItemChecked(1, true);
            GraphicsCardNumberOfMonitorsList.SetItemChecked(2, true);
            GraphicsCardMemoryBusWidthMax.Text = "384";
            GraphicsCardMemoryBusWidthMin.Text = "32";

            RAMBacklightNo.Checked = true;
            RAMBacklightYes.Checked = true;
            RAMMemoryList.SetItemChecked(0, false);
            RAMMemoryList.SetItemChecked(1, false);
            RAMMemoryList.SetItemChecked(2, false);
            for (int i = 3; i < 8; i++)
            {
                RAMMemoryList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 4; i++)
            {
                RAMMemoryTypeList.SetItemChecked(i, false);
            }
            RAMMemoryTypeList.SetItemChecked(4, true);
            for (int i = 0; i < 21; i++)
            {
                RAMFabricatorList.SetItemChecked(i, true);
            }

            PowerSupplyBacklightNo.Checked = true;
            PowerSupplyBacklightYes.Checked = true;
            PowerSupplyDetachableCablesNo.Checked = true;
            PowerSupplyDetachableCablesYes.Checked = true;
            PowerSupplyWireBraidingNo.Checked = true;
            PowerSupplyWireBraidingYes.Checked = true;
            for (int i = 0; i < 25; i++)
            {
                PowerSupplyFabricatorList.SetItemChecked(i, true);
            }

            CorpsWindowNo.Checked = true;
            CorpsWindowYes.Checked = true;
            for (int i = 0; i < 31; i++)
            {
                CorpsFabricatorList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 7; i++)
            {
                CorpsMainColorList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 10; i++)
            {
                CorpsBacklightList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 9; i++)
            {
                if (i != 4)
                {
                    CorpsFrameSizeList.SetItemChecked(i, true);
                }
                else
                {
                    CorpsFrameSizeList.SetItemChecked(i, false);
                }
            }

            HDDLevelOfNoiseMin.Text = "21";
            HDDLevelOfNoiseMax.Text = "27";
            HDDDataExchangeRateMin.Text = "126";
            HDDDataExchangeRateMax.Text = "160";
            for (int i = 0; i < 4; i++)
            {
                HDDBufferSizeList.SetItemChecked(i, true);
            }
            for (int i = 0; i < 3; i++)
            {
                HDDFabricatorList.SetItemChecked(i, true);
            }
            HDDMemoryList.SetItemChecked(0, false);
            HDDMemoryList.SetItemChecked(1, true);
            HDDMemoryList.SetItemChecked(2, true);
            HDDMemoryList.SetItemChecked(3, true);
            HDDMemoryList.SetItemChecked(4, true);

            for (int i = 0; i < 32; i++)
            {
                SSDFabricatorList.SetItemChecked(i, true);
            }
            SSDMemoryMin.Text = "120";
            SSDMemoryMax.Text = "7680";
            SSDWriteSpeedMin.Text = "290";
            SSDWriteSpeedMax.Text = "540";
            SSDReadSpeedMin.Text = "450";
            SSDReadSpeedMax.Text = "3480";
        }
    }
}
