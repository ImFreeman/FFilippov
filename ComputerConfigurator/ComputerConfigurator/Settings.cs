using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComputerConfigurator
{
    public class Settings
    {
        public int[] Price = new int[2];

        //Для игрового ПК:
        public string[] ForGamingPC = new string[2];

        //CPU настройки
        public string[] CPUFabricator=new string[2];
        public string[] CPUCores = new string[7];
        public string[] CPUGraphicCore = new string[2];
        public string[] CPUMemoryType = new string[2];
        public int[] CPUBaseFrequency = new int[2];
        public string[] CPUMultithreading = new string[2];        
        

        //MotherВoard настройки
        public string[] MotherBoardFabricator = new string[6];
        public string[] MotherBoardCPUType = new string[2];
        public string[] MotherBoardMemoryType = new string[4];
        public string[] MotherBoardNumberOfMemorySlots = new string[3];
        public string[] MotherBoardNumberOfPCIEx16Slots = new string[6];
        public string[] MotherBoardNumberOfM2Slots = new string[6];
        public string[] MotherBoardWiFiAdapter = new string[2];
        public string[] MotherBoardBuildInCPU = new string[2];

        //GrapgicsCard настройки
        public string[] ProfCard = new string[2];
        public string[] GraphicsCardFabricator = new string[10];
        public string[] GraphicsCardMemory = new string[6];
        public string[] GraphicsCardMemoryType = new string[4];
        public string[] GraphicsCardFabricatorOfGPU = new string[2];
        public string[] GraphicsCardNumberOfMonitors = new string[3];
        public string[] GraphicsCardPCIExpress = new string[3];
        public int[] GraphicsCardMemoryBusWidth = new int[2];

        //RAM настройки
        public string[] RAMFabricator = new string[21];
        public string[] RAMBacklight = new string[2];
        public string[] RAMMemory = new string[8];
        public string[] RAMMemoryType = new string[5];

        //PowerSupply настройки
        public string[] PowerSupplyFabricator = new string[25];
        public string[] PowerSupplyWireBraiding = new string[2];
        public string[] PowerSupplyBacklight = new string[2];
        public string[] PowerSupplyDetachableCables = new string[2];

        //Corps настройки
        public string[] CorpsFabricator = new string[31];
        public string[] CorpsWindow = new string[2];
        public string[] CorpsMainColor = new string[7];
        public string[] CorpsBacklight = new string[10];
        public string[] CorpsFrameSize = new string[9];

        //Необходимость накопителей
        public bool HDDRequired;
        public bool SSDRequired;

        //HHD настройки
        public string[] HDDMemory = new string[5];
        public int[] HDDLevelOfNoise = new int[2];
        public int[] HDDDataExchangeRate = new int[2];
        public string[] HDDFabricator = new string[3];
        public string[] HDDBufferSize = new string[4];


        //SSD настройки
        public string[] SSDFabricator = new string[32];
        public int[] SSDMemory = new int[2];
        public int[] SSDWriteSpeed = new int[2];
        public int[] SSDReadSpeed = new int[2];
    }
}
