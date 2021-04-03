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

            

            int j = 2;
            Excel.Range range;
            range = (Excel.Range)sheet.Cells[2, 1];
            while(range.Value2!="end")
            {
                CPU proc=new CPU();
                proc.Name= (string)range.Value2;
                range = (Excel.Range)sheet.Cells[j, 2];
                proc.Price = range.Value2;
                range = (Excel.Range)sheet.Cells[j, 3];
                proc.Fabricator= range.Value2;
                range = (Excel.Range)sheet.Cells[j, 4];
                proc.NumberOfCores = Convert.ToInt32(range.Value2);
                range = (Excel.Range)sheet.Cells[j, 5];
                proc.GraphicCore = range.Value2;
                range = (Excel.Range)sheet.Cells[j, 6];
                proc.MemoryType = range.Value2;
                range = (Excel.Range)sheet.Cells[j, 7];
                proc.BaseFrequency = Convert.ToInt32(range.Value2);
                range = (Excel.Range)sheet.Cells[j, 8];
                proc.Multithreading = range.Value2;
                range = (Excel.Range)sheet.Cells[j, 9];
                proc.Socket = range.Value2;
                range = (Excel.Range)sheet.Cells[j, 10];
                proc.ForGamingPC = range.Value2;
                range = (Excel.Range)sheet.Cells[j, 11];
                proc.Site = range.Value2;

                for(int i=0;i<=settings.CPUFabricator.Length;i++)
                {
                    if(settings.CPUFabricator[i]==proc.Fabricator)
                    {
                        for(int c=0;c<settings.CPUCores.Length;c++)
                        {
                            if(settings.CPUCores[c]==Convert.ToString(proc.NumberOfCores))
                            {
                                for(int q=0;q<settings.CPUGraphicCore.Length;q++)
                                {
                                    if(settings.CPUGraphicCore[q]==proc.GraphicCore)
                                    {
                                        for(int t=0;t<settings.CPUMemoryType.Length;t++)
                                        {
                                            if(settings.CPUMemoryType[t]==proc.MemoryType)
                                            {
                                                if((proc.BaseFrequency>=settings.CPUBaseFrequency[0])&& (proc.BaseFrequency <= settings.CPUBaseFrequency[1]))
                                                {
                                                    for(int x=0;x<settings.CPUMultithreading.Length;x++)
                                                    {
                                                        if(settings.CPUMultithreading[x]==proc.Multithreading)
                                                        {
                                                            for(int y=0;y<settings.ForGamingPC.Length;y++)
                                                            {
                                                                if(settings.ForGamingPC[y]==proc.ForGamingPC)
                                                                {
                                                                    Computer BufferComp = new Computer();
                                                                    BufferComp.processor = proc;
                                                                    //теперь материнка
                                                                    Excel.Application excel_app2 = new Excel.Application();
                                                                    Excel.Workbook workbook2 = excel_app2.Workbooks.Open(
                                                                        Path.GetFullPath("../../Resources/MotherBoard"),
                                                                        Type.Missing, true, Type.Missing, Type.Missing,
                                                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                                        Type.Missing, Type.Missing);

                                                                    // Получить первый рабочий лист.
                                                                    Excel.Worksheet sheet2 = (Excel.Worksheet)workbook2.Sheets[1];
                                                                    Excel.Range range2;
                                                                    int r = 2;
                                                                    range2 = (Excel.Range)sheet2.Cells[2, 5];
                                                                    while(range2.Value2!="end")
                                                                    {
                                                                        if(range2.Value2==BufferComp.processor.Socket)
                                                                        {
                                                                            MotherBoard Mom = new MotherBoard();
                                                                            Mom.Socket = range2.Value2;
                                                                            range2 = (Excel.Range)sheet2.Cells[r, 1];
                                                                            Mom.Name = range2.Value2;
                                                                            range2 = (Excel.Range)sheet2.Cells[r, 2];
                                                                            Mom.Price = range2.Value2;
                                                                            range2 = (Excel.Range)sheet2.Cells[r, 3];
                                                                            Mom.Fabricator = range2.Value2;
                                                                            range2 = (Excel.Range)sheet2.Cells[r, 4];
                                                                            Mom.CPUType = range2.Value2;
                                                                            range2 = (Excel.Range)sheet2.Cells[r, 6];
                                                                            Mom.Chipset = range2.Value2;
                                                                            range2 = (Excel.Range)sheet2.Cells[r, 7];
                                                                            Mom.NumberOfPCIEx16Slots = Convert.ToInt32(range2.Value2);
                                                                            range2 = (Excel.Range)sheet2.Cells[r, 8];
                                                                            Mom.NumberOfM2Slots = Convert.ToInt32(range2.Value2);
                                                                            range2 = (Excel.Range)sheet2.Cells[r, 9];
                                                                            Mom.WiFiAdapter = range2.Value2;
                                                                            range2 = (Excel.Range)sheet2.Cells[r, 10];
                                                                            Mom.BuildInCPU = range2.Value2;
                                                                            range2 = (Excel.Range)sheet2.Cells[r, 11];
                                                                            Mom.ForGamingPC = range2.Value2;
                                                                            range2 = (Excel.Range)sheet2.Cells[r, 12];
                                                                            Mom.MemoryType = range2.Value2;
                                                                            range2 = (Excel.Range)sheet2.Cells[r, 13];
                                                                            Mom.NumberOfMemorySlots = Convert.ToInt32(range2.Value2);
                                                                            range2 = (Excel.Range)sheet2.Cells[r, 14];
                                                                            Mom.Site = range2.Value2;
                                                                            
                                                                            for(int w=0;w<settings.MotherBoardFabricator.Length;w++)
                                                                            {
                                                                                if(settings.MotherBoardFabricator[w]==Mom.Fabricator)
                                                                                {
                                                                                    for(int e=0;e<settings.MotherBoardCPUType.Length;e++)
                                                                                    {
                                                                                        if(settings.MotherBoardCPUType[e]==Mom.CPUType)
                                                                                        {
                                                                                            for(int s=0;s<settings.MotherBoardNumberOfPCIEx16Slots.Length;s++)
                                                                                            {
                                                                                                if(settings.MotherBoardNumberOfPCIEx16Slots[s]==Convert.ToString(Mom.NumberOfPCIEx16Slots))
                                                                                                {
                                                                                                    for(int d=0;d<settings.MotherBoardNumberOfM2Slots.Length;d++)
                                                                                                    {
                                                                                                        if(settings.MotherBoardNumberOfM2Slots[d]==Convert.ToString(Mom.NumberOfM2Slots))
                                                                                                        {
                                                                                                            for(int z=0;z<settings.MotherBoardNumberOfMemorySlots.Length;z++)
                                                                                                            {
                                                                                                                if(settings.MotherBoardNumberOfMemorySlots[z]==Convert.ToString(Mom.NumberOfMemorySlots))
                                                                                                                {
                                                                                                                    for(int g=0;g<settings.MotherBoardWiFiAdapter.Length;g++)
                                                                                                                    {
                                                                                                                        if(settings.MotherBoardWiFiAdapter[g]==Mom.WiFiAdapter)
                                                                                                                        {
                                                                                                                            for(int f=0;f<settings.MotherBoardBuildInCPU.Length;f++)
                                                                                                                            {
                                                                                                                                if(settings.MotherBoardBuildInCPU[f]==Mom.BuildInCPU)
                                                                                                                                {
                                                                                                                                    for(int b=0;b<settings.MotherBoardMemoryType.Length;b++)
                                                                                                                                    {
                                                                                                                                        if(settings.MotherBoardMemoryType[b]==Mom.MemoryType)
                                                                                                                                        {
                                                                                                                                            for(int k=0;k<settings.ForGamingPC.Length;k++)
                                                                                                                                            {
                                                                                                                                                if(settings.ForGamingPC[k]==Mom.ForGamingPC)
                                                                                                                                                {
                                                                                                                                                    BufferComp.motherBoard = Mom;
                                                                                                                                                    //оперативка
                                                                                                                                                    Excel.Application excel_app3 = new Excel.Application();
                                                                                                                                                    Excel.Workbook workbook3 = excel_app3.Workbooks.Open(
                                                                                                                                                        Path.GetFullPath("../../Resources/RAM"),
                                                                                                                                                        Type.Missing, true, Type.Missing, Type.Missing,
                                                                                                                                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                                                                                                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                                                                                                                        Type.Missing, Type.Missing);

                                                                                                                                                    // Получить первый рабочий лист.
                                                                                                                                                    Excel.Worksheet sheet3 = (Excel.Worksheet)workbook3.Sheets[1];
                                                                                                                                                    Excel.Range range3;
                                                                                                                                                    int l = 2;
                                                                                                                                                    range3 = (Excel.Range)sheet3.Cells[2, 6];
                                                                                                                                                    while (range3.Value2 != "end")
                                                                                                                                                    {
                                                                                                                                                        if(range3.Value2==BufferComp.motherBoard.MemoryType)
                                                                                                                                                        {
                                                                                                                                                            RAM ozu = new RAM();
                                                                                                                                                            ozu.MemoryType = range3.Value2;
                                                                                                                                                            range3 = (Excel.Range)sheet3.Cells[l, 1];
                                                                                                                                                            ozu.Name = range3.Value2;
                                                                                                                                                            range3 = (Excel.Range)sheet3.Cells[l, 2];
                                                                                                                                                            ozu.Price = range3.Value2;
                                                                                                                                                            range3 = (Excel.Range)sheet3.Cells[l, 3];
                                                                                                                                                            ozu.Fabricator = range3.Value2;
                                                                                                                                                            range3 = (Excel.Range)sheet3.Cells[l, 4];
                                                                                                                                                            ozu.Backlight = range3.Value2;
                                                                                                                                                            range3 = (Excel.Range)sheet3.Cells[l, 5];
                                                                                                                                                            ozu.Memory = Convert.ToInt32(range3.Value2);
                                                                                                                                                            range3 = (Excel.Range)sheet3.Cells[l, 7];
                                                                                                                                                            ozu.ForGamingPC = range3.Value2;
                                                                                                                                                            range3 = (Excel.Range)sheet3.Cells[l, 8];
                                                                                                                                                            ozu.Site = range3.Value2;

                                                                                                                                                            for(int o=0;o<settings.RAMFabricator.Length;o++)
                                                                                                                                                            {
                                                                                                                                                                if(settings.RAMFabricator[o]==ozu.Fabricator)
                                                                                                                                                                {
                                                                                                                                                                    for(int p=0;p<settings.RAMBacklight.Length;p++)
                                                                                                                                                                    {
                                                                                                                                                                        if(settings.RAMBacklight[p]==ozu.Backlight)
                                                                                                                                                                        {
                                                                                                                                                                            for(int m=0;m<settings.RAMMemory.Length;m++)
                                                                                                                                                                            {
                                                                                                                                                                                if(settings.RAMMemory[m]==Convert.ToString(ozu.Memory))
                                                                                                                                                                                {
                                                                                                                                                                                    for(int u=0;u<settings.RAMMemoryType.Length;u++)
                                                                                                                                                                                    {
                                                                                                                                                                                        if(settings.RAMMemoryType[u]==ozu.MemoryType)
                                                                                                                                                                                        {
                                                                                                                                                                                            for(int v=0;v<settings.ForGamingPC.Length;v++)
                                                                                                                                                                                            {
                                                                                                                                                                                                if(settings.ForGamingPC[v]==ozu.ForGamingPC)
                                                                                                                                                                                                {
                                                                                                                                                                                                    BufferComp.ram = ozu;
                                                                                                                                                                                                    
                                                                                                                                                                                                    //видеокарта
                                                                                                                                                                                                    Excel.Application excel_app4 = new Excel.Application();
                                                                                                                                                                                                    Excel.Workbook workbook4 = excel_app4.Workbooks.Open(
                                                                                                                                                                                                        Path.GetFullPath("../../Resources/GraphicsCard"),
                                                                                                                                                                                                        Type.Missing, true, Type.Missing, Type.Missing,
                                                                                                                                                                                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                                                                                                                                                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                                                                                                                                                                        Type.Missing, Type.Missing);

                                                                                                                                                                                                    // Получить первый рабочий лист.
                                                                                                                                                                                                    Excel.Worksheet sheet4 = (Excel.Worksheet)workbook4.Sheets[1];
                                                                                                                                                                                                    Excel.Range range4;
                                                                                                                                                                                                    int n = 2;
                                                                                                                                                                                                    range4 = (Excel.Range)sheet4.Cells[2, 1];
                                                                                                                                                                                                    while (range4.Value2 != "end")
                                                                                                                                                                                                    {
                                                                                                                                                                                                        GraphicsCard vid = new GraphicsCard();
                                                                                                                                                                                                        vid.Name = range4.Value2;
                                                                                                                                                                                                        range4 = (Excel.Range)sheet4.Cells[n, 2];
                                                                                                                                                                                                        vid.Price = range4.Value2;
                                                                                                                                                                                                        range4 = (Excel.Range)sheet4.Cells[n, 3];
                                                                                                                                                                                                        vid.Fabricator= range4.Value2;
                                                                                                                                                                                                        range4 = (Excel.Range)sheet4.Cells[n, 4];
                                                                                                                                                                                                        vid.RecommendedEnergy = Convert.ToInt32(range4.Value2);
                                                                                                                                                                                                        range4 = (Excel.Range)sheet4.Cells[n, 5];
                                                                                                                                                                                                        vid.Memory = Convert.ToInt32(range4.Value2);
                                                                                                                                                                                                        range4 = (Excel.Range)sheet4.Cells[n, 6];
                                                                                                                                                                                                        vid.MemoryType = range4.Value2;
                                                                                                                                                                                                        range4 = (Excel.Range)sheet4.Cells[n, 7];
                                                                                                                                                                                                        vid.FabricatorOfGPU = range4.Value2;
                                                                                                                                                                                                        range4 = (Excel.Range)sheet4.Cells[n, 8];
                                                                                                                                                                                                        vid.NumberOfMonitors = Convert.ToInt32(range4.Value2);
                                                                                                                                                                                                        range4 = (Excel.Range)sheet4.Cells[n, 9];
                                                                                                                                                                                                        vid.PCIExpress = range4.Value2;
                                                                                                                                                                                                        range4 = (Excel.Range)sheet4.Cells[n, 10];
                                                                                                                                                                                                        vid.MemoryBusWidth = Convert.ToInt32(range4.Value2);
                                                                                                                                                                                                        range4 = (Excel.Range)sheet4.Cells[n, 11];
                                                                                                                                                                                                        vid.ForGamingPC = range4.Value2;
                                                                                                                                                                                                        range4 = (Excel.Range)sheet4.Cells[n, 12];
                                                                                                                                                                                                        vid.ProfessionalGraphicsCard = range4.Value2;
                                                                                                                                                                                                        range4 = (Excel.Range)sheet4.Cells[n, 13];
                                                                                                                                                                                                        vid.Site = range4.Value2;

                                                                                                                                                                                                        for(int qw=0;qw< settings.GraphicsCardFabricator.Length;qw++)
                                                                                                                                                                                                        {
                                                                                                                                                                                                            if(settings.GraphicsCardFabricator[qw]==vid.Fabricator)
                                                                                                                                                                                                            {
                                                                                                                                                                                                                for(int qe=0;qe<settings.GraphicsCardMemory.Length;qe++)
                                                                                                                                                                                                                {
                                                                                                                                                                                                                    if(settings.GraphicsCardMemory[qe]==Convert.ToString(vid.Memory))
                                                                                                                                                                                                                    {
                                                                                                                                                                                                                        for(int az=0;az<settings.GraphicsCardMemoryType.Length;az++)
                                                                                                                                                                                                                        {
                                                                                                                                                                                                                            if(settings.GraphicsCardMemoryType[az]==vid.MemoryType)
                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                for(int ax=0;ax<settings.GraphicsCardFabricatorOfGPU.Length;ax++)
                                                                                                                                                                                                                                {
                                                                                                                                                                                                                                    if(settings.GraphicsCardFabricatorOfGPU[ax]==vid.FabricatorOfGPU)
                                                                                                                                                                                                                                    {
                                                                                                                                                                                                                                        for(int ac=0; ac<settings.GraphicsCardNumberOfMonitors.Length;ac++)
                                                                                                                                                                                                                                        {
                                                                                                                                                                                                                                            if(settings.GraphicsCardNumberOfMonitors[ac]==Convert.ToString(vid.NumberOfMonitors))
                                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                                for(int wq=0;wq<settings.GraphicsCardPCIExpress.Length;wq++)
                                                                                                                                                                                                                                                {
                                                                                                                                                                                                                                                    if(settings.GraphicsCardPCIExpress[wq]==vid.PCIExpress)
                                                                                                                                                                                                                                                    {
                                                                                                                                                                                                                                                        if((vid.MemoryBusWidth>=Convert.ToInt32(settings.GraphicsCardMemoryBusWidth[0]))&& (vid.MemoryBusWidth <= Convert.ToInt32(settings.GraphicsCardMemoryBusWidth[1])))
                                                                                                                                                                                                                                                        {
                                                                                                                                                                                                                                                            for(int zx=0;zx<settings.ForGamingPC.Length;zx++)
                                                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                                                if(settings.ForGamingPC[zx]==vid.ForGamingPC)
                                                                                                                                                                                                                                                                {
                                                                                                                                                                                                                                                                    for(int ed=0;ed<settings.ProfCard.Length;ed++)
                                                                                                                                                                                                                                                                    {
                                                                                                                                                                                                                                                                        if(settings.ProfCard[ed]==vid.ProfessionalGraphicsCard)
                                                                                                                                                                                                                                                                        {
                                                                                                                                                                                                                                                                            BufferComp.graphicsCard = vid;
                                                                                                                                                                                                                                                                            //блок питания
                                                                                                                                                                                                                                                                            Excel.Application excel_app5 = new Excel.Application();
                                                                                                                                                                                                                                                                            Excel.Workbook workbook5 = excel_app5.Workbooks.Open(
                                                                                                                                                                                                                                                                                Path.GetFullPath("../../Resources/PowerSupply"),
                                                                                                                                                                                                                                                                                Type.Missing, true, Type.Missing, Type.Missing,
                                                                                                                                                                                                                                                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                                                                                                                                                                                                                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                                                                                                                                                                                                                                                Type.Missing, Type.Missing);

                                                                                                                                                                                                                                                                            // Получить первый рабочий лист.
                                                                                                                                                                                                                                                                            Excel.Worksheet sheet5 = (Excel.Worksheet)workbook5.Sheets[1];
                                                                                                                                                                                                                                                                            Excel.Range range5;
                                                                                                                                                                                                                                                                            int po = 2;
                                                                                                                                                                                                                                                                            range5 = (Excel.Range)sheet5.Cells[2, 4];
                                                                                                                                                                                                                                                                            while (Convert.ToString(range5.Value2) != "end")
                                                                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                                                                if(range5.Value2==BufferComp.graphicsCard.RecommendedEnergy)
                                                                                                                                                                                                                                                                                {
                                                                                                                                                                                                                                                                                    PowerSupply power = new PowerSupply();
                                                                                                                                                                                                                                                                                    power.Energy = Convert.ToInt32(range5.Value2);
                                                                                                                                                                                                                                                                                    range5 = (Excel.Range)sheet5.Cells[po, 1];
                                                                                                                                                                                                                                                                                    power.Name = range5.Value2;
                                                                                                                                                                                                                                                                                    range5 = (Excel.Range)sheet5.Cells[po, 2];
                                                                                                                                                                                                                                                                                    power.Price = range5.Value2;
                                                                                                                                                                                                                                                                                    range5 = (Excel.Range)sheet5.Cells[po, 3];
                                                                                                                                                                                                                                                                                    power.Fabricator= range5.Value2;
                                                                                                                                                                                                                                                                                    range5 = (Excel.Range)sheet5.Cells[po, 5];
                                                                                                                                                                                                                                                                                    power.WireBraiding = range5.Value2;
                                                                                                                                                                                                                                                                                    range5 = (Excel.Range)sheet5.Cells[po, 6];
                                                                                                                                                                                                                                                                                    power.Backlight = range5.Value2;
                                                                                                                                                                                                                                                                                    range5 = (Excel.Range)sheet5.Cells[po, 7];
                                                                                                                                                                                                                                                                                    power.DetachableCables = range5.Value2;
                                                                                                                                                                                                                                                                                    range5 = (Excel.Range)sheet5.Cells[po, 8];
                                                                                                                                                                                                                                                                                    power.Site = range5.Value2;

                                                                                                                                                                                                                                                                                    for(int qx=0;qx<settings.PowerSupplyFabricator.Length;qx++)
                                                                                                                                                                                                                                                                                    {
                                                                                                                                                                                                                                                                                        if(settings.PowerSupplyFabricator[qx]==power.Fabricator)
                                                                                                                                                                                                                                                                                        {
                                                                                                                                                                                                                                                                                            for(int pl=0;pl<settings.PowerSupplyWireBraiding.Length;pl++)
                                                                                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                                                                                if(settings.PowerSupplyWireBraiding[pl]==power.WireBraiding)
                                                                                                                                                                                                                                                                                                {
                                                                                                                                                                                                                                                                                                    for(int hl=0;hl<settings.PowerSupplyBacklight.Length;hl++)
                                                                                                                                                                                                                                                                                                    {
                                                                                                                                                                                                                                                                                                        if(settings.PowerSupplyBacklight[hl]==power.Backlight)
                                                                                                                                                                                                                                                                                                        {
                                                                                                                                                                                                                                                                                                            for(int go=0; go<settings.PowerSupplyDetachableCables.Length;go++)
                                                                                                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                                                                                                if(settings.PowerSupplyDetachableCables[go]==power.DetachableCables)
                                                                                                                                                                                                                                                                                                                {
                                                                                                                                                                                                                                                                                                                    BufferComp.powerSupply = power;
                                                                                                                                                                                                                                                                                                                    //теперь корпус
                                                                                                                                                                                                                                                                                                                    Excel.Application excel_app6 = new Excel.Application();
                                                                                                                                                                                                                                                                                                                    Excel.Workbook workbook6 = excel_app6.Workbooks.Open(
                                                                                                                                                                                                                                                                                                                        Path.GetFullPath("../../Resources/Corps"),
                                                                                                                                                                                                                                                                                                                        Type.Missing, true, Type.Missing, Type.Missing,
                                                                                                                                                                                                                                                                                                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                                                                                                                                                                                                                                                                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                                                                                                                                                                                                                                                                                        Type.Missing, Type.Missing);

                                                                                                                                                                                                                                                                                                                    // Получить первый рабочий лист.
                                                                                                                                                                                                                                                                                                                    Excel.Worksheet sheet6 = (Excel.Worksheet)workbook6.Sheets[1];
                                                                                                                                                                                                                                                                                                                    Excel.Range range6;
                                                                                                                                                                                                                                                                                                                    int op = 2;
                                                                                                                                                                                                                                                                                                                    range6 = (Excel.Range)sheet6.Cells[2, 1];
                                                                                                                                                                                                                                                                                                                    while (range6.Value2 != "end")
                                                                                                                                                                                                                                                                                                                    {
                                                                                                                                                                                                                                                                                                                        Corps corps = new Corps();

                                                                                                                                                                                                                                                                                                                        corps.Name = range6.Value2;
                                                                                                                                                                                                                                                                                                                        range6 = (Excel.Range)sheet6.Cells[op, 2];
                                                                                                                                                                                                                                                                                                                        corps.Price = range6.Value2;
                                                                                                                                                                                                                                                                                                                        range6 = (Excel.Range)sheet6.Cells[op, 3];
                                                                                                                                                                                                                                                                                                                        corps.Fabricator = range6.Value2;
                                                                                                                                                                                                                                                                                                                        range6 = (Excel.Range)sheet6.Cells[op, 4];
                                                                                                                                                                                                                                                                                                                        corps.MainColor = range6.Value2;
                                                                                                                                                                                                                                                                                                                        range6 = (Excel.Range)sheet6.Cells[op, 5];
                                                                                                                                                                                                                                                                                                                        corps.Window = range6.Value2;
                                                                                                                                                                                                                                                                                                                        range6 = (Excel.Range)sheet6.Cells[op, 6];
                                                                                                                                                                                                                                                                                                                        corps.Backlight = range6.Value2;
                                                                                                                                                                                                                                                                                                                        range6 = (Excel.Range)sheet6.Cells[op, 7];
                                                                                                                                                                                                                                                                                                                        corps.FrameSize = range6.Value2;
                                                                                                                                                                                                                                                                                                                        range6 = (Excel.Range)sheet6.Cells[op, 8];
                                                                                                                                                                                                                                                                                                                        corps.ForGamingPC = range6.Value2;
                                                                                                                                                                                                                                                                                                                        range6 = (Excel.Range)sheet6.Cells[op, 9];
                                                                                                                                                                                                                                                                                                                        corps.Site = range6.Value2;

                                                                                                                                                                                                                                                                                                                        for(int qqq=0;qqq<settings.CorpsFabricator.Length;qqq++)
                                                                                                                                                                                                                                                                                                                        {
                                                                                                                                                                                                                                                                                                                            if(settings.CorpsFabricator[qqq]==corps.Fabricator)
                                                                                                                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                                                                                                                for(int asd=0;asd<settings.CorpsMainColor.Length;asd++)
                                                                                                                                                                                                                                                                                                                                {
                                                                                                                                                                                                                                                                                                                                    if(settings.CorpsMainColor[asd]==corps.MainColor)
                                                                                                                                                                                                                                                                                                                                    {
                                                                                                                                                                                                                                                                                                                                        for(int qwe=0;qwe<settings.CorpsWindow.Length;qwe++)
                                                                                                                                                                                                                                                                                                                                        {
                                                                                                                                                                                                                                                                                                                                            if(settings.CorpsWindow[qwe]==corps.Window)
                                                                                                                                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                                                                                                                                for(int hj=0;hj<settings.CorpsBacklight.Length;hj++)
                                                                                                                                                                                                                                                                                                                                                {
                                                                                                                                                                                                                                                                                                                                                    if(settings.CorpsBacklight[hj]==corps.Backlight)
                                                                                                                                                                                                                                                                                                                                                    {
                                                                                                                                                                                                                                                                                                                                                        for(int rot=0;rot<settings.CorpsFrameSize.Length;rot++)
                                                                                                                                                                                                                                                                                                                                                        {
                                                                                                                                                                                                                                                                                                                                                            if(settings.CorpsFrameSize[rot]==corps.FrameSize)
                                                                                                                                                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                                                                                                                                                for(int pc=0;pc<settings.ForGamingPC.Length;pc++)
                                                                                                                                                                                                                                                                                                                                                                {
                                                                                                                                                                                                                                                                                                                                                                    if(settings.ForGamingPC[pc]==corps.ForGamingPC)
                                                                                                                                                                                                                                                                                                                                                                    {
                                                                                                                                                                                                                                                                                                                                                                        BufferComp.corps = corps;
                                                                                                                                                                                                                                                                                                                                                                        //HDD и SSD
                                                                                                                                                                                                                                                                                                                                                                        if(settings.HDDRequired)
                                                                                                                                                                                                                                                                                                                                                                        {
                                                                                                                                                                                                                                                                                                                                                                            Excel.Application excel_app7 = new Excel.Application();
                                                                                                                                                                                                                                                                                                                                                                            Excel.Workbook workbook7 = excel_app7.Workbooks.Open(
                                                                                                                                                                                                                                                                                                                                                                                Path.GetFullPath("../../Resources/HDD"),
                                                                                                                                                                                                                                                                                                                                                                                Type.Missing, true, Type.Missing, Type.Missing,
                                                                                                                                                                                                                                                                                                                                                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                                                                                                                                                                                                                                                                                                                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                                                                                                                                                                                                                                                                                                                                                Type.Missing, Type.Missing);

                                                                                                                                                                                                                                                                                                                                                                            // Получить первый рабочий лист.
                                                                                                                                                                                                                                                                                                                                                                            Excel.Worksheet sheet7 = (Excel.Worksheet)workbook7.Sheets[1];
                                                                                                                                                                                                                                                                                                                                                                            Excel.Range range7;
                                                                                                                                                                                                                                                                                                                                                                            int ad = 2;
                                                                                                                                                                                                                                                                                                                                                                            range7 = (Excel.Range)sheet7.Cells[2, 1];
                                                                                                                                                                                                                                                                                                                                                                            while (range7.Value2 != "end")
                                                                                                                                                                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                                                                                                                                                                HDD disk = new HDD();
                                                                                                                                                                                                                                                                                                                                                                                disk.Name = range7.Value2;
                                                                                                                                                                                                                                                                                                                                                                                range7 = (Excel.Range)sheet7.Cells[ad, 2];
                                                                                                                                                                                                                                                                                                                                                                                disk.Price = range7.Value2;
                                                                                                                                                                                                                                                                                                                                                                                range7 = (Excel.Range)sheet7.Cells[ad, 3];
                                                                                                                                                                                                                                                                                                                                                                                disk.Fabricator = range7.Value2;
                                                                                                                                                                                                                                                                                                                                                                                range7 = (Excel.Range)sheet7.Cells[ad, 4];
                                                                                                                                                                                                                                                                                                                                                                                disk.Memory = Convert.ToString(range7.Value2);
                                                                                                                                                                                                                                                                                                                                                                                range7 = (Excel.Range)sheet7.Cells[ad, 5];
                                                                                                                                                                                                                                                                                                                                                                                disk.LevelOfNoise = range7.Value2;
                                                                                                                                                                                                                                                                                                                                                                                range7 = (Excel.Range)sheet7.Cells[ad, 6];
                                                                                                                                                                                                                                                                                                                                                                                disk.DataExchangeRate = Convert.ToInt32(range7.Value2);
                                                                                                                                                                                                                                                                                                                                                                                range7 = (Excel.Range)sheet7.Cells[ad, 7];
                                                                                                                                                                                                                                                                                                                                                                                disk.BufferSize = Convert.ToString(range7.Value2);
                                                                                                                                                                                                                                                                                                                                                                                range7 = (Excel.Range)sheet7.Cells[ad, 8];
                                                                                                                                                                                                                                                                                                                                                                                disk.Site = range7.Value2;

                                                                                                                                                                                                                                                                                                                                                                                for(int us=0;us<settings.HDDFabricator.Length;us++)
                                                                                                                                                                                                                                                                                                                                                                                {
                                                                                                                                                                                                                                                                                                                                                                                    if(settings.HDDFabricator[us]==disk.Fabricator)
                                                                                                                                                                                                                                                                                                                                                                                    {
                                                                                                                                                                                                                                                                                                                                                                                        for(int sus=0; sus<settings.HDDMemory.Length;sus++)
                                                                                                                                                                                                                                                                                                                                                                                        {
                                                                                                                                                                                                                                                                                                                                                                                            if(settings.HDDMemory[sus]==disk.Memory)
                                                                                                                                                                                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                                                                                                                                                                                if((disk.LevelOfNoise>=settings.HDDLevelOfNoise[0])&&(disk.LevelOfNoise <= settings.HDDLevelOfNoise[1]))
                                                                                                                                                                                                                                                                                                                                                                                                {
                                                                                                                                                                                                                                                                                                                                                                                                    if ((disk.DataExchangeRate >= settings.HDDDataExchangeRate[0]) && (disk.DataExchangeRate<= settings.HDDDataExchangeRate[1]))
                                                                                                                                                                                                                                                                                                                                                                                                    {
                                                                                                                                                                                                                                                                                                                                                                                                        for(int amg=0;amg<settings.HDDBufferSize.Length;amg++)
                                                                                                                                                                                                                                                                                                                                                                                                        {
                                                                                                                                                                                                                                                                                                                                                                                                            if(settings.HDDBufferSize[amg]==disk.BufferSize)
                                                                                                                                                                                                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                                                                                                                                                                                                BufferComp.hdd = disk;
                                                                                                                                                                                                                                                                                                                                                                                                                
                                                                                                                                                                                                                                                                                                                                                                                                                break;
                                                                                                                                                                                                                                                                                                                                                                                                            }
                                                                                                                                                                                                                                                                                                                                                                                                        }break;
                                                                                                                                                                                                                                                                                                                                                                                                    }
                                                                                                                                                                                                                                                                                                                                                                                                    else { break; }
                                                                                                                                                                                                                                                                                                                                                                                                }
                                                                                                                                                                                                                                                                                                                                                                                                else { break; }
                                                                                                                                                                                                                                                                                                                                                                                            }
                                                                                                                                                                                                                                                                                                                                                                                        } break;
                                                                                                                                                                                                                                                                                                                                                                                    }
                                                                                                                                                                                                                                                                                                                                                                                }

                                                                                                                                                                                                                                                                                                                                                                                ad++;
                                                                                                                                                                                                                                                                                                                                                                                range7 = (Excel.Range)sheet7.Cells[ad, 1];
                                                                                                                                                                                                                                                                                                                                                                            }

                                                                                                                                                                                                                                                                                                                                                                        }
                                                                                                                                                                                                                                                                                                                                                                    }
                                                                                                                                                                                                                                                                                                                                                                } break;
                                                                                                                                                                                                                                                                                                                                                            }
                                                                                                                                                                                                                                                                                                                                                        }break;
                                                                                                                                                                                                                                                                                                                                                    }
                                                                                                                                                                                                                                                                                                                                                }break;
                                                                                                                                                                                                                                                                                                                                            }
                                                                                                                                                                                                                                                                                                                                        }break;
                                                                                                                                                                                                                                                                                                                                    }
                                                                                                                                                                                                                                                                                                                                } break;
                                                                                                                                                                                                                                                                                                                            }
                                                                                                                                                                                                                                                                                                                        } //break;

                                                                                                                                                                                                                                                                                                                        op++;
                                                                                                                                                                                                                                                                                                                        range6 = (Excel.Range)sheet6.Cells[op, 1];
                                                                                                                                                                                                                                                                                                                    }
                                                                                                                                                                                                                                                                                                                }
                                                                                                                                                                                                                                                                                                            }break;
                                                                                                                                                                                                                                                                                                        }
                                                                                                                                                                                                                                                                                                    }break;
                                                                                                                                                                                                                                                                                                }
                                                                                                                                                                                                                                                                                            }break;
                                                                                                                                                                                                                                                                                        }
                                                                                                                                                                                                                                                                                    }
                                                                                                                                                                                                                                                                                }
                                                                                                                                                                                                                                                                                po++;
                                                                                                                                                                                                                                                                                range5 = (Excel.Range)sheet5.Cells[po, 4];
                                                                                                                                                                                                                                                                            }
                                                                                                                                                                                                                                                                        }
                                                                                                                                                                                                                                                                    } break;
                                                                                                                                                                                                                                                                }
                                                                                                                                                                                                                                                            }break;
                                                                                                                                                                                                                                                        }
                                                                                                                                                                                                                                                        else { break; }
                                                                                                                                                                                                                                                    }
                                                                                                                                                                                                                                                } break;
                                                                                                                                                                                                                                            }
                                                                                                                                                                                                                                        } break;
                                                                                                                                                                                                                                    }
                                                                                                                                                                                                                                } break;
                                                                                                                                                                                                                            }
                                                                                                                                                                                                                        } break;
                                                                                                                                                                                                                    }
                                                                                                                                                                                                                } break;
                                                                                                                                                                                                            }
                                                                                                                                                                                                        } n++; range4 = (Excel.Range)sheet4.Cells[n, 1]; //break;
                                                                                                                                                                                                    }
                                                                                                                                                                                                }
                                                                                                                                                                                            } break;
                                                                                                                                                                                        }
                                                                                                                                                                                    } break;
                                                                                                                                                                                }
                                                                                                                                                                            } break;
                                                                                                                                                                        }
                                                                                                                                                                    } break;
                                                                                                                                                                }
                                                                                                                                                            } //break;
                                                                                                                                                        }
                                                                                                                                                        l++;
                                                                                                                                                        range3 = (Excel.Range)sheet3.Cells[l, 6];
                                                                                                                                                    }
                                                                                                                                                }
                                                                                                                                            } break;
                                                                                                                                        }
                                                                                                                                    } break;
                                                                                                                                }
                                                                                                                            } break;
                                                                                                                        }
                                                                                                                    } break;
                                                                                                                }
                                                                                                            } break;
                                                                                                        }
                                                                                                    } break;
                                                                                                }
                                                                                            } break;
                                                                                        }
                                                                                    } break;
                                                                                }
                                                                            }
                                                                            //break;//
                                                                        }
                                                                        r++;
                                                                        range2 = (Excel.Range)sheet2.Cells[r, 5];
                                                                    }
                                                                    break;//
                                                                }
                                                            }
                                                            break;
                                                        }
                                                    }
                                                }
                                                else { break; }
                                            }
                                        }
                                        break;
                                    }
                                }
                                break;
                            }
                        }
                        break;
                    }
                }
                j++;
                range = (Excel.Range)sheet.Cells[j, 1];
            }
            //label1.Text = proc.Name;
            //label2.Text = proc.Fabricator;
            //double buffer = range.Value2;
            //label1.Text = Convert.ToString(buffer);
        }

        private void MotherBoardExtraction(MotherBoard Mom, int j)
        {

        }

        private Computer computer = new Computer();
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
    }
}
