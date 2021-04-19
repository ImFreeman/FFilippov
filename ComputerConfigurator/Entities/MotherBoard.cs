using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComputerConfigurator
{
    class MotherBoard : ComputerComponent
    {


        /// <summary>
        /// Вид процессора (AMD или Intel)
        /// </summary>
        public string CPUType { get; set; }

        /// <summary>
        /// Тип памяти RAM
        /// </summary>
        public string MemoryType { get; set; }

        /// <summary>
        /// Количество слотов памяти RAM
        /// </summary>
        public string NumberOfMemorySlots { get; set; }

        /// <summary>
        /// Количество слотов PCI-E x16
        /// </summary>
        public string NumberOfPCIEx16Slots { get; set; }

        /// <summary>
        /// Количество слотов М.2
        /// </summary>
        public string NumberOfM2Slots { get; set; }

        /// <summary>
        /// Наличие Wi-Fi адаптера
        /// </summary>
        public string WiFiAdapter { get; set; }

        /// <summary>
        /// Наличие встроенного CPU
        /// </summary>
        public string BuildInCPU { get; set; }

        /// <summary>
        /// Для игрового ПК
        /// </summary>
        public string ForGamingPC { get; set; }

        /// <summary>
        /// Сокет процессора
        /// </summary>
        public string Socket { get; set; }

        /// <summary>
        /// Чипсет процессора
        /// </summary>
        public string Chipset { get; set; }
    }
}
