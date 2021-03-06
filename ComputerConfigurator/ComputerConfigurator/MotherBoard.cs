using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComputerConfigurator
{
    class MotherBoard
    {
        public string Site { get; set; }

        /// <summary>
        /// Название
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Цена комплектующего
        /// </summary>
        public double Price { get; set; }

        /// <summary>
        /// Производитель
        /// </summary>
        public string Fabricator { get; set; }


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
        public int NumberOfMemorySlots { get; set; }

        /// <summary>
        /// Количество слотов PCI-E x16
        /// </summary>
        public int NumberOfPCIEx16Slots { get; set; }

        /// <summary>
        /// Количество слотов М.2
        /// </summary>
        public int NumberOfM2Slots { get; set; }

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
