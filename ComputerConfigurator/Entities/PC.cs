using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComputerConfigurator
{
    class PC
    {
        public MotherBoard motherBoard = new MotherBoard();
        public CPU processor = new CPU();
        public RAM ram = new RAM();
        public GraphicsCard graphicsCard = new GraphicsCard();
        public PowerSupply powerSupply = new PowerSupply();
        public Corps corps = new Corps();
        public HDD hdd = new HDD();
        public SSD ssd = new SSD();

        /// <summary>
        /// Цена ПК
        /// </summary>
        public double Price { get; set; }
    }
}
