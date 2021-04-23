using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComputerConfigurator
{
    /// <summary>
    /// Сборка персонального компьютера
    /// </summary>
    class PC
    {
        public MotherBoard motherBoard = new MotherBoard()
        {
            Name="NULL"
        };
        public CPU processor = new CPU()
        {
            Name = "NULL",
            Price = 0
        };
        public RAM ram = new RAM()
        {
            Name = "NULL"
        };
        public GraphicsCard graphicsCard = new GraphicsCard()
        {
            Name = "NULL"
        };
        public PowerSupply powerSupply = new PowerSupply()
        {
            Name = "NULL"
        };
        public Corps corps = new Corps()
        {
            Name = "NULL"
        };
        public HDD hdd = new HDD()
        {
            Name = "NULL",
            Price=0
        };
        public SSD ssd = new SSD()
        {
            Name = "NULL",
            Price=0
        };
        public double TotalPrice()
        {
            double totalPrice = 0;
            totalPrice += motherBoard.Price;
            totalPrice += processor.Price;
            totalPrice += ram.Price;
            totalPrice += graphicsCard.Price;
            totalPrice += powerSupply.Price;
            totalPrice += corps.Price;
            totalPrice += hdd.Price;
            totalPrice += ssd.Price;
            return totalPrice;
        }

        /// <summary>
        /// Цена ПК
        /// </summary>
    }
}
