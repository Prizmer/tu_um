using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

using Drivers.LibMeter;

namespace Drivers.UMDriver
{
    public class UMCommon
    {
        #region CRC Вычисление

        private UInt16 crc16_update_poly(UInt16 crc, byte a)
        {
            crc ^= a;
            for (byte i = 0; i < 8; ++i)
            {
                if ((crc & 1) == 1)
                    crc = (UInt16)((crc >> 1) ^ 0xA001);
                else
                    crc = (UInt16)(crc >> 1);
            }
            return crc;
        }
        private UInt16 crc16_calc_poly(byte[] buf, int len, UInt16 crc)
        {

            for (int i = 0; i < len; i++)
                crc = crc16_update_poly(crc, buf[i]);

            return crc;
        }

        public byte[] StringToByteArray(string hex)
        {
            return Enumerable.Range(0, hex.Length)
                                .Where(x => x % 2 == 0)
                                .Select(x => Convert.ToByte(hex.Substring(x, 2), 16))
                                .ToArray();
        }

        public byte[] makeCRC(byte[] buf)
        {
            UInt16 crc = 0x40BF;
            UInt16 resCrcNumb = crc16_calc_poly(buf, buf.Length, crc);

            char[] CRCCharArr = Convert.ToString(resCrcNumb, 16).ToUpper().ToCharArray();
            int zerosNeeded = 4 - CRCCharArr.Length;

            List<char> charList = new List<char>();
            for (int i = 0; i < zerosNeeded; i++)
                charList.Add('0');
            charList.AddRange(CRCCharArr);

            CRCCharArr = charList.ToArray();

            List<byte> CRCByteList = new List<byte>();
            byte[] CRCASCIICharBytes = Encoding.ASCII.GetBytes(CRCCharArr);

            CRCByteList.Add(CRCASCIICharBytes[2]);
            CRCByteList.Add(CRCASCIICharBytes[3]);
            CRCByteList.Add(CRCASCIICharBytes[0]);
            CRCByteList.Add(CRCASCIICharBytes[1]);

            return CRCByteList.ToArray();
        }

        #endregion


    }


    public enum UM_VERSION
    {
        UM31,
        UM40,
        UNKNOWN
    }

    public enum MeterModels
    {
        USPD = 0,
        M200 = 1,
        M230 = 3,
        SET4TM = 4,
        M203 = 31,
        M206 = 32,
        MConcentrator = 91,
        PulsarRadio = 93,
        Impulse = 121
    }

    public enum InterfaceModels
    {
        CAN1,
        CAN2,
        CAN3,
        RS485_2,
        RS485_1,
        RS232,
        CONCENTRATOR
    }


    public struct ValueUM
    {
        public float value;
        public DateTime dt;
        public int address;
        public int channel;
        public string caption;
        public string name;
        public string meterSN;
    }

    public struct MetersTableEntry
    {
        public int id;
        public int networkAddr;
        public string interfaceType;

        public string meterName;
  
        public string passFormat;
        public string pass1;
        public string pass2;
    }

    public interface UMDriver : IMeter
    {
        bool ReadDailyValues2(ushort param, ushort tarif, ref float recordValue);
        bool getSlicesValuesForID(int id, DateTime dt_start, DateTime dt_end, out List<RecordPowerSlice> rpsVals);
        bool getMetersTableEntriesNumber(out int cnt);
        bool getDailyValuesForID(int id, DateTime dt, out List<ValueUM> umVals);
        bool getDailyValuesForID(int id, out List<ValueUM> umVals);
        bool parseSingleSliceString(string sliceString, ref RecordPowerSlice powerSlice);
        bool getMetersTable(ref DataTable metersTable);
    }

}
