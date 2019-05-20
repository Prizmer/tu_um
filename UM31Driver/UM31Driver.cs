using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Threading;
using System.Threading.Tasks;

using PollingLibraries.LibPorts;
using Drivers.LibMeter;




namespace Drivers.UMDriver
{
    public class UM31Driver : CMeter, UMDriver, IMeter
    {
        //пароль к самой ум
        string password = "00000000";

        const string SEPARATOR = ",";
        const int MESSAGE_TAIL_SIZE_BYTES = 6;

        public int softwareVersion = 22;

        Dictionary<ushort, string> currCorrelationDict = new Dictionary<ushort, string>();
        Dictionary<ushort, string> monthlyCorrelationDict = new Dictionary<ushort, string>();

        List<ValueUM> listOfCurrentValues = new List<ValueUM>();
        List<ValueUM> listOfMonthlyValues = new List<ValueUM>();

        public UM31Driver()
        {
            monthlyCorrelationDict.Add(0, "MA+");
            monthlyCorrelationDict.Add(1, "MA-");
            monthlyCorrelationDict.Add(2, "MR+");
            monthlyCorrelationDict.Add(3, "MR-");

            foreach (ushort k in monthlyCorrelationDict.Keys)
                currCorrelationDict.Add(k, monthlyCorrelationDict[k].Replace("M", ""));
        }

        int meterId = 1;
        bool meterIdParsingResult = false;
        UM_VERSION umVersion = UM_VERSION.UM31;

        UMCommon umCommon = new UMCommon();

        private List<byte> wrapCmd(string cmd, string prms = "")
        {
            List<byte> fullCmdList = new List<byte>();

            string CRCFeedString = this.password + SEPARATOR + cmd + prms;
            byte[] CRCFeedASCIIArr = ASCIIEncoding.ASCII.GetBytes(CRCFeedString);
            fullCmdList.AddRange(CRCFeedASCIIArr);

            byte[] CRCASCIIByteArr = umCommon.makeCRC(CRCFeedASCIIArr);
            string CRCASCIIStr = Encoding.ASCII.GetString(CRCASCIIByteArr).Replace("-", "");
            fullCmdList.AddRange(CRCASCIIByteArr);

            byte[] stopSignArr = new byte[] { 0x0A, 0x0A };
            string stopSignString = ASCIIEncoding.ASCII.GetString(stopSignArr);
            fullCmdList.AddRange(stopSignArr);

            //для наглядности
            string strASCIICmd = CRCFeedString + CRCASCIIStr + stopSignString;

            return fullCmdList;
        }

        #region Служебные

        public bool readDiagInfo()
        {
            List<byte> cmd = wrapCmd("RDIAGN");

            byte[] incommingData = new byte[1];
            m_vport.WriteReadData(FindPacketSignature, cmd.ToArray(), ref incommingData, cmd.Count, -1);

            if (incommingData.Length < MESSAGE_TAIL_SIZE_BYTES + 1) return false;
            byte[] cuttedIncommingData = new byte[incommingData.Length - MESSAGE_TAIL_SIZE_BYTES];
            Array.Copy(incommingData, 0, cuttedIncommingData, 0, cuttedIncommingData.Length);

            string answ = ASCIIEncoding.ASCII.GetString(cuttedIncommingData);
            WriteToLog("DIAGN: " + answ);

            if (!answ.Contains("ERROR")) return false;
            return true;
        }



        public bool readUMSerial(ref string serial_number)
        {
            List<byte> cmd = wrapCmd("UM_READ_ID");

            byte[] incommingData = new byte[1];
            m_vport.WriteReadData(FindPacketSignature, cmd.ToArray(), ref incommingData, cmd.Count, -1);

            if (incommingData.Length < MESSAGE_TAIL_SIZE_BYTES + 1) return false;
            byte[] cuttedIncommingData = new byte[incommingData.Length - MESSAGE_TAIL_SIZE_BYTES];
            Array.Copy(incommingData, 0, cuttedIncommingData, 0, cuttedIncommingData.Length);

            string answ = ASCIIEncoding.ASCII.GetString(cuttedIncommingData);
            WriteToLog("readUMSerial: " + answ);

            if (!answ.Contains("UM_ID")) return false;

            serial_number = answ.Replace("UM_ID=", "");

            return true;
        }

        public void sendAbort()
        {
            List<byte> cmd = wrapCmd("ABORT");
            byte[] incommingData = new byte[1];
            m_vport.WriteReadData(FindPacketSignature, cmd.ToArray(), ref incommingData, cmd.Count, -1);

        }

        public bool readUMName(out UM_VERSION um_version)
        {
            um_version = UM_VERSION.UNKNOWN;
            List<byte> cmd = wrapCmd("RDIAGN");

            byte[] incommingData = new byte[1];
            m_vport.WriteReadData(FindPacketSignature, cmd.ToArray(), ref incommingData, cmd.Count, -1);

            if (incommingData.Length < MESSAGE_TAIL_SIZE_BYTES + 1) return false;
            byte[] cuttedIncommingData = new byte[5];
            Array.Copy(incommingData, 0, cuttedIncommingData, 0, 5);

            string umName = ASCIIEncoding.ASCII.GetString(cuttedIncommingData);

            if (umName == "UM-31") um_version = UM_VERSION.UM31;
            else if (umName == "UM-40") um_version = UM_VERSION.UM40;

            return true;
        }

        public bool readSWVersion(out int version)
        {

            version = 22;

            List<byte> cmd = wrapCmd("GETSWVER");

            byte[] incommingData = new byte[1];
            m_vport.WriteReadData(FindPacketSignature, cmd.ToArray(), ref incommingData, cmd.Count, -1);

            if (incommingData.Length < MESSAGE_TAIL_SIZE_BYTES + 1) return false;
            byte[] cuttedIncommingData = new byte[incommingData.Length - MESSAGE_TAIL_SIZE_BYTES];
            Array.Copy(incommingData, 0, cuttedIncommingData, 0, cuttedIncommingData.Length);

            string answ = ASCIIEncoding.ASCII.GetString(cuttedIncommingData);
            WriteToLog("SoftwareVersion: " + answ);

            if (!answ.Contains("SW")) return false;

            return int.TryParse(answ.Replace("SW=", ""), out version);
        }

        #endregion

        #region Таблица приборов

        public bool getMetersTableEntriesNumber(out int cnt)
        {
            int emergencyNumberOfRecords = 1;
            cnt = emergencyNumberOfRecords;

            // длина таблицы - кол-во приборов, подключенных к УМ,
            // воспользуемся READITEMPARAM (настройка измеряемых параметров)
            List<byte> cmd = wrapCmd("READITEMPARAM");

            int attempts = 0;

            TRY_AGAIN:

            byte[] incommingData = new byte[1];
            m_vport.WriteReadData(FindPacketSignature, cmd.ToArray(), ref incommingData, cmd.Count, -1);

            string str = ASCIIEncoding.ASCII.GetString(incommingData);
            WriteToLog("READITEMPARAM, att " + attempts + ": " + str);

            // ITEMPARAM=010;[TARIF];[AP];[AM];[RP];[RM]
            if (!str.Contains("ITEMPARAM") || incommingData.Length < 13)
            {
                if (attempts < 2)
                {
                    m_vport.Close();
                    attempts++;
                    Thread.Sleep(2000);
                    goto TRY_AGAIN;
                }
                else
                {
                    WriteToLog("READITEMPARAM, att " + attempts + 
                        ": от прибора не получен ответ ITEMPARAM= в требуемом виде..: " + str);

                    return false;
                }
            }


            string[] strItems = str.Replace("ITEMPARAM=", "").Split(';');
    
            if (strItems.Length == 0)
            {
                WriteToLog("READITEMPARAM: strItems.Length == 0");
                return false;
            }

            // ITEMPARAM=[LEN];[TARIF];[AP];[AM];[RP];[RM]
            string countStr = strItems[0];
            if (!int.TryParse(countStr, out cnt)) return false;

            return true;
        }

        private bool parseMetersTableEntry(byte[] entry, out List<string> metersTableEntryList, ref MetersTableEntry metersTableEntry)
        {
            metersTableEntry = new MetersTableEntry();
            metersTableEntryList = new List<string>();

            string str = ASCIIEncoding.ASCII.GetString(entry);
            if (!str.Contains("TABLE")) return false;

            string[] strItems = str.Replace("TABLE=", "").Split(';');

            WriteToLog("parseMetersTableEntry, entry: " + str);

            if (strItems.Length < 5) return false;

            int moveIndex = 0;
            //if (softwareVersion < 22) moveIndex = 1;


            metersTableEntry.id = int.Parse(strItems[0]);
            metersTableEntry.meterName = ((MeterModels)int.Parse(strItems[moveIndex + 3])).ToString();
            metersTableEntry.networkAddr = int.Parse(strItems[moveIndex + 1]);
            //metersTableEntry.interfaceType = ((InterfaceModels)int.Parse(strItems[moveIndex + 2])).ToString();
            metersTableEntry.pass1 = strItems[moveIndex + 5];
            //metersTableEntry.pass2 = strItems[moveIndex + 4];

            metersTableEntryList.Add(metersTableEntry.id.ToString());
            metersTableEntryList.Add(metersTableEntry.meterName);
            metersTableEntryList.Add(metersTableEntry.networkAddr.ToString());
            // metersTableEntryList.Add(metersTableEntry.interfaceType);
            metersTableEntryList.Add(metersTableEntry.pass1);
            //  metersTableEntryList.Add(metersTableEntry.pass2);

            return true;
        }

        public bool getMetersTable(ref DataTable metersTable)
        {
            if (!OpenLinkCanal()) return false;

            // readSWVersion(out softwareVersion);

            int metersCount = 0;
            if (!getMetersTableEntriesNumber(out metersCount)) return false;

            metersTable = new DataTable();
            DataTable metersDt = metersTable;

            DataColumn tmpCol = new DataColumn();
            tmpCol.Caption = "ID";
            tmpCol.ColumnName = "colID";
            metersDt.Columns.Add(tmpCol);

            tmpCol = new DataColumn();
            tmpCol.Caption = "Модель";
            tmpCol.ColumnName = "colMeterModel";
            metersDt.Columns.Add(tmpCol);

            tmpCol = new DataColumn();
            tmpCol.Caption = "Сетевой адрес";
            tmpCol.ColumnName = "colNetwAddr";
            metersDt.Columns.Add(tmpCol);

            //tmpCol = new DataColumn();
            //tmpCol.Caption = "Интерфейс";
            //tmpCol.ColumnName = "colInterfaceName";
            //metersDt.Columns.Add(tmpCol);

            tmpCol = new DataColumn();
            tmpCol.Caption = "П1";
            tmpCol.ColumnName = "colPass1";
            metersDt.Columns.Add(tmpCol);

            //tmpCol = new DataColumn();
            //tmpCol.Caption = "П2";
            //tmpCol.ColumnName = "colPass2";
            //metersDt.Columns.Add(tmpCol);

            DataRow captionRow = metersDt.NewRow();
            for (int i = 0; i < metersDt.Columns.Count; i++)
                captionRow[i] = metersDt.Columns[i].Caption;
            metersDt.Rows.Add(captionRow);

            for (int i = 0; i < metersCount; i++)
            {
                DataRow dRow = metersDt.NewRow();
                List<byte> cmd = wrapCmd("READTABL=" + i);

                byte[] incommingData = new byte[1];
                m_vport.WriteReadData(FindPacketSignature, cmd.ToArray(), ref incommingData, cmd.Count, -1);
                string answ = ASCIIEncoding.ASCII.GetString(incommingData);

                WriteToLog("getMetersTable,  запись таблицы №" + i + ": " + answ);

                if (!answ.Contains("TABL")) continue;

                MetersTableEntry mte = new MetersTableEntry();
                List<string> mteList = new List<string>();

                if (!parseMetersTableEntry(incommingData, out mteList, ref mte)) continue;
                if (mteList.Count == 0) continue;

                for (int j = 0; j < mteList.Count; j++)
                    dRow[j] = mteList[j];

                metersDt.Rows.Add(dRow);
            }

            if (metersDt.Rows.Count == 0) return false;


            return true;
        }

        #endregion


        #region Реализация интерфейса СО

        public void Init(uint address, string pass, VirtualPort data_vport)
        {
            m_address = address;
            this.m_vport = data_vport;

            meterIdParsingResult = true;
            meterIdParsingResult = int.TryParse(pass, out meterId);

            if (!meterIdParsingResult)
            {
                meterId = (int)address;
                WriteToLog("Init драйвера, не удалось распарсить пароль, использую сетевой адрес " + meterId);
            }
            else
            {
                WriteToLog("Init драйвера, использую то, что в пароле " + meterId);
            }


            //очистка временных списков ОЧЕНь ВАЖНО для данного прибора
            listOfCurrentValues.Clear();
            listOfMonthlyValues.Clear();
        }

        private int FindPacketSignature(Queue<byte> queue)
        {
            //определим конец пакета

            if (queue.Count < 2) return 0;

            byte last = queue.Dequeue();
            byte preLast = queue.Dequeue();
            if (last == preLast && preLast == 0x0a) return 1;
            else
                return 0;
        }

        public bool OpenLinkCanal()
        {
            for (int i = 0; i < 1; i++)
            {
                if (readDiagInfo()) return true;
            }

            WriteToLog("OpenLinkCanal: false");
            return false;
        }



        public bool ReadCurrentValues(ushort param, ushort tarif, ref float recordValue)
        {
            return false;
        }

        public bool ReadDailyValues(DateTime dt, ushort param, ushort tarif, ref float recordValue)
        {
            return false;
        }

        public bool ReadMonthlyValues(DateTime dt, ushort param, ushort tarif, ref float recordValue)
        {
            return false;
        }

        public bool ReadSerialNumber(ref string serial_number)
        {
            return false;
        }

        #endregion

        #region Неиспользуемые методы интерфейса

        public List<byte> GetTypesForCategory(CommonCategory common_category)
        {
            return null;
        }

        public bool ReadDailyValues(uint recordId, ushort param, ushort tarif, ref float recordValue)
        {
            return false;
        }

        public bool ReadPowerSlice(DateTime dt_begin, DateTime dt_end, ref List<RecordPowerSlice> listRPS, byte period)
        {
            return false;
        }

        public bool ReadPowerSlice(ref List<SliceDescriptor> sliceUniversalList, DateTime dt_end, SlicePeriod period)
        {
            return false;
        }

        public bool ReadSliceArrInitializationDate(ref DateTime lastInitDt)
        {
            return false;
        }

        public bool SyncTime(DateTime dt)
        {
            return false;
        }

        #endregion

        #region Неиспользуемые публичные методы поддержки драйвера UMRTU


        public bool ReadDailyValues2(ushort param, ushort tarif, ref float recordValue)
        {
            throw new NotImplementedException();
        }

        public bool getSlicesValuesForID(int id, DateTime dt_start, DateTime dt_end, out List<RecordPowerSlice> rpsVals)
        {
            throw new NotImplementedException();
        }

        public bool getDailyValuesForID(int id, DateTime dt, out List<ValueUM> umVals)
        {
            throw new NotImplementedException();
        }

        public bool getDailyValuesForID(int id, out List<ValueUM> umVals)
        {
            throw new NotImplementedException();
        }

        public bool parseSingleSliceString(string sliceString, ref RecordPowerSlice powerSlice)
        {
            throw new NotImplementedException();
        }

        #endregion

    }


}
