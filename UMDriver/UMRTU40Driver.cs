﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;

using PollingLibraries.LibPorts;
using Drivers.LibMeter;

using System.Threading;

namespace Drivers.UMDriver
{
    public class UMRTU40Driver: CMeter, IMeter, UMDriver
    {
        //пароль к самой ум
        string password = "00000000";

        const string SEPARATOR = ",";
        const int MESSAGE_TAIL_SIZE_BYTES = 6;

        Dictionary<ushort, string> currCorrelationDict = new Dictionary<ushort, string>();
        Dictionary<ushort, string> dailyCorrelationDict = new Dictionary<ushort, string>();
        Dictionary<ushort, string> monthlyCorrelationDict = new Dictionary<ushort, string>();
        Dictionary<ushort, string> slicesCorrelationDict = new Dictionary<ushort, string>();

        UMCommon umCommon = new UMCommon();

        public UMRTU40Driver()
        {
            dailyCorrelationDict.Add(0, "dA+");
            dailyCorrelationDict.Add(1, "dA-");
            dailyCorrelationDict.Add(2, "dR+");
            dailyCorrelationDict.Add(3, "dR-");

            foreach (ushort k in dailyCorrelationDict.Keys)
                monthlyCorrelationDict.Add(k, dailyCorrelationDict[k].Replace("d", "M"));


            foreach (ushort k in dailyCorrelationDict.Keys)
                currCorrelationDict.Add(k, dailyCorrelationDict[k].Replace("d", ""));


            slicesCorrelationDict.Add(0, "DPAp");
            slicesCorrelationDict.Add(1, "DPAm");
            slicesCorrelationDict.Add(2, "DPRp");
            slicesCorrelationDict.Add(3, "DPRm");

        }

        private void FormSelectLocalIp_Load(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        int meterId = 1;
        bool meterIdParsingResult = false;
        UM_VERSION umVersion = UM_VERSION.UM40;

        public void Init(uint address, string pass, VirtualPort data_vport)
        {
            m_address = address;
            this.m_vport = data_vport;

            meterIdParsingResult = true;


            //
            meterIdParsingResult = int.TryParse(pass, out meterId);

            if (!meterIdParsingResult)
            {
                meterId = (int)address;
                WriteToLog("Init драйвера, не удалось распарсить пароль, использую сетевой адрес " + meterId);
            } else
            {
                WriteToLog("Init драйвера, использую то, что в пароле " + meterId);
            }


            //this.password = pass.Length > 0 ? pass : "00000000";



            //очистка временных списков ОЧЕНь ВАЖНО для данного прибора
            listOfDailyValues.Clear();
            listOfMonthlyValues.Clear();
        }



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



        public int softwareVersion = 22;


        #region Служебные

        public void sendAbort()
        {
            List<byte> cmd = wrapCmd("ABORT");
            byte[] incommingData = new byte[1];
            m_vport.WriteReadData(FindPacketSignature, cmd.ToArray(), ref incommingData, cmd.Count, -1);

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
            int emergencyNumberOfRecords = 4;

            cnt = 0;
            //определим длину таблицы с приборами
            List<byte> cmd = wrapCmd("RTABLLEN");

            int attempts = 0;

            TRY_AGAIN:

            byte[] incommingData = new byte[1];
            m_vport.WriteReadData(FindPacketSignature, cmd.ToArray(), ref incommingData, cmd.Count, -1);



            if (incommingData.Length < 11)
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
                    cnt = emergencyNumberOfRecords;
                    return true;
                }

            }

            string answ = ASCIIEncoding.ASCII.GetString(incommingData);
            WriteToLog("getMetersTableEntriesNumber: " + answ);

            if (!answ.Contains("TABLLEN"))
            {
                return false;
            }

            string countStr = answ.Substring(8, 3);
            if (!int.TryParse(countStr, out cnt)) return false;

            return true;
        }

        private bool parseMetersTableEntry(byte[] entry, out List<string> metersTableEntryList, ref MetersTableEntry metersTableEntry)
        {
            metersTableEntry = new MetersTableEntry();
            metersTableEntryList = new List<string>();

            string str = ASCIIEncoding.ASCII.GetString(entry);
            if (!str.Contains("TABLEX")) return false;

            string[] strItems = str.Replace("TABLEX=", "").Split(';');

            WriteToLog(str);

            if (strItems.Length < 5) return false;

            int moveIndex = 0;
            if (softwareVersion < 22) moveIndex = 1;


            metersTableEntry.id = int.Parse(strItems[0]);
            metersTableEntry.meterName = ((MeterModels)int.Parse(strItems[moveIndex + 1])).ToString();
            metersTableEntry.networkAddr = int.Parse(strItems[moveIndex + 2]);
            metersTableEntry.interfaceType = ((InterfaceModels)int.Parse(strItems[moveIndex + 6])).ToString();
            metersTableEntry.pass1 = strItems[moveIndex + 3];
            metersTableEntry.pass2 = strItems[moveIndex + 4];

            metersTableEntryList.Add(metersTableEntry.id.ToString());
            metersTableEntryList.Add(metersTableEntry.meterName);
            metersTableEntryList.Add(metersTableEntry.networkAddr.ToString());
            // metersTableEntryList.Add(metersTableEntry.interfaceType);
            //  metersTableEntryList.Add(metersTableEntry.pass1);
            //  metersTableEntryList.Add(metersTableEntry.pass2);

            return true;
        }

        public bool getMetersTable(ref DataTable metersTable)
        {
            if (!OpenLinkCanal()) return false;

            readSWVersion(out softwareVersion);

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

            //tmpCol = new DataColumn();
            //tmpCol.Caption = "П1";
            //tmpCol.ColumnName = "colPass1";
            //metersDt.Columns.Add(tmpCol);

            //tmpCol = new DataColumn();
            //tmpCol.Caption = "П2";
            //tmpCol.ColumnName = "colPass2";
            //metersDt.Columns.Add(tmpCol);

            DataRow captionRow = metersDt.NewRow();
            for (int i = 0; i < metersDt.Columns.Count; i++)
                captionRow[i] = metersDt.Columns[i].Caption;
            metersDt.Rows.Add(captionRow);

            //счет с 1го у них
            for (int i = 0; i < metersCount; i++)
            {
                DataRow dRow = metersDt.NewRow();
                List<byte> cmd = wrapCmd("READTABLEX=" + i);
                int attempts = 0;

                TRY_AGAIN:
                byte[] incommingData = new byte[1];
                m_vport.WriteReadData(FindPacketSignature, cmd.ToArray(), ref incommingData, cmd.Count, -1);
                string answ = ASCIIEncoding.ASCII.GetString(incommingData);

                WriteToLog("getMetersTable, сам запрос таблицы: " + answ);

                if (!answ.Contains("TABLEX")) continue;



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

        #region Значения

        private bool getDateFromAnswString(string answ, ref DateTime date)
        {
            //получает первую дату из поданой части ответа

            date = new DateTime();

            try
            {
                string pattern = "DT\\s.*<";
                Regex reg = new Regex(pattern);
                string res = reg.Match(answ).Groups[0].Value.Replace("DT\n", "").Replace("<", "");

                CultureInfo provider = CultureInfo.InvariantCulture;
                string dateStr = res.Remove(res.Length - 5);
                string syncrFlagStr = res.Remove(0, res.Length - 4).Remove(2, 2);
                string winterFlagStr = res.Remove(0, res.Length - 1);

                bool areClockSincronized = syncrFlagStr == "02" ? true : false;
                bool isWinterTime = winterFlagStr == "0" ? true : false;

                DateTime dt = DateTime.ParseExact(dateStr, "dd.MM.yy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);

                date = dt;

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private bool getMetersSNFromAnswString(string answ, ref string serialNumberStr)
        {
            string pattern = ";\\d*<";

            try
            {
                Regex reg = new Regex(pattern);
                serialNumberStr = reg.Match(answ).Groups[0].Value.Replace(";", "").Replace("<", "");
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private bool getRecordsDictionary(string answ, ref Dictionary<string, float> records, float coefficient = 1000)
        {
            Dictionary<string, float> tmpDict = new Dictionary<string, float>();

            string startSign = "\n<";
            string endSign = "<";

            int index = answ.IndexOf(startSign);

            while (index != -1)
            {
                int valFirstLetterIndex = index + startSign.Length;

                int secondIndex = answ.IndexOf(endSign, valFirstLetterIndex + 1);
                if (secondIndex == -1) break;

                int valLastLetterIndex = secondIndex - endSign.Length;
                int valStrLength = valLastLetterIndex - valFirstLetterIndex + 1;

                string tmpRecordStr = answ.Substring(valFirstLetterIndex, valStrLength);
                string[] tmpRecordArr = tmpRecordStr.Split('\n');

                string tmpValue = tmpRecordArr[1];
                float resVal = -1;

                if (!float.TryParse(tmpValue, NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out resVal))
                    if (tmpValue != "?")
                        return false;

                tmpDict.Add(tmpRecordArr[0], resVal / coefficient);

                index = answ.IndexOf(startSign, secondIndex);
            }

            records = tmpDict;
            return true;

        }

        //используется в обработчике интерфейса
        public bool getDailyValuesForID(int id, DateTime dt, out List<ValueUM> umVals)
        {
            umVals = new List<ValueUM>();

            string cmdStr = "READDAY=" + dt.ToString("yy") + "." + dt.ToString("MM") + "." + dt.ToString("dd") +
                ";" + id;
            List<byte> cmd = wrapCmd(cmdStr);

            byte[] incommingData = new byte[1];
            m_vport.WriteReadData(FindPacketSignature, cmd.ToArray(), ref incommingData, cmd.Count, -1);

            string answ = ASCIIEncoding.ASCII.GetString(incommingData);
            // WriteToLog("getDailyValuesForID_ANSW: " + answ); 


            List<string> recordStringsForDates = new List<string>();

            int endIndex = answ.IndexOf("\nEND");
            if (endIndex == -1) return false;

            string tmpMeterSerial = "";
            getMetersSNFromAnswString(answ, ref tmpMeterSerial);


            int indexDt = answ.IndexOf("<DT");
            while (indexDt != -1)
            {
                int tmpIndexDt = answ.IndexOf("<DT", indexDt + 1);
                string tmpVal = "";
                try
                {           
                    if (tmpIndexDt == -1)
                    {
                        tmpVal = answ.Substring(indexDt, endIndex - indexDt + 1);
                    }
                    else
                    {
                        tmpVal = answ.Substring(indexDt, tmpIndexDt - indexDt + 1);
                    }
                }
                catch (Exception ex)
                {
                    // WriteToLog("getDailyValuesForID, ex: " + ex.Message);
                    return false ;
                }

                indexDt = tmpIndexDt;
                recordStringsForDates.Add(tmpVal);
            }

            if (recordStringsForDates.Count > 1)
                WriteToLog("Суточные: на данную дату пришло несколько значений, возможно расходятся часы");
            if (recordStringsForDates.Count == 0) return false;

            DateTime recordDt = new DateTime();
            if (!getDateFromAnswString(recordStringsForDates[0], ref recordDt)) return false;

            string selectedRecordString = recordStringsForDates[0];

            //получим блок TD
            string tdStartSign = "<TD\n";
            int tdIndex = selectedRecordString.IndexOf(tdStartSign);

            int secondIndex = selectedRecordString.IndexOf(tdStartSign, tdIndex + tdStartSign.Length);
            if (secondIndex == -1)
            {
                secondIndex = endIndex;
            }
            else
            {
                WriteToLog("Внимание, несколько тегов TD!");
            }

            string tdString = selectedRecordString.Substring(tdIndex, selectedRecordString.Length - tdIndex - 1);

            Dictionary<string, float> recordsDict = new Dictionary<string, float>();
            if (!getRecordsDictionary(tdString, ref recordsDict)) return false;


            //if (recordsDict.Count != 20) return false;

            int cnt = 0;

            foreach (string s in recordsDict.Keys)
            {
                ValueUM tmpVal = new ValueUM();
                tmpVal.dt = recordDt;
                tmpVal.name = s;
                tmpVal.value = recordsDict[s];
                tmpVal.meterSN = tmpMeterSerial;

                umVals.Add(tmpVal);
                cnt++;
            }

            if (umVals.Count > 0)
            {
                WriteToLog("getDailyValuesForID, val0: " + umVals[0].dt + ": " + umVals[0].value);
                return true;
            }
            else
            {
                WriteToLog("getDailyValuesForID, umVals.Count < 0");
                return false;
            }
        }

        public bool getDailyValuesForID(int id, out List<ValueUM> umVals)
        {
            umVals = new List<ValueUM>();

            string cmdStr = "READCENG=" + id;
            List<byte> cmd = wrapCmd(cmdStr);

            byte[] incommingData = new byte[1];
            m_vport.WriteReadData(FindPacketSignature, cmd.ToArray(), ref incommingData, cmd.Count, -1);

            string answ = ASCIIEncoding.ASCII.GetString(incommingData);

            WriteToLog(answ);

            List<string> recordStringsForDates = new List<string>();

            int endIndex = answ.IndexOf("\nEND");
            if (endIndex == -1)
            {
                //может быть из-за слишком малого ожидания: лучше ставить ожидание порядка 10 секунд 
                //и расчитывать на то, что признак конца ответа остановит цикл ожидания
                WriteToLog("Суточные: не обнаружен признак конца сообщения, возможно следует увеличить время ожидания ответа");
                return false;
            }

            string tmpMeterSerial = "";
            getMetersSNFromAnswString(answ, ref tmpMeterSerial);


            int indexDt = answ.IndexOf("<DT");
            while (indexDt != -1)
            {
                int tmpIndexDt = answ.IndexOf("<DT", indexDt + 1);
                string tmpVal = "";
                if (tmpIndexDt == -1)
                {
                    tmpVal = answ.Substring(indexDt, endIndex - indexDt + 1);
                }
                else
                {
                    tmpVal = answ.Substring(indexDt, tmpIndexDt - indexDt + 1);
                }

                indexDt = tmpIndexDt;
                recordStringsForDates.Add(tmpVal);
            }

            if (recordStringsForDates.Count > 1)
                WriteToLog("Суточные: на данную дату пришло несколько значений, возможно расходятся часы");
            if (recordStringsForDates.Count == 0) return false;

            DateTime recordDt = new DateTime();
            if (!getDateFromAnswString(recordStringsForDates[0], ref recordDt)) return false;

            string selectedRecordString = recordStringsForDates[0];

            //получим блок TD
            string tdStartSign = "<TD\n";
            int tdIndex = selectedRecordString.IndexOf(tdStartSign);

            int secondIndex = selectedRecordString.IndexOf(tdStartSign, tdIndex + tdStartSign.Length);
            if (secondIndex == -1)
            {
                secondIndex = endIndex;
            }
            else
            {
                WriteToLog("Внимание, несколько тегов TD!");
            }

            string tdString = selectedRecordString.Substring(tdIndex, selectedRecordString.Length - tdIndex - 1);


            Dictionary<string, float> recordsDict = new Dictionary<string, float>();
            if (!getRecordsDictionary(tdString, ref recordsDict)) return false;


            //why?
            //if (recordsDict.Count != 20) return false;

            int cnt = 0;

            foreach (string s in recordsDict.Keys)
            {
                ValueUM tmpVal = new ValueUM();
                tmpVal.dt = recordDt;
                tmpVal.name = s;
                tmpVal.value = recordsDict[s];
                tmpVal.meterSN = tmpMeterSerial;

                umVals.Add(tmpVal);
                cnt++;
            }

            if (umVals.Count > 0)
                return true;
            else
                return false;
        }

        public bool getMonthlyValuesForID(int id, DateTime dt, out List<ValueUM> umVals)
        {
            umVals = new List<ValueUM>();

            string cmdStr = "READMONTH=" + dt.ToString("MM") + "." + dt.ToString("yy") +
                ";" + id;
            List<byte> cmd = wrapCmd(cmdStr);

            byte[] incommingData = new byte[1];
            m_vport.WriteReadData(FindPacketSignature, cmd.ToArray(), ref incommingData, cmd.Count, -1);

            string answ = ASCIIEncoding.ASCII.GetString(incommingData);

            List<string> recordStringsForDates = new List<string>();

            int endIndex = answ.IndexOf("\nEND");
            if (endIndex == -1) return false;

            string tmpMeterSerial = "";
            getMetersSNFromAnswString(answ, ref tmpMeterSerial);

            int indexDt = answ.IndexOf("<DT");
            while (indexDt != -1)
            {
                int tmpIndexDt = answ.IndexOf("<DT", indexDt + 1);
                string tmpVal = "";
                if (tmpIndexDt == -1)
                {
                    tmpVal = answ.Substring(indexDt, endIndex - indexDt + 1);
                }
                else
                {
                    tmpVal = answ.Substring(indexDt, tmpIndexDt - indexDt + 1);
                }

                indexDt = tmpIndexDt;
                recordStringsForDates.Add(tmpVal);
            }

            if (recordStringsForDates.Count > 1)
                WriteToLog("Месячные: на данную дату пришло несколько значений, возможно расходятся часы");
            if (recordStringsForDates.Count == 0) return false;

            DateTime recordDt = new DateTime();
            if (!getDateFromAnswString(recordStringsForDates[0], ref recordDt)) return false;

            string selectedRecordString = recordStringsForDates[0];

            //получим блок TD
            string tdStartSign = "<TD\n";
            int tdIndex = selectedRecordString.IndexOf(tdStartSign);

            int secondIndex = selectedRecordString.IndexOf(tdStartSign, tdIndex + tdStartSign.Length);
            if (secondIndex == -1)
            {
                secondIndex = endIndex;
            }
            else
            {
                WriteToLog("Внимание, несколько тегов TD!");
            }

            string tdString = selectedRecordString.Substring(tdIndex, selectedRecordString.Length - tdIndex - 1);

            Dictionary<string, float> recordsDict = new Dictionary<string, float>();
            if (!getRecordsDictionary(tdString, ref recordsDict)) return false;


            if (recordsDict.Count != 20) return false;

            int cnt = 0;

            foreach (string s in recordsDict.Keys)
            {
                if (cnt == 5) break;
                if (!s.Contains("MA+")) continue;

                ValueUM tmpVal = new ValueUM();
                tmpVal.dt = recordDt;
                tmpVal.name = s;
                tmpVal.value = recordsDict[s];
                tmpVal.meterSN = tmpMeterSerial;

                umVals.Add(tmpVal);
                cnt++;
            }


            //разбор значений

            return true;
        }

        //добавлены при интеграции получасовок
        public bool parseSingleSliceString(string sliceString, ref RecordPowerSlice powerSlice)
        {
            powerSlice = new RecordPowerSlice();

            //содержит символ \n (0x0A)...от него надо избавиться в данном блоке
            sliceString = sliceString.Replace("\n", "");

            //здесь происходит разбор фрагмента ответа от DT до следующего DT
            //sliceString = "<DT.27.07.17 00:00:00 02 0<.<TD.1<.<FL.W;;;<.<DPAp.0.0025<.<DPAm.?<.<DPRp.0.0035<.<DPRm.?<.";
            int dateIdx = sliceString.IndexOf("DT");
            string dateStr = sliceString.Substring(dateIdx + 2, 17);

            //!!! дата не учитывает сезон лето/зима, хотя драйвер присылает эту информацию
            DateTime sliceDt = new DateTime();

            if (!DateTime.TryParseExact(dateStr, "dd.MM.yy HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out sliceDt))
            {
                WriteToLog("Ошибка преобразования даты в методе parseSingleSliceString: " + sliceString);
                WriteToLog("parseSingleSliceString bytes: " + BitConverter.ToString(ASCIIEncoding.ASCII.GetBytes(sliceString)));
                return false;
            }

            List<int> valueIndexesList = new List<int>();
            valueIndexesList.Add(sliceString.IndexOf("DPAp"));
            valueIndexesList.Add(sliceString.IndexOf("DPAm"));
            valueIndexesList.Add(sliceString.IndexOf("DPRp"));
            valueIndexesList.Add(sliceString.IndexOf("DPRm"));

            int valueCaptionLength = "DPap.".Length;

            List<string> valueStringsList = new List<string>();


            powerSlice.date_time = sliceDt;
            powerSlice.period = 30;

            for (int i = 0; i < valueIndexesList.Count; i++)
            {
                valueStringsList.Add("-1");
                if (valueIndexesList[i] == -1) continue;

                int tmpValueStartIndex = valueIndexesList[i] + valueCaptionLength - 1;
                int tmpValueEndIndex = sliceString.IndexOf("<", tmpValueStartIndex);
                var tmpValString = sliceString.Substring(tmpValueStartIndex, tmpValueEndIndex - tmpValueStartIndex);
                valueStringsList[i] = tmpValString.Replace("?", "0");
            }

            try
            {
                powerSlice.APlus = float.Parse(valueStringsList[0], CultureInfo.InvariantCulture);
                powerSlice.AMinus = float.Parse(valueStringsList[1], CultureInfo.InvariantCulture);
                powerSlice.RPlus = float.Parse(valueStringsList[2], CultureInfo.InvariantCulture);
                powerSlice.RMinus = float.Parse(valueStringsList[3], CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {
                WriteToLog("Ошибка преобразования строк в методе parseSingleSliceString: " + ex.Message);
                return false;
            }

            return true;
        }
        public bool getSlicesValuesForID(int id, DateTime dt_start, DateTime dt_end, out List<RecordPowerSlice> rpsVals)
        {
            rpsVals = new List<RecordPowerSlice>();

            string cmdStr = "READSTATE=" + dt_start.ToString("yy.MM.dd HH:mm:ss", CultureInfo.InvariantCulture) + " " + Convert.ToInt32(dt_start.IsDaylightSavingTime()).ToString() +
                " " + dt_end.ToString("yy.MM.dd HH:mm:ss", CultureInfo.InvariantCulture) + " " + Convert.ToInt32(dt_end.IsDaylightSavingTime()).ToString() +
                ";" + id;
            List<byte> cmd = wrapCmd(cmdStr);

            //byte[] testCmd = ASCIIEncoding.ASCII.GetBytes("00000000,READSTATE=17.07.27 00:00:00 0 17.07.27 22:02:13 043BC\n\n");

            byte[] incommingData = new byte[1];
            m_vport.WriteReadData(FindPacketSignature, cmd.ToArray(), ref incommingData, cmd.Count, -1);

            string answ = ASCIIEncoding.ASCII.GetString(incommingData);
            WriteToLog("Получасовки ответ: " + answ);

            int endIndex = answ.IndexOf("\nEND");
            if (endIndex == -1) return false;

            //разбор значений
            int indexDt = answ.IndexOf("<DT");
            if (indexDt == -1) return false;

            while (indexDt != -1)
            {
                int tmpIndexDt = answ.IndexOf("<DT", indexDt + 1);
                string tmpVal = "";
                if (tmpIndexDt == -1)
                {
                    tmpVal = answ.Substring(indexDt, endIndex - indexDt + 1);
                }
                else
                {
                    tmpVal = answ.Substring(indexDt, tmpIndexDt - indexDt + 1);
                }

                indexDt = tmpIndexDt;

                RecordPowerSlice rps = new RecordPowerSlice();
                if (!parseSingleSliceString(tmpVal, ref rps))
                {
                    WriteToLog("Ошибка в методе разбора строки с одной получасовкой: " + tmpVal + "; метод getSlicesValuesForID завершен");
                    return false;
                }

                rpsVals.Add(rps);
            }

            return true;
        }

        public bool findValueInListByName(string name, List<ValueUM> values, out ValueUM result)
        {
            result = values.Find((x) => { return x.name == name; });

            if (result.name == null) return false;
            //todo
            return true;
        }


        #endregion


        #region Реализация интерфейса СО

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
            string serial = "";
            for (int i = 0; i < 1; i++)
            {
                if (readUMSerial(ref serial)) return true;
            }

            WriteToLog("OpenLinkCanal: false");
            return false;
        }

        public bool ReadSerialNumber(ref string serial_number)
        {

            //WriteToLog("<< ReadSerialNumber: start, meterid: " + meterId);

            serial_number = "";

            List<ValueUM> valList = new List<ValueUM>();
            if (listOfDailyValues.Count > 0) valList = listOfDailyValues;
            else if (listOfMonthlyValues.Count > 0) valList = listOfMonthlyValues;
            else if (!getDailyValuesForID(meterId, DateTime.Now.Date, out valList) || valList.Count == 0) return false;

            //WriteToLog("ReadSerialNumber: daily values, cnt: " + valList.Count);
            serial_number = valList[0].meterSN;

            //WriteToLog("sn[last]: " + valList[valList.Count - 1].meterSN);
            //WriteToLog(">> ReadSerialNumber: end, sn[0]: " + serial_number);
            return true;
        }


        List<ValueUM> listOfDailyValues = new List<ValueUM>();
        public bool ReadDailyValues(DateTime dt, ushort param, ushort tarif, ref float recordValue)
        {
            if (listOfDailyValues == null || listOfDailyValues.Count == 0)
            {
                if (!getDailyValuesForID(meterId, dt, out listOfDailyValues)) return false;
            }

            string paramName = dailyCorrelationDict[param];
            string fullParamName = paramName + tarif.ToString();

            ValueUM val = new ValueUM();
            if (!findValueInListByName(fullParamName, listOfDailyValues, out val)) return false;

            recordValue = val.value;

            return true;
        }
        public bool ReadDailyValues2(ushort param, ushort tarif, ref float recordValue)
        {
            if (listOfDailyValues == null || listOfDailyValues.Count == 0)
            {
                if (!getDailyValuesForID(meterId, out listOfDailyValues)) return false;
            }


            string paramName = currCorrelationDict[param];
            string fullParamName = paramName + tarif.ToString();

            ValueUM val = new ValueUM();
            if (!findValueInListByName(fullParamName, listOfDailyValues, out val)) return false;

            recordValue = val.value;

            return true;
        }

        List<ValueUM> listOfMonthlyValues = new List<ValueUM>();
        public bool ReadMonthlyValues(DateTime dt, ushort param, ushort tarif, ref float recordValue)
        {
            //вся эта шелуха нужна для того, чтобы не дергать UM несколько раз при запросе разных параметров
            //одного типа. Например месячных параметров на запрос - выдается 5 штук (по пяти каналам). Они сохраняются в список.
            //потом просто подсасываются из списка
            if (listOfMonthlyValues == null || listOfMonthlyValues.Count == 0)
            {
                if (!getMonthlyValuesForID(meterId, dt, out listOfMonthlyValues)) return false;
            }

            string paramName = monthlyCorrelationDict[param];
            string fullParamName = paramName + tarif.ToString();

            ValueUM val = new ValueUM();
            if (!findValueInListByName(fullParamName, listOfMonthlyValues, out val)) return false;

            recordValue = val.value;

            return true;
        }

        public bool ReadPowerSlice(DateTime dt_begin, DateTime dt_end, ref List<RecordPowerSlice> listRPS, byte period)
        {
            return getSlicesValuesForID(meterId, dt_begin, dt_end, out listRPS);
        }

        #endregion

        #region Неиспользуемые методы интерфейса

        public List<byte> GetTypesForCategory(CommonCategory common_category)
        {
            return null;
        }

        public bool ReadCurrentValues(ushort param, ushort tarif, ref float recordValue)
        {
            return false;
        }

        public bool ReadDailyValues(uint recordId, ushort param, ushort tarif, ref float recordValue)
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

    }
}
