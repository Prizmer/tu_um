using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Drivers.UMDriver
{
    public static class UMCommon
    {
        public enum UM_VERSION
        {
            UM31,
            UM40,
            UNKNOWN
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
    }

    public enum UM_VERSION
    {
        UM31,
        UM40,
        UNKNOWN
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
}
