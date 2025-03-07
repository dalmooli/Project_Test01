using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ICDIBasic
{
    public class LogClass
    {
        public byte seccSts;        // SECC Status
        public byte Psoc;           // Present SOC
        public ushort Pmeter;       // OCPP power meter

        public byte PmSts;          // Power module Status

        public ushort Tvolt;        // Target Volt
        public uint Ovolt;        // Output Volt
        public uint Ocurrent;     // Output Current
        public ushort Opwr;         // Output Power
    }
}
