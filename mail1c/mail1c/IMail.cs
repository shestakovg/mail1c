﻿using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace mail1c
{
    [Guid("36637F39-9209-419C-AC11-CC32C6E2BFBE")]
    interface IMail
    {
        [DispId(1)]
        void ConnectIMAP(string server, int port, string username, string password);
        void CloseConnection();
        bool SendMessage(string to, string tc, string[] attachments);
    }

    [Guid("A4420A21-152A-4D2A-9FCB-7FE20E5D1B71"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IMailEvents
    {

    }
}
