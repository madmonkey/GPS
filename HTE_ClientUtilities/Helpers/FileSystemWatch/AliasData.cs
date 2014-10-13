using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Runtime.InteropServices;

namespace FileSystemWatch
{
    [GuidAttribute("270c2dca-8d3e-4b56-8d83-5d5c3ef122be"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IAliasData
    {
        [DispId(0)]
        string AliasId { get; }
        [DispId(1)]
        string PhysicalLookUp { get; }
        [DispId(2)]
        string AddressLookUp { get; }
        [DispId(3)]
        string AppLookUp { get; }
        [DispId(4)]
        string Device { get; }
        [DispId(5)]
        string Alias { get; }
        [DispId(6)]
        string Comments { get; }
    }

    public class AliasData : IAliasData
    {
        public string AliasId { get; private set; }
        public string PhysicalLookUp{ get; private set; }
        public string AddressLookUp{ get; private set; }
        public string AppLookUp{ get; private set; }
        public string Device{ get; private set; }
        public string Alias{ get; private set; }
        public string Comments{ get; private set; }

        public AliasData(DataRow dataRow)
        {
            AliasId = dataRow["AliasID"].ToString();
            PhysicalLookUp = dataRow["PhysicalLookUp"].ToString();
            AddressLookUp = dataRow["AddressLookUp"].ToString();
            AppLookUp = dataRow["AppLookUp"].ToString();
            Device = dataRow["Device"].ToString();
            Alias = dataRow["Alias"].ToString();
            Comments = dataRow["Comments"].ToString();
        }
    }
}
