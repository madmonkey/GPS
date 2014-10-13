cd..
cd resources
cd type library
regtlb GPS.TLB
cd..
cd..
cd hte_clientutilities
regsvr32 HTE_ClientConfiguration.dll /s
HTE_ClientUtilities.exe /regserver

cd helpers
regsvr32 HTE_FlowControl.ocx /s
regsvr32 HTE_TransConfig.ocx /s
regsvr32 HTE_MCSConfig.ocx /s
regsvr32 HTE_ServerConfig.ocx /s
regsvr32 HTE_SocketConfig.ocx /s
regsvr32 HTE_ComPortConfig.ocx /s
regsvr32 HTE_TabView6.ocx /s
regsvr32 HTE_Entity.dll /s
regsvr32 HTE_GPSData.dll /s
regsvr32 HTE_UDP_Config.ocx /s
regsvr32 SunGardHTEParser.dll /s
cd..
cd Processes
regsvr32 HTE_Translation.dll /s
regsvr32 HTE_MCS_CARRIER.dll /s
regsvr32 HTE_RawSocket.dll /s
regsvr32 HTE_ServerSocket.dll /s
regsvr32 HTE_ComPort.dll /s
regsvr32 HTE_Emulator.dll /s
regsvr32 HTE_UDP_Transport.dll /s
@pause