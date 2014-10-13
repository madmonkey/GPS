If you wish to debug the compiled version it is best that you run the application from the HTE_ClientUtilities folder.

Since the application relies on a particular client directory structure:

APPLICATION
	|
	|_ PROCESSES
	|
	|_ HELPER
	|
	|_ DATA

It uses this structure to determine where the processes exist and where our shared configuration file is kept.

The Application folder would consist of the HTE_GPS.EXE, HTE_CLIENTCONFIGURATION.DLL, HTE_CLIENTSERVICES.DLL, HTE_CLIENTUTILITIES.EXE, CONFIG.XML, SERVICECONTROLLER.EXE AND SUNGARDGPSSERVICE.EXE.
The Processes folder would consist of any DLL or EXE that supports the HTE_GPS.PROCESS interface and has a creatable interface named PROCESS.
The Helper folder consists of any UI that is required to properly configure a Process - additionally it may contain stylesheets, shared application dlls or ocxs.
The Data folder consists of the data repository for managing identities. NOTE: It is imperative that the registry key HKEY_LOCAL_MACHINE\SOFTWARE\HTE\MODULAR GPS\Install_Path 
points to the application root - since several components attempt to access the datastore they use that value to determine relative to themseleves where the correct location is.

NOTE: You can NOT add the project HTE_ClientUtilities to a group project since it is an ActiveX EXE type. You may however run the application in a separate
instance of the IDE in order to debug the Process.

In order to debug as an IDE Project you will most likely have to "mirror" the projects you wish to debug inside of the Utilities folder.

Dependancies will be registered from the RegisterDependancies.bat file; however, you may still have to manually register you GPS.tlb (located in the resources folder). 
Additionally you might have to register the HTE_SystemUtility.dll (since it is a shared dll - it is not auto-magically registered by the batch file.)

Several shared components have been re-used in the making of this project - HTE_PubData.dll, HTE_Properties.dll, ccrpTimers.dll.



