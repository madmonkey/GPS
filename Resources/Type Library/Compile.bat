PATH C:\Program Files\Microsoft Visual Studio\VC98\Bin;%PATH%
call vcvars32
midl.exe "./GPS.idl" /tlb "./GPS.tlb" /Oicf
pause