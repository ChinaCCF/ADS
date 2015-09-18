%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\installutil.exe ADS.exe
Net Start ADS
sc config ADS start= auto
