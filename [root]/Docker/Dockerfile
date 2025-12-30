FROM mcr.microsoft.com/dotnet/framework/aspnet:4.8-windowsservercore-ltsc2019
SHELL ["powershell", "-command"]

RUN Install-WindowsFeature Web-ASP;
RUN Install-WindowsFeature Web-CGI;
RUN Install-WindowsFeature Web-ISAPI-Ext;
RUN Install-WindowsFeature Web-ISAPI-Filter;
RUN Install-WindowsFeature Web-Includes;
RUN Install-WindowsFeature Web-HTTP-Errors;
RUN Install-WindowsFeature Web-Common-HTTP;
RUN Install-WindowsFeature Web-Performance;
RUN Install-WindowsFeature WAS;
RUN Import-module IISAdministration;
RUN md c:/msi;

RUN Invoke-WebRequest 'https://aka.ms/vs/15/release/vc_redist.x64.exe' -OutFile c:/msi/vc_redist.x64.exe; 
RUN Start-Process "c:/msi/vc_redist.x64.exe" -ArgumentList "/quiet", "/norestart" -Wait -NoNewWindow

RUN Invoke-WebRequest 'https://go.microsoft.com/fwlink/?linkid=2345415' -OutFile c:/msi/msodbcsql18.msi; 
RUN Start-Process "msiexec.exe" -ArgumentList "/i", "c:\msi\msodbcsql18.msi", "IACCEPTMSODBCSQLLICENSETERMS=YES", "ADDLOCAL=ALL", "/qn" -Wait -NoNewWindow

EXPOSE 80
RUN Remove-Item -Recurse C:\inetpub\wwwroot\*
RUN & c:\windows\system32\inetsrv\appcmd.exe unlock config /section:system.webServer/asp
RUN & c:\windows\system32\inetsrv\appcmd.exe unlock config /section:system.webServer/handlers
RUN & c:\windows\system32\inetsrv\appcmd.exe unlock config /section:system.webServer/modules

WORKDIR C:/inetpub/wwwroot/
COPY ASP/* C:/inetpub/wwwroot/