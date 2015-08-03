using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AdysTech.IISLogAnalytics
{
    //ref://https://www.microsoft.com/technet/prodtechnol/WindowsServer2003/Library/IIS/676400bc-8969-4aa7-851a-9319490a9bbb.mspx?mfr=true
    class IISLogEntry
    {
       public static readonly string propClientIPAddress = "c-ip";
       public static readonly string propUserName = "cs-username";
       public static readonly string propServiceNameandInstanceNumber = "s-sitename";
       public static readonly string propServerName = "s-computername";
       public static readonly string propServerIPAddress = "s-ip";
       public static readonly string propServerPort = "s-port";
       public static readonly string propMethod = "cs-method";
       public static readonly string propURIStem = "cs-uri-stem";
       public static readonly string propURIQuery = "cs-uri-query";
       public static readonly string propHTTPStatus = "sc-status";
       public static readonly string propWin32Status = "sc-win32-status";
       public static readonly string propBytesSent = "sc-bytes";
       public static readonly string propBytesReceived = "cs-bytes";
       public static readonly string propTimeTaken = "time-taken";
       public static readonly string propProtocolVersion = "cs-version";
       public static readonly string propHost = "cs-host";
       public static readonly string propUserAgent = "cs(User-Agent)";
       public static readonly string propCookie = "cs(Cookie)";
       public static readonly string propReferrer = "cs(Referrer)";
       public static readonly string propProtocolSubstatus = "sc-substatus";

       public DateTime TimeStamp { get; set; }
       public string ClientIPAddress { get; set; }
       public string UserName { get; set; }
       public string ServiceNameandInstanceNumber { get; set; }
       public string ServerName { get; set; }
       public string ServerIPAddress { get; set; }
       public int ServerPort { get; set; }
       public string Method { get; set; }
       public string URIStem { get; set; }
       public string URIQuery { get; set; }
       public int HTTPStatus { get; set; }
       public int Win32Status { get; set; }
       public int BytesSent { get; set; }
       public int BytesReceived { get; set; }
       public int TimeTaken { get; set; }
       public string ProtocolVersion { get; set; }
       public string Host { get; set; }
       public string UserAgent { get; set; }
       public string Cookie { get; set; }
       public string Referrer { get; set; }
       public string ProtocolSubstatus { get; set; }
    }
}
