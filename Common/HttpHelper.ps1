$global:headers = new-object system.collections.hashtable;
$global:urlCookies = new-object System.Collections.Hashtable;

function ClearCookies()
{
    $global:urlCookies = new-object System.Collections.Hashtable;
}

function SetCookiesUnformatted($url, $cookies)
{
    SetCookies $url $(ParseCookies $cookies);
}

function SetCookies($url, $incoming)
{
    $u = new-object Uri($url);

    if ($global:urlCookies.Contains($u.Host))
    {        
        $current = $global:urlCookies[$u.Host];

        foreach ($key in $incoming.Keys)
        {
            if ($current.ContainsKey($key))
            {
                $current[$key] = $incoming[$key];
            }
            else
            {
                $current.Add($key, $incoming[$key]);
            }
        }
    }
    else
    {
        $global:urlCookies.Add($u.Host, $incoming);
    }

    return;
}

function GetCookies($url)
{
    if ($global:nocookies)
    {
        return "";
    }

    $u = new-object Uri($url);

    if ($global:urlCookies.Contains($u.Host))
    {
        return $(FormatCookie $global:urlCookies[$u.Host]);
    }

    return "";
}

function FormatCookie($ht)
{
    $cookie = "";

    foreach ($key in $ht.Keys)
    {
        $cookie += $key.Trim() + "=" + $ht[$key] + "; ";
    }

    return $cookie;
}

function ParseCookies($cookieString)
{
    $ht = new-object System.Collections.Hashtable;    
    $cookies = $cookieString.Split(';');

    foreach ($c in $cookies)
    {
        try
        {
            $c1 = $c.Replace("HttpOnly,", "");
            $c1 = $c1.Replace("version=1,", "");

            if ($c1.Trim().Contains("HttpOnly"))
            {
                continue;
            }

            if ($c1.Trim().tolower().Contains("httponly"))
            {
                continue;
            }

            if ($c1.Trim().tolower().Contains("httponly"))
            {
                continue;
            }

            if ($c1.tolower().Contains("expires="))
            {
                continue;
            }

            if ($c1.Contains("Max-Age="))
            {
                continue;
            }

            if ($c1.Contains("domain="))
            {
                continue;
            }

            if ($c1.Contains("secure="))
            {
                continue;
            }

            if ($c1.Contains("path="))
            {
                continue;
            }

            if ($c1.Trim() -eq "secure")
            {
                continue;
            }

            if ($c1.Trim().StartsWith("secure,"))
            {
                $c1 = $c1.substring(8);
            }

            if ($c1.Trim().StartsWith("httponly,"))
            {
                $c1 = $c1.substring(10);
            }

            if ($c1.Contains("="))
            {
                try
                {                    
                    $value = $c1.Substring($c1.IndexOf("=") + 1);
                    $name = $c1.Substring(0, $c1.IndexOf("=")).trim();

                    if ($ht.ContainsKey($name))
                    {
                        $ht[$name] = $value;
                    }
                    else
                    {
                        $ht.Add($name, $value);
                    }
                }
                catch
                {
                }
            }
        }
        catch
        {
        }
    }

    return $ht;
}

function ParseValue($line, $startToken, $endToken)
{
    if ($startToken -eq $null)
    {
        return "";
    }

    if ($startToken -eq "")
    {
        return $line.substring(0, $line.indexof($endtoken));
    }
    else
    {
        try
        {
            $rtn = $line.substring($line.indexof($starttoken));
            return $rtn.substring($startToken.length, $rtn.indexof($endToken, $startToken.length) - $startToken.length).replace("`n","").replace("`t","");
        }
        catch [System.Exception]
        {
            $message = "Could not find $starttoken"
            #write-host $message -ForegroundColor Yellow
        }
    }

}

function PopulateO365FormVariables($res, $isMultiPart){

    $digest = ParseValue $res "id=`"__REQUESTDIGEST`" value=`"" "`""
    $viewState = ParseValue $res "id=`"__VIEWSTATE`" value=`"" "`""
    $sideBySideToken = ParseValue $res "id=`"SideBySideToken`" value=`"" "`""
    $viewStateGen = ParseValue $res "id=`"__VIEWSTATEGENERATOR`" value=`"" "`""
    $eventValidation = ParseValue $res "id=`"__EVENTVALIDATION`" value=`"" "`""

    if ($res.contains("|updatePanel|") -and $res.contains("__VIEWSTATE|"))
    {
        $viewState = ParseValue $res "__VIEWSTATE|" "|";
        $viewStateGen = ParseValue $res "__VIEWSTATEGENERATOR|" "|";
        $eventValidation = ParseValue $res "__EVENTVALIDATION|" "|";
    }

    ###################################
    #
    #  Basic webpart junk...
    #
    ###################################

    $req = BuildFormData "MSOWebPartPage_PostbackSource" "" $isMultiPart
    $req = $req.replace("&", "")

    $req += BuildFormData "MSOTlPn_SelectedWpId" "" $isMultiPart
    $req += BuildFormData "MSOTlPn_View" "0" $isMultiPart
    $req += BuildFormData "MSOTlPn_ShowSettings" "False" $isMultiPart
    $req += BuildFormData "MSOGallery_SelectedLibrary" "" $isMultiPart
    $req += BuildFormData "MSOGallery_FilterString" "" $isMultiPart
    $req += BuildFormData "MSOTlPn_Button" "none" $isMultiPart
    $req += BuildFormData "__LASTFOCUS" "" $isMultiPart
    $req += BuildFormData "MSOSPWebPartManager_DisplayModeName" "Browse" $isMultiPart
    $req += BuildFormData "MSOSPWebPartManager_ExitingDesignMode" "false" $isMultiPart
    $req += BuildFormData "__EVENTTARGET" "" $isMultiPart
    $req += BuildFormData "__EVENTARGUMENT" "" $isMultiPart
    $req += BuildFormData "MSOWebPartPage_Shared" "" $isMultiPart
    $req += BuildFormData "MSOLayout_LayoutChanges" "" $isMultiPart
    $req += BuildFormData "MSOLayout_InDesignMode" "" $isMultiPart
    $req += BuildFormData "MSOSPWebPartManager_OldDisplayModeName" "Browse" $isMultiPart
    $req += BuildFormData "MSOSPWebPartManager_StartWebPartEditingName" "false" $isMultiPart
    $req += BuildFormData "MSOSPWebPartManager_EndWebPartEditing" "false" $isMultiPart            

    ###################################
    #
    #  SPO and ASP.NET goodness....
    #
    ###################################

    if (!$isMultiPart)
    {
        #$digest = [System.Web.HttpUtility]::UrlEncode($digest)                
        #$viewState = [System.Web.HttpUtility]::UrlEncode($viewstate)     
        #$eventValidation = [System.Web.HttpUtility]::UrlEncode($eventValidation)     
    }              

    $req += BuildFormData "__REQUESTDIGEST" $digest $isMultiPart  
    $req += BuildFormData "SideBySideToken" $sidebysidetoken $isMultiPart                     
    $req += BuildFormData "__VIEWSTATE" $viewState $isMultiPart            
    $req += BuildFormData "__VIEWSTATEGENERATOR" $viewStateGen $isMultiPart                
    $req += BuildFormData "__EVENTVALIDATION" $eventValidation $isMultiPart                

    return $req
}

function BuildFormData($name, $value, $isMultiPart, $contentType, $filename){

    $return = "";

    if ($isMultiPart)
    {
        $return = "-----------------------------902713772214`n"
        $return += "Content-Disposition: form-data; name=`"$name`""

        if ($filename){
            $return += "; filename=`"$filename`""
        }
        
        $return += "`n"        

        if ($contentType)
        {
            $return += "Content-Type: $contentType`n"
        }

        $return += "`n"
        $return += "$value`n"
    }
    else
    {
        $value = UrlEncode $value;
        $return = "&$name=$value"
    }

    return $return;
}

function UrlEncode($in)
{
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.web")
    $val = [system.web.httputility]::urlencode($in)
    return $val
}

function UrlDecode($in)
{
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.web")
    $val = [system.web.httputility]::urldecode($in)
    return $val
}

function HtmlEncode($in)
{
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.web")
    $val = [system.web.httputility]::htmlencode($in)
    return $val
}

function HtmlDecode($in)
{
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.web")
    $val = [system.web.httputility]::htmldecode($in)
    return $val
}

function GetAllCookies($strcookies)
{
    
}

$global:headers = new-object System.Collections.Hashtable
$global:videoBuffer = 2097152;
$global:currentMark = -1; 
$global:nextMark = -1;
$global:doChucks = $false;

function DoGet($url, $strCookies)
{    
    $cookies = new-object system.net.CookieContainer;
    
    try
    {
        $uri = new-object uri($url);
    }
    catch
    {
        write-host $($_.message + ":" + $url);
        return;
    }
    
    $httpReq = [system.net.HttpWebRequest]::Create($uri)
    
    if ($httpReq.GetType().Name -eq "FileWebRequest")
    {
        write-host $($_.message + ":" + $url);
        return;
    }    

    $httpReq.Accept = "text/html, application/xhtml+xml, */*"
    $httpReq.method = "GET"   
    
    if ($global:httptimeout)
    {        
        $httpReq.Timeout = $global:httptimeout;
    }

    if ($global:language)
    {
        $httpReq.Headers["Accept-Language"] = $global:language;
        $global:language = $Null;
    }

    if ($global:useragent)
    {
        $httpReq.useragent = $global:useragent;
        $global:useragent = $null;
    }
    else
    {
        #$httpReq.useragent = "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36";
        $httpReq.useragent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:58.0) Gecko/20100101 Firefox/58.0";
    }

    if ($global:referer)
    {
        $httpReq.Referer = $global:referer;
        $global:referer = $null;
    }

    if ($global:accept)
    {
        $httpReq.Accept = $global:accept;
        $global:accept = $null;
    }

    if ($global:connection)
    {
        $sp = $httpreq.ServicePoint;
        $prop = $sp.GetType().GetProperty("HttpBehaviour", [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic);
        $prop.SetValue($sp, [byte]0, $null);

        #$httpReq.keepalive = $true;     
        $global:connection = $null;
    }
    
    $httpReq.AllowAutoRedirect = $global:allowautoredirect;    
    
    #allow us to override the cookies if we have done so...
    if ($strCookies.length -gt 0)
    {
        $httpReq.Headers.add("Cookie", $strCookies);
    }
    else
    {
        $cookie = GetCookies($url);

        if (![string]::IsNullOrEmpty($cookie))
        {
            $httpreq.Headers.Add("Cookie", $cookie);
        }    
    }

    foreach($key in $global:headers.keys)
    {
        $httpReq.Headers.add($key, $global:headers[$key]);
    }

    $global:headers.Clear();

    if ($url.contains(".mp4"))
    {
        write-host $url;
    }

    if ($url.contains("post"))
    {
        write-host $url;
    }

    if (!$url.endswith(".mp4#_=_"))
    {
        #clear the buffer...
        $global:fileBuffer = $null;  
        $global:currentMark = 0; 
        $global:nextMark = 0;
    }

    if (($url.endswith(".mp4#_=_") -and ($global:currentMark -eq 0 -or $global:currentMark -eq $null)) -or $global:doChucks)
    {
        [int]$global:currentMark = 0; #MB
        [int]$global:nextMark = 0 + [int]$global:videoBuffer #1MB;
    }

    if (($url.endswith(".mp4#_=_") -or $global:doChucks) -and $global:nextMark -ne 0)
    {        
        #add the extra headers to download a part of the file...
        $httpReq.addrange("bytes",$global:currentMark,$global:nextMark);
    }

    [string]$results = ProcessResponse $httpReq;    

    if ($global:contentRange)
    {
        if ($global:nextMark -eq $global:maxMark-1)
        {
            return;
        }

        $vals = $global:contentRange.replace("bytes","").split("-");
        $global:lastMark = $vals[0];        
        $vals = $vals[1].split("/");
        [int]$global:currentMark = $vals[0];
        [int]$global:maxMark = $vals[1];

        #set the current marks...
        [int]$global:currentMark += 1;
        [int]$global:nextMark += [int]$global:videoBuffer;

        if ($global:nextMark -ge $global:maxMark-1)
        {
            $global:nextMark = $global:maxMark-1;
            $global:done = $true;
        }

        DoGet $url $strCookies;

        if ($global:done)
        {
            $global:contentRange = $null;
            return;
        }
    }
    else
    {
        $global:currentMark = $null;
    }
    
    return $results
}

$global:fileName = ""
$global:fileBuffer = $null;

$global:videoMaxSize = 2000000;  #2MB download at a time...
$global:videoPointer = 0;

$global:location = "";

$global:contentRange = "";

function ProcessResponse($req)
{
    #use them all...
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Ssl3 -bor [System.Net.SecurityProtocolType]::Tls;

    if ($global:ignoreSSL)
    {
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true};
        #[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;
    }

    $global:httpCode = -1;
    $global:fileName = ""
    #$global:fileBuffer = $null;    

    $urlFileName = $req.RequestUri.Segments[$req.RequestUri.Segments.Length - 1];            
    $response = "";            

    try
    {
        $res = $req.GetResponse();

        $mimeType = $res.ContentType;
        $statusCode = $res.StatusCode.ToString();
        $global:httpCode = [int]$res.StatusCode;
        $cookieC = $res.Cookies;
        $resHeaders = $res.Headers;  
        $global:rescontentLength = $res.ContentLength;
                                
        try
        {
            $global:location = $res.Headers["Location"].ToString();
        }
        catch
        {
        }

        try
        {
            $global:contentRange = $res.Headers["Content-Range"].ToString();

            $vals = $global:contentRange.replace("bytes","").split("-");
            $global:lastMark = $vals[0];        
            $vals = $vals[1].split("/");
            $global:currentMark = $vals[0];
            $global:maxMark = $vals[1];

            #set the content length...
            $global:rescontentLength = $global:maxMark;
        }
        catch
        {
        }

        try
        {
            $rawCookies = $res.Headers["set-cookie"].ToString();

            SetCookiesUnformatted $res.ResponseUri.ToString() $rawCookies;
        }
        catch 
        {
        }

        $global:fileName = "";
        $length = 0;

        try
        {
            $global:fileName = $res.Headers["Content-Disposition"].ToString();

            if ($global:fileName -ne "attachment")
            {
                $global:fileName = $global:fileName.Replace("attachment; filename=", "").Replace("""", "");

                if ($global:filename.contains("filename="))
                {
                    $global:filename = ParseValue $global:fileName "filename=" ";";
                }

                $length = $res.ContentLength;
            }
            else
            {
                $global:fileName = "";
            }
        }
        catch
        {

        }        

        if ($global:fileName.Length -gt 0)
        {
            $bufferSize = 10240;
            $buffer = new-object byte[] $buffersize;

            $strm = $res.GetResponseStream();  
            
            if ($global:fileBuffer -eq $Null)
            {          
                $global:bytesRead = 0;
                $global:fileBuffer = new-object byte[] $($res.ContentLength);
                $global:ms = new-object system.io.MemoryStream (,$global:fileBuffer);
            }
            
            while (($bytesRead = $strm.Read($buffer, 0, $bufferSize)) -ne 0)
            {
                $global:ms.Write($buffer, 0, $bytesRead);
            } 

            $global:ms.Close();
            $strm.Close();
        }
        else
        {
            $responseStream = $res.GetResponseStream();
            $contentType = $res.Headers["Content-Type"];

            if ($res.ContentEncoding.ToLower().Contains("gzip"))
            {
                $responseStream = new-object System.IO.Compression.GZipStream($responseStream, [System.IO.Compression.CompressionMode]::Decompress);
            }
            
            if ($res.ContentEncoding.ToLower().Contains("deflate"))
            {
                $responseStream = new-object System.IO.Compression.DeflateStream($responseStream, [System.IO.Compression.CompressionMode]::Decompress);
            }

            switch($contentType)
            {
                {$_ -in "image/gif","image/png","image/jpeg","video/mp4"}
                {
                    $bufferSize = 409600;
                    $buffer = new-object byte[] $buffersize;                                        
                    $bytesRead = 0;

                    if ($global:fileBuffer -eq $Null)
                    {          
                        $global:bytesRead = 0;
                        $global:fileBuffer = new-object byte[] $($global:rescontentlength);
                        $global:ms = new-object system.io.MemoryStream (,$global:fileBuffer);
                    }
                                        
                    while (($bytesRead = $responseStream.Read($buffer, 0, $bufferSize)) -ne 0)
                    {
                        $global:ms.Write($buffer, 0, $bytesRead);
                        $global:bytesRead += $bytesRead;
                    } 
                    
                    <#
                    if ($global:bytesRead -eq $global:maxMark)
                    {
                        #$global:ms.Close();
                    }
                    #>

                    $responseStream.Close();

                    if ($global:fileName.Length -eq 0)
                    {                        
                        $global:fileName = $req.requesturi.segments[$req.requesturi.segments.length-1];

                        if ($contentType -eq "video/mp4")
                        {
                            $global:fileName += ".mp4";
                        }
                    }

                    }
                default{
                    $reader = new-object system.io.StreamReader($responseStream, [System.Text.Encoding]::Default);                    
                    $response = $reader.ReadToEnd();                            
                    }
            }

            $res.Close();
            $responseStream.Close();

            $req = $null;
            $proxy = $null;
        }
    }
    catch
    {
        $res = $_.Exception.InnerException.Response;
        $global:httpCode = $_.Exception.InnerException.HResult;
        $global:contentRange = $null;

        try
        {
            $responseStream = $res.GetResponseStream();
            $statusCode = $res.StatusCode.ToString();
            $global:httpCode = [int]$res.StatusCode;
            $reader = new-object system.io.StreamReader($responseStream, [System.Text.Encoding]::Default);                    
            $response = $reader.ReadToEnd();                            
            return $response;
        }
        catch
        {
            $global:httperror = $_.exception.message;

            write-host "Error getting response from $($req.RequestUri)";
            return $null;
        }

        if ($res.ContentEncoding.ToLower().Contains("gzip"))
        {
            $responseStream = new GZipStream($responseStream, [System.IO.Compression.CompressionMode]::Decompress);
        }
        
        if ($res.ContentEncoding.ToLower().Contains("deflate"))
        {
            $responseStream = new DeflateStream($responseStream, [System.IO.Compression.CompressionMode]::Decompress);
        }

        $reader = new-object System.IO.StreamReader($responseStream, [System.Text.Encoding]::Default);
        $response = $reader.ReadToEnd();                
    }    

    return $response;
}

$contentType = "application/x-www-form-urlencoded"
$overrideContentType = $null
$useXRequestWith = $false

function DoHttpSendAction($action, $url, $post, $strCookies )
{
    $encoding = new-object system.text.asciiencoding
    $buf = $encoding.GetBytes($post)
    $uri = new-object uri($url);
    $httpReq = [system.net.HttpWebRequest]::Create($uri)
    $httpReq.AllowAutoRedirect = $false
    $httpReq.method = $action;
    #$httpReq.Referer = ""
    $httpReq.contentlength = $buf.length

    $httpReq.Accept = "text/html, application/xhtml+xml, */*"
    #$httpReq.ContentType = "application/x-www-form-urlencoded"
    $httpReq.headers.Add("Accept-Language", "en-US")
    $httpReq.UserAgent = "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0; LEN2)"

    #allow us to override the cookies if we have done so...
    if ($strCookies)
    {
        $httpReq.Headers.add("Cookie", $strCookies);
    }
    else
    {
        $cookie = GetCookies($url);

        if (![string]::IsNullOrEmpty($cookie))
        {
            $httpreq.Headers.Add("Cookie", $cookie);
        }    
    }

    if ($global:referer)
    {
        $httpReq.Referer = $global:referer;
        $global:referer = $null;
    }

    if ($global:useragent)
    {
        $httpReq.useragent = $global:useragent;
        $global:useragent = $null;
    }
    
    if ($global:overrideContentType)
    {
        $httpReq.ContentType = $overrideContentType
        $global:overrideContentType = $null
    }
    else
    {
        $httpReq.ContentType = "application/x-www-form-urlencoded"
    }

    if ($global:accept)
    {
        $httpReq.Accept = $global:accept;
        $global:accept = $null;
    }

    if ($global:connection)
    {
        $httpReq.keepalive = $true;     
        $global:connection = $null;
    }

    if ($digest)
    {
        $httpReq.headers.Add("X-RequestDigest", $digest)
    }

    if ($useXRequestWith)
    {
        $httpReq.headers.Add("X-Requested-With", "XMLHttpRequest")
        $useXRequestWith = $false
    }

    foreach($key in $global:headers.keys)
    {
        $httpReq.Headers.add($key, $global:headers[$key]);
    }

    $global:headers.Clear();
    
    $stream = $httpReq.GetRequestStream()

    [void]$stream.write($buf, 0, $buf.length)
    $stream.close()

    [string]$results = ProcessResponse $httpReq;       

    return $results
}

function DoPost($url, $post, $strCookies )
{    
    DoHttpSendAction "POST" $url $post $strCookies
}

function DoPut($url, $post, $strCookies )
{    
    DoHttpSendAction "PUT" $url $post $strCookies
}

function DoGetCache($url, $strCookies, $cs)
{
    $fileName = GetCacheFileName $cs;

    if ([System.IO.File]::Exists($fileName) -and !$cs.Overwrite)
    {
        $global:headers.Clear();
        return [System.IO.File]::ReadAllText($fileName);
    }
    else
    {        
        $html = DoGet $url $strCookies;
        [System.IO.File]::AppendAllText($fileName, $html);
        return $html;
    }        
}