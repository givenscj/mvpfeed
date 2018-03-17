function UnblockAndAdd($path, $suffix)
{
    try
    {
        if ($path.endswith("dll"))
        {
            Unblock-File -Path $path      
            
            write-host "Loading assembly [$Path]";
              
            Add-Type -Path $path        
        }
        else
        {        
            if ($suffix)
            {
                write-host "Loading assembly [$path $suffix]";
                $dll = [System.Reflection.Assembly]::Load($path + $suffix);                
            }
            else
            {
                write-host "Loading assembly [$path]";
                $dll = [System.Reflection.Assembly]::LoadWithPartialName($path);
            }
        }
    }
    catch
    {
        foreach($ex in $_.Exception.LoaderExceptions)
        {
            write-host $ex.Message;
        }
		if (!$dll)
        {
            #try to load from the path...
            $path = $runPath + "\" + $path + ".dll";

            write-host "Loading assembly [$path]";
            Add-Type -Path $path
        }
    }
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

$libraryCache = New-Object System.Collections.Hashtable
$folderCache = New-Object System.Collections.Hashtable

function GetListAndFolder($context, $libraryName)
{
    $key = $context.url + "|$libraryname"

    if ($libraryCache.containskey($key))
    {
        $list = $libraryCache[$key]
    }

    if ($folderCache.containskey($key))
    {
        $rootFolder = $folderCache[$key]
    }

    if (!$libraryCache -or !$rootFOlder)
    {
        #get the target library
        $List = $Context.Web.Lists.GetByTitle($libraryName)
        $rootfolder = $list.RootFolder
        $Context.Load($List)
        $Context.Load($rootfolder)

        try
        {
            $Context.ExecuteQuery()   

            $libraryCache.Add($key, $list);
            $folderCache.add($key, $rootFolder);
        }
        catch
        {
                #most likely the username/password or site url is wrong...
                LogImportError $error[0] $appConfig $doc  

                WriteProgress $PID_FILE -activity "Processing File" "Processing $($doc.id): Error : occured"  7

                #no use continuing...
                return $(SetDocumentResultError $docResult $($error[0].Exception.Message));
        }
    }
}

$fieldCache = new-object System.Collections.Hashtable
$fieldsCache = new-object System.Collections.Hashtable

function GetListField($context, $list, $fieldName)
{
    $fieldskey = $list.id.tostring()
    $fieldkey = $list.id.tostring() + "|" + $fieldName

    if ($fieldCache.containskey($fieldkey))
    {
        return $fieldCache[$fieldkey]
    }

    if ($fieldsCache.containskey($fieldskey))
    {
        try
        {
            $fields = $fieldsCache[$fieldskey];
            $field = $fields.GetByInternalNameOrTitle($fieldName);        
            $context.load($field);
            $context.ExecuteQuery();
            $fieldCache.Add($fieldkey, $field);
        }
        catch
        {
            if ($error[0].Exception.Message.contains("does not exist. It may have been deleted by another user."))
            {
                $fieldCache.add($fieldkey, $null);
            }
        }
    }   
    else
    { 
        try
        {
            $fields = $list.Fields
            $field = $fields.GetByInternalNameOrTitle($fieldName);
            $context.load($fields);
            $context.load($field);
            $context.ExecuteQuery();

            $fieldsCache.Add($fieldskey, $fields);
            $fieldCache.Add($fieldkey, $field);
        }
        catch
        {
        
        }
    }

    return $field;
}

function TokenReplace($inValue, $hashtable)
{
    $tempValue = $inValue    

    foreach($key in $hashtable.keys)
    {
        $tempValue = $tempValue.replace("{$key}", $hashtable[$key])
    }    

    $tempValue = $tempValue.replace("{Date}", [System.DateTime]::Now);    
    $tempValue = $tempValue.replace("{AppName}", $appName);    

    return $tempValue
}

$contextCache = new-object System.Collections.Hashtable

function GetContext($siteUrl, $appConfig)
{   
    #try to pull from cache...
    if ($contextCache[$siteUrl])
    {
        return $contextCache[$siteUrl]
    }


    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
    $context.RequestTimeout = $config.COnfiguration.Source.TimeOut;
    $secure = ConvertTo-SecureString $($config.Configuration.destination.Password) -AsPlainText -force
                    
    $Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($($config.Configuration.Destination.Username), $secure)
    $Creds = New-Object System.Net.NetworkCredential($($config.Configuration.Destination.Username), $($config.Configuration.destination.Password))

    $Context.Credentials = $Creds  

    $contextCache.Add($siteurl, $context);

    return $Context
}

$termCache = new-object System.Collections.Hashtable

function GetTermIdForTerm($term, $termSetId, $context, $mmsSource, $autocreate, $parent, $termIdHint)
{
    if ($term.InnerText)
    {
        $term = $term.InnerText;
    }

    if (!$term -or $term.length -eq 0)
    {
        return $Null;
    }

    $termId = $null

    $termCacheKey = $term + "|" + $termSetId;

    if ($termCache.ContainsKey($termCacheKey))
    {
        $termMatches = $termCache[$termCacheKey]
    }
    else
    {
        $session = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($context);

        #default to the MMS Service App layer...
        $ts = $session.GetDefaultKeywordsTermStore();

        if ($mmsSource -eq "DefaultKeywords")
        {
            $ts = $session.GetDefaultKeywordsTermStore();
        }
    
        if ($mmsSource -eq "DefaultSiteCollection")
        {
            $ts = $session.GetDefaultSiteCollectionTermStore();
        }
    
        $tset = $ts.getTermSet($termSetId);    

        $lmi = new-object Microsoft.SharePoint.Client.Taxonomy.LabelMatchInformation($context);
        $lmi.lcid = 1033
        $lmi.TrimUnavailable = $true;
        $lmi.TermLabel = $term;    
    
        $termMatches = $tset.GetTerms($lmi);
        $context.Load($session);
        $context.Load($ts);    
        $context.Load($tset);
        $context.Load($termMatches);
        $context.ExecuteQuery();

        $termCache.Add($termCacheKey, $termMatches);
    }    

    if ($termMatches -and $termmatches.Count -gt 0)
    {        
        if ($termMatches.count -gt 1)
        {
            if ($termIdHint)
            {
                $enum = $termMatches.GetEnumerator();

                while($enum.MoveNext())
                {
                    if ($enum.Current.Id.tostring() -eq $termIdHint)
                    {
                        return $enum.Current.Id;
                    }    
                }   
            }

            if (!$parent)
            {
                $enum = $termMatches.GetEnumerator();

                while($enum.MoveNext())
                {
                    $sItem = $enum.Current

                    #get the term
                    $context.load($sItem);
                    $context.ExecuteQuery();

                    #output it...
                    $sItem                    
                }   

                return
            }
            else
            {
                #there may be more than one...find the one with the right parent...
                $parentId = GetTermIdForTerm $parent $termSetId $context $mmsSource $autocreate                

                $enum = $termMatches.GetEnumerator();

                while($enum.MoveNext())
                {
                    $sItem = $enum.Current

                    #get the term
                    $context.load($sItem);
                    $context.load($sItem.Parent);
                    $context.ExecuteQuery();

                    if ($parentId.Count)
                    {
                        foreach($pTerm in $parentId)
                        {
                            if ($sItem.Parent.Id -eq $pTerm.Id)
                            {
                                return $sItem.id;
                            }

                            if ($sItem.Parent.Id -eq $pTerm)
                            {
                                return $sItem.id;
                            }
                        }
                    }                    
                }
            }
        }
        else
        {
            return $termmatches.item(0).Id;
        }
    }
    else
    {
        WriteProgress $PID_MIGRATION "Migrating Documents" "MMS Term [$term] was not found" 99

        if ($autocreate)
        {
            WriteProgress $PID_MIGRATION "Migrating Documents" "Creating a new MMS Term for $term" 99

            $term = (Get-Culture).textinfo.totitlecase($term.tolower())            
            $newTerm = CreateMMSTerm $context $tset $term

            if (!$newTerm)
            {
                $ex = new-object System.Exception("MMS Term [$term] could not be created");
                throw $ex
            }
            else
            {
                return $newTerm;
            }
        }
    }
}

function CreateMMSTerm($context, $ts, $term)
{    
    try
    {
        $TermAdd = $Ts.CreateTerm($Term,1033,[System.Guid]::NewGuid().toString())
        $Context.Load($TermAdd)
        $Context.ExecuteQuery()            
    }
    catch
    {
    }

    if ($termadd)
    {
        return $termadd.id;
    }

    #something errored...
    return $Null;
}

function ParseSharePointColumnName($targetName)
{
    $name = $targetName.replace(" ","_x0020_")
    $name = $name.replace("(","_x0028_")
    $name = $name.replace(")","_x0029_")
    $name = $name.replace("-","_x002d_")

    return $name
}

function WriteToVerbose($line)
{
    if (!$line.endswith("`r`n"))
    {
        $line = $line + "`r`n"
    }

    #log to log file...        
    $wroteLog = $false;

    while(!$wroteLog)
    {
        try
        {
            [System.IO.FIle]::AppendAllText($verboseLogFile, $line);
            $wroteLog = $true;
        }
        catch
        {            
        }
    }
}

function WriteProgress($id, $activity, $status, $percentComplete, $doNotLog)
{
    if ($percentComplete -lt 100)
    {
        Write-Progress -id $id -activity $activity -status $status -percentComplete $percentComplete
    }
    else
    {
        Write-Progress -id $id -activity $activity -status $status -Completed
    }

    if (!$doNotLog)
    {
        $line = [System.DateTime]::Now.ToString("hh:mm:ss tt") + "`t" + $id.tostring() + "`t" + $activity + "`t" + $status + "`t" + $percentComplete.tostring() + "`r`n"
        WriteToVerbose $line
    }
}

function GetFirst100($path)
{
    $fs = new-object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read);
    $br = new-object System.IO.BinaryReader($fs);
    $bytes = [Byte[]] (,0xFF * 100)
    [void]$br.read($bytes, 0, 100);
    $line = [System.Text.Encoding]::UTF8.GetString($bytes);
    return $line.replace("`r","").replace("`n","");
}

function CheckDocType($doc)
{
    $fi = new-object system.io.fileinfo($doc.path);

    if (!$fi.Exists)
    {
        return;
    }

    #if the line wasn't populated for some reason...get it.
    if (!$doc.firstline)
    {
        $line = GetFirst100 $($doc.Path)

        #$line = Get-Content $doc.Path -First 1 -ea SilentlyContinue

        if ($line.length -gt 200)
        {
            $doc.firstline = $line.substring(0,200)
        }
        else
        {
            $doc.firstline = $line
        }
    }

    if ([int]$doc.firstline.tochararray()[1] -eq 5)
    {
        $doc.DBFileType = "pcx"
        $doc.filetype = "pcx"
    }

    if ($doc.firstline.contains("COM"))
    {
        $doc.DbFileType = "com"        
        $doc.FileType = "com"        
    }

    if ($doc.firstline.contains("PDF"))
    {
        $doc.DbFileType = "pdf"        
        $doc.FileType = "pdf"                
    }

    if ($doc.firstline.contains("HTML"))
    {
        $doc.DbFileType = "html"        
        $doc.FileType = "html"                
    }

    if ($doc.firstline.contains("JFIF"))
    {
        $doc.DbFileType = "jfif"        
        $doc.FileType = "jfif"                
    }

    if ($doc.firstline.contains("FFL1.0"))
    {
        $doc.DbFileType = "ffl"        
        $doc.FileName = $doc.firstLine.replace("FFL1.0","")     

        $vals = $doc.FileName.split(".")           
        $doc.FileType = $vals[$vals.length-1]        
    }    

    if ($doc.firstline.substring(0,2).startswith("BM"))
    {
        $doc.DbFileType = "bmp"        
        $doc.FileType = "bmp"        
    }    

    if ($doc.firstline.contains("II*"))
    {
        $doc.DbFileType = "tif"        
        $doc.FileType = "tif"        
    }    

    if ($doc.firstline.contains("GIF"))
    {
        $doc.DbFileType = "gif"        
        $doc.FileType = "gif"
    }    

    if ($doc.firstline.contains("MM*"))
    {
        $doc.DbFileType = "tif"        
        $doc.FileType = "tif"        
    }    

    if ($doc.firstline.contains("\rtf1\"))
    {
        $doc.DbFileType = "rtf"        
        $doc.FileType = "rtf"        
    }    

    if (!$doc.FileType -or $doc.FileType -eq "")
    {
        $doc.dbFileType = "unknown"
        $doc.FileType = "unknown"
        LogDocTypeError $doc
    }

    if (!$doc.FileName -or $doc.FileName -eq "")
    {        
        $doc.FileName = $doc.id + "." + $doc.FileType
    }    
}

function ConfigTokenReplace($inValue)
{
    $ht = new-object System.Collections.Hashtable

    foreach($token in $config.Configuration.Source.Tokens.Token)
    {
        $ht.Add($token.Name, $token.Value)
    }

    $val = TokenReplace $inValue $ht
    return $val
}

function CheckForKeyPress()
{
    if ($host.ui.rawui.KeyAvailable)
    {
        write-host "Key press detected, enter a command (runorganizer, loadconfig, exit):"
        
        #wait for input...
        $cmd = $host.ui.ReadLine();

        #do something...
        switch($cmd.tolower())
        {
            "runorganizer" {
                
            }
            "loadconfig" {
                [xml]$config = get-content "$configFile" -raw
            }
            "exit"{
                exit;
            }
        }
    }    
}

Function TimedPrompt($prompt,$secondsToWait){   
    Write-Host -NoNewline $prompt
    $secondsCounter = 0
    $subCounter = 0
    While ( (!$host.ui.rawui.KeyAvailable) -and ($count -lt $secondsToWait) ){
        start-sleep -m 10
        $subCounter = $subCounter + 10
        if($subCounter -eq 1000)
        {
            $secondsCounter++
            $subCounter = 0
            Write-Host -NoNewline "."
        }       
        If ($secondsCounter -eq $secondsToWait) { 
            Write-Host "`r`n"
            return $false;
        }
    }
    Write-Host "`r`n"
    return $true;
}

function CleanConfigFile($configFile)
{
    $cacheApp = Import-Clixml $configFile

    foreach($doc in $cacheApp.Documents)
    {
        if ($doc.FirstLine.length -gt 200)
        {
            $doc.FirstLine = $doc.FirstLine.Substring(0,200);
        }
    }

    SaveAppToCache $cacheApp
}

function ProcessColumn($appConfig, $context, $list, $doc, $columnName)
{
    $column = new-object Column
    $column.Value = $doc.metadata[$columnName]
    $column.Name = $columnName;    
    $column.NewValue = $doc.metadata[$columnName]
    $column.NewName = $columnName;    

    #get the col config; if it exists...        
    $colConfig = $null
    foreach($temp in $appConfig.Columns.Column)
    {             
        if ($temp.Name -eq $columnName)
        {
            $colConfig = $temp
            break
        }
    }        
    
    #do the column mapping rules...
    if ($colConfig)
    {            
        #check to see if a format handler is present
        $column.formatHandler = $colConfig.FormatHandler
        $column.Type = $colConfig.Type;
        
        #possibly have a different column name in the target, set it here
        if ($colConfig.MappedName)
        {
            $column.NewName = $colConfig.MappedName                
        }            
        
        if ($colConfig.Ignore -eq "true")
        {
            $column.Ignore = $true
        }
            
        #set to ignore empty values...
        if ($colConfig.IgnoreEmpty -and $column.Value -eq "")
        {
            $column.Ignore = $true
        }

        if ($colConfig.IgnoreEmpty -and !$column.Value)
        {
            $column.Ignore = $true
        }

        #possibly have some type of value mapping...set it here.
        switch($colConfig.Type)
        {
            "Static"{
                $column.NewValue = $colConfig.Value
            }
            "MappedValue"{
                foreach($tValue in $colConfig.Values.Value)
                {
                    if ($tValue.Source -eq $value)
                    {
                        $column.NewValue = $tValue.Destination
                    }
                }
            }
            "Calculated"{
                #do token replacement...
                $tempValue = $colConfig.Value

                while($tempValue.Contains("{"))
                {
                    $token = ParseValue $tempvalue "{" "}"
                    $tempValue = $tempValue.replace($token, $doc.metadata[$token])
                }

                $column.NewValue = $tempvalue
            }
            "ManagedMetadata" {

                $tempCol = $col

                if ($colConfig.MappedName)
                {
                    $tempCol = $colConfig.MappedName
                }

                #check to see if parent is needed...
                if ($colConfig.Value -and $colConfig.Value.Attributes)
                {
                    $parent = $colConfig.Value.Attributes["Parent"].Value
                }

                try
                {
                    $column.NewValue = SetManagedMetaDataField $context $list $uploadFile $tempCol $($column.value) $($colConfig.MMSSource) $($colConfig.AutoCreate) $parent
                }
                catch
                {
                    LogImportError $error[0] $appConfig $doc $uploadfile
                }
                    
                $valueWasSet = $true;
            }
        }                            
    }

    #blank dates don't work for CSOM...
    if($colConfig.Type -and $colConfig.Type -eq "DateTime")
    {
        if (!$column.NewValue -or $column.NewValue.length -eq 0)
        {
            $column.Ignore = $true;
        }
    }

    #parse for sharepointyness...
    $column.NewName = ParseSharePointColumnName $column.NewName    

    return $column
}

function CleanValueForExport($value)
{
    $value = $value.Replace($([char]0xFFFF).tostring(), ' ');

    return $value
}

function AddAppColumn($app, $name)
{
    if (!$app.Columns.ContainsKey($name))
    {
        $app.Columns.Add($name, $name)
    }
}

function AddDocColumn($app, $doc, $name)
{
    AddAppColumn $app $name

    if (!$doc.Metadata.ContainsKey($name))
    {
        $doc.Metadata.Add($name, $name)
    }
}

function Write-ColorText
{
    # DO NOT SPECIFY param(...)
    #    we parse colors ourselves.

    $allColors = ("-Black",   "-DarkBlue","-DarkGreen","-DarkCyan","-DarkRed","-DarkMagenta","-DarkYellow","-Gray",
                  "-Darkgray","-Blue",    "-Green",    "-Cyan",    "-Red",    "-Magenta",    "-Yellow",    "-White",
                   "-Foreground")
    
    $color = "Foreground"
    $nonewline = $false

    foreach($arg in $args)
    {
        if ($arg -eq "-nonewline")
        { 
            $nonewline = $true 
        }
        elseif ($allColors -contains $arg)
        {
            $color = $arg.substring(1)
        }
        else
        {
            if ($color -eq "Foreground")
            {
                Write-Host $arg -nonewline
            }
            else
            {
                Write-Host $arg -foreground $color -nonewline
            }
        }
    }

    Write-Host -nonewline:$nonewline
}

function Invoke-LoadMethod() {
param(
$ClientObject = $(throw "Please provide an Client Object instance on which to invoke the generic method")
)
$ctx = $ClientObject.Context
$load = [Microsoft.SharePoint.Client.ClientContext].GetMethod("Load")
$type = $ClientObject.GetType()
$clientObjectLoad = $load.MakeGenericMethod($type)
$clientObjectLoad.Invoke($ctx,@($ClientObject,$null))
}

function AddCSOM(){

     #Load SharePoint client dlls
     $a = [System.Reflection.Assembly]::LoadFile(    "$myDllPath\Microsoft.SharePoint.Client.dll")
     $at = [System.Reflection.Assembly]::LoadFile(    "$myDllPath\Microsoft.SharePoint.Client.Taxonomy.dll")
     $ar = [System.Reflection.Assembly]::LoadFile(    "$myDllPath\Microsoft.SharePoint.Client.Runtime.dll")
     $up = [System.Reflection.Assembly]::LoadFile(    "$myDllPath\Microsoft.SharePoint.Client.UserProfiles.dll")
    
     if( !$a ){
         $a = [System.Reflection.Assembly]::LoadWithPartialName(        "Microsoft.SharePoint.Client")
     }
     if( !$ar ){
         $ar = [System.Reflection.Assembly]::LoadWithPartialName(        "Microsoft.SharePoint.Client.Runtime")
     }
     if( !$up ){
         $up = [System.Reflection.Assembly]::LoadWithPartialName(        "Microsoft.SharePoint.Client.UserProfiles")
     }
    
     if( !$a -or !$ar ){
         throw         "Could not load Microsoft.SharePoint.Client.dll or Microsoft.SharePoint.Client.Runtime.dll"
     }
    
    
     #Add overload to the client context.
     #Define new load method without type argument
     $csharp =     "
      using Microsoft.SharePoint.Client;
      namespace SharepointClient
      {
          public class PSClientContext: ClientContext
          {
              public PSClientContext(string siteUrl)
                  : base(siteUrl)
              {
              }

              public void Load2(ClientObject objectToLoad)
              {
                  base.Load(objectToLoad);
              }
          }
      }"
    
     $assemblies = @( $a.FullName, $ar.FullName, $at.FullName,  "System.Core")
     Add-Type -TypeDefinition $csharp -ReferencedAssemblies $assemblies
}

function JavaDateToDateTime($msec)
{
    $date = new-object System.DateTime(1970, 1, 1, 0, 0, 0, 0);
    $date = $date.AddMilliseconds($msec);
    return $date;
}

function GetEpoch()
{
    $date1 = Get-Date -Date "01/01/1970"
    $date2 = Get-Date
    return (New-TimeSpan -Start $date1 -End $date2).TotalSeconds;
}

function GetCacheFileName($cs)
{
    $cachePath = "c:\temp\HttpCache\";

    try
    {
        [Directory]::CreateDirectory($cachePath);
    }
    catch { }

    $fileName = "";
    $prefix = "";

    $check = [DateTime]::Now;

    if ($cs.DateOverride -ne [DateTime]::MinValue)
    {
        $check = $cs.DateOverride;
    }

    switch ([MyHttp.CacheFrequency]$cs.Frequency)
    {
        "Daily"
        {
            $prefix = "D" + $check.ToString("MM-dd-yyyy");
            break;
        }
        "Hourly"
        {
            $prefix = "H" + $check.ToString("hh-MM-dd-yyyy");
            break;
        }
        "Monthly"
        {
            $prefix = "M" + $check.ToString("MM-yyyy");
            break;
        }
        "Weekly"
        {
            $prefix = "W" + $check.ToString("w");
            break;
        }
        "Yearly"
        {
            $prefix = "Y" + $check.ToString("yyyy");
            break;
        }
    }

    $fileName = $cachePath + $prefix + "_" + $cs.Category + "_" + $cs.Id + ".html";

    return $fileName;
}