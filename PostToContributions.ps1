<#########################################################
#
#  Contribution Areas
#



##############################################################>

<#########################################################
#
#  Activity Types
#



##############################################################>


function LoadConfig()
{
    $strJson = Get-Content "PostToContributions.config" -Raw;
    $global:Config = ConvertFrom-json $strJson;
}

function Initalize()
{
    $global:username = read-host "What is your MS ID username?";
    $global:password = read-host "What is your MS ID password?";

    $global:defaultReach = $global:config.BlogReach;
    $global:defaultArea = $global:config.AwardCategory;

    $global:startDate = GetMvpStartDate;
    $global:endDate = $global:startDate.addyears(1);    

    Login $global:username $global:password;

    #load all the activities
    LoadActivities;

    #load contribution areas
    GetContributionAreas;

    #load all activity types
    LoadActivityTypes;
}

$global:posts = new-object system.collections.hashtable;

function ProcessBlog($url)
{    
    $global:posts = GetBlog $url $global:defaultReach;

    foreach($post in $global:posts.values)
    {    
        if ($post.PostDate -ge $global:startdate -and $post.PostDate -le $global:enddate)
        {
            $exists = $global:mvpActivities.Values | where {$_.Url -eq $post.Url};

            if (!$exists)
            {
                AddPostToMVPProfile($post);
            }
            else
            {
                write-host "BLOG found in profile [$($post.PostDate)] : [$($post.title)]"; 
            }
        }
    }    
}

function ProcessOutlookCalendar()
{
}

function ProcessGoogleCalendar()
{
}

function ProcessSPSEvents($username)
{
}

function ProcessMeetups($username)
{
}

$global:youtube = new-object system.collections.hashtable;

function ProcessYouTubeChannel($channelUrl)
{
    $global:youtube.Clear();
    
    $html = DoGet $channelUrl;

    $json = ParseValue $html "window[`"ytInitialData`"] = " "}};";
    $json += "}}"

    $obj = ConvertFrom-Json $json;
    
    foreach($item in $obj.contents.twoColumnBrowseResultsRenderer.tabs[0].tabRenderer.content.sectionListRenderer.contents[0].itemSectionRenderer.contents[0].shelfRenderer.content.horizontalListRenderer.items)
    {
        $post = new-object post;
        $post.Title = $item.gridVideoRenderer.title.simpleText;
        $videoId = $item.gridVideoRenderer.videoId;    
        $post.Quantity = $item.gridVideoRenderer.viewCountText.simpleText.replace("views","").trim();

        $url = "https://www.youtube.com/watch?v=$videoId";
        $html = DoGet $url;

        $post.url = $url;
        $post.postdate = ParseValue $html "Published on" "`"";

        AddYouTubeToMVPProfile($post);
    }
}

function ProcessYouTubeUser($userUrl)
{
    $global:youtube.Clear();
    
    $html = DoGet $userUrl;

    $json = ParseValue $html "window[`"ytInitialData`"] = " "}};";
    $json += "}}"

    $obj = ConvertFrom-Json $json;
    
    foreach($item in $obj.contents.twoColumnBrowseResultsRenderer.tabs[1].tabRenderer.content.sectionListRenderer.contents[0].itemSectionRenderer.contents[0].gridRenderer.items)
    {
        $post = new-object post;
        $post.Title = $item.gridVideoRenderer.title.simpleText;
        $videoId = $item.gridVideoRenderer.videoId;        

        $url = "https://www.youtube.com/watch?v=$videoId";
        $html = DoGet $url;

        $post.postdate = ParseValue $html "Published on" "<";

        AddYouTubeToMVPProfile($post);
    }
}

function ParseYouTubeDate($val)
{
    
}

function ProcessTechCommunity($userId)
{
    #https://techcommunity.microsoft.com/t5/user/viewprofilepage/user-id/
    
    $page = 0;
    $count = 0;

    while($page -lt 5)
    {
        $url = "https://techcommunity.microsoft.com/gxcuf89792/plugins/custom/microsoft/o365/custom-messages-feed-showmore-conversations"
        $post = "userId=$userId&pageOffsetValue=$count&type=startedConversations";
        $html = DoPost $url $post;

        $htmlDoc = new-object HtmlAgilityPack.HtmlDocument;
        $htmlDoc.LoadHtml($html);
        $homeNode = $htmlDoc.DocumentNode;

        $nodes = $homeNode.SelectNodes(".//li");

        foreach($node in $nodes)
        {
            $post = new-object post;
            $post.Title = ParseValue $node.InnerHtml "subject-text`">" "<";
            $url = ParseValue $node.innerhtml "subject-link`" href=`"" "`"";
            $post.url = "https://techcommunity.microsoft.com/" + $url;
            $post.postdate = ParseValue $html "time`">,&nbsp;" "<";
            $post.reach = $(ParseValue $html "views`">" "<").replace(",","").replace("Views","").trim();
            $post.quantity = 1;

            AddPostToMVPProfile($post);
            
            $count++;
        }

        $page++;
    }
}

function LinkedInPosts($username)
{
    #https://techcommunity.microsoft.com/t5/user/viewprofilepage/user-id/56
}

function ProcessMSDN($username)
{
    $html = DoGet "https://social.msdn.microsoft.com/Profile/$username/activity";
    $userid = ParseValue $html "" "";

    $userid = "2785652e-ff0d-47ba-b804-a46ab9b907e9";
    $startDate = "636213064270000000";

    $url = "https://api.recognition.microsoft.com/v2/users/$userId/reputations/year?startDate=$startDate&callback=userreputationCallback";

    $html = DoGet $url;
    $json = ConvertFrom-Json $html;
    
    #add the points to the profile with the 12/31/xxxx date...
    $post = new Post;
    $post.PostDate = "12/31/2018";
    
    AddMsdnToMVPProfile($post);
}

function ProcessGitHub($userName)
{
    $year = 2017;
    GetGitHubActivityForYear $username $year;

    foreach($a in $global:github.values)
    {
        if (!$exists)
        {
            if ($a.PostDate -ge $global:startDate -and $a.PostDate -le $global:enddate)
            {
                AddGitHubToMVPProfile($a);
            }
        }
    }
}

function ProcessAmazonBooks($authorUrl)
{
    GetAmazonBooks $authorUrl;

    foreach($a in $global:books.values)
    {
        if (!$exists)
        {            
            if ($a.PostDate -ge $global:startDate -and $a.PostDate -le $global:enddate)
            {
                AddBookToMVPProfile($a);
            }
        }
    }
}



function GetMvpStartDate()
{
    $date = [datetime]::now;
    $year = $date.Year;

    if ($date.month -ge 6)
    {
        return [Datetime]::parse($date.tostring("07/01/$year"));
    }
    else
    {
        return [Datetime]::Parse($date.tostring("07/01/$($year-1)"));
    }
}

function GetDocsActivity($username)
{
    
}

$global:github = new-object system.collections.hashtable;

function GetGitHubActivityForYear($username, $year)
{
    $global:github.Clear();    

    $startdate = [Datetime]::parse("1/1/$year");
    $enddate = [Datetime]::parse("1/1/$year");

    for($i=1;$i-le12;$i++)
    {        
        $date = [datetime]::parse("$i/1/$year");
        $startofmonth = Get-Date $date -day 1 -hour 0 -minute 0 -second 0;
        $endofmonth = (($startofmonth).AddMonths(1).AddSeconds(-1));

        $start = $startofmonth.tostring("yyyy-MM-dd");
        $end = $endofMonth.tostring("yyyy-MM-dd");

        GetGitHubActivity $username $start $end;
    }
}

function GetGitHubActivity($username, $start, $end)
{    
    #generic get
    $url = "https://github.com/$username" + "?tab=overview&from=$start&to=$end&_pjax=%23js-contribution-activity"
    $html = DoGet $url;

    $results = new-object System.Collections.Hashtable;

    #commits get
    $url = "https://github.com/users/$username/created_commits?from=$start&to=$end";
    $html = DoGet $url;

    $htmlDoc = new-object HtmlAgilityPack.HtmlDocument;
    $htmlDoc.LoadHtml($html);
    $homeNode = $htmlDoc.DocumentNode;

    $items = $homeNode.SelectNodes(".//li[@class='ml-0 py-1']");
    
    foreach($item in $items)
    {
        $as = $item.SelectNodes(".//a");

        $a0 = $as[0];
        $a1 = $as[1];

        $post = new-object Post;
        $post.url = "http://github.com" + $a0.attributes["href"].value;
        $val = $a1.innertext.replace("commits","").trim();
        $val = $val.replace("commit","");
        $postDate = ParseValue $item.OuterHtml "since=" "&"
        $post.title = "Contributions - " + $post.url;
        $post.PostDate = $postDate;
        $post.Quantity = $val;
        
        if (!$global:github.ContainsKey($post.url))
        {
            $global:github.add($post.url, $post);
        }
}

    #repos created
    $url = "https://github.com/users/$username/created_repositories?from=$start&to=$end";
    $html = DoGet $url;

    $htmlDoc = new-object HtmlAgilityPack.HtmlDocument;
    $htmlDoc.LoadHtml($html);
    $homeNode = $htmlDoc.DocumentNode;

    $items = $homeNode.SelectNodes(".//li");

    foreach($item in $items)
    {
        $as = $item.SelectNodes(".//a");

        $a0 = $as[0];
        $a1 = $as[1];

        $post = new-object Post;
        $post.url = "http://github.com" + $a0.attributes["href"].value;
        $post.title = "Created - " + $post.url;
        $val = $a1.innertext.replace("commits","").trim();
        $val = $val.replace("commit","");
        $post.Quantity = $val;        
        $post.Reach = 1;
        $postDate = ParseValue $item.OuterHtml "since=" "&"
        $post.PostDate = $postDate;

        if (!$global:github.ContainsKey($post.url))
        {
            $global:github.add($post.url, $post);
        }
    }    
}

$global:books = new-object System.Collections.Hashtable;

function GetAmazonBooks($authorUrl)
{    
    $html = DoGet $authorUrl;

    $htmlDoc = new-object HtmlAgilityPack.HtmlDocument;
    $htmlDoc.LoadHtml($html);
    $homeNode = $htmlDoc.DocumentNode;

    $books = $homeNode.SelectNodes(".//li");

    foreach($book in $Books)
    {
        if ($book.attributes["data-asin"])
        {
            $id = $book.attributes["data-asin"].value;
            $url = "https://www.amazon.com/blah/dp/$id"
            $bookInfoHtml = DoGet $url;

            $htmlDoc2 = new-object HtmlAgilityPack.HtmlDocument;
            $htmlDoc2.LoadHtml($bookInfoHtml);
            $homeNode2 = $htmlDoc2.DocumentNode;

            $details = $homeNode2.SelectNodes(".//div[@class='content']");

            if ($details.Count -gt 1)
            {
                foreach($node in $details)
                {
                    if ($node.innertext.contains("ISBN"))
                    {
                        $bookDetails = $node;
                        break;
                    }
                }
            }
            else
            {
                $bookdetails = $details[0];
            }

            if ($bookDetails)
            {
                $post = new-object Post;
                $post.url = $url;
                $post.postdate = $(ParseValue $bookDetails.OuterHtml "(" ")");
                $post.Title = $homeNode2.SelectSingleNode(".//span[@id='productTitle']").InnerText;
                $post.Quantity = 1;
                $post.Reach = 1;
                
                if (!$global:books.ContainsKey($post.url))
                {
                    $global:books.add($post.url, $post);
                }
            }
        }
    }    
}

function ByteArraysAreEqual($one, $two)
{
   $BYTES_TO_READ = $one.length;
 
   if ($one.Length -ne $two.Length)
   {
        return $false;
   }  

    $iterations = 1;
 
   for ($i = 0; $i -lt $iterations; $i = $i + 1)
   {       
       if ([BitConverter]::ToInt64($one, 0) -ne 
           [BitConverter]::ToInt64($two, 0))
       {
           return $false;
       }
   }
  
   return $true;
}

function AddActivityToMVPProfile($activity, $visibility, $activityTypeId, $technologyId)
{
    $html = DoGet "https://mvp.microsoft.com/en-us/MyProfile/EditActivity";
    $requestToken = ParseValue $html "__RequestVerificationToken`" type=`"hidden`" value=`"" "`"";

    $post = "__RequestVerificationToken=$RequestToken";
    $post += "&formchanged=false";
    $post += "&PrivateSiteId=0";
    $post += "&ActivityType.Id=$activityTypeId";
    $post += "&ApplicableTechnology.Id=$technologyId";
    $post += "&ApplicableTechnology.Name=$technologyId";
    $post += "&ActivityVisibility.Id=$visibility";
    $post += "&DateOfActivity=$($activity.PostDate.tostring("MM-d-yyyy"))";  #12-8-2017
    $title = UrlDecode($activity.Title);
    $title = HtmlDecode($title);
    $post += "&TitleOfActivity=$title";
    $purl = UrlDecode($activity.Url);
    $purl = HtmlDecode($purl);
    $post += "&ReferenceUrl=$purl";
    $post += "&Description=$title";
    $post += "&SecondAnnualQuantity=$($activity.Quantity)";
    $post += "&AnnualQuantity=$($activity.Quantity)";
    $post += "&AnnualReach=$($activity.Reach)";    

    $url = "https://mvp.microsoft.com/en-us/MyProfile/SaveActivity";
    $results = DoPost $url $post;

    $json = ConvertFrom-Json $results;

    if ($json.result -eq "success")
    {
        $global:mvpActivities.Add($activity.url, $activity);
    }
}

function LookupVisibilityId($in)
{
    switch($in.tolower())
    {
        "public"
        {
            return 299600000;
        }
        "everyone"
        {
            return 299600000;
        }
        "microsoft"
        {
            return 100000000;
        }
        "mvp"
        {
            return 100000001;
        }
        "private"
        {
            return 100000000;
        }
    }
}

function LookupTechnologyId($in)
{
    $id = $global:contributionareas[$in];

    if (!$id)
    {
        foreach($key in $global:contributionareas.keys)
        {
            $id = $global:contributionareas[$key];

            if ($key.contains($in))
            {
                return $id;
            }
        }
    }

    if (!$id)
    {
        #grab the first on of the personal areas
        foreach($key in $global:personalareas.keys)
        {
            return $global:personalareas[$key];
        }
    }

    return $id;    
}

function LookupActivityId($in)
{
    $id = $global:activities[$in];

    if (!$id)
    {
        foreach($key in $global:activities.keys)
        {
            $id = $global:activities[$key];

            if ($key.contains($in))
            {
                return $id;
            }
        }
    }
    return $id;    
}

function AddGitHubToMVPProfile($post)
{    
    write-host "CODE SAMPLE [$($post.Url)] - Adding [$($post.Title)]";

    $v = LookupVisibilityId "Everyone";
    $t = LookupTechnologyId $global:defaultArea;
    $a = LookupActivityId "Code Samples";
    
    AddActivityToMVPProfile $post $v $a $t;
}

function AddBookToMVPProfile($post)
{    
    write-host "BOOK [$($post.Url)] - Adding [$($post.Title)]";

    $v = LookupVisibilityId "Everyone";
    $t = LookupTechnologyId $global:defaultArea;
    $a = LookupActivityId "Book (Author)";
    
    AddActivityToMVPProfile $post $v $a $t;
}

function AddMsdnToMVPProfile($post)
{    
    write-host "MSDN [$($post.Url)] - Adding [$($post.Title)]";

    $v = LookupVisibilityId "Everyone";
    $t = LookupTechnologyId $global:defaultArea;
    $a = LookupActivityId "Book (Author)";
    
    AddActivityToMVPProfile $post $v $a $t;
}

function AddYouTubeToMVPProfile($post)
{    
    write-host "YouTube [$($post.Url)] - Adding [$($post.Title)]";

    $v = LookupVisibilityId "Everyone";
    $t = LookupTechnologyId $global:defaultArea;
    $a = LookupActivityId "Video";
    
    AddActivityToMVPProfile $post $v $a $t;
}

function AddPostToMVPProfile($post)
{    
    write-host "BLOG [$($post.PostDate)] - Adding [$($post.Title)]";

    $v = LookupVisibilityId "Everyone";
    $t = LookupTechnologyId $global:defaultArea;
    $a = LookupActivityId "Post";
    
    AddActivityToMVPProfile $post $v $a $t;
}

$global:mvpActivities = new-object system.collections.hashtable;

function LoadActivities()
{
    $url = "https://mvp.microsoft.com/en-us/MyProfile/EditActivity";
    $html = DoGet $url;
    
    $results = ParseValue $html "all_activities = " "baseMapUrl";
    $results = $results.trim().substring(0, $results.trim().length-1);

    $json = ConvertFrom-Json $results;

    foreach($activity in $json)
    {
        if($activity.ReferenceUrl)
        {
            $post = new-object Post;
            $post.PostDate = $activity.DateOfActivityFormatted;
            $post.id = $activity.privatesiteid;
            $post.Quantity = $activity.AnnualQuantity;
            $post.Title = $activity.TitleOfActivity;
            $post.Url = $activity.ReferenceUrl;

            $global:mvpActivities.Add($post.id, $post);
        }
    }
}

$global:activities = new-object system.collections.hashtable;

function LoadActivityTypes()
{
    $url = "https://mvp.microsoft.com/en-us/MyProfile/EditActivity";
    $html = DoGet $url;
    $requestToken = ParseValue $html "__RequestVerificationToken`" type=`"hidden`" value=`"" "`"";

    $url = "https://mvp.microsoft.com/en-us/MyProfile/AddActivityDialog";
    $post = "__RequestVerificationToken=$RequestToken";
    $html = DoPost $url $post;

    $htmlDoc = new-object HtmlAgilityPack.HtmlDocument;
    $htmlDoc.LoadHtml($html);
    $homeNode = $htmlDoc.DocumentNode;

    $table = $homeNode.SelectSingleNode(".//select[@id='activityTypeSelector']");

    $html = $table.InnerHtml;

    while($html.contains("<option"))
    {
        $id = ParseValue $html "value=`"" "`"";
        $name = ParseValue $html ">" "<";

        $global:activities.add($name.trim(), $id.trim());

        $html = $html.substring($html.IndexOf("<option") + 7);

    }    
}

$global:personalareas = new-object system.collections.hashtable;
$global:contributionareas = new-object system.collections.hashtable;

function GetContributionAreas()
{
    $url = "https://mvp.microsoft.com/MyProfile/GetContributionAreas";
    $results = DoGet $url;

    $json = ConvertFrom-Json $results;

    foreach($area in $json[0].Contributions)
    {
        foreach($ca in $area.contributionarea)
        {
            $global:contributionAreas.add($ca.name,$ca.Id);
            $global:personalareas.add($ca.name,$ca.Id);
        }
    }

    foreach($area in $json[1].Contributions)
    {
        foreach($ca in $area.contributionarea)
        {
            $global:contributionAreas.add($ca.name,$ca.Id);
        }
    }
}

function GetMVPId()
{    
    $profile = Get-MvpProfile;

    $global:mvpId = $profile.MvpId;
    $global:mvpTechnology = $profile.AwardCategoryDisplay;
}

function Ping()
{
    $html = DoGet "https://mvp.microsoft.com/en-us/Account/SignIn" $global:strCookies;

    if ($global:location.contains("login"))
    {
        return $false;
    }

    SetCookiesUnformatted "https://mvp.microsoft.com" $global:strCookies;

    return $true;
}

function Login($username, $password)
{
    $global:strCookies = Get-content "mvp.cookie" -ea SilentlyContinue;

    $loggedIn = Ping;

    if (!$global:strCookies -or !$loggedIn)
    {
        $url = "https://mvp.microsoft.com/en-us/Account/SignIn";
        $html = DoGet $url;    
        $url = $global:location;    

        #get some cookie action..
        $html = DoGet $url;    

        #get the PPFT value
        $ppft = ParseValue $html "name=`"PPFT`"" ">";
        $ppft = ParseValue $ppft "value=`"" "`"";    
    
        $url = ParseValue $html "urlPost:'" "'";

        $uaid = $global:urlCookies["login.live.com"]["uaid"];

        #make the call to "getcredential type"...just in case   
        $global:referer = $url; 
        $global:headers.add("Origin","https://login.live.com");
        $global:headers.add("client-request-id",$uaid);
        $global:headers.add("hpgid","33");
        $global:headers.add("hpgact","0");
        $global:overrideContentType = "application/json";

        $credsurl = $url.replace("https://login.live.com/ppsecure/post.srf","https://login.live.com/GetCredentialType.srf");
        $post = "{`"username`":`"$username`",`"uaid`":`"$uaid`",`"isOtherIdpSupported`":false,`"checkPhones`":false,`"isRemoteNGCSupported`":true,`"isCookieBannerShown`":false,`"isFidoSupported`":false,`"flowToken`":`"$ppft`"}";
        $results = DoPost $credsurl $post;    
        $json = ConvertFrom-Json $results;

        $sessionid = $json.Credentials.RemoteNgcParams.SessionIdentifier;

        $global:location = $null;
        #do the post...
        $post = "i13=0&login=$username&loginfmt=$username&type=11&LoginOptions=3&lrt=&lrtPartition=&hisRegion=&hisScaleUnit=&passwd=$password&ps=2&psRNGCDefaultType=1&psRNGCEntropy=&psRNGCSLK=$sessionId&psFidoAllowList=&canary=&ctx=&PPFT=$ppft&PPSX=Pas&NewUser=1&FoundMSAs=&fspost=0&i21=0&CookieDisclosure=0&i2=1&i17=0&i18=__ConvergedLoginPaginatedStrings%7C1%2C__ConvergedLogin_PCore%7C1%2C&i19=27286";
        $html = DoPost $url $post;        

        #get the PPFT value
        $ppft = ParseValue $html "sFT:'" "'";        

        if ($global:location)
        {
            $html = DoGet $global:location;
        }
        else
        {
            $code = ParseValue $html ",f:'" "'";
            write-host "Please approve $code auth request";

            $slk = ParseValue $html "AI:'" "'";
            $url = ParseValue $html "urlPost:'" "'";

            #session state url...for polling...but uses image weirdness...so just try to hit the end point...
            $sessionStateUrl = ParseValue $html ",S:'" "'";

            #add the slk and dt part..
            $sessionStateUrl += "&slk=$slk&dt=1512513152619";

            #might have to wait for the token to verify...
            while($true)
            {   
                $global:referer = $url;

                #check session...                
                $html = DoGet $sessionStateUrl;

                #check for valid image...
                #$bytes = [System.IO.File]::ReadAllBytes("c:\temp\success.gif");
                #$equal = ByteArraysAreEqual $global:fileBuffer $bytes;
            
                #ends up the value in byte 32 is the indicator of SUCCESS!!!!
                if ($global:filebuffer[32] -eq 10)
                {
                    $post = "type=22&request=&mfaLastPollStart=&mfaLastPollEnd=&login=$(UrlEncode($username))"
                    $post += "&PPFT=$(UrlEncode($ppft))"
                    $post += "&sacxt=1&purpose=eOTT_OneTimePassword"
                    $post += "&SLK=$(UrlEncode($slk))"
                    $post += "&i2=&i17=0&i18=__ConvergedSAStrings%7C1%2C__ConvergedSA_Core%7C1%2C&i19=7572";

                    $global:referer = $url; 
                    $html = DoPost $url $post;

                    $formAction = ParseValue $html "id=`"fmHF`" action=`"" "`"";                            

                    if ($formAction)
                    {              
                        $napExp = ParseValue $html "id=`"NAPExp`" value=`"" "`"";
                        $nap = ParseValue $html "id=`"NAP`" value=`"" "`"";
                        $anon = ParseValue $html "id=`"ANON`" value=`"" "`"";
                        $anonExp = ParseValue $html "id=`"ANONExp`" value=`"" "`"";
                        $t = ParseValue $html "id=`"t`" value=`"" "`"";
                          
                        $post = "NAPExp=$napExp&NAP=$nap&ANON=$anon&ANONExp=$anonExp&t=$t";
                        $html = DoPost $formAction $post;

                        break;
                    }               
                }          
            
                #check for verification...
                start-sleep 2;                   
            }
        }    

        remove-item "mvp.cookie" -ea SilentlyContinue;

        set-content "mvp.cookie" $(GetCookies "https://mvp.microsoft.com");
    }

    $url = "https://mvp.microsoft.com/en-us/MyProfile/EditProgramInfo";
    $html = DoGet $url;

    $htmlDoc = new-object HtmlAgilityPack.HtmlDocument;
    $htmlDoc.LoadHtml($html);
    $homeNode = $htmlDoc.DocumentNode;

    $table = $homeNode.SelectSingleNode(".//table[@class='raListTable']");

    $body = $table.ChildNodes[1];
    

    foreach($node in $body.ChildNodes)
    {
        if($node.childnodes[1].innertext -eq "Award Categories")
        {
            $global:mvpTechnologyId = $node.childnodes[3].innertext;
            #$global:mvpTechnologyId = "70c301bb-189a-e411-93f2-9cb65495d3c4";
        }

        if($node.childnodes[1].innertext -eq "MVP ID")
        {
            $global:mvpId = $node.childnodes[3].innertext;
        }
    }        
}

<####################################################
#
#
#   Originally wanted to use the MVP API, but getting the subscription key will be a bit too much for some...
#
#
####################################################>
<#
#https://mvpapi.portal.azure-api.net/docs/services/580eb8bfac2551138cf5da27/operations/580eb8bfac25510f0c09c1a5
#https://github.com/lazywinadmin/MVP
Install-Module -name MVP;

$SubscriptionKey = 'blah'
Set-MVPConfiguration -SubscriptionKey $SubscriptionKey
#>

#load helper dlls
Add-Type -Path "C:\Users\givenscj\OneDrive\My Scripts\Dlls\HtmlAgilityPack.dll";

$global:scriptcommonpath = "C:\Users\givenscj\OneDrive\My Scripts\Common";
$global:scriptpath = "C:\Users\givenscj\OneDrive\My Scripts\MVP";

cd $global:scriptPath;

#add helper files...
. "$global:scriptcommonpath\Util.ps1"
. "$global:scriptcommonpath\HttpHelper.ps1"
. "$global:scriptcommonpath\Blogs.ps1"

. "$global:scriptpath\Classes.ps1"

LoadConfig;

Initalize;

#https://techcommunity.microsoft.com/t5/user/viewprofilepage/user-id/62590
ProcessTechCommunity $config.TechCommunityId;

#$channelurl = "https://www.youtube.com/user/TEDtalksDirector/videos";
ProcessYouTubeChannel $config.YouTubeChannel;

#$authorUrl = "https://www.amazon.com/Ted-Pattison/e/B001JSBY5C";
ProcessAmazonBooks $config.AmazonUrl;

#$githubId= "https://github.com/givenscj";
ProcessGitHub $config.GitHubId;

<######################################################################################

NOTE:  You need a respectable blog platform that has a pagable ATOM/RSS feed...Medium does not count...

#######################################################################################>

ProcessBlog $config.BlogUrl;
