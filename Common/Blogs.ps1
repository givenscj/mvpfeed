function ConvertXmlEntryToPost($xml)
{
    $post = new-object Post;
    $post.Id = $xml.Id;
    $post.PostDate = $xml.published;

    foreach($link in $xml.link)
    {
        if ($link.rel -eq "alternate")
        {
            $post.Url = $link.href;
        }
    }
    
    $post.Title = $xml.title.'#cdata-section';
    $post.Quantity = 1;
    $post.Reach = 1000;

    return $post;

}

function ConvertRssEntryToPost($rss)
{
    $post = new-object Post;
    $post.Id = $rss.link;
    $post.PostDate = $rss.pubDate;
    $post.Url = $rss.Link;
    $post.Title = $rss.title.'#cdata-section';
    $post.Quantity = 1;
    $post.Reach = 1000;

    return $post;
}

$global:posts = new-object system.collections.hashtable;

function GetBlogPage($blogRssFeedUrl, $type, $page)
{
    if ($page -ne 1)
    {
        switch($type)
        {
            "Rss"
            {
                $blogRssFeedUrl += "/$page";
            }
            "Atom"
            {
                $blogRssFeedUrl += "?paged=$page";
            }
        }               
    }

    [xml]$hsg = Invoke-WebRequest $blogRssFeedUrl;

    $user = "system";
    
    #not atom
    if ($hsg.rss)
    {
        $type = "Rss";

        foreach($post in $hsg.rss.channel.item)
        {            
            $post = ConvertRssEntryToPost $post;
            
            if (!$global:posts.ContainsKey($post.url))
            {
                $global:posts.Add($post.url, $post);
            }
        }
    }
    else
    {
        $type = "Atom";

        #atom...
        foreach($post in $hsg.DocumentElement.entry)
        {
            $post = ConvertXmlEntryToPost $post;
            
            if (!$global:posts.ContainsKey($post.url))
            {
                $global:posts.Add($post.url, $post);
            }
        }
    }

    $type;
    $global:posts;
}

function GetBlog($blogRssFeedUrl)
{
    $page = 1;

    $lastCount = 0;

    while ($newcount -ne $lastCount)
    {        
        $lastCount = $global:posts.Count;

        $vals = GetBlogPage $blogRssFeedUrl $type $page;    
        $type = $vals[0];

        $newCount = $global:posts.Count;

        $page++;
    }

    return $global:posts;
}