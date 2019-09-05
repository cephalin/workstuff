$Style = "<style>
table {
  border-collapse: collapse;
}

table, th, td {
  border: 1px solid black;
  padding: 3px;
}

th {
  background-color: #4CAF50;
  color: white;
}

tr:nth-child(even) {background-color: #f2f2f2;}
</style>"

$Issues = Find-GitHubIssue -State open -Labels app-service/svc -Repo MicrosoftDocs/azure-docs
$Issues += Find-GitHubIssue -State open -Labels app-service-web/svc -Repo MicrosoftDocs/azure-docs
$Formatted = ($Issues | Sort-Object -Descending -Property @{Expression={$_.Assignee.login}; Descending=$true}, @{Expression={$_.created_at} ;Descending=$false} | select @{Name="Issue#"; Expression={"<a href='" + $_.html_url + "'>" + $_.number + "</a>"}},@{Name="Assignee"; Expression={$_.Assignee.login}},@{Name="Days Old"; Expression={((Get-Date) - [DateTime] $_.created_at).Days}})

$Intro = "<p>Dear team,</p>
<p>If you're on the To line, then you have an outstanding GitHub issue or PR in App Service docs. As shown in
the table below there are currently <b style=""color:rgb(255,0,0);"">$($Formatted.Count)</b> open issues and PRs 
labeled app-service/svc or app-service-web/svc, <i>many of which are quite stale and in urgent need of your 
attention</i>. Please do your part in responding to the customer queries in docs and move the issues/PRs toward 
the closed/merged state.</p>
<ul>
<li>If you need help making an update to the docs, I'd be happy to help as long as you 
provide the necessary technical information in the issue comments and add a comment with <b>#reassign:cephalin</b>.</li>
<li>If you no longer own a specific area, please reassign to its new owner by adding a comment with <b>#reassign:github-user-handle</b>.</li>
<li>If you feel you've resolved the customer's query and wish to close an issue, just add a comment with 
<b>#please-close</b>, and the issue will be closed.</li>
</ul>
<p> Many thanks in advance!</p>
<p>Cephas Lin</p>"

Add-Type -AssemblyName System.Web
Add-type -assembly "Microsoft.Office.Interop.Outlook"
$Outlook = New-Object -ComObject Outlook.Application
$Mail = $Outlook.CreateItem(0)
$Mail.To = ($Formatted | select Assignee | Foreach {"$($_.Assignee)"} | Get-Unique | Out-String)
$Mail.To = (((((($Mail.To -replace "ggailey777", "Glenn.Gailey") -replace "msangapu-msft", "msangapu") -replace "mattchenderson", "Matthew.Henderson") -replace "Jen7714", "Jennifer.Lee") -replace "SnehaAgrawal-MSFT", "Agrawal.Sneha") -replace "Grace-MacJones-MSFT", "Grace.Macjones") -replace "AjayKumar-MSFT", "Kumar.Ajay"
$Mail.CC = "george.wallace; msangapu; stefsch"
$Mail.Subject = "Outstanding GitHub issues in App Service docs"
$Mail.HTMLBody = ([System.Web.HttpUtility]::HtmlDecode(($Formatted | ConvertTo-Html -Head $Style -PreContent $Intro | Out-String)))
$Mail.Display()
