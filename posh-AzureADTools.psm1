$Global:msAD = $null
$Global:msGraph = $null
$Global:AzureADPowerShellClientId = '1950a258-227b-4e31-a9cf-717495945fc2'
$Global:OAuth2OobReplyUri = 'urn:ietf:wg:oauth:2.0:oob'
$Global:OAuth2AutoReplyUri = 'urn:ietf:wg:oauth:2.0:oob:auto'
$Global:MicrosoftGraphEndpointUri = 'https://graph.microsoft.com'

function Set-MyAzureADEnvironment {
	$Global:msAD = Get-AzureADServicePrincipal -Filter "AppId eq '00000002-0000-0000-c000-000000000000'"
	$Global:msGraph = Get-AzureADServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"
}

function Get-MyAzureADBasicPermissionSet {
	$requiredResourceAccess = New-Object -TypeName "Microsoft.Open.AzureAD.Model.RequiredResourceAccess"
	$requiredResourceAccess.ResourceAppId = $msAD.AppId
	$userReadPermission = $msad.Oauth2Permissions | Where-Object {$_.Value -eq 'User.Read'}
	$userRead = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList $userReadPermission.Id,"Scope"
	$requiredResourceAccess.ResourceAccess = $userRead
	return $requiredResourceAccess
}

function New-MyAzureADApplicationRegistration {
    [CmdletBinding()]
	param(
	[Parameter(Position=0)]
    [string]$DisplayName,
    [Parameter(Position=1)]
    [string]$HomePage,
    [Parameter(Position=2)]
    [string]$IdentifierUri = $HomePage,
    [Parameter(Position=3)]
	[Array]$ReplyUrls = $HomePage,
    [Parameter(Position=4)]
    [boolean]$ImplicitFlow = $false,
	[Parameter(Position=5)]
    [Microsoft.Open.AzureAD.Model.RequiredResourceAccess]$Permissions = $null
	)
	[Collections.Generic.List[String]]$Urls = $ReplyUrls
	$application = New-AzureADApplication -DisplayName $DisplayName -HomePage $HomePage -IdentifierUris $IdentifierUri -ReplyUrls $Urls -Oauth2AllowImplicitFlow $ImplicitFlow -RequiredResourceAccess $Permissions
	$principal = New-AzureADServicePrincipal -AccountEnabled $true -AppId $application.AppId -DisplayName $application.DisplayName -Tags {WindowsAzureActiveDirectoryIntegratedApp} -AppRoleAssignmentRequired $true
	return $application
}
function Get-MyAzureADGraphAuthorizationToken {
	$azureAdContext = Get-AzureADCurrentSessionInfo
	$domain = $azureAdContext.TenantId
	$authority = "https://login.windows.net/$domain"
	$authorizationContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
	$platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList "Always"
	
	$authorizationToken = $authorizationContext.AcquireTokenAsync($MicrosoftGraphEndpointUri, $AzureADPowerShellClientId, $OAuth2OobReplyUri, $platformParameters).Result
	return @{'Content-Type'='application\json';'Authorization'=$authorizationToken.CreateAuthorizationHeader()}
}	
function Get-MyAzureADGraphObjects {
	[CmdletBinding()]
    [OutputType([int])]
    Param
    (
        # Authorization Token from Azure AD
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $AuthHeader,
		[Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
		[string]
		$ObjectCollection,
		[Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
		[string]
		$ApiVersion = "beta",
		[Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
		[string[]]
		$Attributes,
		[Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
		[string]
		$Filter,
		[string]
		$EndPoint = "graph.microsoft.com",
		$Top = "999",
		[switch]
		$AsDeltaQuery,
		[string]
		$DeltaLink
    )

	$results = @{}
	$results.Values = $null
	$results.DeltaLink = $null

	$select = If($Attributes -eq $null) {""} Else {"select=$Attributes&"}
	$maxResults = "top=$Top"
	$deltaQuery = If($AsDeltaQuery) {"/delta"} Else {""}

	if ($DeltaLink -like $null)
	{
		$uri = "https://${EndPoint}/${ApiVersion}/${ObjectCollection}${deltaQuery}?${select}${maxResults}"
	}
	else
	{
		$uri = $DeltaLink
	}

	$expressionResult = $null
	try
	{
		$expressionResult = Invoke-Expression "Invoke-RestMethod -Method Get -Uri $Uri -Headers $AuthHeader"
	}
	catch
	{
		write-error $_
	}

	if ($expressionResult)
	{
		$nextPage = $null
		do 
		{
			$results.Values += $expressionResult.value
			if (Get-Member -inputobject $expressionResult -name '@odata.nextlink' -MemberType Properties)
			{
				$nextPage = $expressionResult.'@odata.nextlink'
				if (Get-Member -inputobject $expressionResult -name '@odata.deltalink' -MemberType Properties)
				{
					$results.deltalink = $expressionResult.'@odata.deltalink'
					Write-Verbose ("Delta Link: {0}" -f $results.deltalink)
				}
				if ($nextPage -notlike $Null)
				{
					Write-Debug ("DEBUG:Getting Next Page of results using Paging URI: {0}" -f $nextPage )
					$cmd = 'Invoke-RestMethod -Method Get -Uri $nextPage -Headers $AuthHeader'
					$expressionResult = $null
					$expressionResult = Invoke-Expression $cmd
				}
			}
			else
			{
				$nextPage = $null
			}
		}
		until ($nextPage -eq $null)

		if (Get-Member -inputobject $x -name '@odata.deltalink' -MemberType Properties)
		{
			$results.deltalink = $x.'@odata.deltalink'
			Write-Verbose ("Delta Link: {0}" -f $results.deltalink)
		}
    }
	
	Write-Output ([pscustomobject]$results)
}