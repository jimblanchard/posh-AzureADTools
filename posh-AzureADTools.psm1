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
		$ObjectType,
		[Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
		[string]
		$APIVersion = "beta",
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
		$UseDeltaQuery,
		[string]
		$DeltaLink
    )

	$results = @{}
	$results.Values = $null
	$results.DeltaLink = $null

	if ($UseDeltaQuery -and ($null -notlike $DeltaLink))
	{
				Write-Verbose "Using DeltaQueryLink!"	
				$uri = $DeltaLink
	}
	else
	{
		if ($Attributes -like $null)
		{
			if ($UseDeltaQuery)
			{
				$uri = ("https://{0}/{1}/{2}s/delta?top={3}" -f $EndPoint,$APIVersion,$ObjectType,$top)
			}
			else
			{
				$uri = ("https://{0}/{1}/{2}s?top={3}" -f $EndPoint,$APIVersion,$ObjectType,$top)
			}
		}
		else
		{
			if ($UseDeltaQuery)
			{
				$selectAttributes = $attributes -join ','
				$uri = ("https://{0}/{1}/{2}s/delta?select={3}&top={4}" -f $Endpoint,$APIVersion,$ObjectType,$selectAttributes,$top)
			}
			else
			{
				$selectAttributes = $attributes -join ','
				$uri = ("https://{0}/{1}/{2}s?select={3}&top={4}" -f $Endpoint,$APIVersion,$ObjectType,$selectAttributes,$top)
			}
		}
	}

	write-debug ("DEBUG:MS Graph URI:{0}" -f $uri)
	$cmd = 'Invoke-RestMethod -Method Get -Uri $Uri -Headers $AuthHeader'
	$statusMsg = "VEBOSE:Invoking Expression $cmd"
	write-verbose $statusMsg
	$activityName = $MyInvocation.InvocationName

	Write-Progress -Id 1 -Activity $activityName -Status $statusMsg
	$x = $null
	try
	{
		$x = Invoke-Expression $cmd
	}
	catch
	{
		write-error $_
	}
	$pagedUri = $Null

	if ($x)
	{
		$i = 1

		do 
		{
		   Write-Verbose ("VERBOSE:Query Paging page {0} for {1}" -f $i++,$ObjectType )
		  
			$results.Values += $x.value
			if (Get-Member -inputobject $x -name '@odata.nextlink' -MemberType Properties)
			{
				$pagedUri = $x.'@odata.nextlink'
				if (Get-Member -inputobject $x -name '@odata.deltalink' -MemberType Properties)
				{
					$results.deltalink = $x.'@odata.deltalink'
					Write-Verbose ("Delta Link: {0}" -f $results.deltalink)
				}
				if ($pagedUri -notlike $Null)
				{
					Write-Debug ("DEBUG:Getting Next Page of results using Paging URI: {0}" -f $pagedUri )
					$cmd = 'Invoke-RestMethod -Method Get -Uri $pagedUri -Headers $AuthHeader'
					$x = $null
					$x = Invoke-Expression $cmd
				}
			}
			else
			{
				$pagedUri = $null
			}
		}
		until ($pagedUri -eq $Null)

		if (Get-Member -inputobject $x -name '@odata.deltalink' -MemberType Properties)
			{
				$results.deltalink = $x.'@odata.deltalink'
				Write-Verbose ("Delta Link: {0}" -f $results.deltalink)
			}
		
    }
	
	Write-Output ([pscustomobject]$results)
}