#
# PaperCut.ps1 - PaperCut Web Services API
#


$Log_MaskableKeys = @(
    # Put a comma-separated list of attribute names here, whose value should be masked before 
    'Password'
)

#
# System functions
#
function Idm-SystemInfo {
    param (
        # Operations
        [switch] $Connection,
        [switch] $TestConnection,
        [switch] $Configuration,
        # Parameters
        [string] $ConnectionParams
    )

    Log info "-Connection=$Connection -TestConnection=$TestConnection -Configuration=$Configuration -ConnectionParams='$ConnectionParams'"

    if ($Connection) {
        @(
            @{
                name = 'Hostname'
                type = 'textbox'
                label = 'Hostname'
                description = 'Hostname for Web Services'
                value = ''
            }
            @{
                name = 'authtoken'
                type = 'textbox'
                label = 'Authentication Token'
                description = ''
                value = ''
            }
            @{
                name = 'pagesize'
                type = 'textbox'
                label = 'Page Size'
                label_indent = $true
                description = 'Number of records per page'
                value = '250'
            }
            @{
                name = 'ignore_ssl_trust'
                type = 'checkbox'
                label = 'Skip SSL Trust Check'
                value = $false                  # Default value of checkbox item
            }
            @{
                name = 'use_proxy'
                type = 'checkbox'
                label = 'Use Proxy'
                description = 'Use Proxy server for requets'
                value = $false                  # Default value of checkbox item
            }
            @{
                name = 'proxy_address'
                type = 'textbox'
                label = 'Proxy Address'
                description = 'Address of the proxy server'
                value = 'http://localhost:8888'
                disabled = '!use_proxy'
                hidden = '!use_proxy'
            }
            @{
                name = 'use_proxy_credentials'
                type = 'checkbox'
                label = 'Use Proxy'
                description = 'Use Proxy server for requets'
                value = $false
                disabled = '!use_proxy'
                hidden = '!use_proxy'
            }
            @{
                name = 'proxy_username'
                type = 'textbox'
                label = 'Proxy Username'
                label_indent = $true
                description = 'Username account'
                value = ''
                disabled = '!use_proxy_credentials'
                hidden = '!use_proxy_credentials'
            }
            @{
                name = 'proxy_password'
                type = 'textbox'
                password = $true
                label = 'Proxy Password'
                label_indent = $true
                description = 'User account password'
                value = ''
                disabled = '!use_proxy_credentials'
                hidden = '!use_proxy_credentials'
            }
            @{
                name = 'nr_of_sessions'
                type = 'textbox'
                label = 'Max. number of simultaneous sessions'
                description = ''
                value = 1
            }
            @{
                name = 'sessions_idle_timeout'
                type = 'textbox'
                label = 'Session cleanup idle time (minutes)'
                description = ''
                value = 1
            }
        )
    }

    if ($TestConnection) {
        
    }

    if ($Configuration) {
        @()
    }

    Log info "Done"
}

function Idm-OnUnload {
}

#
# Object CRUD functions
#
$Properties = @{
    "User" = @(
        @{ name = 'Username';                              options = @('default','key')                      }
        @{ name = 'balance';                              options = @('default')                      }
        @{ name = 'primary-card-number';                              options = @('default')                      }
        @{ name = 'secondary-card-number';                              options = @('default')                      }
        @{ name = 'department';                              options = @('default')                      }
        @{ name = 'disabled-print';                              options = @('default')                      }
        @{ name = 'email';                              options = @('default')                      }
        @{ name = 'full-name';                              options = @('default')                      }
        @{ name = 'internal';                              options = @('default')                      }
        @{ name = 'notes';                              options = @('default')                      }
        @{ name = 'office';                              options = @('default')                      }
        @{ name = 'delegated-users';                              options = @('default')                      }
        @{ name = 'delegated-groups';                              options = @('default')                      }
        @{ name = 'print-stats.job-count';                              options = @('default')                      }
        @{ name = 'print-stats.page-count';                              options = @('default')                      }
        @{ name = 'net-stats.data-mb';                              options = @('default')                      }
        @{ name = 'net-stats.time-hours';                              options = @('default')                      }
        @{ name = 'restricted';                              options = @('default')                      }
        @{ name = 'home';                              options = @('default')                      }
        @{ name = 'unauthenticated';                              options = @('default')                      }
        @{ name = 'username-alias';                              options = @('default')                      }
        @{ name = 'dont-hold-jobs-in-release-station';                              options = @('default')                      }
        @{ name = 'dont-apply-printer-filter-rules';                              options = @('default')                      }
        @{ name = 'printer-cost-adjustment-rate-percent';                              options = @('default')                      }
        @{ name = 'dont-archive';                              options = @('default')                      }
        @{ name = 'auto-release-jobs';                              options = @('default')                      }
        @{ name = 'overdraft-amount';                              options = @('default')                      }
        @{ name = 'account-selection.mode';                              options = @('default')                      }
        @{ name = 'account-selection.can-charge-personal';                              options = @('default')                      }
        @{ name = 'account-selection.can-charge-shared-from-list';                              options = @('default')                      }
        @{ name = 'account-selection.can-charge-shared-by-pin';                              options = @('default')                      }
        @{ name = 'other-emails';                              options = @('default')                      }
        @{ name = 'auto-shared-account';                              options = @('default')                      }

    )
}

function Idm-UsersRead {
    param (
        [switch] $GetMeta,
        [string] $SystemParams,
        [string] $FunctionParams
    )
    $Class = "User"
    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {

        Get-ClassMetaData -SystemParams $SystemParams -Class $Class
    }
    else {
        
        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        $properties = $function_params.properties

        if ($properties.length -eq 0) {
            $properties = ($Global:Properties.$Class | Where-Object { $_.options.Contains('default') }).name
        }

        # Assure key is the first column
        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name
        $properties = @($key) + @($properties | Where-Object { $_ -ne $key })

        try { 
                    $results = [System.Collections.ArrayList]@()
                    
                    # Gather Usernames
                    $userList = [System.Collections.ArrayList]@()

                    while($true) {

                        $xmlRequest = '<?xml version="1.0" encoding="UTF-8"?>
                        <methodCall>
                            <methodName>api.listUserAccounts</methodName>
                            <params>
                                <param>
                                    <value>
                                        <string>{0}</string>
                                    </value>
                                </param>
                                <param>
                                    <value>
                                        <int>{1}</int>
                                    </value>
                                </param>
                                <param>
                                    <value>
                                        <int>{2}</int>
                                    </value>
                                </param>
                            </params>
                        </methodCall>' -f $system_params.authtoken, $userList.Count, $system_params.pagesize

                        Log info ("Retrieving User List - Offset ($($userList.Count))")
                        $response = Invoke-PaperCutRequest -system_params $system_params -function_params $function_params -Body $xmlRequest
                        
                        Log info "Retrieving User List - Returned ($($response.methodResponse.params.param.value.array.data.value.count))"

                        if($response.methodResponse.params.param.value.array.data.value.count -lt 1) {
                            break
                        }

                        foreach($item in $response.methodResponse.params.param.value.array.data.value) {
                            [void]$userList.Add($item)

                            #[void]$results.Add([PSCustomObject]@{ "Username" = $item})
                        }
                    }   
                    
                    Log info "Processing Users"
                    # Retrieve User Properties
                    $i = 0
                    foreach($user in $userList) {
                        $i++
                        Log info "Processing user $($i) of $($userList.count)"
                        
                        $userObject = Get-UserProperties -system_params $system_params -function_params $function_params -username $user
                        
                        [void]$results.Add($userObject)
                    }
                
            }
            catch {
                Log error "Failed: $_"
                Write-Error $_
            }

            $results
    }

    Log info "Done"
}

function Idm-UsersRename {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'update'
            parameters = @(
                @{ name = 'Username';       allowance = 'mandatory'   }
                @{ name = 'NewUsername';    allowance = 'mandatory' }
                @{ name = '*';              allowance = 'prohibited' }
            )
        }
    }
    else {
        #
        # Execute function
        #
        $system_params = ConvertFrom-Json2 $SystemParams
        $function_params   = ConvertFrom-Json2 $FunctionParams

        $properties = $function_params.Clone()

        try {
            $xmlRequest = '<?xml version="1.0" encoding="UTF-8"?>
                        <methodCall>
                            <methodName>api.renameUserAccount</methodName>
                            <params>
                                <param>
                                    <value>
                                        <string>{0}</string>
                                    </value>
                                </param>
                                <param>
                                    <value>
                                        <string>{1}</string>
                                    </value>
                                </param>
                                <param>
                                    <value>
                                        <string>{2}</string>
                                    </value>
                                </param>
                            </params>
                        </methodCall>' -f $system_params.authtoken, $properties.Username, $properties.NewUsername

            LogIO info 'Idm-UsersRename' -In @system_params -Properties $properties

            $response = Invoke-PaperCutRequest -system_params $system_params -function_params $function_params -Body $xmlRequest
            
            LogIO info 'Idm-UsersRename' -Out -Response $response
            
            # Out new object
            [PSCustomObject]@{
                Username = $properties.NewUsername
            }
        }
        catch {
            LogIO error 'Idm-UsersRename' -Out -Response $_
            Write-Error $_
        }
    }
    Log info "Done"
}

# https://www.papercut.com/help/manuals/ng-mf/common/tools-web-services/#:~:text=set%20to%20%27individual%27-,api.setUserProperties,-Set%20multiple%20user
<#
balance: the user's balance, unformatted, e.g. "1234.56".
primary-card-number
secondary-card-number
department
disabled-print: true if the user's printing is disabled, otherwise false
email
full-name
internal: true if this is an internal user, otherwise false
notes
office
delegated-users: the user names the given user can release print jobs for
delegated-groups: the group names the given user can release print jobs for
print-stats.job-count: total number of print jobs from this user, unformatted, e.g. "1234"
print-stats.page-count: total number of pages printed by this user, unformatted, e.g. "1234"
net-stats.data-mb: total 'net MB used by this user, unformatted, e.g. "1234.56"
net-stats.time-hours: total 'net hours used by this user, unformatted, e.g. "1234.56"
restricted: true if this user's printing is restricted, false if they are unrestricted.
home: the user's home folder (a double-quoted UNC path)
unauthenticated: true if the user is an unauthenticated user, {otherwise false
username-alias: The alias for a given user
dont-hold-jobs-in-release-station: true if the user's jobs will bypass all release station queues, otherwise false
dont-apply-printer-filter-rules: true if the user's jobs will bypass printer filter settings, otherwise false
printer-cost-adjustment-rate-percent: The percentage modifier for the user's job costs, unformatted, e.g. "123.45". If the flag to enable adjustment is not set, returns -1
dont-archive: true if the user's jobs will not be archived, false otherwise.
auto-release-jobs: true if the user's jobs will always release on device login, otherwise false
overdraft-amount: the user's individual overdraft amount, unformatted, e.g. "1234.56". Note this amount is in use only when the user account is restricted and overdraft mode is set to 'individual'
#>
function Idm-UsersUpdate {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'update'
            parameters = @(
                @{ name = 'Username';                               allowance = 'mandatory'   }
                @{ name = 'balance';                                allowance = 'optional'    }
                @{ name = 'primary-card-number';                    allowance = 'optional'    }
                @{ name = 'secondary-card-number';                  allowance = 'optional'    }
                @{ name = 'department';                             allowance = 'optional'    }
                @{ name = 'disabled-print';                         allowance = 'optional'    }
                @{ name = 'email';                                  allowance = 'optional'    }
                @{ name = 'full-name';                              allowance = 'optional'    }
                @{ name = 'internal';                               allowance = 'optional'    }
                @{ name = 'notes';                                  allowance = 'optional'    }
                @{ name = 'office';                                 allowance = 'optional'    }
                @{ name = 'delegated-users';                        allowance = 'optional'    }
                @{ name = 'delegated-groups';                       allowance = 'optional'    }
                @{ name = 'print-stats.job-count';                  allowance = 'optional'    }
                @{ name = 'print-stats.page-count';                 allowance = 'optional'    }
                @{ name = 'net-stats.data-mb';                      allowance = 'optional'    }
                @{ name = 'net-stats.time-hours';                   allowance = 'optional'    }
                @{ name = 'restricted';                             allowance = 'optional'    }
                @{ name = 'home';                                   allowance = 'optional'    }
                @{ name = 'unauthenticated';                        allowance = 'optional'    }
                @{ name = 'username-alias';                         allowance = 'optional'    }
                @{ name = 'dont-hold-jobs-in-release-station';      allowance = 'optional'    }
                @{ name = 'dont-apply-printer-filter-rules';        allowance = 'optional'    }
                @{ name = 'printer-cost-adjustment-rate-percent';   allowance = 'optional'    }
                @{ name = 'dont-archive';                           allowance = 'optional'    }
                @{ name = 'auto-release-jobs';                      allowance = 'optional'    }
                @{ name = 'overdraft-amount';                       allowance = 'optional'    } 
                @{ name = '*';                                      allowance = 'prohibited'  }
            )
        }
    }
    else {
        #
        # Execute function
        #
        $system_params = ConvertFrom-Json2 $SystemParams
        $function_params   = ConvertFrom-Json2 $FunctionParams
        
        $properties = $function_params.Clone()
        $properties.remove('Username')

        # Parse Properties to update.
        $propertiesXML = '';
        foreach($prop in $properties.GetEnumerator()) {
            $propertiesXML += @'
    <value>
        <array>
            <data>
                <value>
                    <string>{0}</string>
                </value>
                <value>
                    <string>{1}</string>
                </value>
            </data>
        </array>
    </value>
'@ -f $prop.name, $prop.value
        } #end foreach
        # Example Updated Property
        <#
        <value>
            <array>
                <data>
                    <value>
                        <string>department</string>
                    </value>
                    <value>
                        <string>newdept</string>
                    </value>
                </data>
            </array>
        </value>
        #>

        try {
            $xmlRequest = '<?xml version="1.0" encoding="UTF-8"?>
                        <methodCall>
                            <methodName>api.setUserProperties</methodName>
                            <params>
                                <param>
                                    <value>
                                        <string>{0}</string>
                                    </value>
                                </param>
                                <param>
                                    <value>
                                        <string>{1}</string>
                                    </value>
                                </param>
                                <param>
                                    <value>
                                        <array>
                                            <data>
                                                {2}
                                            </data>
                                        </array>
                                    </value>
                                </param>
                            </params>
                        </methodCall>' -f $system_params.authtoken, $properties.Username, $propertiesXML
            
            LogIO info 'Idm-UsersUpdate' -In -Parameters $function_params
            
            $response = Invoke-PaperCutRequest -system_params $system_params -function_params $function_params -Body $xmlRequest
            
            # Output is a boolean.
            if ($response = 1) {
                
                LogIO info 'Idm-UsersUpdate' -Out -Response $response
            
            } else {
            
                #Log error ("Failed to Update Properties for ({0}), response: ({1})" -f $properties.Username, ($response | ConvertTo-Json -Depth 5))
                LogIO error 'Idm-UsersUpdate' -Out -Response $response
                throw ("Failed to Update Properties for ({0}), response: ({1})" -f $properties.Username, ($response | ConvertTo-Json -Depth 5))
            
            }
            # Out new object
            Get-UserProperties -system_params $system_params -function_params $function_params -username $function_params.Username
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }
    Log info "Done"
}

# 
# Functions
# 

function Get-UserProperties {
    param (
        [hashtable] $system_params,
        [hashtable] $function_params,
        [string] $username
    )

    $xmlRequest = '<?xml version="1.0" encoding="UTF-8"?>
        <methodCall>
            <methodName>api.getUserProperties</methodName>
            <params>
                <param>
                    <value>
                        <string>{0}</string>
                    </value>
                </param>
                <param>
                    <value>
                        <string>{1}</string>
                    </value>
                </param>
                <param>
                    <value>
                        <array>
                            <data>
                                <value>
                                    <string>balance</string>
                                </value>
                                <value>
                                    <string>primary-card-number</string>
                                </value>
                                <value>
                                    <string>secondary-card-number</string>
                                </value>
                                <value>
                                    <string>department</string>
                                </value>
                                <value>
                                    <string>disabled-print</string>
                                </value>
                                <value>
                                    <string>email</string>
                                </value>
                                <value>
                                    <string>full-name</string>
                                </value>
                                <value>
                                    <string>internal</string>
                                </value>
                                <value>
                                    <string>notes</string>
                                </value>
                                <value>
                                    <string>office</string>
                                </value>
                                <value>
                                    <string>delegated-users</string>
                                </value>
                                <value>
                                    <string>delegated-groups</string>
                                </value>
                                <value>
                                    <string>print-stats.job-count</string>
                                </value>
                                <value>
                                    <string>print-stats.page-count</string>
                                </value>
                                <value>
                                    <string>net-stats.data-mb</string>
                                </value>
                                <value>
                                    <string>net-stats.time-hours</string>
                                </value>
                                <value>
                                    <string>restricted</string>
                                </value>
                                <value>
                                    <string>home</string>
                                </value>
                                <value>
                                    <string>unauthenticated</string>
                                </value>
                                <value>
                                    <string>username-alias</string>
                                </value>
                                <value>
                                    <string>dont-hold-jobs-in-release-station</string>
                                </value>
                                <value>
                                    <string>dont-apply-printer-filter-rules</string>
                                </value>
                                <value>
                                    <string>printer-cost-adjustment-rate-percent</string>
                                </value>
                                <value>
                                    <string>dont-archive</string>
                                </value>
                                <value>
                                    <string>auto-release-jobs</string>
                                </value>
                                <value>
                                    <string>overdraft-amount</string>
                                </value>
                                <value>
                                    <string>account-selection.mode</string>
                                </value>
                                <value>
                                    <string>account-selection.can-charge-personal</string>
                                </value>
                                <value>
                                    <string>account-selection.can-charge-shared-from-list</string>
                                </value>
                                <value>
                                    <string>account-selection.can-charge-shared-by-pin</string>
                                </value>
                                <value>
                                    <string>other-emails</string>
                                </value>
                                <value>
                                    <string>auto-shared-account</string>
                                </value>
                            </data>
                        </array>
                    </value>
                </param>
            </params>
        </methodCall>' -f $system_params.authtoken, $username

        $response = Invoke-PaperCutRequest -system_params $system_params -function_params $function_params -Body $xmlRequest
        $attributes = $response.methodResponse.params.param.value.array.data.value
        $userObject = [PSCustomObject]@{
            "Username" = $username
            "balance" = $attributes[0]
            "primary-card-number" = $attributes[1]
            "secondary-card-number" = $attributes[2]
            "department" = $attributes[3]
            "disabled-print" = $attributes[4]
            "email" = $attributes[5]
            "full-name" = $attributes[6]
            "internal" = $attributes[7]
            "notes" = $attributes[8]
            "office" = $attributes[9]
            "delegated-users" = $attributes[10]
            "delegated-groups" = $attributes[11]
            "print-stats_job-count" = $attributes[12]
            "print-stats_page-count" = $attributes[13]
            "net-stats_data-mb" = $attributes[14]
            "net-stats_time-hours" = $attributes[15]
            "restricted" = $attributes[16]
            "home" = $attributes[17]
            "unauthenticated" = $attributes[18]
            "username-alias" = $attributes[19]
            "dont-hold-jobs-in-release-station"  = $attributes[20]
            "dont-apply-printer-filter-rules" = $attributes[21]
            "printer-cost-adjustment-rate-percent" = $attributes[22]
            "dont-archive" = $attributes[23]
            "auto-release-jobs" = $attributes[24]
            "overdraft-amount" = $attributes[25]
            "account-selection_mode" = $attributes[26]
            "account-selection_can-charge-personal" = $attributes[27]
            "account-selection_can-charge-shared-from-list" = $attributes[28]
            "account-selection_can-charge-shared-by-pin" = $attributes[29]
            "other-emails" = $attributes[30]
            "auto-shared-account" = $attributes[31]
        }
        
        $userObject
}

function Invoke-PaperCutRequest {
    param (
        [hashtable] $system_params,
        [hashtable] $function_params,
        [string] $Body
    )

    if($system_params.ignore_ssl_trust) {
        add-type @"
    using System.Net;
    using System.Security.Cryptography.X509Certificates;
    public class TrustAllCertsPolicy : ICertificatePolicy {
        public bool CheckValidationResult(
            ServicePoint srvPoint, X509Certificate certificate,
            WebRequest request, int certificateProblem) {
            return true;
        }
    }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Ssl3, [Net.SecurityProtocolType]::Tls, [Net.SecurityProtocolType]::Tls11, [Net.SecurityProtocolType]::Tls12
    }
    
    $uri = "https://{0}:9192/rpc/api/xmlrpc" -f $system_params.hostname

   
    $headers= @{
        'Content-Type' = 'text/xml;charset=UTF-8'
    }
    
    try {
        $splat = @{
            Method = "POST"
            Uri = $uri
            Headers = $headers
            Body = $Body
        }

        if($system_params.use_proxy)
        {
            Log info ("Using proxy for PaperCut Request")
            $splat["Proxy"] = $system_params.proxy_address

            if($system_params.use_proxy_credentials)
            {
                $splat["proxyCredential"] = New-Object System.Management.Automation.PSCredential ($system_params.proxy_username, (ConvertTo-SecureString $system_params.proxy_password -AsPlainText -Force) )
            }
        }

        $response = Invoke-RestMethod @splat -ErrorAction Stop
        $result = [xml]$response
    }
    catch [System.Net.WebException] {
       
        try {
            $reader = New-Object System.IO.StreamReader -ArgumentList $_.Exception.Response.GetResponseStream()
            $response = $reader.ReadToEnd()
            $reader.Close()

            $result = ([xml]$response)
        }
        catch {}
        
        $message = "Error : $($_)"
        Log error $message
        Write-Error $_
    }
    catch {
        $message = "Error : $($_)"
        Log error $message
        Write-Error $_
    }
    finally {
        Write-Output $result
    }
}
