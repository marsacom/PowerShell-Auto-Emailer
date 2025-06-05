#Module needed to connect to backend SQL server for API info
Import-Module SqlServer

#Global vars for auth info we will need
$global:client_id = ""
$global:client_secret = ""
$global:bearer = ""
$global:tenant = ""

$server= ""
$db= ""

#The user/pass used to auth to SQL
$user = $Env:VERIFY_BACKUP_USER
$pass = $Env:VERIFY_BACKUP_PASS

#Create an object we can use for passing credentials to dbatools when connecting to SQL
$credentials = Get-Credential -Credential (New-Object System.Management.Automation.PSCredential($user, ((ConvertTo-SecureString $pass -AsPlainText -Force))))
 
#Path to the log file
$log = ""

#These are the sender and the recipient(s) of the email
$msgSender = ""

#Depending on what day it is, Tuesday or Friday we are going to send a different message so set the corresponding vars to their respective values
If (((Get-Date).DayOfWeek) -eq "Tuesday") {
    $global:customer = ""
    $global:service1 = ""
    $global:service2 = ""
    $global:service3 = ""
    $global:service4 = ""
    $global:msgRecipient = ""
    $global:technician = ""
}else{
    $global:customer = ""
    $global:service1 = ""
    $global:service2 = ""
    $global:service3 = ""
    $global:service4 = ""
    $global:msgRecipient = ""
    $global:technician = ""
}

$htmlMsg = @" 
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Weekly Backup Verification Required</title>
</head>
<body style="margin:0; padding:0; background-color:#f4f4f4;">
  <table align="center" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; max-width:600px; background-color:#1e1e1e; font-family:Arial, sans-serif; color:#ffffff;">
    <tr>
      <td style="padding:20px;">
        <h2 style="color:#ffffff; margin-top:0;">Weekly Backup Verification Required</h2>
        <p style="color:#dcdcdc;">Good morning $company,</p>
        <p style="color:#dcdcdc;">
          This is your weekly reminder to verify that all backup systems have completed successfully for
          <strong>$customer</strong>. Please confirm the status of the following services:
        </p>

        <ul style="color:#ffffff; padding-left:20px;">
          <li>$service1</li>
          <li>$service2</li>
          <li>$service3</li>
          <li>$service4</li>
          <li>Endpoint Workstation Backups (if applicable)</li>
        </ul>

        <p style="color:#dcdcdc;">
          Ensure that all scheduled jobs have succeeded and no errors are present in the logs.
          If issues are found, address problems as needed and rerun backups.
        </p>
      </td>
    </tr>
    <tr>
      <td style="padding:10px; text-align:center; font-size:12px; color:#888888;">
        This is an automated message. Please do not reply.
      </td>
    </tr>
  </table>
</body>
</html>
"@

Function GetAPIInfo {
    Write-Host $Env:VERIFY_BACKUP_USER
    Write-Host $Env:VERIFY_BACKUP_PASS
    try {        
        $keys = @("", "", "") 
        foreach ($key in $keys) {
            $query = ""
            Write-Host $query
            $result = Invoke-Sqlcmd -Credential $credentials -ServerInstance $server -Database $db -Query $query -TrustServerCertificate
            if ($key -eq '') {
                $global:client_id = $result.TokenValue
                Write-Output "Client ID: $global:client_id"
            } elseif ($key -eq '') {
                $global:client_secret = $result.TokenValue
                Write-Output "Client Secret: $global:client_secret"
            } elseif ($key -eq '') {
                $global:tenant = $result.TokenValue
                Write-Output "Tenant: $global:tenant"
            }
        }          
    } catch {
        $err = $_.Exception.Message
        $stack = $_.Exception.StackTrace
        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"    
        Add-Content -Path $log -Value "$ts - ERROR encountered while attempting to obtain API credentials: $err, $stack"
    }
}

Function Get-Token {
    GetAPIInfo
        $url = "https://login.microsoftonline.com/$tenant/oauth2/v2.0/token"

        $body = @{
            client_id = $client_id
            scope = "https://graph.microsoft.com/.default"
            client_secret = $client_secret
            grant_type = "client_credentials"
        }
    try {    
        $response = Invoke-WebRequest -Method POST -Uri $url -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing

        $global:bearer = ($response.Content | ConvertFrom-Json).access_token
        return $bearer
    } catch {
        $err = $_.Exception.Message
        $stack = $_.Exception.StackTrace
        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"    
        Add-Content -Path $log -Value "$ts - ERROR encountered while attempting to obtain token from Microsoft Graph API: $err, $stack"
    }
}

Function Send-Messsage {
    Get-Token
    $url = "https://graph.microsoft.com/v1.0/users/$msgSender/sendMail"

    $headers = @{
        'Content-Type' = "application\json"
        'Authorization' = "Bearer $bearer" 
    }

    $email = @{
        message = @{
            subject = "AUTO : Verify Backup Status"
            body = @{
                contentType = "HTML"
                content = $htmlMsg
            }
            toRecipients = @(
                @{
                    emailAddress = @{
                        address = $msgRecipient
                    }
                }
            )
            from = @{
                emailAddress = @{
                    address = $msgSender
                }
            }
        }
    }

    try {
        $emailJSON = $email  | ConvertTo-Json -Depth 100
        $response = Invoke-WebRequest -Uri $url -Method POST -Headers $headers -Body $emailJSON -ContentType "application/json"
        $status = $response.StatusCode
        Write-Output "STATUS : $status"
        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        if ($status -ge 200 -and $status -lt 300) {
            Add-Content -Path $log -Value "$ts - Successfully sent email with status: $status"
        } else {
            Add-Content -Path $log -Value "$ts - Encountered error sending email, status: $status"
        }

    } catch {
        $err = $_.Exception.Message
        $stack = $_.Exception.StackTrace
        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"    
        Add-Content -Path $log -Value "$ts - ERROR encountered while attempging to send email: $err, $stack"
    }
}

#Main
Send-Messsage
