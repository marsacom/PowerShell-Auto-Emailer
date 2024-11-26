Import-Module SqlServer #Connect to backend SQL server for API info

#Global vars for auth info we will need
$global:client_id = ""
$global:client_secret = ""
$global:bearer = ""
$global:tenant = ""

$mode = "PROD" #Change from DEV to PROD
$company = "YOUR-COMPANY-NAME" #Your company name here

$db = ""
$server = ""
$serverDEV = "YOUR-DEV-SERVER"
$dbDEV = "YOUR-DEV-DATABASE"
$serverPROD = "YOUR-PROD-SERVER"
$dbPROD = "YOUR-PROD-DATABASE"

#The user/pass used to auth to SQL
$user = $Env:VERIFY_BACKUP_USER
$pass = $Env:VERIFY_BACKUP_PASS

#These are the sender and the recipient(s) of the email
$msgSender = "YOUR-EMAIL-SENDER"
$msgRecipient = "YOUR-EMAIL-RECIPIENT"
$msgCC = "YOUR-EMAIL-CC"

#Path to the log file you wish to use
$log

#Depending on what day it is, Tuesday or Friday we are going to send a different message so set the corresponding vars to their respective values
If (((Get-Date).DayOfWeek) -eq "Tuesday") {
    $global:customer = "YOUR-CUSTOMER"
    $global:service1 = ""
    $global:service2 = ""
    $global:service3 = ""
}else{
    $global:customer = ""
    $global:service1 = ""
    $global:service2 = ""
    $global:service3 = ""
}


#HTML to be sent to the recipients
$htmlMsg = @" 
<!DOCTYPE html>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">

<head>
 <meta charset="UTF-8" />
 <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
 <!--[if !mso]><!-- -->
 <meta http-equiv="X-UA-Compatible" content="IE=edge" />
 <!--<![endif]-->
 <meta name="viewport" content="width=device-width, initial-scale=1.0" />
 <meta name="format-detection" content="telephone=no" />
 <meta name="format-detection" content="date=no" />
 <meta name="format-detection" content="address=no" />
 <meta name="format-detection" content="email=no" />
 <meta name="x-apple-disable-message-reformatting" />
 <link href="https://fonts.googleapis.com/css?family=Nunito+Sans:ital,wght@0,200;0,400;0,700" rel="stylesheet" />
 <title>Thrivology</title>
 <!-- Made with Postcards Email Builder by Designmodo -->
 <style>
 html,
         body {
             margin: 0 !important;
             padding: 0 !important;
             min-height: 100% !important;
             width: 100% !important;
             -webkit-font-smoothing: antialiased;
         }
 
         * {
             -ms-text-size-adjust: 100%;
         }
 
         #outlook a {
             padding: 0;
         }
 
         .ReadMsgBody,
         .ExternalClass {
             width: 100%;
         }
 
         .ExternalClass,
         .ExternalClass p,
         .ExternalClass td,
         .ExternalClass div,
         .ExternalClass span,
         .ExternalClass font {
             line-height: 100%;
         }
 
         table,
         td,
         th {
             mso-table-lspace: 0 !important;
             mso-table-rspace: 0 !important;
             border-collapse: collapse;
         }
 
         u + .body table, u + .body td, u + .body th {
             will-change: transform;
         }
 
         body, td, th, p, div, li, a, span {
             -webkit-text-size-adjust: 100%;
             -ms-text-size-adjust: 100%;
             mso-line-height-rule: exactly;
         }
 
         img {
             border: 0;
             outline: 0;
             line-height: 100%;
             text-decoration: none;
             -ms-interpolation-mode: bicubic;
         }
 
         a[x-apple-data-detectors] {
             color: inherit !important;
             text-decoration: none !important;
         }
 
         .pc-gmail-fix {
             display: none;
             display: none !important;
         }
 
         .body .pc-project-body {
             background-color: transparent !important;
         }
 
         @media (min-width: 621px) {
             .pc-lg-hide {
                 display: none;
             } 
 
             .pc-lg-bg-img-hide {
                 background-image: none !important;
             }
         }
 </style>
 <style>
 @media (max-width: 620px) {
 .pc-project-body {min-width: 0px !important;}
 .pc-project-container {width: 100% !important;}
 .pc-sm-hide {display: none !important;}
 .pc-sm-bg-img-hide {background-image: none !important;}
 .pc-w620-padding-30-0-0-0 {padding: 30px 0px 0px 0px !important;}
 .pc-w620-fontSize-58px {font-size: 58px !important;}
 .pc-w620-padding-60-20-10-20 {padding: 60px 20px 10px 20px !important;}
 table.pc-w620-spacing-0-0-0-0 {margin: 0px 0px 0px 0px !important;}
 td.pc-w620-spacing-0-0-0-0,th.pc-w620-spacing-0-0-0-0{margin: 0 !important;padding: 0px 0px 0px 0px !important;}
 
 .pc-w620-gridCollapsed-1 > tbody,.pc-w620-gridCollapsed-1 > tbody > tr,.pc-w620-gridCollapsed-1 > tr {display: inline-block !important;}
 .pc-w620-gridCollapsed-1.pc-width-fill > tbody,.pc-w620-gridCollapsed-1.pc-width-fill > tbody > tr,.pc-w620-gridCollapsed-1.pc-width-fill > tr {width: 100% !important;}
 .pc-w620-gridCollapsed-1.pc-w620-width-fill > tbody,.pc-w620-gridCollapsed-1.pc-w620-width-fill > tbody > tr,.pc-w620-gridCollapsed-1.pc-w620-width-fill > tr {width: 100% !important;}
 .pc-w620-gridCollapsed-1 > tbody > tr > td,.pc-w620-gridCollapsed-1 > tr > td {display: block !important;width: auto !important;padding-left: 0 !important;padding-right: 0 !important;margin-left: 0 !important;}
 .pc-w620-gridCollapsed-1.pc-width-fill > tbody > tr > td,.pc-w620-gridCollapsed-1.pc-width-fill > tr > td {width: 100% !important;}
 .pc-w620-gridCollapsed-1.pc-w620-width-fill > tbody > tr > td,.pc-w620-gridCollapsed-1.pc-w620-width-fill > tr > td {width: 100% !important;}
 .pc-w620-gridCollapsed-1 > tbody > .pc-grid-tr-first > .pc-grid-td-first,pc-w620-gridCollapsed-1 > .pc-grid-tr-first > .pc-grid-td-first {padding-top: 0 !important;}
 .pc-w620-gridCollapsed-1 > tbody > .pc-grid-tr-last > .pc-grid-td-last,pc-w620-gridCollapsed-1 > .pc-grid-tr-last > .pc-grid-td-last {padding-bottom: 0 !important;}
 
 .pc-w620-gridCollapsed-0 > tbody > .pc-grid-tr-first > td,.pc-w620-gridCollapsed-0 > .pc-grid-tr-first > td {padding-top: 0 !important;}
 .pc-w620-gridCollapsed-0 > tbody > .pc-grid-tr-last > td,.pc-w620-gridCollapsed-0 > .pc-grid-tr-last > td {padding-bottom: 0 !important;}
 .pc-w620-gridCollapsed-0 > tbody > tr > .pc-grid-td-first,.pc-w620-gridCollapsed-0 > tr > .pc-grid-td-first {padding-left: 0 !important;}
 .pc-w620-gridCollapsed-0 > tbody > tr > .pc-grid-td-last,.pc-w620-gridCollapsed-0 > tr > .pc-grid-td-last {padding-right: 0 !important;}
 
 .pc-w620-tableCollapsed-1 > tbody,.pc-w620-tableCollapsed-1 > tbody > tr,.pc-w620-tableCollapsed-1 > tr {display: block !important;}
 .pc-w620-tableCollapsed-1.pc-width-fill > tbody,.pc-w620-tableCollapsed-1.pc-width-fill > tbody > tr,.pc-w620-tableCollapsed-1.pc-width-fill > tr {width: 100% !important;}
 .pc-w620-tableCollapsed-1.pc-w620-width-fill > tbody,.pc-w620-tableCollapsed-1.pc-w620-width-fill > tbody > tr,.pc-w620-tableCollapsed-1.pc-w620-width-fill > tr {width: 100% !important;}
 .pc-w620-tableCollapsed-1 > tbody > tr > td,.pc-w620-tableCollapsed-1 > tr > td {display: block !important;width: auto !important;}
 .pc-w620-tableCollapsed-1.pc-width-fill > tbody > tr > td,.pc-w620-tableCollapsed-1.pc-width-fill > tr > td {width: 100% !important;box-sizing: border-box !important;}
 .pc-w620-tableCollapsed-1.pc-w620-width-fill > tbody > tr > td,.pc-w620-tableCollapsed-1.pc-w620-width-fill > tr > td {width: 100% !important;box-sizing: border-box !important;}
 }
 </style>
 <!--[if !mso]><!-- -->
 <style>
 @font-face { font-family: 'Nunito Sans'; font-style: normal; font-weight: 700; src: url('https://fonts.gstatic.com/s/nunitosans/v15/pe1mMImSLYBIv1o4X1M8ce2xCx3yop4tQpF_MeTm0lfGWVpNn64CL7U8upHZIbMV51Q42ptCp5F5bxqqtQ1yiU4GMS5XvVUj.woff') format('woff'), url('https://fonts.gstatic.com/s/nunitosans/v15/pe1mMImSLYBIv1o4X1M8ce2xCx3yop4tQpF_MeTm0lfGWVpNn64CL7U8upHZIbMV51Q42ptCp5F5bxqqtQ1yiU4GMS5XvVUl.woff2') format('woff2'); } @font-face { font-family: 'Nunito Sans'; font-style: normal; font-weight: 200; src: url('https://fonts.gstatic.com/s/nunitosans/v15/pe1mMImSLYBIv1o4X1M8ce2xCx3yop4tQpF_MeTm0lfGWVpNn64CL7U8upHZIbMV51Q42ptCp5F5bxqqtQ1yiU4GVilXvVUj.woff') format('woff'), url('https://fonts.gstatic.com/s/nunitosans/v15/pe1mMImSLYBIv1o4X1M8ce2xCx3yop4tQpF_MeTm0lfGWVpNn64CL7U8upHZIbMV51Q42ptCp5F5bxqqtQ1yiU4GVilXvVUl.woff2') format('woff2'); }
 </style>
 <!--<![endif]-->
 <!--[if mso]>
    <style type="text/css">
        .pc-font-alt {
            font-family: Arial, Helvetica, sans-serif !important;
        }
    </style>
    <![endif]-->
 <!--[if gte mso 9]>
    <xml>
        <o:OfficeDocumentSettings>
            <o:AllowPNG/>
            <o:PixelsPerInch>96</o:PixelsPerInch>
        </o:OfficeDocumentSettings>
    </xml>
    <![endif]-->
</head>

<body class="body pc-font-alt" style="width: 100% !important; min-height: 100% !important; margin: 0 !important; padding: 0 !important; line-height: 1.5; color: #2D3A41; mso-line-height-rule: exactly; -webkit-font-smoothing: antialiased; -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; font-variant-ligatures: normal; text-rendering: optimizeLegibility; -moz-osx-font-smoothing: grayscale; background-color: #ffffff;" bgcolor="#ffffff">
 <table class="pc-project-body" style="table-layout: fixed; min-width: 600px; background-color: #ffffff;" bgcolor="#ffffff" width="100%" border="0" cellspacing="0" cellpadding="0" role="presentation">
  <tr>
   <td align="center" valign="top">
    <table class="pc-project-container" align="center" width="600" style="width: 600px; max-width: 600px;" border="0" cellpadding="0" cellspacing="0" role="presentation">
     <tr>
      <td class="pc-w620-padding-30-0-0-0" style="padding: 20px 0px 20px 0px;" align="left" valign="top">
       <table border="0" cellpadding="0" cellspacing="0" role="presentation" width="100%" style="width: 100%;">
        <tr>
         <td valign="top">
          <!-- BEGIN MODULE: Personal Letter -->
          <table width="100%" border="0" cellspacing="0" cellpadding="0" role="presentation">
           <tr>
            <td class="pc-w620-spacing-0-0-0-0" style="padding: 0px 0px 0px 0px;">
             <table width="100%" border="0" cellspacing="0" cellpadding="0" role="presentation">
              <tr>
               <td valign="top" class="pc-w620-padding-60-20-10-20" style="padding: 20px 40px 20px 40px; border-radius: 0px; background-color: transparent;" bgcolor="transparent">
                <table width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation">
                 <tr>
                  <td align="left" valign="top" style="padding: 0px 0px 40px 0px;">
                   <table border="0" cellpadding="0" cellspacing="0" role="presentation" width="100%" style="border-collapse: separate; border-spacing: 0; margin-right: auto; margin-left: auto;">
                    <tr>
                     <td valign="top" align="left" style="padding: 0px 0px 0px 0px;">
                      <div class="pc-font-alt pc-w620-fontSize-58px" style="line-height: 107%; letter-spacing: -0.2px; font-family: 'Nunito Sans', Arial, Helvetica, sans-serif; font-size: 50px; font-weight: 200; font-variant-ligatures: normal; color: #000a28; text-align: left; text-align-last: left;">
                       <div><span>Verify Backup Reminder</span>
                       </div>
                      </div>
                     </td>
                    </tr>
                   </table>
                  </td>
                 </tr>
                </table>
                <table width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation">
                 <tr>
                  <td align="left" valign="top" style="padding: 0px 0px 0px 0px;">
                   <table border="0" cellpadding="0" cellspacing="0" role="presentation" width="100%" style="border-collapse: separate; border-spacing: 0; margin-right: auto; margin-left: auto;">
                    <tr>
                     <td valign="top" align="left" style="padding: 0px 0px 0px 0px;">
                      <div class="pc-font-alt" style="line-height: 160%; letter-spacing: 0px; font-family: 'Nunito Sans', Arial, Helvetica, sans-serif; font-size: 16px; font-weight: normal; font-variant-ligatures: normal; color: #000a28; text-align: left; text-align-last: left;">
                       <div><span>Hello $company, </span>
                       </div>
                       <div><span>﻿</span>
                       </div>
                       <div><span>This is an automated email to remind technicians to verify backup services are running.</span>
                       </div>
                       <div><span>&#xFEFF;</span>
                       </div>
                       <div><span>Please ensure that all backups are up and running for $customer.</span>
                       </div>
                       <div><span>﻿</span>
                       </div>
                       <div><span style="font-weight: 700;font-style: normal;line-height: 250%;">Backup Services:</span>
                       </div>
                       <ol style="margin: 0; padding: 0 0 0 20px; list-style: arabic;">
                        <li><span style="font-weight: 700;font-style: normal;">$service1</span>
                        </li>
                        <li><span style="font-weight: 700;font-style: normal;">$service2</span>
                        </li>
                        <li><span style="font-weight: 700;font-style: normal;">$service3</span>
                        </li>
                       </ol>
                      </div>
                     </td>
                    </tr>
                   </table>
                  </td>
                 </tr>
                </table>
                <table width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="width: 100%;">
                 <tr>
                  <td valign="top" style="padding: 40px 0px 40px 0px;">
                   <table width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="margin: auto;">
                    <tr>
                     <!--[if gte mso 9]>
                    <td height="1" valign="top" style="line-height: 1px; font-size: 1px; border-bottom: 1px solid #cccccc;">&nbsp;</td>
                <![endif]-->
                     <!--[if !gte mso 9]><!-- -->
                     <td height="1" valign="top" style="line-height: 1px; font-size: 1px; border-bottom: 1px solid #cccccc;">&nbsp;</td>
                     <!--<![endif]-->
                    </tr>
                   </table>
                  </td>
                 </tr>
                </table>
                <table width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation">
                 <tr>
                  <td align="left" valign="top" style="padding: 0px 0px 20px 0px;">
                   <table border="0" cellpadding="0" cellspacing="0" role="presentation" width="100%" style="border-collapse: separate; border-spacing: 0; margin-right: auto; margin-left: auto;">
                    <tr>
                     <td valign="top" align="left" style="padding: 0px 0px 0px 0px;">
                      <div class="pc-font-alt" style="line-height: 20px; letter-spacing: -0.2px; font-family: 'Nunito Sans', Arial, Helvetica, sans-serif; font-size: 14px; font-weight: normal; font-variant-ligatures: normal; color: #000a28; text-align: left; text-align-last: left;">
                       <div><span>NOTE : This is an automated email from $msgSender, this mailbox is not monitored and does not reply to messages.</span>
                       </div>
                      </div>
                     </td>
                    </tr>
                   </table>
                  </td>
                 </tr>
                </table>
               </td>
              </tr>
             </table>
            </td>
           </tr>
          </table>
          <!-- END MODULE: Personal Letter -->
         </td>
        </tr>
       </table>
      </td>
     </tr>
    </table>
   </td>
  </tr>
 </table>
</body>
</html>
"@

If ($mode -eq "DEV"){
    $server = $serverDEV
    $db = $dbDEV
}else{
    $server = $serverPROD
    $db = $dbPROD
}

Function ConnectToSQL{
    try {
        Invoke-Sqlcmd -TrustServerCertificate -ServerInstance $server -Database $db -Username $user -Password $pass 
        Write-Output "Connecting to SQL Server/DB : $server.$db..."
    } catch {
        $err = $_.Exception.Message
        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"    
        Add-Content -Path $log -Value "$ts - ERROR: $err`n"
    }
}

Function GetAPIInfo {
    ConnectToSQL
    try {
        $query = "SELECT TokenName, TokenValue FROM APITokens WHERE TokenName LIKE 'VerifyBackup%';"
        $result = Invoke-Sqlcmd -TrustServerCertificate -Query $query -ServerInstance $server -Database $db

        if ($result[0][0] -eq "VerifyBackupClientID") {
            $global:client_id = $result[0][1]
            Write-Output "Client ID : $client_id"
        }else{
            Write-Output "ERROR: Unable to get Client ID..."
        }
        if ($result[1][0] -eq "VerifyBackupClientSecret") {
            $global:client_secret = $result[1][1]
            Write-Output "Client Secret : $client_secret"
        }else{
            Write-Output "ERROR: Unable to get Client Secret..."
        }
        if ($result[2][0] -eq "VerifyBackupTenantID") {
            $global:tenant = $result[2][1]
            Write-Output "Tenant ID : $tenant"
        }else{
            Write-Output "ERROR: Unable to get Tenant ID..."
        }
    } catch {
        $err = $_.Exception.Message
        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"    
        Add-Content -Path $log -Value "$ts - ERROR: $err`n"
    }
}

Function Get-Token {
    GetAPIInfo
    try {
        $url = "https://login.microsoftonline.com/$tenant/oauth2/v2.0/token"

        $body = @{
            client_id = $client_id
            scope = "https://graph.microsoft.com/.default"
            client_secret = $client_secret
            grant_type = "client_credentials"
        }

        $response = Invoke-WebRequest -Method POST -Uri $url -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing

        $global:bearer = ($response.Content | ConvertFrom-Json).access_token
        return $bearer
    } catch {
        $err = $_.Exception.Message
        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"    
        Add-Content -Path $log -Value "$ts - ERROR: $err`n"
    }
}

Function Send-Messsage {
    Get-Token
    try {
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
                    content = $htmlMsg#"This is an automated email on behalf of Adevity to remind tecnicians to verify ALL backup services are running for support customers..."
                }
                toRecipients = @(
                    @{
                        emailAddress = @{
                            address = $msgRecipient
                        }
                    }
                )
                ccRecipients = @(
                    @{
                        emailAddress = @{
                            address = $msgCC
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

        $emailJSON = $email  | ConvertTo-Json -Depth 100
        Invoke-RestMethod -Uri $url -Method POST -Headers $headers -Body $emailJSON -ContentType "application/json"
    } catch {
        $err = $_.Exception.Message
        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"    
        Add-Content -Path $log -Value "$ts - ERROR: $err`n"
    }
}

#Main
try {
    Send-Messsage
    Write-Output "Email sent to $msgRecipient successfully"
}catch{
    Write-Output "Failed to send email to $msgRecipient"
}