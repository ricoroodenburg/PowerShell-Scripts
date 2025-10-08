Function Upload-FileToIncident {

    Param (

        [Parameter(Mandatory = $true)]
        [string]$URL,

        [Parameter(Mandatory = $true)]
        [string]$Username,

        [Parameter(Mandatory = $true)]
        [string]$ApplicationPassword,

        [Parameter(Mandatory = $true)]
        [string]$IncidentNumber,

        [Parameter(Mandatory = $true)]
        [string]$AttachmentPath


    )

    Begin {

        # Check if the file exists
        if (-not (Test-Path $AttachmentPath)) {
            Write-Error "Error: File not found at $AttachmentPath."
            return
        }

        # Convert the attachment file to base64
        $FileContent = [Convert]::ToBase64String((Get-Content -Path $AttachmentPath -Encoding Byte))
        
        $Headers = @{
        
            Authorization = "Basic " + [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($Username):$($ApplicationPassword)"))
            
        }

        $body = @"
--BOUNDARY
Content-Disposition: form-data; name="file"; filename="$($(Get-Item $AttachmentPath).Name)"
Content-Type: application/plain;charset=utf-8
Content-Transfer-Encoding: base64

$fileContent
--BOUNDARY--
"@

    }

    Process {

        Try {
            
            Invoke-RestMethod -Uri ($URL + "/tas/api/incidents/number/$IncidentNumber/attachments") -Headers $Headers -Method Post -Body $body -ContentType "multipart/form-data; boundary=BOUNDARY"
                
        } Catch {

            # Error...
            Write-Debug ("$(Get-TimeStamp) >> Error while adding a new incident within TOPdesk, see error code below..")
            Write-Debug ("$(Get-TimeStamp) StatusCode:" + $_.Exception.Response.StatusCode.value__ )
            Write-Debug ("$(Get-TimeStamp) StatusDescription:" + $_.Exception.Response.StatusDescription)
            Write-Error $_
            Throw

        }

    }

}
